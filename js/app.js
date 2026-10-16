/* Procesador de tablas: lee CSV, HTML o Excel y los muestra con el formato del laboratorio.
   Todo corre en el navegador; ningún archivo se envía a un servidor. */

const VACIOS = new Set(["", "-x-", "#VALUE!", "#N/A"]);

const $ = (id) => document.getElementById(id);

// ---------------------------------------------------------------- utilidades

function limpiar(texto) {
  return String(texto ?? "").replace(/\s+/g, " ").trim();
}

/** Quita filas y columnas completamente vacías de una matriz de textos. */
function recortar(matriz) {
  const filas = matriz.filter((f) => f.some((c) => limpiar(c) !== ""));
  if (!filas.length) return [];
  const ancho = Math.max(...filas.map((f) => f.length));
  const columnasUtiles = [];
  for (let c = 0; c < ancho; c++) {
    if (filas.some((f) => limpiar(f[c]) !== "")) columnasUtiles.push(c);
  }
  return filas.map((f) => columnasUtiles.map((c) => limpiar(f[c])));
}

/** Decodifica texto respetando el charset (los reportes del equipo vienen en windows-1252). */
function decodificar(buffer) {
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(buffer);
  } catch {
    return new TextDecoder("windows-1252").decode(buffer);
  }
}

// ---------------------------------------------------------------- lectores

/** Reporte "Tabla resumen de resultados" del analizador: tabla de maquetación por posiciones,
    con encabezados escalonados y dividida en páginas. Se reconstruye en una sola tabla. */
function leerReporteAnalizador(doc) {
  const filas = [...doc.querySelectorAll("tr")].map((tr) => {
    let col = 0;
    const celdas = [];
    for (const td of tr.children) {
      const texto = limpiar(td.textContent);
      if (texto) celdas.push([col, texto]);
      col += parseInt(td.getAttribute("colspan") || "1", 10);
    }
    return celdas;
  });

  const analitos = [];
  const metodos = {};
  const muestras = new Map();
  let columnas = [];
  let encabezado = null;

  for (const fila of filas) {
    const textos = fila.map(([, t]) => t);
    if (!textos.length || textos[0].includes("Tabla resumen")) continue;

    if (textos.includes("Muestras")) {
      encabezado = fila.filter(([, t]) => t !== "Muestras");
    } else if (encabezado && textos[0].startsWith("Tipo de")) {
      columnas = encabezado.sort((a, b) => a[0] - b[0]).map(([, t]) => t);
      columnas.forEach((nombre, i) => {
        metodos[nombre] = textos[i + 1] || "";
        if (!analitos.includes(nombre)) analitos.push(nombre);
      });
      encabezado = null;
    } else if (encabezado) {
      encabezado.push(...fila);
    } else if (textos.length >= 2 && textos.length - 1 === columnas.length) {
      const [id, ...valores] = textos;
      if (!muestras.has(id)) muestras.set(id, {});
      columnas.forEach((nombre, i) => (muestras.get(id)[nombre] = valores[i]));
    }
    // Filas de una celda = numeración de página: se ignoran
  }

  const encabezados = ["MUESTRA", ...analitos.map((a) => ({ texto: a, sub: metodos[a] }))];
  const cuerpo = [...muestras].map(([id, datos]) => [id, ...analitos.map((a) => datos[a] ?? "")]);
  return [{ nombre: "Tabla resumen de resultados", encabezados, cuerpo }];
}

/** Convierte una <table> HTML genérica en matriz, expandiendo colspan/rowspan. */
function tablaAMatriz(tabla) {
  const matriz = [];
  [...tabla.rows].forEach((tr, r) => {
    matriz[r] = matriz[r] || [];
    let c = 0;
    for (const td of tr.cells) {
      while (matriz[r][c] !== undefined) c++;
      const cs = parseInt(td.getAttribute("colspan") || "1", 10);
      const rs = parseInt(td.getAttribute("rowspan") || "1", 10);
      const texto = limpiar(td.textContent);
      for (let i = 0; i < rs; i++) {
        matriz[r + i] = matriz[r + i] || [];
        for (let j = 0; j < cs; j++) matriz[r + i][c + j] = i === 0 && j === 0 ? texto : "";
      }
      c += cs;
    }
  });
  return matriz.map((f) => Array.from(f, (v) => v ?? ""));
}

function matrizATabla(nombre, matriz) {
  const datos = recortar(matriz);
  if (datos.length < 1) return null;
  const [encabezados, ...cuerpo] = datos;
  return { nombre, encabezados, cuerpo };
}

function leerHTML(buffer) {
  const doc = new DOMParser().parseFromString(decodificar(buffer), "text/html");
  const texto = doc.body ? doc.body.textContent : "";
  if (texto.includes("Tabla resumen de resultados") && texto.includes("Muestras")) {
    return leerReporteAnalizador(doc);
  }
  // Solo tablas "hoja" (sin tablas anidadas dentro) para no duplicar datos
  const tablas = [...doc.querySelectorAll("table")].filter((t) => !t.querySelector("table"));
  return tablas
    .map((t, i) => {
      const titulo = limpiar(t.caption?.textContent) || `Tabla ${i + 1}`;
      return matrizATabla(titulo, tablaAMatriz(t));
    })
    .filter(Boolean);
}

function leerHojaDeCalculo(buffer, esTexto) {
  const libro = esTexto
    ? XLSX.read(decodificar(buffer), { type: "string" })
    : XLSX.read(buffer, { type: "array", cellDates: true });
  return libro.SheetNames.map((nombre) => {
    const matriz = XLSX.utils.sheet_to_json(libro.Sheets[nombre], {
      header: 1, raw: false, defval: "", blankrows: false,
    });
    return matrizATabla(nombre, matriz);
  }).filter(Boolean);
}

async function leerArchivo(archivo) {
  const ext = archivo.name.split(".").pop().toLowerCase();
  const buffer = await archivo.arrayBuffer();
  if (ext === "html" || ext === "htm") return leerHTML(buffer);
  if (["csv", "tsv", "txt"].includes(ext)) {
    // Un CSV es una sola tabla: se nombra con el archivo en lugar de "Sheet1"
    return leerHojaDeCalculo(buffer, true).map((t) => ({ ...t, nombre: archivo.name }));
  }
  if (["xlsx", "xlsm", "xls", "ods"].includes(ext)) return leerHojaDeCalculo(buffer, false);
  throw new Error(`Formato no soportado: .${ext}`);
}

// ---------------------------------------------------------------- render

function crear(tag, props = {}, ...hijos) {
  const el = document.createElement(tag);
  Object.assign(el, props);
  el.append(...hijos);
  return el;
}

function renderTabla(tabla, numero) {
  const filaEnc = crear("tr");
  for (const enc of tabla.encabezados) {
    const th = crear("th", { scope: "col" });
    if (typeof enc === "object") {
      th.append(enc.texto);
      if (enc.sub) th.append(crear("span", { className: "metodo", textContent: enc.sub }));
    } else {
      th.textContent = enc || "";
    }
    filaEnc.append(th);
  }

  const tbody = crear("tbody");
  for (const fila of tabla.cuerpo) {
    const tr = crear("tr");
    fila.forEach((valor, i) => {
      if (i === 0) {
        tr.append(crear("th", { scope: "row", textContent: valor }));
      } else if (VACIOS.has(valor)) {
        tr.append(crear("td", { className: "vacio", textContent: "—" }));
      } else {
        tr.append(crear("td", { textContent: valor }));
      }
    });
    tbody.append(tr);
  }

  return crear(
    "section", {},
    crear("h2", { className: "seccion", textContent: `${numero}-${tabla.nombre}` }),
    crear("div", { className: "tabla-contenedor" },
      crear("table", { className: "resultados" }, crear("thead", {}, filaEnc), tbody)),
  );
}

function fechaDesdeNombre(nombre) {
  // "12,10,26.htm.html" -> "12/10/26"
  const m = nombre.match(/^(\d{1,2})[,._-](\d{1,2})[,._-](\d{2,4})/);
  return m ? `${m[1]}/${m[2]}/${m[3]}` : new Date().toLocaleDateString("es-MX");
}

function mostrarReporte(archivos, tablas) {
  $("fecha").textContent = fechaDesdeNombre(archivos[0].name);
  $("generado").textContent = new Date().toLocaleString("es-MX", { dateStyle: "short", timeStyle: "short" });

  const conDato = tablas.reduce(
    (n, t) => n + t.cuerpo.reduce((m, f) => m + f.slice(1).filter((v) => !VACIOS.has(v)).length, 0), 0);
  const ficha = $("ficha");
  ficha.replaceChildren();
  for (const [dt, dd] of [
    ["Archivo:", archivos.map((a) => a.name).join(", ")],
    ["Tablas:", tablas.length],
    ["Filas:", tablas.reduce((n, t) => n + t.cuerpo.length, 0)],
    ["Resultados:", conDato],
  ]) {
    ficha.append(crear("div", {}, crear("dt", { textContent: dt }), crear("dd", { textContent: dd })));
  }

  ultimo = { archivos, tablas };
  $("tablas").replaceChildren(...tablas.map((t, i) => renderTabla(t, i + 2)));
  document.title = `Registro de análisis · ${archivos[0].name}`;
  $("carga").hidden = true;
  $("reporte").hidden = false;
  window.scrollTo(0, 0);
}

async function procesar(lista) {
  const archivos = [...lista];
  if (!archivos.length) return;
  const error = $("error");
  error.hidden = true;
  try {
    const tablas = [];
    for (const archivo of archivos) {
      const encontradas = await leerArchivo(archivo);
      // Con varios archivos, se antepone el nombre del archivo a cada tabla
      if (archivos.length > 1) encontradas.forEach((t) => (t.nombre = `${archivo.name} · ${t.nombre}`));
      tablas.push(...encontradas);
    }
    if (!tablas.length) throw new Error("No se encontraron tablas con datos en el archivo.");
    mostrarReporte(archivos, tablas);
  } catch (e) {
    error.textContent = e.message || String(e);
    error.hidden = false;
  }
}

// ---------------------------------------------------------------- PDF

let ultimo = null; // { archivos, tablas } del último procesamiento

async function descargarPDF() {
  if (!ultimo) return;
  const boton = $("btn-pdf");
  boton.disabled = true;
  boton.textContent = "Generando PDF…";
  try {
    const texto = (sel) => limpiar(document.querySelector(sel)?.textContent);
    const info = {
      titulo: texto("#reporte .titulo h1") || "REGISTRO DE ANÁLISIS",
      subtitulo: texto("#reporte .titulo p"),
      fecha: texto("#fecha"),
      generado: texto("#generado"),
      nota: texto("#reporte .banda-inferior p"),
      firmas: [...document.querySelectorAll("#reporte .nombre-firma")].map((el) => limpiar(el.innerText)),
      ficha: [...document.querySelectorAll("#ficha > div")].map((d) => [
        limpiar(d.querySelector("dt").textContent), limpiar(d.querySelector("dd").textContent)]),
    };
    const base = ultimo.archivos[0].name.replace(/\.[^.]+$/, "").replace(/\.htm$/i, "");
    await generarPDF(info, ultimo.tablas, VACIOS, `Registro de analisis - ${base}.pdf`);
  } catch (e) {
    alert(`No se pudo generar el PDF: ${e.message || e}`);
  } finally {
    boton.disabled = false;
    boton.textContent = "Descargar PDF";
  }
}

// ---------------------------------------------------------------- eventos

const zona = $("zona");
["dragenter", "dragover"].forEach((ev) =>
  zona.addEventListener(ev, (e) => { e.preventDefault(); zona.classList.add("activa"); }));
["dragleave", "drop"].forEach((ev) =>
  zona.addEventListener(ev, (e) => { e.preventDefault(); zona.classList.remove("activa"); }));
zona.addEventListener("drop", (e) => procesar(e.dataTransfer.files));
$("archivo").addEventListener("change", (e) => { procesar(e.target.files); e.target.value = ""; });

// Evita que soltar un archivo fuera de la zona lo abra en el navegador
window.addEventListener("dragover", (e) => e.preventDefault());
window.addEventListener("drop", (e) => e.preventDefault());

$("btn-nuevo").addEventListener("click", () => {
  $("reporte").hidden = true;
  $("carga").hidden = false;
});
$("btn-pdf").addEventListener("click", descargarPDF);
