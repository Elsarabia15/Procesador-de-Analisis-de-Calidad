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

// ---------------------------------------------------------------- catálogos

/** Etiquetas del menú "Nombre del vino". Las que se agregan con "Otro…" se guardan en este navegador. */
const PRODUCTOS = ["Tablas 1-4-5", "Tablas 2", "Nudo", "Ensamble Arenal", "Ensamble Colina", "CDD Fusión de tintas"];
const CLAVE_PRODUCTOS = "productos-personalizados";
const OTRO = "__otro__";

/** Métodos del analizador BioSystems Y15 (catálogo de enología de BioSystems).
    Cada parámetro con método lleva un número volado y el pie del reporte explica solo los métodos
    de los parámetros que tienen datos en las muestras seleccionadas. Las fuentes del PDF solo tienen
    los volados ¹²³, por eso hay a lo más tres notas. */
const METODOS = [
  {
    patron: /ACETIC|MALIC|GLUCOSE|FRUCTOSE/i,
    unidad: "g/L",
    nota: () => "Enzimático (NADH), determinación diferencial con lectura a 340 nm.",
  },
  {
    patron: /SULFITE|SO2/i,
    unidad: "mg/L",
    nota: (parametros) => {
      const total = parametros.some((p) => /TOTAL/i.test(p));
      const libre = parametros.some((p) => !/TOTAL/i.test(p));
      return [
        libre && "SO2 libre: colorimétrico con pararosanilina y formaldehído, lectura a 560 nm",
        total && "SO2 total: colorimétrico con DTNB, lectura a 405 nm",
      ].filter(Boolean).join("; ") + ".";
    },
  },
  {
    patron: /ABS\s*\d|ICM/i,
    unidad: "",
    nota: () => "Espectrofotométrico: absorbancia a 420, 520 y 620 nm de la muestra diluida en buffer.",
  },
];
const VOLADOS = ["¹", "²", "³"];

function metodoDe(parametro) {
  return METODOS.find((m) => m.patron.test(parametro)) ?? null;
}

/** Numera los métodos en el orden en que aparecen los parámetros.
    Devuelve la marca de cada parámetro y las líneas del pie. */
function marcasMetodo(parametros) {
  const usados = [];
  for (const p of parametros) {
    const m = metodoDe(p);
    if (m && !usados.includes(m)) usados.push(m);
  }
  const marca = (p) => VOLADOS[usados.indexOf(metodoDe(p))] ?? "";
  const notas = usados.map((m, i) => `${VOLADOS[i]} ${m.nota(parametros.filter((p) => metodoDe(p) === m))}`);
  return { marca, notas };
}

function productosGuardados() {
  try {
    return JSON.parse(localStorage.getItem(CLAVE_PRODUCTOS)) || [];
  } catch {
    return [];
  }
}

let productos = [...new Set([...PRODUCTOS, ...productosGuardados()])];

function agregarProducto(nombre) {
  if (productos.includes(nombre)) return;
  productos.push(nombre);
  try {
    localStorage.setItem(CLAVE_PRODUCTOS, JSON.stringify(productos.filter((p) => !PRODUCTOS.includes(p))));
  } catch {
    // Sin almacenamiento disponible: la etiqueta solo queda en esta sesión
  }
}

function llenarProductos(seleccionado = "") {
  $("d-producto").replaceChildren(
    crear("option", { value: "", textContent: "— Elige una etiqueta —" }),
    ...productos.map((p) => crear("option", { value: p, textContent: p })),
    crear("option", { value: OTRO, textContent: "Otro… (escribir)" }),
  );
  $("d-producto").value = seleccionado;
}

// ---------------------------------------------------------------- render

/** Estado del último archivo procesado:
    { archivos, tablas, muestras, modo, sel: { general: Set, etiqueta: Set }, observaciones: { parámetro: texto } }
    El modo se decide solo: "etiqueta" si el archivo trae una sola muestra, "general" si trae varias.
    Cada modo tiene su propia selección de muestras (clave "tabla:fila"). */
let ultimo = null;

function crear(tag, props = {}, ...hijos) {
  const el = document.createElement(tag);
  Object.assign(el, props);
  el.append(...hijos);
  return el;
}

function casilla(clave) {
  const input = crear("input", { type: "checkbox" });
  input.dataset.clave = clave;
  return input;
}

function textoEncabezado(enc) {
  return typeof enc === "object" ? enc.texto : enc || "";
}

function celdaValor(valor) {
  return VACIOS.has(valor)
    ? crear("td", { className: "vacio", textContent: "—" })
    : crear("td", { textContent: valor });
}

function renderTabla(tabla, t, numero) {
  const filaEnc = crear("tr");
  for (const enc of tabla.encabezados) {
    const th = crear("th", { scope: "col" });
    th.dataset.parametro = textoEncabezado(enc);
    if (typeof enc === "object") {
      th.append(enc.texto, crear("sup", { className: "marca" }));
      if (enc.sub) th.append(crear("span", { className: "metodo", textContent: enc.sub }));
    } else {
      th.append(enc || "", crear("sup", { className: "marca" }));
    }
    filaEnc.append(th);
  }

  const tbody = crear("tbody");
  tabla.cuerpo.forEach((fila, r) => {
    const tr = crear("tr");
    fila.forEach((valor, i) => {
      if (i === 0) {
        // La casilla elige si la muestra sale en el reporte
        tr.append(crear("th", { scope: "row" },
          crear("label", { className: "casilla" }, casilla(`${t}:${r}`), valor)));
      } else {
        tr.append(celdaValor(valor));
      }
    });
    tbody.append(tr);
  });

  return crear(
    "section", {},
    crear("h2", { className: "seccion", textContent: `${numero}-${tabla.nombre}` }),
    crear("div", { className: "tabla-contenedor" },
      crear("table", { className: "resultados" }, crear("thead", {}, filaEnc), tbody)),
  );
}

/** Tablas del reporte general, solo con las muestras seleccionadas. */
function tablasGenerales() {
  const sel = ultimo.sel.general;
  return ultimo.tablas
    .map((tabla, t) => ({ ...tabla, cuerpo: tabla.cuerpo.filter((_, r) => sel.has(`${t}:${r}`)) }))
    .filter((tabla) => tabla.cuerpo.length);
}

/** Parámetros con al menos un resultado en las muestras seleccionadas, en orden de aparición. */
function parametrosGenerales(tablas) {
  const parametros = [];
  for (const tabla of tablas) {
    tabla.encabezados.forEach((enc, c) => {
      const nombre = textoEncabezado(enc);
      if (c > 0 && !parametros.includes(nombre) && tabla.cuerpo.some((f) => !VACIOS.has(f[c] ?? ""))) {
        parametros.push(nombre);
      }
    });
  }
  return parametros;
}

/** Marca de método de cada parámetro y notas del pie, según el modo y las muestras seleccionadas. */
function metodosActuales() {
  return ultimo.modo === "etiqueta"
    ? marcasMetodo(construirVertical().filas.map((f) => f.parametro))
    : marcasMetodo(parametrosGenerales(tablasGenerales()));
}

function renderMetodos() {
  const { marca, notas } = metodosActuales();
  document.querySelectorAll("#reporte th[data-parametro] > .marca").forEach((sup) => {
    sup.textContent = marca(sup.parentElement.dataset.parametro);
  });
  $("notas").replaceChildren(...notas.map((n) => crear("p", { textContent: n })));
}

function renderFicha() {
  const tablas = tablasGenerales();
  const conDato = tablas.reduce(
    (n, t) => n + t.cuerpo.reduce((m, f) => m + f.slice(1).filter((v) => !VACIOS.has(v)).length, 0), 0);
  const ficha = $("ficha");
  ficha.replaceChildren();
  for (const [dt, dd] of [
    ["Archivo:", ultimo.archivos.map((a) => a.name).join(", ")],
    ["Tablas:", tablas.length],
    ["Muestras:", tablas.reduce((n, t) => n + t.cuerpo.length, 0)],
    ["Resultados:", conDato],
  ]) {
    ficha.append(crear("div", {}, crear("dt", { textContent: dt }), crear("dd", { textContent: dd })));
  }
}

/** Reporte por etiqueta: una fila por parámetro y una columna por lote seleccionado.
    Se omiten los parámetros sin ningún resultado en esos lotes. */
function construirVertical() {
  const lotes = ultimo.muestras.filter((m) => ultimo.sel.etiqueta.has(m.clave));
  const valores = new Map(); // parámetro -> valor de cada lote
  lotes.forEach((lote, i) => {
    const tabla = ultimo.tablas[lote.t];
    tabla.encabezados.slice(1).forEach((enc, c) => {
      const nombre = textoEncabezado(enc);
      if (!valores.has(nombre)) valores.set(nombre, Array(lotes.length).fill(""));
      valores.get(nombre)[i] = tabla.cuerpo[lote.r][c + 1] ?? "";
    });
  });
  const filas = [...valores]
    .filter(([, v]) => v.some((x) => !VACIOS.has(x)))
    .map(([parametro, v]) => ({ parametro, unidad: metodoDe(parametro)?.unidad ?? "", valores: v }));
  // Con un solo lote la columna se llama VALOR, como en el formato de Excel. La metodología va en el pie.
  const encabezados = ["PARÁMETRO", ...(lotes.length === 1 ? ["VALOR"] : lotes.map((l) => l.id)), "UNIDAD", "OBSERVACIONES"];
  return { lotes, filas, encabezados };
}

function renderVertical() {
  const { filas, encabezados } = construirVertical();
  const titulo = crear("h2", { className: "seccion", textContent: "2-Análisis" });

  const thead = crear("thead", {}, crear("tr", {}, ...encabezados.map((e) => crear("th", { scope: "col", textContent: e }))));
  const tbody = crear("tbody");
  for (const fila of filas) {
    // Observaciones: la única celda editable de la tabla
    const obs = crear("td", {
      className: "texto observacion", contentEditable: "true", spellcheck: true,
      textContent: ultimo.observaciones[fila.parametro] ?? "",
    });
    obs.dataset.parametro = fila.parametro;
    tbody.append(crear("tr", {},
      crear("th", { scope: "row", className: "texto" }, fila.parametro, crear("sup", { className: "marca" })),
      ...fila.valores.map(celdaValor),
      fila.unidad ? crear("td", { textContent: fila.unidad }) : celdaValor(""),
      obs));
  }

  $("vertical").replaceChildren(crear("section", {},
    titulo,
    crear("div", { className: "tabla-contenedor" }, crear("table", { className: "resultados vertical" }, thead, tbody)),
  ));
  tbody.querySelectorAll("th").forEach((th, i) => (th.dataset.parametro = filas[i].parametro));
  tbody.addEventListener("input", (e) => {
    const celda = e.target.closest(".observacion");
    if (!celda) return;
    if (!limpiar(celda.textContent)) celda.replaceChildren(); // Chrome deja un <br> al borrar todo
    ultimo.observaciones[celda.dataset.parametro] = limpiar(celda.textContent);
  });
}

/** Refleja la selección del modo actual en las casillas y el contenido del reporte. */
function sincronizar() {
  const sel = ultimo.sel[ultimo.modo];
  document.querySelectorAll("#reporte input[data-clave]").forEach((c) => {
    c.checked = sel.has(c.dataset.clave);
    c.closest("tr").classList.toggle("excluida", !c.checked);
  });
  if (ultimo.modo === "general") renderFicha();
  else renderVertical();
  renderMetodos();
}

/** Quita la basura que a veces se teclea en el equipo al final del ID ("26113,,,", "26111}"). */
function idLimpio(id) {
  return id.replace(/[^\p{L}\p{N}]+$/u, "");
}

function mostrarModo() {
  const etiqueta = ultimo.modo === "etiqueta";
  $("sec-resumen").hidden = etiqueta;
  $("tablas").hidden = etiqueta;
  $("sec-descripcion").hidden = !etiqueta;
  $("vertical").hidden = !etiqueta;
  $("firmas").hidden = !etiqueta; // El reporte general no lleva firmas
  sincronizar();
}

/** Fecha del análisis en formato de <input type="date">: la del nombre que pone el equipo
    ("12,10,26.htm.html" -> "2026-10-12", día/mes/año) o la de hoy si el nombre no la trae. */
function fechaDesdeNombre(nombre) {
  const dos = (n) => String(n).padStart(2, "0");
  const m = nombre.match(/^(\d{1,2})[,._-](\d{1,2})[,._-](\d{2,4})/);
  if (m) return `${m[3].length === 2 ? "20" + m[3] : m[3]}-${dos(m[2])}-${dos(m[1])}`;
  const hoy = new Date();
  return `${hoy.getFullYear()}-${dos(hoy.getMonth() + 1)}-${dos(hoy.getDate())}`;
}

function mostrarReporte(archivos, tablas) {
  $("d-fecha-analisis").value = fechaDesdeNombre(archivos[0].name);
  $("generado").textContent = new Date().toLocaleString("es-MX", { dateStyle: "short", timeStyle: "short" });

  const muestras = tablas.flatMap((tabla, t) =>
    tabla.cuerpo.map((fila, r) => ({ clave: `${t}:${r}`, id: fila[0], t, r, tabla: tabla.nombre })));
  // Un archivo con una sola muestra es el análisis de un lote: se abre como reporte por etiqueta
  const unLote = muestras.length === 1;
  ultimo = {
    archivos, tablas, muestras,
    modo: unLote ? "etiqueta" : "general",
    // General: todas incluidas de inicio (se quitan con las casillas). Por etiqueta: su único lote.
    sel: {
      general: new Set(muestras.map((m) => m.clave)),
      etiqueta: new Set(unLote ? [muestras[0].clave] : []),
    },
    observaciones: {},
  };
  if (unLote) $("d-lote").value = idLimpio(muestras[0].id);
  $("tablas").replaceChildren(...tablas.map((t, i) => renderTabla(t, i, i + 2)));
  mostrarModo();
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

function fechaCorta(iso) {
  // "2026-02-22" -> "22/02/2026"
  const m = iso.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? `${m[3]}/${m[2]}/${m[1]}` : "";
}

/** Campos de "1-Descripción muestra" en el orden del formulario (dos por renglón).
    La etiqueta de cada campo se toma del formulario para que el PDF diga lo mismo que la vista previa. */
function datosDescripcion() {
  return [...document.querySelectorAll("#sec-descripcion .descripcion > label")].map((label) => {
    const campo = label.querySelector("input:not([hidden]), select");
    let valor = limpiar(campo.value);
    if (campo.id === "d-producto" && valor === OTRO) valor = limpiar($("d-producto-otro").value);
    if (campo.type === "date") valor = fechaCorta(campo.value);
    return [limpiar(label.querySelector("span").textContent), valor];
  });
}

async function descargarPDF() {
  if (!ultimo) return;
  const texto = (sel) => limpiar(document.querySelector(sel)?.textContent);
  const info = {
    titulo: texto("#reporte .titulo h1") || "REGISTRO DE ANÁLISIS",
    subtitulo: texto("#reporte .titulo p"),
    folio: texto("#folio"),
    generado: texto("#generado"),
  };

  const { marca, notas } = metodosActuales();
  info.notas = notas;
  let tablas, nombreArchivo;
  if (ultimo.modo === "etiqueta") {
    const { lotes, filas, encabezados } = construirVertical();
    if (!lotes.length) return alert("Elige al menos un lote para el reporte por etiqueta.");
    const pares = datosDescripcion();
    info.ficha = { titulo: "1-Descripción muestra", pares, porFila: 2 };
    info.orientacion = "portrait";
    info.firmas = [...document.querySelectorAll("#firmas > div")].map((d) => ({
      etiqueta: limpiar(d.querySelector("span").textContent),
      // Se conservan los renglones (nombre y cargo)
      nombre: d.querySelector(".nombre-firma").innerText.split("\n").map(limpiar).filter(Boolean).join("\n"),
    }));
    tablas = [{
      nombre: "Análisis",
      encabezados,
      // Observación vacía como espacio duro: celda en blanco en lugar de "-"
      cuerpo: filas.map((f) => [f.parametro + marca(f.parametro), ...f.valores, f.unidad,
        ultimo.observaciones[f.parametro] || "\u00A0"]),
      estilosColumna: {
        0: { fontStyle: "bold", halign: "left" },
        [encabezados.length - 1]: { halign: "left", cellWidth: 65 },
      },
    }];
    const [producto, cosecha] = [pares[0][1], limpiar($("d-cosecha").value)];
    nombreArchivo = `Registro de analisis - ${[producto || "etiqueta", cosecha].filter(Boolean).join(" ")}.pdf`;
  } else {
    tablas = tablasGenerales().map((t) => ({
      ...t,
      encabezados: t.encabezados.map((enc, c) => {
        if (c === 0) return enc;
        const m = marca(textoEncabezado(enc));
        return typeof enc === "object" ? { ...enc, texto: enc.texto + m } : enc + m;
      }),
    }));
    if (!tablas.length) return alert("Selecciona al menos una muestra para el reporte.");
    info.ficha = {
      titulo: "1-Resumen",
      pares: [...document.querySelectorAll("#ficha > div")].map((d) => [
        limpiar(d.querySelector("dt").textContent), limpiar(d.querySelector("dd").textContent)]),
      porFila: 4,
    };
    const base = ultimo.archivos[0].name.replace(/\.[^.]+$/, "").replace(/\.htm$/i, "");
    nombreArchivo = `Registro de analisis - ${base}.pdf`;
  }

  const boton = $("btn-pdf");
  boton.disabled = true;
  boton.textContent = "Generando PDF…";
  try {
    await generarPDF(info, tablas, VACIOS, nombreArchivo.replace(/[\\/:*?"<>|]/g, "-"));
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

// Casillas de la columna Muestra
$("reporte").addEventListener("change", (e) => {
  const clave = e.target.dataset.clave;
  if (!clave) return;
  const sel = ultimo.sel[ultimo.modo];
  if (e.target.checked) sel.add(clave);
  else sel.delete(clave);
  sincronizar();
});

// Al borrar todo el nombre, Chrome deja un <br>: se vacía para que vuelva a verse el texto guía
document.querySelectorAll(".nombre-firma").forEach((el) =>
  el.addEventListener("input", () => { if (!limpiar(el.textContent)) el.replaceChildren(); }));

// Menú de etiquetas: "Otro…" muestra un campo para escribir una nueva
llenarProductos();
$("d-producto").addEventListener("change", (e) => {
  const otro = e.target.value === OTRO;
  $("d-producto-otro").hidden = !otro;
  if (otro) $("d-producto-otro").focus();
});
$("d-producto-otro").addEventListener("keydown", (e) => {
  if (e.key === "Enter") e.target.blur();
});
$("d-producto-otro").addEventListener("change", (e) => {
  const nombre = limpiar(e.target.value);
  if (!nombre) return;
  agregarProducto(nombre);
  llenarProductos(nombre);
  e.target.value = "";
  e.target.hidden = true;
});
