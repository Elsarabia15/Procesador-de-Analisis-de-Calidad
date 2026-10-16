/* Genera el PDF para imprimir con el formato del laboratorio
   (mismo diseño que "Formato análisis vino e ingredientes_TTA.xlsx"). */

const COLOR = {
  verdeOscuro: [0, 61, 17],   // #003D11
  verde: [0, 123, 59],        // #007B3B
  grisFila: [242, 242, 242],  // #F2F2F2
  grisTexto: [95, 95, 95],
  grisVacio: [178, 178, 178],
  negro: [0, 0, 0],
  blanco: [255, 255, 255],
};

const MARGEN = 12;       // mm
const ALTO_BANDA = 26;   // banda superior
const ALTO_PIE = 16;     // banda inferior + empresa

/** Las fuentes estándar de PDF solo cubren Latin-1: se sustituyen caracteres fuera de ese rango. */
function latin1(texto) {
  return String(texto ?? "")
    .replace(/[–—]/g, "-")
    .replace(/[‘’]/g, "'")
    .replace(/[“”]/g, '"')
    .replace(/[^\x00-\xFF]/g, "?");
}

/** Reduce el logo (1730 px) para que el PDF no pese de más. */
async function cargarLogo(src, lado = 240) {
  const img = new Image();
  img.src = src;
  await img.decode();
  const canvas = document.createElement("canvas");
  canvas.width = canvas.height = lado;
  canvas.getContext("2d").drawImage(img, 0, 0, lado, lado);
  return canvas.toDataURL("image/png");
}

function dibujarEncabezado(doc, info, logo) {
  const ancho = doc.internal.pageSize.getWidth();
  doc.setFillColor(...COLOR.verdeOscuro);
  doc.rect(0, 0, ancho, ALTO_BANDA, "F");
  doc.addImage(logo, "PNG", MARGEN, 3, 20, 20, "logo"); // alias: se incrusta una sola vez

  // Recuadro Fecha / Generado (como las celdas H2:I4 del Excel)
  const x = ancho - MARGEN - 62, y = 6, h = 7;

  // Título centrado en el espacio libre entre el logo y el recuadro; se reduce si no cabe
  const izquierda = MARGEN + 24, derecha = x - 10;
  const centro = (izquierda + derecha) / 2;
  doc.setTextColor(...COLOR.blanco);
  doc.setFont("courier", "normal");
  let tamano = 16;
  doc.setFontSize(tamano);
  while (tamano > 9 && doc.getTextWidth(latin1(info.titulo)) > derecha - izquierda) doc.setFontSize(--tamano);
  doc.text(latin1(info.titulo), centro, 12, { align: "center" });
  doc.setFontSize(Math.min(10, tamano));
  doc.text(latin1(info.subtitulo), centro, 18, { align: "center" });

  doc.setDrawColor(...COLOR.blanco);
  doc.setLineWidth(0.3);
  doc.line(x - 6, 5, x - 6, ALTO_BANDA - 5);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(9);
  [["Fecha:", info.fecha], ["Generado:", info.generado]].forEach(([etq, valor], i) => {
    doc.rect(x, y + i * h, 22, h);
    doc.rect(x + 22, y + i * h, 40, h);
    doc.text(etq, x + 2, y + i * h + 4.8);
    doc.text(latin1(valor), x + 42, y + i * h + 4.8, { align: "center" });
  });
}

function dibujarPie(doc, info, pagina, total) {
  const ancho = doc.internal.pageSize.getWidth();
  const alto = doc.internal.pageSize.getHeight();
  const y = alto - ALTO_PIE;
  doc.setFillColor(...COLOR.verdeOscuro);
  doc.rect(0, y, ancho, 7, "F");
  doc.setTextColor(...COLOR.blanco);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(8);
  doc.text(latin1(info.nota), MARGEN, y + 4.6);
  doc.text(`Página ${pagina} de ${total}`, ancho - MARGEN, y + 4.6, { align: "right" });

  doc.setTextColor(...COLOR.negro);
  doc.setFont("courier", "normal");
  doc.setFontSize(8);
  doc.text(["TABLAS, TALLER AGRICOLA S. DE R.L. DE C.V.", "WWW.TABLAS.COM.MX"], ancho - MARGEN, y + 11, {
    align: "right",
  });
}

function tituloSeccion(doc, texto, y) {
  doc.setFont("courier", "bold");
  doc.setFontSize(12);
  doc.setTextColor(...COLOR.verde);
  doc.text(latin1(texto), MARGEN, y);
  return y + 3;
}

/**
 * info: { titulo, subtitulo, fecha, generado, nota, firmas: [nombre, nombre], ficha: [[etq, valor]...] }
 * tablas: [{ nombre, encabezados, cuerpo }] (mismo formato que produce app.js)
 */
async function generarPDF(info, tablas, vacios, nombreArchivo) {
  const { jsPDF } = window.jspdf;
  const maxColumnas = Math.max(...tablas.map((t) => t.encabezados.length));
  const doc = new jsPDF({ orientation: maxColumnas > 7 ? "landscape" : "portrait", unit: "mm", format: "letter", compress: true });
  const logo = await cargarLogo("img/logo.png");
  const inicioContenido = ALTO_BANDA + 8;
  const margenTabla = { top: inicioContenido + 2, bottom: ALTO_PIE + 6, left: MARGEN, right: MARGEN };

  // ---- 1-Resumen
  let y = tituloSeccion(doc, "1-Resumen", inicioContenido);
  const celdasFicha = info.ficha.flatMap(([etq, valor]) => [
    { content: latin1(etq), styles: { fontStyle: "bold", halign: "right" } },
    latin1(valor),
  ]);
  doc.autoTable({
    startY: y,
    margin: margenTabla,
    body: [celdasFicha],
    theme: "plain",
    styles: { font: "helvetica", fontSize: 9, fillColor: COLOR.grisFila, cellPadding: 2 },
  });
  y = doc.lastAutoTable.finalY + 10;

  // ---- Tablas de datos
  tablas.forEach((tabla, i) => {
    if (y > doc.internal.pageSize.getHeight() - margenTabla.bottom - 25) {
      doc.addPage();
      y = inicioContenido;
    }
    y = tituloSeccion(doc, `${i + 2}-${tabla.nombre}`, y);

    const head = [tabla.encabezados.map((enc) =>
      typeof enc === "object" ? latin1(enc.sub ? `${enc.texto}\n${enc.sub}` : enc.texto) : latin1(enc))];
    const body = tabla.cuerpo.map((fila) => fila.map((v, c) =>
      c > 0 && vacios.has(v) ? { content: "-", styles: { textColor: COLOR.grisVacio } } : latin1(v)));

    // Tamaño de letra según el número de columnas, para que la tabla quepa en el ancho de la hoja
    const columnas = tabla.encabezados.length;
    const tamano = columnas > 14 ? 6.5 : columnas > 10 ? 7.5 : 9;

    doc.autoTable({
      startY: y,
      margin: margenTabla,
      head,
      body,
      theme: "plain",
      // Hasta 20 columnas se ajustan al ancho de la hoja (el texto se parte en líneas);
      // más que eso continúa en otra página repitiendo la primera columna.
      horizontalPageBreak: columnas > 20,
      horizontalPageBreakRepeat: 0,
      styles: {
        font: "helvetica", fontSize: tamano, halign: "center", valign: "middle",
        cellPadding: columnas > 14 ? 0.9 : 1.6, lineColor: COLOR.negro, lineWidth: 0, overflow: "linebreak",
      },
      headStyles: { fillColor: COLOR.grisFila, fontStyle: "bold" },
      alternateRowStyles: { fillColor: COLOR.grisFila },
      columnStyles: { 0: { fontStyle: "bold" } },
      didParseCell: (d) => {
        // Las líneas verticales van a la izquierda de cada celda (excepto la primera columna):
        // así el relleno gris de la celda vecina no las tapa.
        d.cell.styles.lineWidth = {
          left: d.column.index > 0 ? 0.2 : 0,
          top: d.section === "head" ? 0.3 : 0,
          bottom: d.section === "head" || d.row.index === d.table.body.length - 1 ? 0.3 : 0,
          right: 0,
        };
      },
    });
    y = doc.lastAutoTable.finalY + 10;
  });

  // ---- Firmas (como las filas 27-29 del Excel)
  const alto = doc.internal.pageSize.getHeight();
  y -= 4;
  if (y > alto - margenTabla.bottom - 24) {
    doc.addPage();
    y = inicioContenido;
  }
  const ancho = doc.internal.pageSize.getWidth();
  const colAncho = (ancho - MARGEN * 2 - 20) / 2;
  doc.setTextColor(...COLOR.negro);
  doc.setDrawColor(...COLOR.negro);
  doc.setLineWidth(0.3);
  ["Elaborado por:", "Aprobado por:"].forEach((etq, i) => {
    const x = MARGEN + i * (colAncho + 20);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(10);
    doc.text(etq, x + colAncho / 2, y + 4, { align: "center" });
    doc.line(x, y + 16, x + colAncho, y + 16);
    if (info.firmas[i]) {
      doc.setFont("helvetica", "bold");
      doc.text(doc.splitTextToSize(latin1(info.firmas[i]), colAncho), x + colAncho / 2, y + 21, { align: "center" });
    }
  });

  // ---- Encabezado y pie en todas las páginas
  const paginas = doc.internal.getNumberOfPages();
  for (let p = 1; p <= paginas; p++) {
    doc.setPage(p);
    dibujarEncabezado(doc, info, logo);
    dibujarPie(doc, info, p, paginas);
  }

  doc.save(nombreArchivo);
}
