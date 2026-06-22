"use strict";

const fs     = require("fs");
const path   = require("path");
const { createCanvas } = require("canvas");

// ══════════════════════════════════════════════════════════════════════════════
//  HELPERS DE FECHA/HORA (exportados para main.js)
// ══════════════════════════════════════════════════════════════════════════════

function excelTimeToString(serial) {
  if (!serial) return "";
  if (isNaN(Number(serial))) return String(serial).trim();
  const totalMinutos = Math.round(Number(serial) * 1440);
  const hh = Math.floor(totalMinutos / 60).toString().padStart(2, "0");
  const mm  = (totalMinutos % 60).toString().padStart(2, "0");
  return `${hh}:${mm}`;
}

function excelDateToString(serial) {
  if (!serial) return "";
  if (isNaN(Number(serial))) return String(serial).trim();
  const date = new Date(Math.round((Number(serial) - 25569) * 86400 * 1000));
  const fc   = new Date(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate());
  return fc.toLocaleDateString("es-CO", { day: "2-digit", month: "2-digit", year: "numeric" });
}

// ══════════════════════════════════════════════════════════════════════════════
//  SELLO PNG  — 2 filas, ancho dinámico, blanco y negro, 10px Arial
//  d = { radicado, fecha, hora, asunto, anexos, radicadoPor }
// ══════════════════════════════════════════════════════════════════════════════

async function generarSelloPNG(d) {
  const H = 60, PADDING = 20;

  const linea1 = `SERVENTEGRAL S.A E.S.P  |  NIT: 828.002.229-2  |  Radicado: ${d.radicado}  ${d.fecha}  ${d.hora}`;

  // Medir ancho real con las fuentes que se usarán al dibujar
  const tmp = createCanvas(1, 1).getContext("2d");

  tmp.font = "bold 10px Arial";
  const a1 = tmp.measureText(linea1).width;

  tmp.font = "bold 10px Arial";
  const wL1 = tmp.measureText("Asunto: ").width;
  tmp.font = "10px Arial";
  const wV1 = tmp.measureText((d.asunto  || "") + "     ").width;
  tmp.font = "bold 10px Arial";
  const wL2 = tmp.measureText("Anexos: ").width;
  tmp.font = "10px Arial";
  const wV2 = tmp.measureText((d.anexos  || "") + "     ").width;
  tmp.font = "bold 10px Arial";
  const wL3 = tmp.measureText("Radicado por: ").width;
  tmp.font = "10px Arial";
  const wV3 = tmp.measureText(d.radicadoPor || "").width;
  const a2  = wL1 + wV1 + wL2 + wV2 + wL3 + wV3;

  const W  = Math.ceil(Math.max(a1, a2)) + PADDING;
  const sw = 40;
  const sh = sw * (W / H);

  const canvas = createCanvas(W, H);
  const ctx    = canvas.getContext("2d");
  ctx.textBaseline = "middle";

  // FILA 1 (y=0, h=30): fondo blanco, bold 10px negro, centrado
  ctx.fillStyle = "#FFFFFF";
  ctx.fillRect(0, 0, W, 30);
  ctx.fillStyle = "#000000";
  ctx.font      = "bold 10px Arial";
  ctx.textAlign = "center";
  ctx.fillText(linea1, W / 2, 15);

  // Línea separadora
  ctx.strokeStyle = "#000000";
  ctx.lineWidth   = 1;
  ctx.beginPath(); ctx.moveTo(0, 30); ctx.lineTo(W, 30); ctx.stroke();

  // FILA 2 (y=30, h=30): fondo blanco, labels bold + valores normal
  ctx.fillStyle = "#FFFFFF";
  ctx.fillRect(0, 30, W, 30);
  ctx.fillStyle = "#000000";
  ctx.textAlign = "left";
  const y2 = 45;

  ctx.font = "bold 10px Arial";
  const lAsunto  = "Asunto: ";
  const aLAsunto = ctx.measureText(lAsunto).width;
  ctx.fillText(lAsunto, 10, y2);
  ctx.font = "10px Arial";
  const vAsunto  = (d.asunto || "") + "     ";
  const aVAsunto = ctx.measureText(vAsunto).width;
  ctx.fillText(vAsunto, 10 + aLAsunto, y2);

  ctx.font = "bold 10px Arial";
  const lAnexos  = "Anexos: ";
  const xAnexos  = 10 + aLAsunto + aVAsunto;
  const aLAnexos = ctx.measureText(lAnexos).width;
  ctx.fillText(lAnexos, xAnexos, y2);
  ctx.font = "10px Arial";
  const vAnexos  = (d.anexos || "") + "     ";
  const aVAnexos = ctx.measureText(vAnexos).width;
  ctx.fillText(vAnexos, xAnexos + aLAnexos, y2);

  ctx.font = "bold 10px Arial";
  const lRadPor  = "Radicado por: ";
  const xRadPor  = xAnexos + aLAnexos + aVAnexos;
  const aLRadPor = ctx.measureText(lRadPor).width;
  ctx.fillText(lRadPor, xRadPor, y2);
  ctx.font = "10px Arial";
  ctx.fillText(d.radicadoPor || "", xRadPor + aLRadPor, y2);

  // Borde exterior
  ctx.strokeStyle = "#000000";
  ctx.lineWidth   = 1;
  ctx.strokeRect(0.5, 0.5, W - 1, H - 1);

  return { pngBuffer: canvas.toBuffer("image/png"), sw, sh };
}

// ══════════════════════════════════════════════════════════════════════════════
//  INSERCIÓN EN PDF  — rotado 90° en margen derecho
// ══════════════════════════════════════════════════════════════════════════════

async function procesarPDF(datosRadicado, docFile, outDir) {
  const { PDFDocument, degrees, PDFName } = require("pdf-lib");

  const fileBytes = fs.readFileSync(docFile);
  const pdfDoc   = await PDFDocument.load(fileBytes);
  const { pngBuffer, sw, sh } = await generarSelloPNG(datosRadicado);
  const pngImage = await pdfDoc.embedPng(pngBuffer);

  const page = pdfDoc.getPages()[0];
  const { width: pw, height: ph } = page.getSize();
  const sy = (ph - sh) / 2;

  // Cargar original en doc separado → evita auto-referencia al incrustar como XObject
  const pdfOrig = await PDFDocument.load(fileBytes);
  const [paginaOrig] = await pdfDoc.embedPdf(pdfOrig, [0]);

  page.node.delete(PDFName.of("CropBox"));
  page.setSize(pw + sw, ph);
  page.node.set(PDFName.of("Contents"), pdfDoc.context.obj([]));

  // Contenido original desplazado hacia la derecha para hacer hueco al sello
  page.drawPage(paginaOrig, { x: sw, y: 0, width: pw, height: ph });

  // Sello en lado izquierdo
  page.drawImage(pngImage, {
    x:      sw,
    y:      sy,
    width:  sh,
    height: sw,
    rotate: degrees(90),
  });

  const outFile = path.join(outDir, path.basename(docFile, ".pdf") + "_radicado.pdf");
  fs.writeFileSync(outFile, await pdfDoc.save());
  return { ok: true, outFile };
}

// ══════════════════════════════════════════════════════════════════════════════
//  FUNCIÓN PRINCIPAL — recibe datosRadicado directamente
// ══════════════════════════════════════════════════════════════════════════════

async function processAll({ datosRadicado, docsFiles, outDir }, onProgress) {
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

  onProgress({ type: "info", msg: `Radicado: ${datosRadicado.radicado} — ${datosRadicado.asunto}` });
  onProgress({ type: "info", msg: `Archivos a procesar: ${docsFiles.length}` });

  let ok = 0, fail = 0;

  for (let i = 0; i < docsFiles.length; i++) {
    const filePath = docsFiles[i];
    onProgress({ type: "progress", current: i + 1, total: docsFiles.length });
    const nombre = path.basename(filePath);
    const tipo   = path.extname(filePath).toLowerCase();
    try {
      let result;
      if (tipo === ".pdf") result = await procesarPDF(datosRadicado, filePath, outDir);
      else                 result = { ok: false, reason: `Tipo no soportado (${tipo})` };

      if (result.ok) {
        ok++;
        onProgress({ type: "ok", msg: `Procesado: ${nombre}` });
      } else {
        fail++;
        onProgress({ type: "error", msg: `Omitido: ${nombre} — ${result.reason}` });
      }
    } catch (err) {
      fail++;
      onProgress({ type: "error", msg: `ERROR en ${nombre}: ${err.message}` });
    }
  }
  return { ok, fail, total: docsFiles.length };
}

module.exports = { processAll, excelDateToString, excelTimeToString };
