"use strict";

const fs     = require("fs");
const path   = require("path");
const XLSX   = require("xlsx");
const PizZip = require("pizzip");
const { createCanvas, loadImage } = require("canvas");

const BASE_DIR    = __dirname;
const XLSX_FILE   = path.join(BASE_DIR, "radicados.xlsx");
const DOCS_DIR    = path.join(BASE_DIR, "documentos");
const OUT_DIR     = path.join(BASE_DIR, "salida");
const LOGO_FILE   = path.join(BASE_DIR, "logo.png");
const CONFIG_FILE = path.join(BASE_DIR, "config.json");

if (!fs.existsSync(OUT_DIR))   fs.mkdirSync(OUT_DIR,  { recursive: true });
if (!fs.existsSync(DOCS_DIR))  fs.mkdirSync(DOCS_DIR, { recursive: true });

// ── config.json ──────────────────────────────────────────────────────────────
function leerConfig() {
  const def = { lado: "derecho", offsetVertical: 914400 };
  if (!fs.existsSync(CONFIG_FILE)) return def;
  try { return { ...def, ...JSON.parse(fs.readFileSync(CONFIG_FILE, "utf8")) }; }
  catch { return def; }
}

// ── namespaces Open XML ───────────────────────────────────────────────────────
const NS = {
  r:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  wp:  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  a:   "http://schemas.openxmlformats.org/drawingml/2006/main",
  pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
};

// ── Generar sello PNG — diseño aprobado 900×278px ────────────────────────────
async function generarSelloPNG(row, logoBuffer) {
  const W = 900, H = 278;
  const canvas = createCanvas(W, H);
  const ctx    = canvas.getContext("2d");

  ctx.fillStyle = "#FFFFFF";
  ctx.fillRect(0, 0, W, H);
  ctx.textAlign    = "center";
  ctx.textBaseline = "middle";

  // Bloque LOGO: x=0, w=160
  const LOGO_W = 160;
  if (logoBuffer) {
    try {
      const img   = await loadImage(logoBuffer);
      const maxSz = 220;
      const scale = Math.min(maxSz / img.width, maxSz / img.height);
      const iw    = img.width  * scale;
      const ih    = img.height * scale;
      ctx.drawImage(img, (LOGO_W - iw) / 2, (H - ih) / 2, iw, ih);
    } catch { /* logo inaccesible */ }
  }

  // Bloque EMPRESA: x=160, w=130
  const EX = LOGO_W, EW = 130;

  ctx.fillStyle = "#1B5E20";
  ctx.fillRect(EX, 0, EW, 104);
  ctx.fillStyle = "#FFFFFF";
  ctx.font = "bold 11px Arial";
  ctx.fillText("SERVENTEGRAL", EX + EW / 2, 42);
  ctx.font = "10px Arial";
  ctx.fillText("S.A. E.S.P",   EX + EW / 2, 73);

  ctx.fillStyle = "#1B5E20";
  ctx.fillRect(EX, 104, EW, 49);
  ctx.fillStyle = "#FFFFFF";
  ctx.font = "9px Arial";
  ctx.fillText("NIT: 828.002.229-2", EX + EW / 2, 128, EW - 6);

  ctx.fillStyle = "#2E7D32";
  ctx.fillRect(EX, 153, EW, 125);
  ctx.fillStyle = "#FFFFFF";
  ctx.font = "bold 11px Arial";
  ctx.fillText("SELLO DE", EX + EW / 2, 195);
  ctx.fillText("RADICADO",  EX + EW / 2, 236);

  // Filas de datos: x=290, una fila = label + valor en la misma línea
  const DX  = LOGO_W + EW;  // 290
  const DW  = W - DX;       // 610
  const PAD = 8;
  const VX  = DX + 140;

  const filas = [
    { label: "N° RADICADO:",   value: String(row.NUMERO_RADICADO  ?? ""), h: 50, bold: true,  size: 20 },
    { label: "FECHA:",         value: String(row.FECHA_RECIBIDO   ?? ""), h: 38, bold: false, size: 12 },
    { label: "RECIBIDO POR:",  value: String(row.RECIBIDO_POR     ?? ""), h: 38, bold: false, size: 12 },
    { label: "ASUNTO:",        value: String(row.ASUNTO           ?? ""), h: 38, bold: false, size: 12 },
    { label: "ANEXOS:",        value: String(row.ANEXOS           ?? ""), h: 38, bold: false, size: 12 },
    { label: "ÁREA RESP:",     value: String(row.AREA_RESPONSABLE ?? ""), h: 38, bold: false, size: 12 },
  ];

  const bgFilas = ["#FFFFFF", "#F1F8E9"];
  let ry = 0;
  ctx.textAlign    = "left";
  ctx.textBaseline = "middle";

  for (let i = 0; i < filas.length; i++) {
    const f = filas[i];
    ctx.fillStyle = bgFilas[i % 2];
    ctx.fillRect(DX, ry, DW, f.h);

    ctx.fillStyle = "#1B5E20";
    ctx.font      = "bold 10px Arial";
    ctx.fillText(f.label, DX + PAD, ry + f.h / 2);

    ctx.fillStyle = "#000000";
    ctx.font      = `${f.bold ? "bold " : ""}${f.size}px Arial`;
    ctx.fillText(f.value, VX, ry + f.h / 2);

    ry += f.h;

    ctx.strokeStyle = "#C8E6C9";
    ctx.lineWidth   = 1;
    ctx.beginPath(); ctx.moveTo(DX, ry); ctx.lineTo(W, ry); ctx.stroke();
  }

  // Separadores verticales
  ctx.strokeStyle = "#C8E6C9";
  ctx.lineWidth   = 1;
  ctx.beginPath(); ctx.moveTo(DX, 0); ctx.lineTo(DX, H); ctx.stroke();
  ctx.beginPath(); ctx.moveTo(EX, 0); ctx.lineTo(EX, H); ctx.stroke();

  // Borde exterior
  ctx.strokeStyle = "#1B5E20";
  ctx.lineWidth   = 2;
  ctx.strokeRect(1, 1, W - 2, H - 2);

  return canvas.toBuffer("image/png");
}

// ── Insertar PNG en el zip del docx ──────────────────────────────────────────
function addStampToZip(zip, pngBuffer) {
  zip.file("word/media/sello_radicado.png", pngBuffer);

  const relPath = "word/_rels/document.xml.rels";
  let   relXml  = zip.file(relPath)?.asText() || "";
  const relId   = "rIdSello001";

  if (!relXml.includes(relId)) {
    relXml = relXml.replace(
      "</Relationships>",
      `<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/sello_radicado.png"/></Relationships>`
    );
    zip.file(relPath, relXml);
  }
  return relId;
}

// ── XML: imagen flotante rotada 90° CCW ──────────────────────────────────────
// CX = 90pt  = 1 143 000 EMU  (grosor del sello visible)
// CY = 680pt = 8 636 000 EMU  (alto del sello visible)
// rot=16200000 = 270° CW = 90° CCW
function buildImageXml(relId, cfg) {
  const CX      = 1_143_000;
  const CY      = 8_636_000;
  const lado    = cfg?.lado === "izquierdo" ? "izquierdo" : "derecho";
  const vOffset = cfg?.offsetVertical ?? 914400;

  const posH = lado === "derecho"
    ? `<wp:positionH relativeFrom="page"><wp:align>right</wp:align></wp:positionH>`
    : `<wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>`;

  return `
<w:p><w:r><w:drawing>
<wp:anchor distT="0" distB="0" distL="0" distR="0"
           simplePos="0" relativeHeight="251658240" behindDoc="0"
           locked="0" layoutInCell="1" allowOverlap="1"
           xmlns:wp="${NS.wp}">
  <wp:simplePos x="0" y="0"/>
  ${posH}
  <wp:positionV relativeFrom="page">
    <wp:posOffset>${vOffset}</wp:posOffset>
  </wp:positionV>
  <wp:extent cx="${CX}" cy="${CY}"/>
  <wp:effectExtent l="0" t="0" r="0" b="0"/>
  <wp:wrapNone/>
  <wp:docPr id="200" name="SelloRadicado"/>
  <wp:cNvGraphicFramePr>
    <a:graphicFrameLocks xmlns:a="${NS.a}" noChangeAspect="1"/>
  </wp:cNvGraphicFramePr>
  <a:graphic xmlns:a="${NS.a}">
    <a:graphicData uri="${NS.pic}">
      <pic:pic xmlns:pic="${NS.pic}">
        <pic:nvPicPr>
          <pic:cNvPr id="200" name="SelloRadicado"/>
          <pic:cNvPicPr/>
        </pic:nvPicPr>
        <pic:blipFill>
          <a:blip r:embed="${relId}" xmlns:r="${NS.r}"/>
          <a:stretch><a:fillRect/></a:stretch>
        </pic:blipFill>
        <pic:spPr>
          <a:xfrm rot="16200000">
            <a:off x="0" y="0"/>
            <a:ext cx="${CY}" cy="${CX}"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </pic:spPr>
      </pic:pic>
    </a:graphicData>
  </a:graphic>
</wp:anchor>
</w:drawing></w:r></w:p>`;
}

// ── Ajustar márgenes (twips: 1mm ≈ 56.69 → 25mm=1418, 40mm=2268) ────────────
function ajustarMargenes(docXml, lado) {
  const right = lado === "izquierdo" ? 1418 : 2268;
  const left  = lado === "izquierdo" ? 2268 : 1418;
  const nuevo = `<w:pgMar w:top="1418" w:right="${right}" w:bottom="1418" w:left="${left}" w:header="709" w:footer="709" w:gutter="0"/>`;

  if (/<w:pgMar\b/.test(docXml))
    return docXml.replace(/<w:pgMar\b[\s\S]*?\/>/, nuevo);
  if (docXml.includes("</w:sectPr>"))
    return docXml.replace("</w:sectPr>", `${nuevo}</w:sectPr>`);
  return docXml;
}

// ── Procesar un documento ─────────────────────────────────────────────────────
async function procesarDocumento(row, logoBuffer, cfg) {
  const docFile = path.join(DOCS_DIR, row.ARCHIVO_DOCUMENTO);
  if (!fs.existsSync(docFile)) {
    console.warn(`  OMITIDO — no existe: ${row.ARCHIVO_DOCUMENTO}`);
    return false;
  }

  const zip        = new PizZip(fs.readFileSync(docFile));
  const pngBuffer  = await generarSelloPNG(row, logoBuffer);
  const relId      = addStampToZip(zip, pngBuffer);
  const docXmlFile = zip.file("word/document.xml");

  if (!docXmlFile) {
    console.warn(`  OMITIDO — sin word/document.xml: ${row.ARCHIVO_DOCUMENTO}`);
    return false;
  }

  let docXml = docXmlFile.asText();
  if (!/<w:body[\s>]/.test(docXml)) {
    console.warn(`  OMITIDO — XML inesperado: ${row.ARCHIVO_DOCUMENTO}`);
    return false;
  }

  docXml = ajustarMargenes(docXml, cfg.lado);
  // Insertar al inicio del body → sello en primera página
  docXml = docXml.replace(/(<w:body[^>]*>)/, `$1${buildImageXml(relId, cfg)}`);
  zip.file("word/document.xml", docXml);

  const base    = path.basename(row.ARCHIVO_DOCUMENTO, path.extname(row.ARCHIVO_DOCUMENTO));
  const outFile = path.join(OUT_DIR, `${base}_radicado.docx`);
  fs.writeFileSync(outFile, zip.generate({ type: "nodebuffer", compression: "DEFLATE" }));
  return outFile;
}

// ── Leer Excel ────────────────────────────────────────────────────────────────
function leerRadicados() {
  if (!fs.existsSync(XLSX_FILE)) {
    console.error(`ERROR: No se encontró radicados.xlsx`);
    process.exit(1);
  }
  const wb   = XLSX.readFile(XLSX_FILE);
  const ws   = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

  rows.forEach(r => {
    if (r.FECHA_RECIBIDO && typeof r.FECHA_RECIBIDO === "number") {
      const d = XLSX.SSF.parse_date_code(r.FECHA_RECIBIDO);
      r.FECHA_RECIBIDO = `${String(d.d).padStart(2,"0")}/${String(d.m).padStart(2,"0")}/${d.y}`;
    }
    r.ENTREGADO_POR    = r.ENTREGADO_POR    ?? "";
    r.AREA_RESPONSABLE = r.AREA_RESPONSABLE ?? "";
  });
  return rows;
}

// ── MAIN ──────────────────────────────────────────────────────────────────────
async function main() {
  console.log("=".repeat(55));
  console.log("  SELLO DE RADICADO — SERVENTEGRAL S.A. E.S.P.");
  console.log("=".repeat(55));

  const cfg = leerConfig();
  console.log(`  Config: margen ${cfg.lado}, ${(cfg.offsetVertical/36000/10).toFixed(1)} cm desde arriba`);

  let logoBuffer = null;
  if (fs.existsSync(LOGO_FILE)) {
    logoBuffer = fs.readFileSync(LOGO_FILE);
    console.log("  Logo:   logo.png cargado");
  } else {
    console.warn("  AVISO:  logo.png no encontrado");
  }

  const rows = leerRadicados();
  console.log(`  Filas:  ${rows.length} registros en Excel\n`);

  // Preview con datos del primer registro
  const prev = rows.length ? { ...rows[0] } : {};
  prev.NUMERO_RADICADO  = prev.NUMERO_RADICADO  || "0001";
  prev.FECHA_RECIBIDO   = prev.FECHA_RECIBIDO   || "02/01/2026";
  prev.RECIBIDO_POR     = prev.RECIBIDO_POR     || "EJEMPLO";
  prev.ENTREGADO_POR    = prev.ENTREGADO_POR    || "EJEMPLO";
  prev.ASUNTO           = prev.ASUNTO           || "TRÁMITE";
  prev.ANEXOS           = prev.ANEXOS           || "0";
  prev.AREA_RESPONSABLE = prev.AREA_RESPONSABLE || "GENERAL";

  const previewPng = await generarSelloPNG(prev, logoBuffer);
  fs.writeFileSync(path.join(OUT_DIR, "preview_sello.png"), previewPng);
  console.log("  Preview: salida/preview_sello.png\n");

  let ok = 0, fail = 0;
  for (const row of rows) {
    if (!row.ARCHIVO_DOCUMENTO) {
      console.warn("  OMITIDO — fila sin ARCHIVO_DOCUMENTO");
      fail++; continue;
    }
    try {
      const out = await procesarDocumento(row, logoBuffer, cfg);
      if (out) {
        console.log(`  ✓ Radicado ${row.NUMERO_RADICADO} → ${path.basename(out)}`);
        ok++;
      } else { fail++; }
    } catch (e) {
      console.error(`  ✗ ERROR ${row.ARCHIVO_DOCUMENTO}: ${e.message}`);
      fail++;
    }
  }

  console.log("\n" + "=".repeat(55));
  console.log(`  ${ok} procesados · ${fail} omitidos/errores`);
  console.log(`  Archivos en: salida/`);
  console.log("=".repeat(55));
}

main().catch(e => { console.error("ERROR FATAL:", e); process.exit(1); });
