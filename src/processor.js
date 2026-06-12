"use strict";

const fs     = require("fs");
const path   = require("path");
const XLSX   = require("xlsx");
const PizZip = require("pizzip");
const { createCanvas, loadImage } = require("canvas");
const { app } = require("electron");

// Raíz de recursos: resources/ en producción, carpeta del proyecto en dev
const BASE = app.isPackaged
  ? process.resourcesPath
  : path.join(__dirname, "..");

// Retorna la ruta real del logo: logo del usuario → logo por defecto → null
function getLogoPath(cfgLogoPath) {
  if (cfgLogoPath && fs.existsSync(cfgLogoPath)) return cfgLogoPath;
  const def = path.join(BASE, "logo.png");
  return fs.existsSync(def) ? def : null;
}

// ══════════════════════════════════════════════════════════════════════════════
//  MAPEO FLEXIBLE DE COLUMNAS
// ══════════════════════════════════════════════════════════════════════════════

function get(raw, ...names) {
  const keys = Object.keys(raw);
  for (const name of names) {
    const n = name.trim().toLowerCase();
    const k = keys.find((k) => k.trim().toLowerCase() === n);
    if (k !== undefined && raw[k] !== null && raw[k] !== undefined) {
      const v = String(raw[k]).trim();
      if (v) return v;
    }
  }
  return "";
}

function getDate(raw, ...names) {
  const keys = Object.keys(raw);
  for (const name of names) {
    const n = name.trim().toLowerCase();
    const k = keys.find((k) => k.trim().toLowerCase() === n);
    if (k !== undefined && raw[k] !== null && raw[k] !== undefined && raw[k] !== "") {
      const v = raw[k];
      if (typeof v === "number") {
        const d = XLSX.SSF.parse_date_code(v);
        return `${String(d.d).padStart(2,"0")}/${String(d.m).padStart(2,"0")}/${d.y}`;
      }
      return String(v).trim();
    }
  }
  return "";
}

function normalizeRow(raw) {
  const radRaw = get(raw,"radicado","numero_radicado","n° radicado","nro radicado","numero radicado");
  return {
    NUMERO_RADICADO:   /^\d+$/.test(radRaw) ? String(parseInt(radRaw,10)).padStart(4,"0") : radRaw,
    RECIBIDO_POR:      get(raw,"quien recibio","quen recibio","recibido_por","recibido por","nombre"),
    ENTREGADO_POR:     get(raw,"entregado_por","entregado por","entregado","quien entrego","quien entregó"),
    ASUNTO:            get(raw,"asunto"),
    ANEXOS:            get(raw,"anexos","tipo de solicitud","entidad"),
    AREA_RESPONSABLE:  get(raw,"area_responsable","area responsable","área responsable","área resp","area"),
    FECHA_RECIBIDO:    getDate(raw,"fecha_recibido","fecha recibido","fecha de ingreso de la solicitud","fecha ingreso","fecha"),
    CONSECUTIVO:       get(raw,"consecutivo","archivo_documento"),
    ARCHIVO_DOCUMENTO: get(raw,"archivo_documento"),
  };
}

function leerRadicados(xlsxPath) {
  const wb  = XLSX.readFile(xlsxPath);
  const ws  = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { defval: "" }).map(normalizeRow);
}

// ══════════════════════════════════════════════════════════════════════════════
//  EMPAREJAMIENTO archivos ↔ filas
// ══════════════════════════════════════════════════════════════════════════════

function buildPairs(rows, docsFiles) {
  const byBase = new Map();
  const byFull = new Map();
  for (const fp of docsFiles) {
    byFull.set(path.basename(fp).toLowerCase(), fp);
    byBase.set(path.basename(fp, path.extname(fp)).toLowerCase(), fp);
  }
  const pairs    = [];
  const usedRows = new Set();
  for (const fp of docsFiles) {
    const base = path.basename(fp, path.extname(fp)).toLowerCase();
    const full = path.basename(fp).toLowerCase();
    let matchRow = null;
    for (let i = 0; i < rows.length; i++) {
      if (usedRows.has(i)) continue;
      const r = rows[i];
      const consec = (r.CONSECUTIVO || "").toLowerCase();
      const adoc   = (r.ARCHIVO_DOCUMENTO || "").toLowerCase();
      if (consec===base || consec===full || adoc===full ||
          path.basename(adoc, path.extname(adoc))===base) {
        matchRow = i; break;
      }
    }
    if (matchRow !== null) {
      usedRows.add(matchRow);
      pairs.push({ row: rows[matchRow], filePath: fp, matched: "key" });
    } else {
      pairs.push({ row: null, filePath: fp, matched: "pending" });
    }
  }
  const freeRows = rows.map((r,i)=>({r,i})).filter(({i})=>!usedRows.has(i)).map(({r})=>r);
  let fi = 0;
  for (const p of pairs) {
    if (p.matched === "pending") { p.row = freeRows[fi] ?? null; p.matched = "order"; fi++; }
  }
  return pairs;
}

// ══════════════════════════════════════════════════════════════════════════════
//  SELLO PNG  (canvas 900×278 — diseño con franja superior + logo izquierdo)
// ══════════════════════════════════════════════════════════════════════════════

async function generarSelloPNG(fields, logoBuffer) {
  const W = 900, H = 278;
  const canvas = createCanvas(W, H);
  const ctx    = canvas.getContext("2d");

  // Fondo blanco base
  ctx.fillStyle = "#FFFFFF";
  ctx.fillRect(0, 0, W, H);

  // ── FRANJA SUPERIOR (y=0, h=35px, ancho completo) ────────────────────────
  ctx.fillStyle = "#1B5E20";
  ctx.fillRect(0, 0, W, 35);

  ctx.fillStyle    = "#FFFFFF";
  ctx.textBaseline = "middle";

  ctx.textAlign = "left";
  ctx.font      = "bold 12px Arial";
  ctx.fillText("SERVENTEGRAL S.A. E.S.P", 15, 17);

  ctx.textAlign = "right";
  ctx.font      = "11px Arial";
  ctx.fillText("NIT: 828.002.229-2", 885, 17);

  // Línea separadora inferior de la franja
  ctx.strokeStyle = "rgba(255,255,255,0.25)";
  ctx.lineWidth   = 1;
  ctx.beginPath(); ctx.moveTo(0, 35); ctx.lineTo(W, 35); ctx.stroke();

  // ── COLUMNA IZQUIERDA (x=0, w=200, y=35..278) ────────────────────────────
  ctx.fillStyle = "#1B5E20";
  ctx.fillRect(0, 35, 200, 243);

  if (logoBuffer) {
    try {
      const img   = await loadImage(logoBuffer);
      const maxW  = 180;
      const maxH  = 223;
      const scale = Math.min(maxW / img.width, maxH / img.height);
      const iw    = img.width  * scale;
      const ih    = img.height * scale;
      const ix    = (200 - iw) / 2;
      const iy    = 35 + (243 - ih) / 2;
      ctx.drawImage(img, ix, iy, iw, ih);
    } catch { /* logo inaccesible */ }
  }

  // Separador derecho columna izquierda
  ctx.strokeStyle = "rgba(255,255,255,0.25)";
  ctx.lineWidth   = 1;
  ctx.beginPath(); ctx.moveTo(200, 35); ctx.lineTo(200, H); ctx.stroke();

  // ── COLUMNA DERECHA: filas de datos (x=200..900, y=35..278) ──────────────
  const filas = [
    { label: "N° RADICADO:",  value: String(fields.NUMERO_RADICADO  ?? ""), h: 43, bold: true,  size: 22 },
    { label: "FECHA:",        value: String(fields.FECHA_RECIBIDO   ?? ""), h: 40, bold: false, size: 12 },
    { label: "RECIBIDO POR:", value: String(fields.RECIBIDO_POR     ?? ""), h: 40, bold: false, size: 12 },
    { label: "ASUNTO:",       value: String(fields.ASUNTO           ?? ""), h: 40, bold: false, size: 12 },
    { label: "ANEXOS:",       value: String(fields.ANEXOS           ?? ""), h: 40, bold: false, size: 12 },
    { label: "ÁREA RESP:",    value: String(fields.AREA_RESPONSABLE ?? ""), h: 40, bold: false, size: 12 },
  ];

  const bgFilas = ["#FFFFFF", "#F1F8E9"];
  ctx.textAlign    = "left";
  ctx.textBaseline = "middle";
  let ry = 35;

  for (let i = 0; i < filas.length; i++) {
    const f = filas[i];

    ctx.fillStyle = bgFilas[i % 2];
    ctx.fillRect(200, ry, 700, f.h);

    ctx.fillStyle = "#1B5E20";
    ctx.font      = "bold 10px Arial";
    ctx.fillText(f.label, 210, ry + f.h / 2);

    ctx.fillStyle = "#000000";
    ctx.font      = `${f.bold ? "bold " : ""}${f.size}px Arial`;
    ctx.fillText(f.value, 370, ry + f.h / 2);

    ry += f.h;

    ctx.strokeStyle = "#C8E6C9";
    ctx.lineWidth   = 1;
    ctx.beginPath(); ctx.moveTo(200, ry); ctx.lineTo(W, ry); ctx.stroke();
  }

  // ── BORDE EXTERIOR ────────────────────────────────────────────────────────
  ctx.strokeStyle = "#1B5E20";
  ctx.lineWidth   = 2;
  ctx.strokeRect(1, 1, 898, 276);

  return canvas.toBuffer("image/png");
}

// ══════════════════════════════════════════════════════════════════════════════
//  SELLO EN DOCX  (PNG flotante, rotado 90° CCW)
// ══════════════════════════════════════════════════════════════════════════════

const NS = {
  r:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  wp:  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  a:   "http://schemas.openxmlformats.org/drawingml/2006/main",
  pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
};

function addSelloPNGToZip(zip, pngBuffer) {
  zip.file("word/media/sello_radicado.png", pngBuffer);
  const relPath = "word/_rels/document.xml.rels";
  let   relXml  = zip.file(relPath)?.asText() || "";
  const relId   = "rIdSello001";
  if (!relXml.includes(relId)) {
    relXml = relXml.replace("</Relationships>",
      `<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/sello_radicado.png"/></Relationships>`);
    zip.file(relPath, relXml);
  }
  return relId;
}

// CX = 90pt  = 1 143 000 EMU (grosor visible post-rotación)
// CY = 680pt = 8 636 000 EMU (alto visible post-rotación)
// rot=16200000 = 270° CW = 90° CCW
function buildImageXml(relId, cfg) {
  const CX      = 1_143_000;
  const CY      = 8_636_000;
  const lado    = cfg?.lado === "izquierdo" ? "izquierdo" : "derecho";
  const vOffset = cfg?.offsetVertical ?? 914400;

  const posH = lado === "derecho"
    ? `<wp:positionH relativeFrom="page"><wp:align>right</wp:align></wp:positionH>`
    : `<wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>`;

  return `<w:p><w:r><w:drawing>
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

// Ajustar márgenes: lado sello=40mm (2268 twips), resto=25mm (1418 twips)
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

function leerConfig(xlsxPath) {
  const def = { lado: "derecho", offsetVertical: 914400 };
  const cfgPath = path.join(path.dirname(xlsxPath), "config.json");
  if (!fs.existsSync(cfgPath)) return def;
  try { return { ...def, ...JSON.parse(fs.readFileSync(cfgPath, "utf8")) }; }
  catch { return def; }
}

// Elimina cualquier sello previo usando indexOf (no regex, 100% fiable)
function limpiarSelloAnterior(docXml) {
  const marker = 'r:embed="rIdSello001"';
  let idx = docXml.indexOf(marker);
  while (idx >= 0) {
    // Buscar <w:p hacia atrás desde idx
    const pOpen = docXml.lastIndexOf('<w:p', idx);
    // Buscar </w:p> hacia adelante desde idx
    const pClose = docXml.indexOf('</w:p>', idx);
    if (pOpen < 0 || pClose < 0) break;
    // Eliminar todo el párrafo del sello
    docXml = docXml.slice(0, pOpen) + docXml.slice(pClose + 6);
    // Buscar si hay otro sello (no debería, pero por seguridad)
    idx = docXml.indexOf(marker);
  }
  return docXml;
}

async function procesarDocx(fields, docFile, outDir, logoBuffer, cfg) {
  const zip     = new PizZip(fs.readFileSync(docFile));
  const xmlFile = zip.file("word/document.xml");
  if (!xmlFile) return { ok: false, reason: "word/document.xml no encontrado" };

  const pngBuffer = await generarSelloPNG(fields, logoBuffer);
  const relId     = addSelloPNGToZip(zip, pngBuffer);

  let docXml = xmlFile.asText();

  // Eliminar sello anterior si existe
  docXml = limpiarSelloAnterior(docXml);

  docXml = ajustarMargenes(docXml, cfg?.lado ?? "derecho");

  // Insertar sello al inicio de <w:body> usando indexOf (no regex)
  const selloXml  = buildImageXml(relId, cfg);
  const bodyIdx   = docXml.indexOf('<w:body');
  if (bodyIdx < 0) return { ok: false, reason: "no se encontró <w:body> en el XML" };
  const bodyClose = docXml.indexOf('>', bodyIdx);      // cierre del tag <w:body...>
  docXml = docXml.slice(0, bodyClose + 1) + selloXml + docXml.slice(bodyClose + 1);

  zip.file("word/document.xml", docXml);
  const outFile = path.join(outDir, path.basename(docFile, ".docx") + "_radicado.docx");
  fs.writeFileSync(outFile, zip.generate({ type: "nodebuffer", compression: "DEFLATE" }));
  return { ok: true, outFile };
}

// ══════════════════════════════════════════════════════════════════════════════
//  SELLO EN PDF  (incrusta el PNG aprobado 900×278 en la PRIMERA página)
// ══════════════════════════════════════════════════════════════════════════════

async function procesarPDF(fields, docFile, outDir, logoBuffer, cfg) {
  const { PDFDocument, degrees } = require("pdf-lib");

  const pdfDoc   = await PDFDocument.load(fs.readFileSync(docFile));
  const pngBuffer = await generarSelloPNG(fields, logoBuffer);
  const pngImage  = await pdfDoc.embedPng(pngBuffer);

  // Siempre primera página
  const page = pdfDoc.getPages()[0];
  const { width: pw, height: ph } = page.getSize();

  const lado = cfg?.lado === "izquierdo" ? "izquierdo" : "derecho";
  const vEmu = cfg?.offsetVertical ?? 914400;

  // EMU → puntos tipográficos  (914400 EMU = 1 pulgada = 72 pt)
  const vPt = vEmu / 12700;

  // Dimensiones visibles del sello (idénticas al DOCX)
  const sw = 1486000 / 12700;   // ≈ 117 pt grosor visible (ratio 900/278 = 3.24)
  const sh = 4800000 / 12700;   // ≈ 378 pt largo visible del sello

  if (lado === "derecho") {
    // Ampliar la página a la derecha para que el contenido original NO se tape
    page.setSize(pw + sw, ph);
    // Con rotate(90°CCW): pivot en (x,y), imagen ocupa x-H…x horizontal, y…y+W vertical
    // x = pw+sw → franja ocupa pw a pw+sw (zona nueva, fuera del contenido original)
    page.drawImage(pngImage, {
      x: pw + sw,
      y: (ph - sh) / 2,
      width:  sh,
      height: sw,
      rotate: degrees(90),
    });
  } else {
    // IZQUIERDO: extender el MediaBox hacia la izquierda — el contenido existente
    // no se desplaza; la zona nueva x=[-sw, 0] queda libre para el sello.
    const { x: mbX, y: mbY } = page.getMediaBox();
    page.setMediaBox(mbX - sw, mbY, pw + sw, ph);
    page.drawImage(pngImage, {
      x: mbX,             // pivot en el borde izquierdo original → sello ocupa [mbX-sw, mbX]
      y: (ph - sh) / 2,
      width:  sh,
      height: sw,
      rotate: degrees(90),
    });
  }

  const outFile = path.join(outDir, path.basename(docFile, ".pdf") + "_radicado.pdf");
  fs.writeFileSync(outFile, await pdfDoc.save());
  return { ok: true, outFile };
}

// ══════════════════════════════════════════════════════════════════════════════
//  FUNCIÓN PRINCIPAL EXPORTADA
// ══════════════════════════════════════════════════════════════════════════════

async function processAll({ xlsxPath, docsFiles, outDir, logoPath, cfg: cfgPassed }, onProgress) {
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

  console.log("=== PROCESSOR v3 (900x278, solo PDF) ===");
  onProgress({ type: "info", msg: "=== Sello v3 — solo PDF, canvas 900×278px, primera página ===" });

  const cfg = cfgPassed ?? leerConfig(xlsxPath);
  onProgress({ type: "info", msg: `Config: margen ${cfg.lado}, ${(cfg.offsetVertical/36000/10).toFixed(1)} cm desde arriba` });

  const resolvedLogo = getLogoPath(logoPath);
  let logoBuffer = null;
  if (resolvedLogo) {
    logoBuffer = fs.readFileSync(resolvedLogo);
    onProgress({ type: "info", msg: `Logo: ${path.basename(resolvedLogo)}` });
  } else {
    onProgress({ type: "warn", msg: "logo.png no encontrado — el sello no tendrá logo." });
  }

  const rows = leerRadicados(xlsxPath);
  onProgress({ type: "info", msg: `Registros en Excel: ${rows.length}` });
  onProgress({ type: "info", msg: `Archivos seleccionados: ${docsFiles.length}` });

  // Calcular el último número de radicado válido y el siguiente a usar
  const filasValidas = rows.filter(r => r.NUMERO_RADICADO && !isNaN(Number(r.NUMERO_RADICADO)));
  const ultimoNum    = filasValidas.length > 0 ? Number(filasValidas[filasValidas.length - 1].NUMERO_RADICADO) : 0;
  onProgress({ type: "info", msg: `Último radicado en Excel: ${ultimoNum} → siguiente: ${ultimoNum + 1}` });

  const pairs = buildPairs(rows, docsFiles);
  let ok = 0, fail = 0;

  for (let i = 0; i < pairs.length; i++) {
    const { row, filePath, matched } = pairs[i];
    onProgress({ type: "progress", current: i + 1, total: pairs.length });
    const nombre = path.basename(filePath);
    if (!row) {
      onProgress({ type: "warn", msg: `Sin fila Excel para: ${nombre} — omitido` });
      fail++; continue;
    }
    // Auto-generar número de radicado: último Excel + posición en el lote
    const nuevoNumero    = String(ultimoNum + i + 1).padStart(4, "0");
    const rowConNumero   = { ...row, NUMERO_RADICADO: nuevoNumero };
    const tipo = path.extname(filePath).toLowerCase();
    try {
      let result;
      if (tipo === ".pdf") result = await procesarPDF(rowConNumero, filePath, outDir, logoBuffer, cfg);
      else                 result = { ok: false, reason: `Solo PDF — ignorado (${tipo})` };

      if (result.ok) { ok++;   onProgress({ type: "ok",    msg: `Procesado: ${nombre} — Radicado ${nuevoNumero}` }); }
      else           { fail++; onProgress({ type: "error", msg: `Omitido: ${nombre} — ${result.reason}` }); }
    } catch (err) {
      fail++; onProgress({ type: "error", msg: `ERROR en ${nombre}: ${err.message}` });
    }
  }
  return { ok, fail, total: pairs.length };
}

module.exports = { processAll };
