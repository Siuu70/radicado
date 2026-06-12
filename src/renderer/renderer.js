"use strict";

// ── Estado ──────────────────────────────────────────────────────────────────
const state = {
  excel:     null,  // ruta al .xlsx
  docsFiles: [],    // array de rutas absolutas (.docx o .pdf)
  out:       null,  // carpeta de salida
  logo:      null,  // ruta logo PNG
  running:   false,
};

// ── Refs DOM ────────────────────────────────────────────────────────────────
const inpExcel      = document.getElementById("inp-excel");
const inpDocsFiles  = document.getElementById("inp-docs-files");
const inpOut        = document.getElementById("inp-out");
const inpLogo       = document.getElementById("inp-logo");

const btnExcel      = document.getElementById("btn-excel");
const btnDocsFiles  = document.getElementById("btn-docs-files");
const btnOut        = document.getElementById("btn-out");
const btnLogo       = document.getElementById("btn-logo");
const btnProcess    = document.getElementById("btn-process");
const btnClear      = document.getElementById("btn-clear");
const btnOpenOut    = document.getElementById("btn-open-out");
const btnClearFiles = document.getElementById("btn-clear-files");

const fileListWrap  = document.getElementById("file-list-wrap");
const fileListEl    = document.getElementById("file-list");
const logEl         = document.getElementById("log");
const progressFill  = document.getElementById("progress-fill");
const progressText  = document.getElementById("progress-text");
const resultText    = document.getElementById("result-text");

// ── Helpers ─────────────────────────────────────────────────────────────────
function pad(n) { return String(n).padStart(2, "0"); }
function ts() {
  const n = new Date();
  return `${pad(n.getHours())}:${pad(n.getMinutes())}:${pad(n.getSeconds())}`;
}

function log(text, type = "info") {
  const line = document.createElement("div");
  const tsSpan = document.createElement("span");
  tsSpan.className = "line-ts";
  tsSpan.textContent = `[${ts()}]`;
  const prefix = { ok: "✓ ", warn: "⚠ ", error: "✗ ", info: "  ", done: "━━ " }[type] ?? "  ";
  line.className = `line-${type}`;
  line.appendChild(tsSpan);
  line.appendChild(document.createTextNode(`${prefix}${text}`));
  logEl.appendChild(line);
  logEl.scrollTop = logEl.scrollHeight;
}

function setProgress(current, total) {
  const pct = total > 0 ? Math.round((current / total) * 100) : 0;
  progressFill.style.width = pct + "%";
  progressText.textContent = `${current} / ${total}`;
}

function shortPath(p) {
  if (!p) return "";
  const parts = p.replace(/\\/g, "/").split("/");
  return parts.length <= 3 ? p : "…/" + parts.slice(-2).join("/");
}

function ext(p) { return p.split(".").pop().toLowerCase(); }

function validateReady() {
  btnProcess.disabled = state.running || !state.excel || state.docsFiles.length === 0 || !state.out;
}

// ── Logo en header ──────────────────────────────────────────────────────────
function updateHeaderLogo(logoPath) {
  if (!logoPath) return;
  const wrap = document.getElementById("header-logo-wrap");
  const placeholder = document.getElementById("header-placeholder");
  let img = wrap.querySelector("img");
  if (!img) {
    img = document.createElement("img");
    img.id = "header-logo";
    wrap.appendChild(img);
  }
  img.style.display = "block";
  img.src = "file:///" + logoPath.replace(/\\/g, "/");
  img.onerror = () => { img.style.display = "none"; placeholder.style.display = "flex"; };
  placeholder.style.display = "none";
}

async function loadDefaultLogo() {
  // Sin logo guardado: el processor usa el logo.png incluido en el exe automáticamente.
  // Solo intentamos mostrarlo en el header para feedback visual.
  state.logo = null;
  inpLogo.value = "Logo incluido (por defecto)";
  const resPath = await window.api.getResourcePath();
  const sep = resPath.includes("/") ? "/" : "\\";
  updateHeaderLogo(resPath + sep + "logo.png");
}

// Actualiza el logo cuando se cambia el Excel (busca logo.png junto al xlsx)
function tryLogoFromExcelDir(excelPath) {
  if (!excelPath) return;
  const sep = excelPath.includes("/") ? "/" : "\\";
  const dir = excelPath.split(sep).slice(0,-1).join(sep);
  const candidate = dir + sep + "logo.png";
  state.logo = candidate;
  inpLogo.value = shortPath(candidate);
  updateHeaderLogo(candidate);
}

// ── Render lista de archivos ─────────────────────────────────────────────────
function renderFileList() {
  const files = state.docsFiles;
  fileListEl.innerHTML = "";

  if (files.length === 0) {
    fileListWrap.style.display = "none";
    inpDocsFiles.value = "";
    return;
  }

  const plural = files.length === 1 ? "archivo" : "archivos";
  inpDocsFiles.value = `${files.length} ${plural} seleccionado(s)`;
  fileListWrap.style.display = "block";

  files.forEach((fp) => {
    const name = fp.replace(/\\/g, "/").split("/").pop();
    const type = ext(fp);

    const item = document.createElement("div");
    item.className = "file-item";
    item.title = fp;

    const icon = document.createElement("span");
    icon.className = "file-icon";
    icon.textContent = type === "pdf" ? "📕" : "📄";

    const nameSpan = document.createElement("span");
    nameSpan.className = "file-name";
    nameSpan.textContent = name;

    const badge = document.createElement("span");
    badge.className = `file-badge ${type}`;
    badge.textContent = type.toUpperCase();

    item.appendChild(icon);
    item.appendChild(nameSpan);
    item.appendChild(badge);
    fileListEl.appendChild(item);
  });
}

// ── Selectores ──────────────────────────────────────────────────────────────
btnExcel.addEventListener("click", async () => {
  const p = await window.api.selectExcel();
  if (p) {
    state.excel = p;
    inpExcel.value = shortPath(p);
    log(`Excel: ${p}`, "info");
    // Buscar logo.png en la misma carpeta del Excel
    tryLogoFromExcelDir(p);
    pvSyncLogo(state.logo);
    validateReady();
  }
});

btnDocsFiles.addEventListener("click", async () => {
  const files = await window.api.selectDocFiles();
  if (files && files.length > 0) {
    state.docsFiles = files;
    renderFileList();
    log(`${files.length} archivo(s) seleccionado(s)`, "info");
    files.forEach((f) => {
      const name = f.replace(/\\/g, "/").split("/").pop();
      const tipo = ext(f).toUpperCase();
      log(`  • [${tipo}] ${name}`, "info");
    });
    validateReady();
  }
});

btnClearFiles.addEventListener("click", () => {
  state.docsFiles = [];
  renderFileList();
  validateReady();
  log("Selección de archivos limpiada.", "info");
});

btnOut.addEventListener("click", async () => {
  const p = await window.api.selectFolder("Carpeta de salida");
  if (p) {
    state.out = p;
    inpOut.value = shortPath(p);
    log(`Carpeta salida: ${p}`, "info");
    validateReady();
  }
});

btnLogo.addEventListener("click", async () => {
  const p = await window.api.selectLogo();
  if (p) {
    state.logo = p;
    inpLogo.value = shortPath(p);
    updateHeaderLogo(p);
    pvSyncLogo(p);
    await window.api.saveConfig({ ...pvCfg, logoPath: p });
    log(`Logo guardado: ${p}`, "info");
  }
});

// ── Limpiar log ──────────────────────────────────────────────────────────────
btnClear.addEventListener("click", () => {
  logEl.innerHTML = "";
  setProgress(0, 0);
  resultText.textContent = "Esperando proceso…";
  resultText.style.color = "#9E9E9E";
});

btnOpenOut.addEventListener("click", () => {
  if (state.out) window.api.openFolder(state.out);
});

// ── Procesar ─────────────────────────────────────────────────────────────────
btnProcess.addEventListener("click", async () => {
  if (state.running) return;
  state.running = true;
  btnProcess.disabled = true;
  document.getElementById("btn-icon").textContent  = "⏳";
  document.getElementById("btn-label").textContent = "Procesando…";
  resultText.textContent = "Procesando…";
  resultText.style.color = "#1E7B34";
  btnOpenOut.disabled = true;
  logEl.innerHTML = "";
  setProgress(0, 0);

  log("Iniciando procesamiento…", "info");

  window.api.removeProgressListener();
  window.api.onProgress((data) => {
    if (data.type === "progress") setProgress(data.current, data.total);
    else log(data.msg, data.type);
  });

  const result = await window.api.processDocuments({
    xlsxPath:  state.excel,
    docsFiles: state.docsFiles,
    outDir:    state.out,
    logoPath:  state.logo,
    cfg:       pvCfg,
  });

  window.api.removeProgressListener();

  if (result.success) {
    setProgress(result.total, result.total);
    log(`Completado: ${result.ok} procesados, ${result.fail} omitidos de ${result.total}`, "done");
    resultText.innerHTML =
      `<span class="stat-ok">✓ ${result.ok} procesados</span>&nbsp;&nbsp;` +
      `<span class="stat-fail">✗ ${result.fail} errores</span>&nbsp;&nbsp;` +
      `de ${result.total} total`;
    btnOpenOut.disabled = false;
  } else {
    log(`Error: ${result.error}`, "error");
    resultText.textContent = `Error: ${result.error}`;
    resultText.style.color = "#D32F2F";
  }

  state.running = false;
  document.getElementById("btn-icon").textContent  = "▶";
  document.getElementById("btn-label").textContent = "Procesar documentos";
  validateReady();
});

// ════════════════════════════════════════════════════════════════════════════
//  TABS
// ════════════════════════════════════════════════════════════════════════════

document.getElementById("tab-btn-procesar").addEventListener("click", () => switchTab("procesar"));
document.getElementById("tab-btn-preview").addEventListener("click",  () => switchTab("preview"));
document.getElementById("hint-go-preview").addEventListener("click",  () => switchTab("preview"));

function switchTab(name) {
  document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
  document.querySelectorAll(".tab-panel").forEach(p => p.classList.remove("active"));
  document.getElementById("tab-btn-" + name).classList.add("active");
  document.getElementById("panel-" + name).classList.add("active");
  if (name === "preview") pvDrawAll();
}

// ════════════════════════════════════════════════════════════════════════════
//  PREVIEW — estado
// ════════════════════════════════════════════════════════════════════════════

let pvLogoImg  = null;
let pvCfg      = { lado: "derecho", offsetVertical: 914400 };
let pvTimer    = null;

// Intentar cargar logo.png de la misma carpeta del logo seleccionado o del excel
function pvTryLoadLogo(logoPath) {
  if (!logoPath) return;
  const img = new Image();
  img.onload  = () => { pvLogoImg = img; pvDrawAll(); };
  img.onerror = () => {};
  img.src = "file:///" + logoPath.replace(/\\/g, "/");
}

function pvDebounceDraw() {
  clearTimeout(pvTimer);
  pvTimer = setTimeout(pvDrawAll, 120);
}

// ── Dibujar sello (diseño aprobado 900×278) ───────────────────────────────
function pvDrawStamp(ctx, data, logo) {
  const W = 900, H = 278;
  ctx.clearRect(0, 0, W, H);
  ctx.fillStyle = "#FFFFFF";
  ctx.fillRect(0, 0, W, H);

  // Franja superior
  ctx.fillStyle = "#1B5E20";
  ctx.fillRect(0, 0, W, 35);
  ctx.fillStyle    = "#FFFFFF";
  ctx.textBaseline = "middle";
  ctx.textAlign    = "left";
  ctx.font         = "bold 12px Arial";
  ctx.fillText("SERVENTEGRAL S.A. E.S.P", 15, 17);
  ctx.textAlign = "right";
  ctx.font      = "11px Arial";
  ctx.fillText("NIT: 828.002.229-2", 885, 17);
  ctx.strokeStyle = "rgba(255,255,255,0.25)";
  ctx.lineWidth   = 1;
  ctx.beginPath(); ctx.moveTo(0, 35); ctx.lineTo(W, 35); ctx.stroke();

  // Columna izquierda (logo)
  ctx.fillStyle = "#1B5E20";
  ctx.fillRect(0, 35, 200, 243);
  if (logo) {
    const maxW = 180, maxH = 223;
    const scale = Math.min(maxW / logo.width, maxH / logo.height);
    const iw = logo.width * scale, ih = logo.height * scale;
    ctx.drawImage(logo, (200 - iw) / 2, 35 + (243 - ih) / 2, iw, ih);
  }
  ctx.strokeStyle = "rgba(255,255,255,0.25)";
  ctx.lineWidth   = 1;
  ctx.beginPath(); ctx.moveTo(200, 35); ctx.lineTo(200, H); ctx.stroke();

  // Filas de datos
  const filas = [
    { label: "N° RADICADO:",  value: data.rad,  h: 43, bold: true,  size: 22 },
    { label: "FECHA:",        value: data.fec,  h: 40, bold: false, size: 12 },
    { label: "RECIBIDO POR:", value: data.rec,  h: 40, bold: false, size: 12 },
    { label: "ASUNTO:",       value: data.asu,  h: 40, bold: false, size: 12 },
    { label: "ANEXOS:",       value: data.ane,  h: 40, bold: false, size: 12 },
    { label: "ÁREA RESP:",    value: data.area, h: 40, bold: false, size: 12 },
  ];
  const bgFilas = ["#FFFFFF", "#F1F8E9"];
  ctx.textAlign    = "left";
  ctx.textBaseline = "middle";
  let ry = 35;
  for (let i = 0; i < filas.length; i++) {
    const f = filas[i];
    ctx.fillStyle = bgFilas[i % 2]; ctx.fillRect(200, ry, 700, f.h);
    ctx.fillStyle = "#1B5E20"; ctx.font = "bold 10px Arial";
    ctx.fillText(f.label, 210, ry + f.h / 2);
    ctx.fillStyle = "#000000"; ctx.font = `${f.bold ? "bold " : ""}${f.size}px Arial`;
    ctx.fillText(String(f.value || ""), 370, ry + f.h / 2);
    ry += f.h;
    ctx.strokeStyle = "#C8E6C9"; ctx.lineWidth = 1;
    ctx.beginPath(); ctx.moveTo(200, ry); ctx.lineTo(W, ry); ctx.stroke();
  }

  // Borde exterior
  ctx.strokeStyle = "#1B5E20";
  ctx.lineWidth   = 2;
  ctx.strokeRect(1, 1, 898, 276);
}

// ── Dibujar hoja A4 con sello rotado ─────────────────────────────────────
function pvDrawA4(stampCanvas) {
  const a4  = document.getElementById("pv-a4");
  const ctx = a4.getContext("2d");
  const SC  = 2, PW = 420, PH = 594;
  const SW  = (1143000 / 36000) * SC;
  const SH  = (8636000 / 36000) * SC;

  ctx.clearRect(0, 0, PW, PH);
  ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, PW, PH);

  const der = pvCfg.lado === "derecho";
  const ml  = (der ? 25 : 40) * SC, mr = (der ? 40 : 25) * SC, cw = PW - ml - mr;
  ctx.fillStyle = "#EBEBEB";
  for (let i = 0; i < 30; i++) {
    const ly = 25 * SC + i * 16; if (ly > PH - 30) break;
    ctx.fillRect(ml, ly, i % 6 === 5 ? cw * .55 : cw, 4);
  }

  ctx.strokeStyle = "#D8D8D8"; ctx.lineWidth = .5; ctx.setLineDash([4, 4]);
  const gx = der ? PW - 40 * SC : 40 * SC;
  ctx.beginPath(); ctx.moveTo(gx, 0); ctx.lineTo(gx, PH); ctx.stroke();
  ctx.setLineDash([]);

  const sy = (pvCfg.offsetVertical / 36000) * SC;
  const sx = der ? PW - SW : 0;
  ctx.save();
  ctx.translate(sx + SW / 2, sy + SH / 2);
  ctx.rotate(der ? -Math.PI / 2 : Math.PI / 2);
  ctx.drawImage(stampCanvas, -SH / 2, -SW / 2, SH, SW);
  ctx.restore();

  if (sy + SH > PH) {
    ctx.fillStyle = "rgba(200,0,0,.2)"; ctx.fillRect(sx, PH - 5, SW, 5);
    ctx.fillStyle = "#C62828"; ctx.font = "10px Arial";
    ctx.textAlign = "center"; ctx.textBaseline = "bottom";
    ctx.fillText("⚠ Sello fuera de página", PW / 2, PH - 2);
  }

  ctx.strokeStyle = "#BBB"; ctx.lineWidth = 1; ctx.strokeRect(0, 0, PW, PH);
}

function pvDrawAll() {
  const data = {
    rad:  document.getElementById("pv-rad").value,
    fec:  document.getElementById("pv-fec").value,
    rec:  document.getElementById("pv-rec").value,
    asu:  document.getElementById("pv-asu").value,
    ane:  document.getElementById("pv-ane").value,
    area: document.getElementById("pv-area").value,
  };
  const sc = document.getElementById("pv-stamp");
  pvDrawStamp(sc.getContext("2d"), data, pvLogoImg);
  pvDrawA4(sc);
}

// ── Controles del preview ────────────────────────────────────────────────

["pv-rad","pv-fec","pv-rec","pv-asu","pv-ane","pv-area"].forEach(id => {
  document.getElementById(id).addEventListener("input", pvDebounceDraw);
});

document.getElementById("pv-btn-der").addEventListener("click", () => pvSetSide("derecho"));
document.getElementById("pv-btn-izq").addEventListener("click", () => pvSetSide("izquierdo"));

function pvSetSide(lado) {
  pvCfg.lado = lado;
  document.getElementById("pv-btn-der").classList.toggle("active", lado === "derecho");
  document.getElementById("pv-btn-izq").classList.toggle("active", lado === "izquierdo");
  pvDrawAll();
}

document.getElementById("pv-slider").addEventListener("input", function () {
  pvCfg.offsetVertical = parseInt(this.value, 10);
  document.getElementById("pv-v-label").textContent =
    (pvCfg.offsetVertical / 36000 / 10).toFixed(1) + " cm";
  pvDrawAll();
});

document.getElementById("pv-save").addEventListener("click", async () => {
  const st = document.getElementById("pv-status");
  const saveData = { ...pvCfg };
  if (state.logo) saveData.logoPath = state.logo;
  const result = await window.api.saveConfig(saveData);
  if (result.ok) {
    st.className = "ok";
    st.textContent = "✓ Configuración guardada.";
  } else {
    st.className = "err";
    st.textContent = "✗ " + result.error;
  }
});

// Sincronizar logo del preview cuando se seleccione en la pestaña Procesar
function pvSyncLogo(logoPath) {
  if (logoPath) pvTryLoadLogo(logoPath);
}

// ── Inicialización ────────────────────────────────────────────────────────────
(async () => {
  log("Aplicación lista.", "info");
  log("Seleccione el Excel de radicados, los documentos (.docx o .pdf) y la carpeta de salida.", "info");

  // Cargar configuración guardada
  const savedCfg = await window.api.getConfig();
  pvCfg = { ...pvCfg, ...savedCfg };
  document.getElementById("pv-slider").value = pvCfg.offsetVertical;
  document.getElementById("pv-v-label").textContent =
    (pvCfg.offsetVertical / 36000 / 10).toFixed(1) + " cm";
  pvSetSide(pvCfg.lado);

  if (savedCfg.logoPath) {
    state.logo = savedCfg.logoPath;
    inpLogo.value = shortPath(savedCfg.logoPath);
    updateHeaderLogo(savedCfg.logoPath);
    pvSyncLogo(savedCfg.logoPath);
  } else {
    await loadDefaultLogo();
  }
  validateReady();
})();
