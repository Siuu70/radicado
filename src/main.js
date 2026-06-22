"use strict";

const { app, BrowserWindow, ipcMain, dialog, shell } = require("electron");
const path = require("path");
const fs   = require("fs");
const XLSX = require("xlsx");
const { processAll, excelDateToString, excelTimeToString } = require("./processor");

let win;

function createWindow() {
  win = new BrowserWindow({
    width: 1080,
    height: 720,
    minWidth: 860,
    minHeight: 620,
    title: "Sello Radicado — SERVENTEGRAL S.A. E.S.P.",
    icon: path.join(__dirname, "..", "assets", "camion.ico"),
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
    },
    backgroundColor: "#F1F8E9",
    autoHideMenuBar: true,
  });

  win.loadFile(path.join(__dirname, "renderer", "index.html"));
}

app.whenReady().then(() => {
  // Primera ejecución: copiar config.json desde resources a userData
  const userCfg = path.join(app.getPath("userData"), "config.json");
  if (!fs.existsSync(userCfg)) {
    const base   = app.isPackaged ? process.resourcesPath : path.join(__dirname, "..");
    const srcCfg = path.join(base, "config.json");
    if (fs.existsSync(srcCfg)) fs.copyFileSync(srcCfg, userCfg);
  }
  createWindow();
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});

// ── IPC: seleccionar archivo Excel ─────────────────────────────────────────
ipcMain.handle("select-excel", async () => {
  const result = await dialog.showOpenDialog(win, {
    title: "Seleccionar archivo de radicados",
    filters: [{ name: "Excel", extensions: ["xlsx", "xls"] }],
    properties: ["openFile"],
  });
  return result.canceled ? null : result.filePaths[0];
});

// ── IPC: seleccionar carpeta ────────────────────────────────────────────────
ipcMain.handle("select-folder", async (_, title) => {
  const result = await dialog.showOpenDialog(win, {
    title,
    properties: ["openDirectory"],
  });
  return result.canceled ? null : result.filePaths[0];
});

// ── IPC: seleccionar archivos .pdf ────────────────────────────────────────
ipcMain.handle("select-doc-files", async () => {
  const result = await dialog.showOpenDialog(win, {
    title: "Seleccionar documentos PDF a radicar",
    filters: [
      { name: "PDF (.pdf)", extensions: ["pdf"] },
    ],
    properties: ["openFile", "multiSelections"],
  });
  return result.canceled ? null : result.filePaths;
});

// ── IPC: seleccionar logo ──────────────────────────────────────────────────
ipcMain.handle("select-logo", async () => {
  const result = await dialog.showOpenDialog(win, {
    title: "Seleccionar logo (PNG)",
    filters: [{ name: "Imagen PNG", extensions: ["png"] }],
    properties: ["openFile"],
  });
  return result.canceled ? null : result.filePaths[0];
});

// ── IPC: abrir carpeta de salida ───────────────────────────────────────────
ipcMain.handle("open-folder", async (_, folderPath) => {
  if (folderPath && fs.existsSync(folderPath)) {
    shell.openPath(folderPath);
  }
});

// ── IPC: ruta resources (para logo incluido en .exe) ──────────────────────
ipcMain.handle("get-resource-path", () => {
  // en desarrollo: carpeta raíz del proyecto
  // en producción (empaquetado): process.resourcesPath
  return app.isPackaged
    ? process.resourcesPath
    : path.join(__dirname, "..");
});

// ── IPC: leer config.json (userData, con fallback a resources) ─────────────
ipcMain.handle("get-config", () => {
  const def     = { lado: "derecho", offsetVertical: 914400 };
  const cfgUser = path.join(app.getPath("userData"), "config.json");
  if (fs.existsSync(cfgUser)) {
    try { return { ...def, ...JSON.parse(fs.readFileSync(cfgUser, "utf8")) }; }
    catch { /* continúa con default */ }
  }
  // Primera ejecución: intentar cargar desde resources (extraResources)
  const base    = app.isPackaged ? process.resourcesPath : path.join(__dirname, "..");
  const cfgRes  = path.join(base, "config.json");
  if (fs.existsSync(cfgRes)) {
    try { return { ...def, ...JSON.parse(fs.readFileSync(cfgRes, "utf8")) }; }
    catch { /* continúa con default */ }
  }
  return def;
});

// ── IPC: guardar config.json en userData ───────────────────────────────────
ipcMain.handle("save-config", async (_, config) => {
  try {
    const cfgPath = path.join(app.getPath("userData"), "config.json");
    fs.writeFileSync(cfgPath, JSON.stringify(config, null, 2), "utf8");
    return { ok: true };
  } catch (err) {
    return { ok: false, error: err.message };
  }
});

// ── IPC: buscar radicado en Excel ──────────────────────────────────────────
ipcMain.handle("buscar-radicado", async (_, { numero, rutaExcel }) => {
  try {
    const wb   = XLSX.readFile(rutaExcel);
    const ws   = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

    const limpias = rows.map(row => {
      const r = {};
      Object.keys(row).forEach(k => {
        r[k.trim()] = typeof row[k] === "string" ? row[k].trim() : row[k];
      });
      return r;
    });

    const fila = limpias.find(r => String(r["RADICADO"]).trim() === String(numero).trim());
    if (!fila) return { encontrado: false };

    return {
      encontrado: true,
      datos: {
        radicado:    String(fila["RADICADO"]).padStart(4, "0"),
        fecha:       excelDateToString(fila["FECHA DE INGRESO DE LA SOLICITUD"]),
        hora:        excelTimeToString(fila["HORA"]),
        asunto:      String(fila["ASUNTO"]        || ""),
        anexos:      String(fila["ANEXOS"]        || ""),
        radicadoPor: String(fila["RADICADO POR"]  || ""),
      }
    };
  } catch (e) {
    return { encontrado: false, error: e.message };
  }
});

// ── IPC: procesar documentos ───────────────────────────────────────────────
ipcMain.handle("process-documents", async (event, opts) => {
  try {
    const result = await processAll(opts, (progress) => {
      event.sender.send("processing-progress", progress);
    });
    return { success: true, ...result };
  } catch (err) {
    return { success: false, error: err.message };
  }
});
