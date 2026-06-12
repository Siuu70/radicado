"use strict";

const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("api", {
  selectExcel:       () => ipcRenderer.invoke("select-excel"),
  selectFolder:      (title) => ipcRenderer.invoke("select-folder", title),
  selectDocFiles:    () => ipcRenderer.invoke("select-doc-files"),
  selectLogo:        () => ipcRenderer.invoke("select-logo"),
  openFolder:        (p) => ipcRenderer.invoke("open-folder", p),
  getResourcePath:   () => ipcRenderer.invoke("get-resource-path"),
  processDocuments:  (opts) => ipcRenderer.invoke("process-documents", opts),
  onProgress:        (cb) => ipcRenderer.on("processing-progress", (_, data) => cb(data)),
  removeProgressListener: () => ipcRenderer.removeAllListeners("processing-progress"),
  getConfig:         () => ipcRenderer.invoke("get-config"),
  saveConfig:        (config) => ipcRenderer.invoke("save-config", config),
});
