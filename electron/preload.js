/**
 * MCS CoWorker — Electron Preload Script
 * Safely exposes IPC channels to the React renderer via contextBridge.
 */

const { contextBridge, ipcRenderer } = require("electron");

// Expose window controls to React
contextBridge.exposeInMainWorld("electronAPI", {
  minimize: () => ipcRenderer.send("window-minimize"),
  maximize: () => ipcRenderer.send("window-maximize"),
  close: () => ipcRenderer.send("window-close"),
  quit: () => ipcRenderer.send("window-quit"),
  onNavigate: (callback) => ipcRenderer.on("navigate", (_, page) => callback(page)),
});

// Signal to the React app that it is running inside Electron
contextBridge.exposeInMainWorld("__ELECTRON__", true);
