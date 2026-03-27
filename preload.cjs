const { contextBridge, ipcRenderer } = require('electron');

// Optionales Metadaten-API fuer spaetere Erweiterungen im Renderer.
contextBridge.exposeInMainWorld('calcalDesktop', {
  isDesktopApp: true,
  platform: process.platform,
  saveMonthlyOverview: (htmlContent) => ipcRenderer.invoke('calcal:save-monthly-overview', htmlContent),
});
