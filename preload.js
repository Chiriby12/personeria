const { contextBridge, ipcRenderer } = require('electron');

// Exponer APIs seguras al renderer
contextBridge.exposeInMainWorld('electronAPI', {
  guardarSolicitud: (data) => ipcRenderer.invoke('guardar-solicitud', data),
  abrirExcel: () => ipcRenderer.invoke('abrir-excel'),
  obtenerRutaExcel: () => ipcRenderer.invoke('obtener-ruta-excel')
});