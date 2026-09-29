const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('click2copy', {
  loadData: () => ipcRenderer.invoke('store:load'),
  saveData: (data) => ipcRenderer.invoke('store:save', data),
  exportData: (data) => ipcRenderer.invoke('store:export', data),
  importData: () => ipcRenderer.invoke('store:import')
});
