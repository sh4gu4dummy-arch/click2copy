const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('click2copy', {
  loadData: () => ipcRenderer.invoke('store:load'),
  getAppVersion: () => ipcRenderer.invoke('app:version'),
  saveData: (data) => ipcRenderer.invoke('store:save', data),
  exportData: (data) => ipcRenderer.invoke('store:export', data),
  importData: () => ipcRenderer.invoke('store:import'),
  loadSession: () => ipcRenderer.invoke('documents:load-session'),
  saveSession: (session) => ipcRenderer.invoke('documents:save-session', session),
  openDocuments: () => ipcRenderer.invoke('documents:open'),
  openDocument: (filePath) => ipcRenderer.invoke('documents:open-path', filePath),
  saveDocument: (filePath, data) => ipcRenderer.invoke('documents:save', filePath, data),
  saveDocumentAs: (data, suggestedName) => ipcRenderer.invoke('documents:save-as', data, suggestedName),
  loadDiagLog: () => ipcRenderer.invoke('diag:log-load'),
  appendDiagLog: (entry) => ipcRenderer.invoke('diag:log-append', entry),
  classifyDiag: (input) => ipcRenderer.invoke('diag:classify', input),
  checkOrigin: () => ipcRenderer.invoke('diag:origin-check'),
  onOpenDocument: (callback) => {
    const listener = (_event, filePath) => callback(filePath);
    ipcRenderer.on('document:open-path', listener);
    return () => ipcRenderer.removeListener('document:open-path', listener);
  },
  onMenuAction: (callback) => {
    const listener = (_event, action) => callback(action);
    ipcRenderer.on('app:menu-action', listener);
    return () => ipcRenderer.removeListener('app:menu-action', listener);
  },
  setWindowTitle: (title) => ipcRenderer.send('window:set-title', title)
});
