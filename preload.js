const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('click2copy', {
  loadData: () => ipcRenderer.invoke('store:load'),
  getAppVersion: () => ipcRenderer.invoke('app:version'),
  saveData: (data) => ipcRenderer.invoke('store:save', data),
  exportData: (data) => ipcRenderer.invoke('store:export', data),
  importData: () => ipcRenderer.invoke('store:import'),
  loadSession: () => ipcRenderer.invoke('documents:load-session'),
  saveSession: (session) => ipcRenderer.invoke('documents:save-session', session),
  loadGridScroll: () => ipcRenderer.invoke('grid-scroll:load'),
  saveGridScroll: (map) => ipcRenderer.invoke('grid-scroll:save', map),
  saveGridScrollSync: (map) => ipcRenderer.sendSync('grid-scroll:save-sync', map),
  getAutosaveLocations: () => ipcRenderer.invoke('documents:autosave-locations'),
  openDocuments: () => ipcRenderer.invoke('documents:open'),
  openDocument: (filePath) => ipcRenderer.invoke('documents:open-path', filePath),
  saveDocument: (filePath, data) => ipcRenderer.invoke('documents:save', filePath, data),
  saveDocumentAs: (data, suggestedName) => ipcRenderer.invoke('documents:save-as', data, suggestedName),
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
  setWindowTitle: (title) => ipcRenderer.send('window:set-title', title),
  onUpdateReady: (callback) => {
    const listener = (_event, payload) => callback(payload);
    ipcRenderer.on('app:update-ready', listener);
    return () => ipcRenderer.removeListener('app:update-ready', listener);
  },
  relaunchApp: () => ipcRenderer.invoke('app:relaunch'),
  backupAndRelaunch: (extra) => ipcRenderer.invoke('app:backup-and-relaunch', extra),
  getBackupSettings: () => ipcRenderer.invoke('backup:get-settings'),
  setBackupSettings: (next) => ipcRenderer.invoke('backup:set-settings', next),
  chooseBackupFolder: () => ipcRenderer.invoke('backup:choose-folder'),
  getUpdateInfo: () => ipcRenderer.invoke('app:update-info')
});
