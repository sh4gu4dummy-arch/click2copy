const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');

const STORE_FILE = 'click2copy-data.json';

const DEFAULT_DATA = {
  tabs: [
    { id: 'tab-1', title: 'Part 1', content: '' },
    { id: 'tab-2', title: 'Part 2', content: '' },
    { id: 'tab-3', title: 'Part 3', content: '' }
  ],
  activeTabId: 'tab-1',
  combinedPrompt: ''
};

function storePath() {
  return path.join(app.getPath('userData'), STORE_FILE);
}

function normalizeData(parsed) {
  if (!parsed || !Array.isArray(parsed.tabs) || parsed.tabs.length === 0) {
    return null;
  }
  const tabs = [];
  for (const t of parsed.tabs) {
    if (!t || typeof t !== 'object') return null;
    if (typeof t.id !== 'string' || !t.id) return null;
    if (typeof t.title !== 'string') return null;
    if (typeof t.content !== 'string') return null;
    tabs.push({ id: t.id, title: t.title, content: t.content });
  }
  let activeTabId = typeof parsed.activeTabId === 'string' ? parsed.activeTabId : tabs[0].id;
  if (!tabs.some((t) => t.id === activeTabId)) {
    activeTabId = tabs[0].id;
  }
  return {
    tabs,
    activeTabId,
    combinedPrompt: typeof parsed.combinedPrompt === 'string' ? parsed.combinedPrompt : ''
  };
}

function loadData() {
  try {
    const file = storePath();
    if (fs.existsSync(file)) {
      const raw = fs.readFileSync(file, 'utf8');
      const parsed = JSON.parse(raw);
      const normalized = normalizeData(parsed);
      if (normalized) return normalized;
    }
  } catch (err) {
    console.error('Failed to load store:', err);
  }
  return structuredClone(DEFAULT_DATA);
}

function saveData(data) {
  try {
    const file = storePath();
    fs.mkdirSync(path.dirname(file), { recursive: true });
    fs.writeFileSync(file, JSON.stringify(data, null, 2), 'utf8');
    return true;
  } catch (err) {
    console.error('Failed to save store:', err);
    return false;
  }
}

let mainWindow = null;

function getParentWindow() {
  return mainWindow || BrowserWindow.getFocusedWindow();
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1000,
    height: 720,
    minWidth: 720,
    minHeight: 520,
    title: 'Click2Copy',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  mainWindow.loadFile('index.html');

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

ipcMain.handle('store:load', () => loadData());

ipcMain.handle('store:save', (_event, data) => saveData(data));

ipcMain.handle('store:export', async (_event, data) => {
  const normalized = normalizeData(data);
  if (!normalized) {
    return { ok: false, error: 'Nothing to export (invalid state).' };
  }

  const result = await dialog.showSaveDialog(getParentWindow(), {
    title: 'Export Click2Copy backup',
    defaultPath: 'click2copy-backup.json',
    filters: [
      { name: 'JSON', extensions: ['json'] },
      { name: 'All Files', extensions: ['*'] }
    ]
  });

  if (result.canceled || !result.filePath) {
    return { ok: false, canceled: true };
  }

  try {
    fs.writeFileSync(result.filePath, JSON.stringify(normalized, null, 2), 'utf8');
    return { ok: true, filePath: result.filePath };
  } catch (err) {
    console.error('Failed to export:', err);
    return { ok: false, error: err.message || 'Failed to write file.' };
  }
});

ipcMain.handle('store:import', async () => {
  const result = await dialog.showOpenDialog(getParentWindow(), {
    title: 'Import Click2Copy backup',
    properties: ['openFile'],
    filters: [
      { name: 'JSON', extensions: ['json'] },
      { name: 'All Files', extensions: ['*'] }
    ]
  });

  if (result.canceled || !result.filePaths || result.filePaths.length === 0) {
    return { ok: false, canceled: true };
  }

  const filePath = result.filePaths[0];
  try {
    const raw = fs.readFileSync(filePath, 'utf8');
    let parsed;
    try {
      parsed = JSON.parse(raw);
    } catch (parseErr) {
      return { ok: false, error: 'File is not valid JSON.' };
    }
    const normalized = normalizeData(parsed);
    if (!normalized) {
      return {
        ok: false,
        error: 'Invalid backup shape. Expected tabs (with id/title/content), activeTabId, and combinedPrompt.'
      };
    }
    return { ok: true, data: normalized, filePath };
  } catch (err) {
    console.error('Failed to import:', err);
    return { ok: false, error: err.message || 'Failed to read file.' };
  }
});
