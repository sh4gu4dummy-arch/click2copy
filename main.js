const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');

const STORE_FILE = 'click2copy-data.json';
const DEFAULT_COLS = 3;
const DEFAULT_ROWS = 8;

function emptyCells(cols, rows) {
  return Array(cols * rows).fill('');
}

function makeDefaultTab(id, title) {
  return {
    id,
    title,
    cols: DEFAULT_COLS,
    rows: DEFAULT_ROWS,
    cells: emptyCells(DEFAULT_COLS, DEFAULT_ROWS)
  };
}

const DEFAULT_DATA = {
  tabs: [
    makeDefaultTab('tab-1', 'Part 1'),
    makeDefaultTab('tab-2', 'Part 2'),
    makeDefaultTab('tab-3', 'Part 3')
  ],
  activeTabId: 'tab-1',
  combinedPrompt: ''
};

function storePath() {
  return path.join(app.getPath('userData'), STORE_FILE);
}

/**
 * Normalize a single tab to the grid schema.
 * Migrates legacy { content: string } into cells[0] at 3×8.
 * Returns null if the tab is invalid.
 */
function normalizeTab(t) {
  if (!t || typeof t !== 'object') return null;
  if (typeof t.id !== 'string' || !t.id) return null;
  if (typeof t.title !== 'string') return null;

  // Legacy: single content string → put in first cell of 3×8 grid
  if (typeof t.content === 'string' && !Array.isArray(t.cells)) {
    const cols = DEFAULT_COLS;
    const rows = DEFAULT_ROWS;
    const cells = emptyCells(cols, rows);
    cells[0] = t.content;
    return { id: t.id, title: t.title, cols, rows, cells };
  }

  let cols = Number.isInteger(t.cols) && t.cols > 0 ? t.cols : DEFAULT_COLS;
  let rows = Number.isInteger(t.rows) && t.rows > 0 ? t.rows : DEFAULT_ROWS;

  let cells;
  if (Array.isArray(t.cells)) {
    cells = t.cells.map((c) => (typeof c === 'string' ? c : ''));
  } else if (Array.isArray(t.grid) && t.grid.every(Array.isArray)) {
    // Alternate 2D shape: grid[row][col]
    rows = t.grid.length || DEFAULT_ROWS;
    cols = t.grid[0] ? t.grid[0].length : DEFAULT_COLS;
    cells = [];
    for (let r = 0; r < rows; r++) {
      const row = t.grid[r] || [];
      for (let c = 0; c < cols; c++) {
        cells.push(typeof row[c] === 'string' ? row[c] : '');
      }
    }
  } else {
    return null;
  }

  const needed = cols * rows;
  if (cells.length < needed) {
    cells = cells.concat(emptyCells(needed - cells.length, 1));
  } else if (cells.length > needed) {
    cells = cells.slice(0, needed);
  }

  return { id: t.id, title: t.title, cols, rows, cells };
}

function normalizeData(parsed) {
  if (!parsed || !Array.isArray(parsed.tabs) || parsed.tabs.length === 0) {
    return null;
  }
  const tabs = [];
  for (const t of parsed.tabs) {
    const normalized = normalizeTab(t);
    if (!normalized) return null;
    tabs.push(normalized);
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
        error:
          'Invalid backup shape. Expected tabs (with id/title and cells grid or legacy content), activeTabId, and combinedPrompt.'
      };
    }
    return { ok: true, data: normalized, filePath };
  } catch (err) {
    console.error('Failed to import:', err);
    return { ok: false, error: err.message || 'Failed to read file.' };
  }
});
