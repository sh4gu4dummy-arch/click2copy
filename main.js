const { app, BrowserWindow, ipcMain, dialog, Menu } = require('electron');
const path = require('path');
const fs = require('fs');

const STORE_FILE = 'click2copy-data.json';
const SESSION_FILE = 'click2copy-session.json';
const DOCUMENT_EXTENSION = 'c2copy';
const DOCUMENT_FORMAT = 'click2copy-document';
const DOCUMENT_VERSION = 1;
const AUTOSAVE_C2COPY = 'latest-autosave.c2copy';
const AUTOSAVE_JSON = 'latest-autosave.json';
const USERDATA_BACKUP_SUBDIR = 'backups';
const DOCUMENTS_BACKUP_FOLDER = 'Click2Copy';
const DEFAULT_COLS = 3;
const DEFAULT_ROWS = 8;
const MIN_COLUMN_WIDTH = 100;
const DEFAULT_SEPARATORS = {
  part: '\\n\\n',
  column: ' | ',
  row: '\\n'
};

function emptyCells(cols, rows) {
  return Array(cols * rows).fill('');
}

function emptySleptCells(cols, rows) {
  return Array(cols * rows).fill(false);
}

function normalizeSleptCells(slept, length) {
  const needed = length > 0 ? length : 0;
  let out;
  if (Array.isArray(slept)) {
    out = slept.map((v) => !!v);
  } else {
    out = Array(needed).fill(false);
  }
  if (out.length < needed) {
    out = out.concat(Array(needed - out.length).fill(false));
  } else if (out.length > needed) {
    out = out.slice(0, needed);
  }
  return out;
}

function makeDefaultTab(id, title) {
  return {
    id,
    title,
    cols: DEFAULT_COLS,
    rows: DEFAULT_ROWS,
    cells: emptyCells(DEFAULT_COLS, DEFAULT_ROWS),
    sleptCells: emptySleptCells(DEFAULT_COLS, DEFAULT_ROWS)
  };
}

const DEFAULT_DATA = {
  tabs: [
    makeDefaultTab('master', 'Master'),
    makeDefaultTab('tab-1', 'Part 1'),
    makeDefaultTab('tab-2', 'Part 2'),
    makeDefaultTab('tab-3', 'Part 3')
  ],
  activeTabId: 'tab-1',
  combinedPrompt: '',
  globalCombined: true,
  partPrompts: {},
  separators: { ...DEFAULT_SEPARATORS },
  confirmedLinks: []
};

function storePath() {
  return path.join(app.getPath('userData'), STORE_FILE);
}

function sessionPath() {
  return path.join(app.getPath('userData'), SESSION_FILE);
}

function userDataBackupDir() {
  return path.join(app.getPath('userData'), USERDATA_BACKUP_SUBDIR);
}

function documentsBackupDir() {
  return path.join(app.getPath('documents'), DOCUMENTS_BACKUP_FOLDER);
}

function ensureDir(dir) {
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

function listAutosaveDirs() {
  const dirs = [];
  try {
    dirs.push(ensureDir(userDataBackupDir()));
  } catch (err) {
    console.error('Failed to prepare userData autosave folder:', err);
  }
  try {
    dirs.push(ensureDir(documentsBackupDir()));
  } catch (err) {
    console.error('Failed to prepare Documents/Click2Copy autosave folder:', err);
  }
  return dirs;
}

function getAutosaveLocations() {
  return {
    userDataStore: storePath(),
    userDataSession: sessionPath(),
    userDataBackupDir: userDataBackupDir(),
    documentsBackupDir: documentsBackupDir(),
    latestC2copyName: AUTOSAVE_C2COPY,
    latestJsonName: AUTOSAVE_JSON
  };
}

function sanitizeFileBase(name) {
  const cleaned = String(name || 'untitled')
    .replace(/[<>:"/\\|?*\x00-\x1f]/g, '_')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/[. ]+$/g, '');
  return cleaned || 'untitled';
}

function documentPayload(normalized) {
  return {
    format: DOCUMENT_FORMAT,
    version: DOCUMENT_VERSION,
    data: normalized
  };
}

/**
 * Write readable recoverable copies for untitled docs (no named Save As path).
 * Named files are left alone here — caller autosaves those to their own path.
 */
function writeUntitledRecoverableBackups(documents, activeDocumentId) {
  const dirs = listAutosaveDirs();
  if (dirs.length === 0) {
    return { ok: false, error: 'No writable autosave backup folder.', paths: [] };
  }

  const paths = [];
  let activeUntitled = null;

  for (const document of documents) {
    if (!document || document.filePath) continue;
    const normalized = normalizeData(document.data);
    if (!normalized) continue;
    const payloadText = JSON.stringify(documentPayload(normalized), null, 2);
    const mirrorText = JSON.stringify(normalized, null, 2);
    const idBase = sanitizeFileBase(document.id);
    const perDocName = `autosave-${idBase}.${DOCUMENT_EXTENSION}`;
    const perDocJson = `autosave-${idBase}.json`;

    for (const dir of dirs) {
      const c2Path = path.join(dir, perDocName);
      const jsonPath = path.join(dir, perDocJson);
      fs.writeFileSync(c2Path, payloadText, 'utf8');
      fs.writeFileSync(jsonPath, mirrorText, 'utf8');
      paths.push(c2Path, jsonPath);
    }

    if (document.id === activeDocumentId || !activeUntitled) {
      activeUntitled = { payloadText, mirrorText };
    }
  }

  if (activeUntitled) {
    for (const dir of dirs) {
      const latestC2 = path.join(dir, AUTOSAVE_C2COPY);
      const latestJson = path.join(dir, AUTOSAVE_JSON);
      fs.writeFileSync(latestC2, activeUntitled.payloadText, 'utf8');
      fs.writeFileSync(latestJson, activeUntitled.mirrorText, 'utf8');
      paths.push(latestC2, latestJson);
    }
  }

  return { ok: true, paths, locations: getAutosaveLocations() };
}

function normalizeColumnWidths(widths, cols) {
  if (!Array.isArray(widths) || widths.length !== cols) return null;
  const normalized = widths.map((width) => {
    const value = Number(width);
    return Number.isFinite(value) && value >= MIN_COLUMN_WIDTH ? value : null;
  });
  return normalized.every((width) => width !== null) ? normalized : null;
}

function normalizeSeparators(separators) {
  const source = separators && typeof separators === 'object' ? separators : {};
  return {
    part: typeof source.part === 'string' ? source.part : DEFAULT_SEPARATORS.part,
    column: typeof source.column === 'string' ? source.column : DEFAULT_SEPARATORS.column,
    row: typeof source.row === 'string' ? source.row : DEFAULT_SEPARATORS.row
  };
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
    const columnWidths = normalizeColumnWidths(t.columnWidths, cols);
    const normalized = {
      id: t.id,
      title: t.title,
      cols,
      rows,
      cells,
      sleptCells: emptySleptCells(cols, rows)
    };
    if (columnWidths) normalized.columnWidths = columnWidths;
    return normalized;
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

  const columnWidths = normalizeColumnWidths(t.columnWidths, cols);
  const sleptCells = normalizeSleptCells(t.sleptCells, needed);
  const normalized = { id: t.id, title: t.title, cols, rows, cells, sleptCells };
  if (columnWidths) normalized.columnWidths = columnWidths;
  return normalized;
}

function toNonNegInt(value) {
  if (typeof value === 'number' && Number.isFinite(value)) {
    const n = Math.trunc(value);
    return n >= 0 ? n : null;
  }
  if (typeof value === 'string' && /^-?\d+$/.test(value.trim())) {
    const n = parseInt(value.trim(), 10);
    return Number.isFinite(n) && n >= 0 ? n : null;
  }
  return null;
}

/** Keep Combined↔cell green links across save/load/session/.c2copy. */
function normalizeConfirmedLinks(raw) {
  if (!Array.isArray(raw)) return [];
  const links = [];
  for (const item of raw) {
    if (!item || typeof item !== 'object') continue;
    if (typeof item.id !== 'string' || !item.id) continue;
    if (typeof item.tabId !== 'string' || !item.tabId) continue;
    if (typeof item.text !== 'string') continue;
    const cellIndex = toNonNegInt(item.cellIndex);
    const start = toNonNegInt(item.start);
    const end = toNonNegInt(item.end);
    if (cellIndex === null || start === null || end === null || end < start) continue;
    const scope = typeof item.scope === 'string' && item.scope ? item.scope : 'global';
    links.push({
      id: item.id,
      tabId: item.tabId,
      cellIndex,
      text: item.text,
      start,
      end,
      scope
    });
  }
  return links;
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
  const masterIndex = tabs.findIndex((tab) =>
    tab.id === 'master' || tab.title.trim().toLowerCase() === 'master'
  );
  if (masterIndex === -1) {
    tabs.unshift(makeDefaultTab('master', 'Master'));
  } else {
    const [master] = tabs.splice(masterIndex, 1);
    master.title = 'Master';
    tabs.unshift(master);
  }
  let activeTabId = typeof parsed.activeTabId === 'string' ? parsed.activeTabId : tabs[0].id;
  if (!tabs.some((t) => t.id === activeTabId)) {
    activeTabId = tabs[0].id;
  }
  const confirmedLinks = normalizeConfirmedLinks(parsed.confirmedLinks);
  return {
    tabs,
    activeTabId,
    combinedPrompt: typeof parsed.combinedPrompt === 'string' ? parsed.combinedPrompt : '',
    globalCombined: typeof parsed.globalCombined === 'boolean' ? parsed.globalCombined : true,
    partPrompts: parsed.partPrompts && typeof parsed.partPrompts === 'object' && !Array.isArray(parsed.partPrompts)
      ? Object.fromEntries(Object.entries(parsed.partPrompts)
        .filter(([, prompt]) => typeof prompt === 'string'))
      : {},
    separators: normalizeSeparators(parsed.separators),
    confirmedLinks
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

function loadSession() {
  try {
    const file = sessionPath();
    if (fs.existsSync(file)) {
      const parsed = JSON.parse(fs.readFileSync(file, 'utf8'));
      if (parsed && Array.isArray(parsed.documents)) {
        const documents = parsed.documents.map((document) => {
          const data = normalizeData(document.data);
          if (!data || typeof document.id !== 'string' || typeof document.title !== 'string') {
            return null;
          }
          return {
            id: document.id,
            title: document.title,
            filePath: typeof document.filePath === 'string' ? document.filePath : null,
            data
          };
        }).filter(Boolean);
        if (documents.length > 0) {
          const activeDocumentId = documents.some((document) => document.id === parsed.activeDocumentId)
            ? parsed.activeDocumentId
            : documents[0].id;
          return { documents, activeDocumentId };
        }
      }
    }
  } catch (err) {
    console.error('Failed to load session:', err);
  }

  return {
    documents: [{
      id: 'document-default',
      title: 'Untitled',
      filePath: null,
      data: loadData()
    }],
    activeDocumentId: 'document-default'
  };
}

function saveSession(session) {
  if (!session || !Array.isArray(session.documents)) {
    return { ok: false, error: 'Cannot save an invalid prompt session.' };
  }
  if (session.documents.length === 0) return { ok: true, skipped: true };

  const documents = [];
  for (const document of session.documents) {
    if (!document || typeof document.id !== 'string' || typeof document.title !== 'string') {
      return { ok: false, error: 'Cannot save an invalid prompt document.' };
    }
    const data = normalizeData(document && document.data);
    if (!data) {
      return { ok: false, error: 'Cannot save an invalid prompt document.' };
    }
    documents.push({
      id: document.id,
      title: document.title,
      filePath: typeof document.filePath === 'string' ? document.filePath : null,
      data
    });
  }

  const activeDocumentId = documents.some((document) => document.id === session.activeDocumentId)
    ? session.activeDocumentId
    : documents[0].id;
  try {
    const file = sessionPath();
    fs.mkdirSync(path.dirname(file), { recursive: true });
    fs.writeFileSync(file, JSON.stringify({ documents, activeDocumentId }, null, 2), 'utf8');
  } catch (err) {
    console.error('Failed to save session:', err);
    return { ok: false, error: err.message || 'Failed to save open prompt documents.' };
  }

  let backup = { ok: true, paths: [], locations: getAutosaveLocations() };
  try {
    backup = writeUntitledRecoverableBackups(documents, activeDocumentId);
  } catch (err) {
    console.error('Failed to write recoverable autosave backups:', err);
    backup = {
      ok: false,
      error: err.message || 'Failed to write recoverable autosave backups.',
      paths: [],
      locations: getAutosaveLocations()
    };
  }

  return {
    ok: true,
    backupOk: !!backup.ok,
    backupError: backup.ok ? undefined : backup.error,
    backupPaths: backup.paths || [],
    locations: backup.locations || getAutosaveLocations()
  };
}

function normalizeDocument(parsed) {
  if (parsed && parsed.format === DOCUMENT_FORMAT && parsed.version === DOCUMENT_VERSION) {
    return normalizeData(parsed.data);
  }
  return normalizeData(parsed);
}

function readDocument(filePath) {
  try {
    const parsed = JSON.parse(fs.readFileSync(filePath, 'utf8'));
    const data = normalizeDocument(parsed);
    if (!data) {
      return { ok: false, error: 'This file is not a valid Click2Copy prompt document.' };
    }
    return { ok: true, filePath, data };
  } catch (err) {
    console.error('Failed to open prompt document:', err);
    return { ok: false, error: err instanceof SyntaxError ? 'File is not valid JSON.' : (err.message || 'Failed to read file.') };
  }
}

function writeDocument(filePath, data) {
  const normalized = normalizeData(data);
  if (!normalized) return { ok: false, error: 'Cannot save an invalid prompt document.' };

  const targetPath = path.extname(filePath).toLowerCase() === `.${DOCUMENT_EXTENSION}`
    ? filePath
    : `${filePath}.${DOCUMENT_EXTENSION}`;
  try {
    fs.mkdirSync(path.dirname(targetPath), { recursive: true });
    fs.writeFileSync(targetPath, JSON.stringify(documentPayload(normalized), null, 2), 'utf8');
    return { ok: true, filePath: targetPath };
  } catch (err) {
    console.error('Failed to save prompt document:', err);
    return { ok: false, error: err.message || 'Failed to save file.' };
  }
}

function saveData(data) {
  const normalized = normalizeData(data);
  if (!normalized) return false;
  try {
    const file = storePath();
    fs.mkdirSync(path.dirname(file), { recursive: true });
    fs.writeFileSync(file, JSON.stringify(normalized, null, 2), 'utf8');
  } catch (err) {
    console.error('Failed to save store:', err);
    return false;
  }

  // Also keep a latest readable .c2copy (+ plain JSON mirror) for recovery outside the app.
  try {
    const dirs = listAutosaveDirs();
    const payloadText = JSON.stringify(documentPayload(normalized), null, 2);
    const mirrorText = JSON.stringify(normalized, null, 2);
    for (const dir of dirs) {
      fs.writeFileSync(path.join(dir, AUTOSAVE_C2COPY), payloadText, 'utf8');
      fs.writeFileSync(path.join(dir, AUTOSAVE_JSON), mirrorText, 'utf8');
    }
  } catch (err) {
    console.error('Failed to mirror latest autosave .c2copy:', err);
  }
  return true;
}

let mainWindow = null;
let mainWindowLoaded = false;
const pendingOpenPaths = [];

function documentPathFromArgs(args) {
  return args.find((arg) => typeof arg === 'string' &&
    path.extname(arg).toLowerCase() === `.${DOCUMENT_EXTENSION}`) || null;
}

function deliverOpenPath(filePath) {
  if (!filePath) return;
  if (!mainWindow || !mainWindowLoaded) {
    pendingOpenPaths.push(filePath);
    return;
  }
  if (mainWindow.isMinimized()) mainWindow.restore();
  mainWindow.focus();
  mainWindow.webContents.send('document:open-path', filePath);
}

const hasSingleInstanceLock = app.requestSingleInstanceLock();
if (!hasSingleInstanceLock) {
  app.quit();
} else {
  app.on('second-instance', (_event, commandLine) => {
    deliverOpenPath(documentPathFromArgs(commandLine));
  });
  app.on('open-file', (event, filePath) => {
    event.preventDefault();
    deliverOpenPath(filePath);
  });
}


/** App-root source files that mean "code updated" (e.g. after git pull). Not userData/Documents autosaves. */
const WATCHED_SOURCE_FILES = [
  'package.json',
  'main.js',
  'preload.js',
  'renderer.js',
  'index.html',
  'styles.css'
];
const SOURCE_POLL_MS = 1500;
const SOURCE_DEBOUNCE_MS = 800;

let startupPackageVersion = null;
let sourceWatchBaseline = new Map();
let sourceUpdateDebounce = null;
const sourceWatchers = [];

function readPackageVersionFromDisk() {
  try {
    const raw = fs.readFileSync(path.join(__dirname, 'package.json'), 'utf8');
    const pkg = JSON.parse(raw);
    return typeof pkg.version === 'string' ? pkg.version : null;
  } catch (err) {
    return null;
  }
}

function snapshotSourceMtimes() {
  const map = new Map();
  for (const name of WATCHED_SOURCE_FILES) {
    const full = path.join(__dirname, name);
    try {
      map.set(name, fs.statSync(full).mtimeMs);
    } catch (err) {
      map.set(name, null);
    }
  }
  return map;
}

function sourceFilesChanged(baseline, current) {
  for (const name of WATCHED_SOURCE_FILES) {
    if (baseline.get(name) !== current.get(name)) return true;
  }
  return false;
}

function buildUpdatePayload() {
  const diskVersion = readPackageVersionFromDisk();
  const runningVersion = startupPackageVersion || app.getVersion();
  return {
    runningVersion,
    diskVersion,
    changed: true
  };
}

function notifySourceUpdateReady() {
  const payload = buildUpdatePayload();
  if (mainWindow && mainWindowLoaded && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send('app:update-ready', payload);
  }
}

function scheduleSourceUpdateNotify() {
  if (sourceUpdateDebounce) clearTimeout(sourceUpdateDebounce);
  sourceUpdateDebounce = setTimeout(() => {
    sourceUpdateDebounce = null;
    const current = snapshotSourceMtimes();
    if (!sourceFilesChanged(sourceWatchBaseline, current)) return;
    sourceWatchBaseline = current;
    notifySourceUpdateReady();
  }, SOURCE_DEBOUNCE_MS);
}

function startSourceWatcher() {
  startupPackageVersion = readPackageVersionFromDisk() || app.getVersion();
  sourceWatchBaseline = snapshotSourceMtimes();

  // Polling is reliable when editors/git replace files (fs.watch can miss renames).
  const pollId = setInterval(() => {
    const current = snapshotSourceMtimes();
    if (sourceFilesChanged(sourceWatchBaseline, current)) {
      scheduleSourceUpdateNotify();
    }
  }, SOURCE_POLL_MS);
  if (typeof pollId.unref === 'function') pollId.unref();

  for (const name of WATCHED_SOURCE_FILES) {
    const full = path.join(__dirname, name);
    try {
      const watcher = fs.watch(full, { persistent: false }, () => {
        scheduleSourceUpdateNotify();
      });
      sourceWatchers.push(watcher);
    } catch (err) {
      // Missing file or unsupported watch; polling still covers updates.
    }
  }
}

function getParentWindow() {
  return mainWindow || BrowserWindow.getFocusedWindow();
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1440,
    height: 720,
    minWidth: 960,
    minHeight: 520,
    title: `Click2Copy v${app.getVersion()}`,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  mainWindow.loadFile('index.html');
  mainWindow.webContents.once('did-finish-load', () => {
    mainWindowLoaded = true;
    while (pendingOpenPaths.length > 0) {
      mainWindow.webContents.send('document:open-path', pendingOpenPaths.shift());
    }
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
    mainWindowLoaded = false;
  });
}

function createApplicationMenu() {
  const menu = Menu.buildFromTemplate([
    {
      label: 'File',
      submenu: [
        { label: 'New Prompt', accelerator: 'CmdOrCtrl+N', click: () => sendMenuAction('new-document') },
        { label: 'Open Prompt Files…', accelerator: 'CmdOrCtrl+O', click: () => sendMenuAction('open-documents') },
        { type: 'separator' },
        { label: 'Save', accelerator: 'CmdOrCtrl+S', click: () => sendMenuAction('save-document') },
        { label: 'Save As…', accelerator: 'CmdOrCtrl+Shift+S', click: () => sendMenuAction('save-document-as') },
        { type: 'separator' },
        { role: 'quit' }
      ]
    },
    { role: 'editMenu' },
    { role: 'viewMenu' },
    { role: 'windowMenu' }
  ]);
  Menu.setApplicationMenu(menu);
}

function sendMenuAction(action) {
  if (mainWindow && mainWindowLoaded) mainWindow.webContents.send('app:menu-action', action);
}

app.whenReady().then(() => {
  if (!hasSingleInstanceLock) return;
  createApplicationMenu();
  createWindow();
  startSourceWatcher();
  const initialPath = documentPathFromArgs(process.argv);
  if (initialPath) pendingOpenPaths.push(initialPath);

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
ipcMain.handle('app:version', () => app.getVersion());
ipcMain.handle('app:relaunch', () => {
  app.relaunch();
  app.exit(0);
});
ipcMain.handle('app:update-info', () => ({
  runningVersion: startupPackageVersion || app.getVersion(),
  diskVersion: readPackageVersionFromDisk()
}));
ipcMain.on('window:set-title', (event, title) => {
  const window = BrowserWindow.fromWebContents(event.sender);
  if (window && typeof title === 'string') window.setTitle(`${title} v${app.getVersion()}`);
});
ipcMain.handle('documents:load-session', () => loadSession());
ipcMain.handle('documents:save-session', (_event, session) => saveSession(session));
ipcMain.handle('documents:autosave-locations', () => getAutosaveLocations());
ipcMain.handle('documents:open', async (event) => {
  const result = await dialog.showOpenDialog(BrowserWindow.fromWebContents(event.sender), {
    title: 'Open Click2Copy prompt',
    properties: ['openFile', 'multiSelections'],
    filters: [{ name: 'Click2Copy prompts', extensions: [DOCUMENT_EXTENSION] }]
  });
  if (result.canceled) return { ok: false, canceled: true };
  return { ok: true, documents: result.filePaths.map(readDocument) };
});
ipcMain.handle('documents:open-path', (_event, filePath) => readDocument(filePath));
ipcMain.handle('documents:save', (_event, filePath, data) => writeDocument(filePath, data));
ipcMain.handle('documents:save-as', async (event, data, suggestedName) => {
  const result = await dialog.showSaveDialog(BrowserWindow.fromWebContents(event.sender), {
    title: 'Save Click2Copy prompt',
    defaultPath: suggestedName || `Untitled.${DOCUMENT_EXTENSION}`,
    filters: [{ name: 'Click2Copy prompts', extensions: [DOCUMENT_EXTENSION] }]
  });
  if (result.canceled || !result.filePath) return { ok: false, canceled: true };
  return writeDocument(result.filePath, data);
});

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
