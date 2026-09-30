const { app, BrowserWindow, ipcMain, dialog, Menu } = require('electron');
const path = require('path');
const fs = require('fs');
const https = require('https');
const dns = require('dns');
const { spawnSync } = require('child_process');

const STORE_FILE = 'click2copy-data.json';
const SESSION_FILE = 'click2copy-session.json';
const DIAG_LOG_FILE = 'click2copy-diag-log.json';
const DIAG_LOG_LIMIT = 100;
const ORIGIN_PROBE_HOST = 'github.com';
const ORIGIN_PROBE_PATH = '/';
const ORIGIN_SLOW_MS = 3000;
const ORIGIN_TIMEOUT_MS = 12000;
const DOCUMENT_EXTENSION = 'c2copy';
const DOCUMENT_FORMAT = 'click2copy-document';
const DOCUMENT_VERSION = 1;
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
    makeDefaultTab('master', 'Master'),
    makeDefaultTab('tab-1', 'Part 1'),
    makeDefaultTab('tab-2', 'Part 2'),
    makeDefaultTab('tab-3', 'Part 3')
  ],
  activeTabId: 'tab-1',
  combinedPrompt: '',
  separators: { ...DEFAULT_SEPARATORS }
};

function storePath() {
  return path.join(app.getPath('userData'), STORE_FILE);
}

function sessionPath() {
  return path.join(app.getPath('userData'), SESSION_FILE);
}

function diagLogPath() {
  return path.join(app.getPath('userData'), DIAG_LOG_FILE);
}

/** Strip userinfo / token-looking query params from URLs before logging. */
function redactUrl(raw) {
  if (raw == null) return '';
  const text = String(raw).trim();
  if (!text) return '';
  try {
    const withScheme = /^[a-zA-Z][a-zA-Z0-9+.-]*:/.test(text) ? text : `https://${text}`;
    const u = new URL(withScheme);
    if (u.username || u.password) {
      u.username = u.username ? '***' : '';
      u.password = u.password ? '***' : '';
    }
    for (const key of [...u.searchParams.keys()]) {
      if (/token|auth|password|secret|key|pat|credential/i.test(key)) {
        u.searchParams.set(key, '***');
      }
    }
    let out = u.toString();
    if (!/^[a-zA-Z][a-zA-Z0-9+.-]*:/.test(text) && out.startsWith('https://')) {
      out = out.slice('https://'.length);
    }
    return out.replace(/:[^/@]+@/g, ':***@');
  } catch {
    return text
      .replace(/:\/\/[^/@:]+(:[^/@]*)?@/g, '://***@')
      .replace(/([?&](?:token|auth|password|secret|key|pat|access_token|authorization)=)[^&\s]+/gi, '$1***');
  }
}

function hostFromUrl(raw) {
  const text = String(raw || '').trim();
  if (!text) return '';
  try {
    const withScheme = /^[a-zA-Z][a-zA-Z0-9+.-]*:/.test(text) ? text : `https://${text}`;
    return new URL(withScheme).hostname || '';
  } catch {
    const m = text.match(/^(?:[a-zA-Z][a-zA-Z0-9+.-]*:\/\/)?(?:[^@/\s]+@)?([^/:?\s]+)/);
    return m ? m[1] : '';
  }
}

function sanitizeSnippet(text, maxLen) {
  const limit = maxLen || 240;
  let s = String(text == null ? '' : text)
    .replace(/\r\n/g, '\n')
    .replace(/:\/\/[^/@\s]+@/g, '://***@')
    .replace(/([?&](?:token|auth|password|secret|key|pat|access_token|authorization)=)[^&\s]+/gi, '$1***')
    .replace(/\bgh[pousr]_[A-Za-z0-9_]{20,}\b/g, 'gh*_***')
    .replace(/\bghp_[A-Za-z0-9]{20,}\b/g, 'ghp_***')
    .replace(/\bgithub_pat_[A-Za-z0-9_]{20,}\b/g, 'github_pat_***')
    .replace(/\bBearer\s+[A-Za-z0-9._\-+=\/]+/gi, 'Bearer ***')
    .trim();
  if (s.length > limit) s = s.slice(0, limit - 1) + '…';
  return s;
}

/**
 * Classify failures so Ash can tell slow Wi-Fi from app/config bugs.
 * Returns { class, label } with plain-language label for the UI.
 */
function classifyDiagFailure(input) {
  const durationMs = Number(input && input.durationMs) || 0;
  const code = String((input && (input.code || input.exitCode)) || '');
  const httpStatus = Number(input && input.httpStatus) || 0;
  const message = String((input && (input.message || input.stderr || input.error)) || '');
  const blob = `${code} ${message}`.toLowerCase();

  if (httpStatus === 401 || httpStatus === 403 ||
      /\b401\b|\b403\b|authentication failed|permission denied|access denied|could not read username|invalid.*credentials|auth failed/.test(blob)) {
    return { class: 'likely_remote_auth_or_permission', label: 'Likely login / permission issue' };
  }
  if (/enotfound|eai_again|enetunreach|ehostunreach|esrch|no route to host|name or service not known|temporary failure in name resolution|could not resolve host|nodename nor servname/.test(blob)) {
    return { class: 'likely_dns_or_unreachable', label: 'Likely DNS or unreachable host' };
  }
  if (/etimedout|esockettimedout|timeout|timed out|aborterror|econnreset|econnaborted|socket hang up|network.?timeout/.test(blob) ||
      (durationMs >= ORIGIN_SLOW_MS && /econnreset|econnrefused|fetch failed|network|hang/.test(blob))) {
    return { class: 'likely_slow_network', label: 'Likely slow Wi‑Fi / network' };
  }
  if (durationMs >= ORIGIN_SLOW_MS && !httpStatus) {
    return { class: 'likely_slow_network', label: 'Likely slow Wi‑Fi / network' };
  }
  if (/econnrefused/.test(blob)) {
    return { class: 'likely_dns_or_unreachable', label: 'Likely DNS or unreachable host' };
  }
  if (/enoent|eacces|eperm|eisdir|enotdir|invalid|parse|syntax|unexpected|cannot find|no such file|typeerror|referenceerror|local git error|not a git repository/.test(blob) ||
      (input && input.local === true)) {
    return { class: 'likely_app_or_config_bug', label: 'Likely app or config bug' };
  }
  if (httpStatus >= 500) {
    return { class: 'likely_slow_network', label: 'Likely slow Wi‑Fi / network' };
  }
  if (httpStatus >= 400) {
    return { class: 'likely_remote_auth_or_permission', label: 'Likely login / permission issue' };
  }
  return { class: 'likely_app_or_config_bug', label: 'Likely app or config bug' };
}

function loadDiagLog() {
  try {
    const file = diagLogPath();
    if (!fs.existsSync(file)) return [];
    const parsed = JSON.parse(fs.readFileSync(file, 'utf8'));
    if (!Array.isArray(parsed)) return [];
    return parsed.slice(-DIAG_LOG_LIMIT);
  } catch (err) {
    console.error('Failed to load diag log:', err);
    return [];
  }
}

function appendDiagLog(entry) {
  try {
    const entries = loadDiagLog();
    const safe = {
      at: entry && entry.at ? String(entry.at) : new Date().toISOString(),
      operation: entry && entry.operation ? String(entry.operation) : 'unknown',
      success: !!(entry && entry.success),
      durationMs: Number.isFinite(Number(entry && entry.durationMs)) ? Math.round(Number(entry.durationMs)) : null,
      host: entry && entry.host ? String(entry.host) : '',
      remoteUrl: entry && entry.remoteUrl ? redactUrl(entry.remoteUrl) : '',
      exitCode: entry && entry.exitCode != null ? entry.exitCode : null,
      httpStatus: entry && entry.httpStatus != null ? entry.httpStatus : null,
      errorClass: entry && entry.errorClass ? String(entry.errorClass) : '',
      errorLabel: entry && entry.errorLabel ? String(entry.errorLabel) : '',
      message: sanitizeSnippet(entry && entry.message, 200),
      detail: sanitizeSnippet(entry && entry.detail, 320),
      online: entry && typeof entry.online === 'boolean' ? entry.online : null
    };
    entries.push(safe);
    while (entries.length > DIAG_LOG_LIMIT) entries.shift();
    const file = diagLogPath();
    fs.mkdirSync(path.dirname(file), { recursive: true });
    fs.writeFileSync(file, JSON.stringify(entries, null, 2), 'utf8');
    return { ok: true, entry: safe, size: entries.length };
  } catch (err) {
    console.error('Failed to append diag log:', err);
    return { ok: false, error: err.message || 'Failed to write diag log' };
  }
}

function runGit(args, cwd, timeoutMs) {
  const started = Date.now();
  const result = spawnSync('git', args, {
    cwd: cwd || process.cwd(),
    encoding: 'utf8',
    timeout: timeoutMs || 8000,
    windowsHide: true
  });
  const durationMs = Date.now() - started;
  const stdout = (result.stdout || '').trim();
  const stderr = (result.stderr || '').trim();
  const error = result.error;
  return {
    durationMs,
    exitCode: error ? (error.code === 'ETIMEDOUT' ? 'ETIMEDOUT' : 1) : result.status,
    stdout,
    stderr: sanitizeSnippet(stderr || (error && error.message) || '', 320),
    error
  };
}

function resolveDns(hostname) {
  return new Promise((resolve) => {
    const started = Date.now();
    dns.lookup(hostname, { all: false }, (err, address, family) => {
      resolve({
        operation: 'dns_lookup',
        success: !err,
        durationMs: Date.now() - started,
        host: hostname,
        address: address || '',
        family: family || null,
        code: err && err.code ? err.code : null,
        message: err ? (err.message || String(err.code || err)) : 'resolved'
      });
    });
  });
}

function httpProbe(hostname, probePath, timeoutMs) {
  return new Promise((resolve) => {
    const started = Date.now();
    let settled = false;
    const finish = (payload) => {
      if (settled) return;
      settled = true;
      resolve(Object.assign({
        operation: 'http_probe',
        host: hostname,
        durationMs: Date.now() - started
      }, payload));
    };
    const req = https.request({
      hostname,
      path: probePath || '/',
      method: 'HEAD',
      timeout: timeoutMs || ORIGIN_TIMEOUT_MS,
      headers: { 'User-Agent': 'Click2Copy-origin-check', Accept: '*/*' }
    }, (res) => {
      res.resume();
      finish({
        success: !!(res.statusCode && res.statusCode < 500),
        httpStatus: res.statusCode || null,
        message: `HTTP ${res.statusCode}`
      });
    });
    req.on('timeout', () => {
      req.destroy(Object.assign(new Error('HTTP probe timed out'), { code: 'ETIMEDOUT' }));
    });
    req.on('error', (err) => {
      finish({
        success: false,
        httpStatus: null,
        code: err && err.code ? err.code : null,
        message: err && err.message ? err.message : 'HTTP probe failed'
      });
    });
    req.end();
  });
}

async function checkOriginReachability() {
  const started = Date.now();
  const steps = [];
  const repoRoot = app.isPackaged ? process.cwd() : path.join(__dirname);
  let remoteUrl = '';
  let remoteHost = ORIGIN_PROBE_HOST;

  const remoteResult = runGit(['remote', 'get-url', 'origin'], repoRoot, 6000);
  steps.push({
    operation: 'list_remotes_origin',
    success: remoteResult.exitCode === 0 && !!remoteResult.stdout,
    durationMs: remoteResult.durationMs,
    exitCode: remoteResult.exitCode,
    message: remoteResult.exitCode === 0 ? 'origin remote ok' : (remoteResult.stderr || 'no origin remote'),
    detail: remoteResult.stderr
  });
  if (remoteResult.exitCode === 0 && remoteResult.stdout) {
    remoteUrl = redactUrl(remoteResult.stdout);
    remoteHost = hostFromUrl(remoteResult.stdout) || ORIGIN_PROBE_HOST;
  }

  const dnsResult = await resolveDns(remoteHost || ORIGIN_PROBE_HOST);
  steps.push(dnsResult);

  const httpResult = await httpProbe(remoteHost || ORIGIN_PROBE_HOST, ORIGIN_PROBE_PATH, ORIGIN_TIMEOUT_MS);
  steps.push(httpResult);

  // Optional lightweight fetch of refs (no auth required for public repos; failures still classify)
  const lsRemote = runGit(['ls-remote', '--heads', 'origin'], repoRoot, ORIGIN_TIMEOUT_MS);
  steps.push({
    operation: 'git_ls_remote',
    success: lsRemote.exitCode === 0,
    durationMs: lsRemote.durationMs,
    host: remoteHost,
    remoteUrl,
    exitCode: lsRemote.exitCode,
    message: lsRemote.exitCode === 0 ? 'git ls-remote ok' : (lsRemote.stderr || 'git ls-remote failed'),
    detail: lsRemote.stderr
  });

  const failed = steps.filter((s) => !s.success);
  const totalMs = Date.now() - started;
  let classification = { class: '', label: '' };
  if (failed.length === 0) {
    if (totalMs >= ORIGIN_SLOW_MS || steps.some((s) => (s.durationMs || 0) >= ORIGIN_SLOW_MS)) {
      classification = { class: 'likely_slow_network', label: 'Likely slow Wi‑Fi / network' };
    }
  } else {
    const worst = failed[failed.length - 1];
    classification = classifyDiagFailure({
      durationMs: worst.durationMs || totalMs,
      code: worst.code || worst.exitCode,
      httpStatus: worst.httpStatus,
      message: worst.message,
      stderr: worst.detail,
      local: worst.operation === 'list_remotes_origin' && /not a git repository|enoent/i.test(String(worst.message || '') + String(worst.detail || ''))
    });
  }

  const success = failed.length === 0;
  const summary = success
    ? (classification.class === 'likely_slow_network'
      ? `Origin reachable but slow (${totalMs} ms)`
      : `Origin reachable (${totalMs} ms)`)
    : `Origin check failed — ${classification.label || 'see details'}`;

  const entry = {
    at: new Date().toISOString(),
    operation: 'origin_check',
    success,
    durationMs: totalMs,
    host: remoteHost,
    remoteUrl,
    exitCode: failed.length ? (failed[failed.length - 1].exitCode || failed[failed.length - 1].code || 1) : 0,
    httpStatus: httpResult.httpStatus,
    errorClass: classification.class,
    errorLabel: classification.label,
    message: summary,
    detail: steps.map((s) => {
      const bits = [s.operation, s.success ? 'ok' : 'fail', `${Math.round(s.durationMs || 0)}ms`];
      if (s.httpStatus) bits.push(`http ${s.httpStatus}`);
      if (s.code) bits.push(String(s.code));
      if (!s.success && s.message) bits.push(sanitizeSnippet(s.message, 80));
      return bits.join(': ');
    }).join(' | ')
  };
  appendDiagLog(entry);

  return {
    ok: true,
    success,
    summary,
    classification,
    remoteUrl,
    host: remoteHost,
    durationMs: totalMs,
    steps: steps.map((s) => ({
      operation: s.operation,
      success: !!s.success,
      durationMs: Math.round(s.durationMs || 0),
      host: s.host || remoteHost,
      exitCode: s.exitCode != null ? s.exitCode : null,
      httpStatus: s.httpStatus != null ? s.httpStatus : null,
      code: s.code || null,
      message: sanitizeSnippet(s.message, 160)
    })),
    entry
  };
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
    const normalized = { id: t.id, title: t.title, cols, rows, cells };
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
  const normalized = { id: t.id, title: t.title, cols, rows, cells };
  if (columnWidths) normalized.columnWidths = columnWidths;
  return normalized;
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
  return {
    tabs,
    activeTabId,
    combinedPrompt: typeof parsed.combinedPrompt === 'string' ? parsed.combinedPrompt : '',
    globalCombined: typeof parsed.globalCombined === 'boolean' ? parsed.globalCombined : true,
    partPrompts: parsed.partPrompts && typeof parsed.partPrompts === 'object' && !Array.isArray(parsed.partPrompts)
      ? Object.fromEntries(Object.entries(parsed.partPrompts)
        .filter(([, prompt]) => typeof prompt === 'string'))
      : {},
    separators: normalizeSeparators(parsed.separators)
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
    return { ok: true };
  } catch (err) {
    console.error('Failed to save session:', err);
    return { ok: false, error: err.message || 'Failed to save open prompt documents.' };
  }
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
    fs.writeFileSync(targetPath, JSON.stringify({
      format: DOCUMENT_FORMAT,
      version: DOCUMENT_VERSION,
      data: normalized
    }, null, 2), 'utf8');
    return { ok: true, filePath: targetPath };
  } catch (err) {
    console.error('Failed to save prompt document:', err);
    return { ok: false, error: err.message || 'Failed to save file.' };
  }
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
ipcMain.on('window:set-title', (event, title) => {
  const window = BrowserWindow.fromWebContents(event.sender);
  if (window && typeof title === 'string') window.setTitle(`${title} v${app.getVersion()}`);
});
ipcMain.handle('documents:load-session', () => loadSession());
ipcMain.handle('documents:save-session', (_event, session) => saveSession(session));
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

ipcMain.handle('diag:log-load', () => ({ ok: true, entries: loadDiagLog() }));
ipcMain.handle('diag:log-append', (_event, entry) => appendDiagLog(entry || {}));
ipcMain.handle('diag:classify', (_event, input) => classifyDiagFailure(input || {}));
ipcMain.handle('diag:origin-check', async () => {
  try {
    return await checkOriginReachability();
  } catch (err) {
    const classification = classifyDiagFailure({
      message: err && err.message,
      code: err && err.code,
      local: true
    });
    const entry = {
      at: new Date().toISOString(),
      operation: 'origin_check',
      success: false,
      durationMs: null,
      host: ORIGIN_PROBE_HOST,
      errorClass: classification.class,
      errorLabel: classification.label,
      message: sanitizeSnippet(err && err.message, 200),
      detail: sanitizeSnippet(err && err.stack, 320)
    };
    appendDiagLog(entry);
    return {
      ok: false,
      success: false,
      summary: 'Origin check crashed — ' + (classification.label || 'app error'),
      classification,
      error: err && err.message ? err.message : 'Origin check failed',
      entry
    };
  }
});
