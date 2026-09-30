(function () {
  'use strict';

  const DEFAULT_COLS = 3;
  const DEFAULT_ROWS = 8;
  const DEFAULT_COLUMN_WIDTH = 160;
  const MIN_COLUMN_WIDTH = 100;
  const DEFAULT_SEPARATORS = {
    part: '\\n\\n',
    column: ' | ',
    row: '\\n'
  };

  const state = {
    documents: [],
    activeDocumentId: null,
    tabs: [],
    activeTabId: null,
    combinedPrompt: '',
    globalCombined: true,
    partPrompts: {},
    separators: {
      part: DEFAULT_SEPARATORS.part,
      column: DEFAULT_SEPARATORS.column,
      row: DEFAULT_SEPARATORS.row
    },
    confirmedLinks: [],
    saveTimer: null
  };
  let autosaveLocations = null;
  let initialized = false;
  let focusedCell = null;
  /** UI-only row filter: 'all' | 'nonempty' | 'included' */
  let gridRowFilter = 'all';
  /** UI-only Column 1 value filter: null = all values; Set of trimmed strings ('' = blank) */
  let col1ValueFilter = null;
  let col1FilterMenuOpen = false;
  const pendingDocumentPaths = [];
  const EVENT_LOG_LIMIT = 40;
  const eventLog = [];
  let statusPanelOpen = false;

  const el = {
    documentBar: document.getElementById('document-bar'),
    appVersion: document.getElementById('app-version'),
    tabBar: document.getElementById('tab-bar'),
    cellGrid: document.getElementById('cell-grid'),
    combinedResizer: document.getElementById('combined-resizer'),
    combinedSection: document.querySelector('.combined-section'),
    masterLibrary: document.getElementById('master-library'),
    masterLibraryItems: document.getElementById('master-library-items'),
    partLabel: document.getElementById('part-label'),
    combined: document.getElementById('combined-prompt'),
    globalCombined: document.getElementById('global-combined'),
    status: document.getElementById('status'),
    statusRow: document.querySelector('.status-row'),
    statusPanel: document.getElementById('status-panel'),
    statusSnapshot: document.getElementById('status-snapshot'),
    statusEvents: document.getElementById('status-events'),
    statusPanelClose: document.getElementById('status-panel-close'),
    btnAdd: document.getElementById('btn-add-tab'),
    btnRename: document.getElementById('btn-rename-tab'),
    btnDelete: document.getElementById('btn-delete-tab'),
    btnAddRow: document.getElementById('btn-add-row'),
    btnAddCol: document.getElementById('btn-add-col'),
    btnFilterAll: document.getElementById('btn-filter-all'),
    btnFilterNonempty: document.getElementById('btn-filter-nonempty'),
    btnFilterIncluded: document.getElementById('btn-filter-included'),
    btnAppend: document.getElementById('btn-append'),
    btnCopy: document.getElementById('btn-copy'),
    btnClear: document.getElementById('btn-clear'),
    btnUndoClear: document.getElementById('btn-undo-clear'),
    btnExport: document.getElementById('btn-export'),
    btnImport: document.getElementById('btn-import'),
    partSeparator: document.getElementById('part-separator'),
    columnSeparator: document.getElementById('column-separator'),
    rowSeparator: document.getElementById('row-separator')
  };

  function uid() {
    return 'tab-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 7);
  }

  function emptyCells(cols, rows) {
    return Array(cols * rows).fill('');
  }

  function emptySleptCells(cols, rows) {
    return Array(cols * rows).fill(false);
  }

  /** Boolean array parallel to cells; missing/short arrays pad with false. */
  function normalizeSleptCells(slept, length) {
    const needed = length > 0 ? length : 0;
    let out;
    if (Array.isArray(slept)) {
      out = slept.map(function (v) { return !!v; });
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

  function makeTab(id, title) {
    return {
      id: id,
      title: title,
      cols: DEFAULT_COLS,
      rows: DEFAULT_ROWS,
      cells: emptyCells(DEFAULT_COLS, DEFAULT_ROWS),
      sleptCells: emptySleptCells(DEFAULT_COLS, DEFAULT_ROWS)
    };
  }

  function makeDefaultData() {
    return {
      tabs: [
        makeTab('master', 'Master'),
        makeTab('tab-1', 'Part 1'),
        makeTab('tab-2', 'Part 2'),
        makeTab('tab-3', 'Part 3')
      ],
      activeTabId: 'tab-1',
      combinedPrompt: '',
      globalCombined: true,
      partPrompts: {},
      separators: Object.assign({}, DEFAULT_SEPARATORS),
      confirmedLinks: []
    };
  }

  function normalizeColumnWidths(widths, cols) {
    if (!Array.isArray(widths) || widths.length !== cols) return null;
    const normalized = widths.map(function (width) {
      const value = Number(width);
      return Number.isFinite(value) && value >= MIN_COLUMN_WIDTH ? value : null;
    });
    return normalized.every(function (width) { return width !== null; })
      ? normalized
      : null;
  }


  /** CSS grid-template columns for tab content cells (no row-gutter). Uses tab.columnWidths. */
  function contentColumnTemplate(tab, equalMinPx) {
    const widths = normalizeColumnWidths(tab.columnWidths, tab.cols);
    if (widths) {
      return widths.map(function (width) { return width + 'px'; }).join(' ');
    }
    const min = Number.isFinite(equalMinPx) ? equalMinPx : 120;
    return 'repeat(' + tab.cols + ', minmax(' + min + 'px, 1fr))';
  }

  function normalizeSeparators(separators) {
    const source = separators && typeof separators === 'object' ? separators : {};
    return {
      part: typeof source.part === 'string' ? source.part : DEFAULT_SEPARATORS.part,
      column: typeof source.column === 'string' ? source.column : DEFAULT_SEPARATORS.column,
      row: typeof source.row === 'string' ? source.row : DEFAULT_SEPARATORS.row
    };
  }

  // The UI stores visible escape sequences so line separators are easy to edit.
  function separatorValue(value) {
    return String(value || '')
      .replace(/\\r/g, '\r')
      .replace(/\\n/g, '\n')
      .replace(/\\t/g, '\t');
  }

  /**
   * Migrate/normalize a tab from storage or import.
   * Legacy { content: string } → first cell of 3×8 grid.
   */
  function normalizeTab(t) {
    if (!t || typeof t !== 'object') return null;
    if (typeof t.id !== 'string' || !t.id) return null;
    if (typeof t.title !== 'string') return null;

    if (typeof t.content === 'string' && !Array.isArray(t.cells)) {
      const cells = emptyCells(DEFAULT_COLS, DEFAULT_ROWS);
      cells[0] = t.content;
      return {
        id: t.id,
        title: t.title,
        cols: DEFAULT_COLS,
        rows: DEFAULT_ROWS,
        cells: cells,
        sleptCells: emptySleptCells(DEFAULT_COLS, DEFAULT_ROWS)
      };
    }

    let cols = Number.isInteger(t.cols) && t.cols > 0 ? t.cols : DEFAULT_COLS;
    let rows = Number.isInteger(t.rows) && t.rows > 0 ? t.rows : DEFAULT_ROWS;
    let cells;

    if (Array.isArray(t.cells)) {
      cells = t.cells.map(function (c) {
        return typeof c === 'string' ? c : '';
      });
    } else {
      return null;
    }

    const needed = cols * rows;
    if (cells.length < needed) {
      cells = cells.concat(Array(needed - cells.length).fill(''));
    } else if (cells.length > needed) {
      cells = cells.slice(0, needed);
    }

    const columnWidths = normalizeColumnWidths(t.columnWidths, cols);
    const sleptCells = normalizeSleptCells(t.sleptCells, needed);
    const normalized = { id: t.id, title: t.title, cols: cols, rows: rows, cells: cells, sleptCells: sleptCells };
    if (columnWidths) normalized.columnWidths = columnWidths;
    return normalized;
  }

  function activeTab() {
    return state.tabs.find(function (t) {
      return t.id === state.activeTabId;
    }) || null;
  }

  function isMasterTab(tab) {
    return !!tab && (tab.id === 'master' || state.tabs[0] === tab);
  }

  function partTabs() {
    return state.tabs.filter(function (tab) { return !isMasterTab(tab); });
  }

  function defaultActiveTabId(tabs) {
    const part = tabs.find(function (tab) {
      return tab.id !== 'master' && tab.title.trim().toLowerCase() !== 'master';
    });
    return (part || tabs[0] || {}).id || null;
  }

  function nextPartTitle() {
    let n = partTabs().length + 1;
    const used = new Set(state.tabs.map(function (tab) { return tab.title.trim().toLowerCase(); }));
    while (used.has(('part ' + n).toLowerCase())) n++;
    return 'Part ' + n;
  }

  function snapshot() {
    return {
      tabs: state.tabs,
      activeTabId: state.activeTabId,
      combinedPrompt: state.combinedPrompt,
      globalCombined: state.globalCombined,
      partPrompts: Object.assign({}, state.partPrompts),
      separators: state.separators,
      confirmedLinks: state.confirmedLinks.map(function (link) {
        return {
          id: link.id,
          tabId: link.tabId,
          cellIndex: link.cellIndex,
          text: link.text,
          start: link.start,
          end: link.end,
          scope: link.scope
        };
      })
    };
  }

  const UNDO_LIMIT = 50;
  const EDIT_COALESCE_MS = 450;
  let undoStack = [];
  let redoStack = [];
  let historySuspended = false;
  let historyCoalescing = false;
  let historyCoalesceTimer = null;
  let lastClearSnapshot = null;

  function cloneDocumentData(data) {
    return JSON.parse(JSON.stringify({
      tabs: data.tabs || [],
      activeTabId: data.activeTabId || null,
      combinedPrompt: typeof data.combinedPrompt === 'string' ? data.combinedPrompt : '',
      globalCombined: typeof data.globalCombined === 'boolean' ? data.globalCombined : true,
      partPrompts: data.partPrompts && typeof data.partPrompts === 'object' ? data.partPrompts : {},
      separators: data.separators || DEFAULT_SEPARATORS,
      confirmedLinks: Array.isArray(data.confirmedLinks) ? data.confirmedLinks : []
    }));
  }

  function cloneCurrentDocument() {
    return cloneDocumentData(snapshot());
  }

  function clearHistory() {
    undoStack = [];
    redoStack = [];
    historyCoalescing = false;
    if (historyCoalesceTimer) {
      clearTimeout(historyCoalesceTimer);
      historyCoalesceTimer = null;
    }
    dismissUndoClear();
  }

  function endHistoryCoalesce() {
    historyCoalescing = false;
    if (historyCoalesceTimer) {
      clearTimeout(historyCoalesceTimer);
      historyCoalesceTimer = null;
    }
  }

  function armHistoryCoalesce() {
    historyCoalescing = true;
    if (historyCoalesceTimer) clearTimeout(historyCoalesceTimer);
    historyCoalesceTimer = setTimeout(endHistoryCoalesce, EDIT_COALESCE_MS);
  }

  /**
   * Snapshot document state onto the undo stack before a mutation.
   * options.coalesce: group bursty edits (cell typing, Combined typing) into one undo step.
   */
  function pushHistory(options) {
    if (historySuspended) return;
    const coalesce = !!(options && options.coalesce);
    if (coalesce) {
      if (historyCoalescing) {
        armHistoryCoalesce();
        return;
      }
      armHistoryCoalesce();
    } else {
      endHistoryCoalesce();
    }
    undoStack.push(cloneCurrentDocument());
    if (undoStack.length > UNDO_LIMIT) undoStack.shift();
    redoStack = [];
  }

  function restoreHistoryEntry(entry) {
    historySuspended = true;
    try {
      applyData(entry);
      const doc = activeDocument();
      if (doc) doc.data = snapshot();
      scheduleSave();
    } finally {
      historySuspended = false;
    }
  }

  function undo() {
    endHistoryCoalesce();
    if (!undoStack.length) {
      setStatus('Nothing to undo');
      return;
    }
    redoStack.push(cloneCurrentDocument());
    if (redoStack.length > UNDO_LIMIT) redoStack.shift();
    restoreHistoryEntry(undoStack.pop());
    setStatus('Undid', 'ok');
  }

  function redo() {
    endHistoryCoalesce();
    if (!redoStack.length) {
      setStatus('Nothing to redo');
      return;
    }
    undoStack.push(cloneCurrentDocument());
    if (undoStack.length > UNDO_LIMIT) undoStack.shift();
    restoreHistoryEntry(redoStack.pop());
    setStatus('Redid', 'ok');
  }

  function activeDocument() {
    return state.documents.find(function (document) {
      return document.id === state.activeDocumentId;
    }) || null;
  }

  function sessionSnapshot() {
    const document = activeDocument();
    if (document) document.data = snapshot();
    return {
      documents: state.documents,
      activeDocumentId: state.activeDocumentId
    };
  }

  function scheduleSave() {
    const document = activeDocument();
    if (document) document.data = snapshot();
    clearTimeout(state.saveTimer);
    state.saveTimer = setTimeout(persist, 250);
  }

  async function persist() {
    if (!window.click2copy) return false;
    try {
      const sessionResult = await window.click2copy.saveSession(sessionSnapshot());
      if (!sessionResult || !sessionResult.ok) {
        setStatus((sessionResult && sessionResult.error) || 'Failed to save open documents', 'err');
        return false;
      }
      if (sessionResult.locations) autosaveLocations = sessionResult.locations;
      if (sessionResult.backupOk === false) {
        pushEvent(sessionResult.backupError || 'Recoverable .c2copy autosave backup failed', 'err');
      }

      const document = activeDocument();
      if (document && document.filePath) {
        // Named Save As path: autosave only there (do not invent another named path).
        const result = await window.click2copy.saveDocument(document.filePath, snapshot());
        if (!result || !result.ok) {
          setStatus((result && result.error) || 'Failed to save prompt file', 'err');
          return false;
        }
      } else {
        // Untitled: userData JSON store + mirrored latest .c2copy / .json in backup folders.
        const saved = await window.click2copy.saveData(snapshot());
        if (!saved) {
          setStatus('Failed to save local prompt data', 'err');
          return false;
        }
      }
      return true;
    } catch (err) {
      console.error(err);
      setStatus('Failed to save', 'err');
      return false;
    }
  }

  function formatClock(date) {
    const d = date instanceof Date ? date : new Date(date);
    const hh = String(d.getHours()).padStart(2, '0');
    const mm = String(d.getMinutes()).padStart(2, '0');
    const ss = String(d.getSeconds()).padStart(2, '0');
    return hh + ':' + mm + ':' + ss;
  }

  function escapeHtml(value) {
    return String(value == null ? '' : value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function describeSelectedCell() {
    const tab = activeTab();
    if (!tab) return 'none';
    if (!focusedCell || focusedCell.tabId !== tab.id ||
        !Number.isInteger(focusedCell.index) ||
        focusedCell.index < 0 || focusedCell.index >= tab.cells.length) {
      return 'none (tab ' + tab.title + ')';
    }
    const row = Math.floor(focusedCell.index / tab.cols) + 1;
    const col = (focusedCell.index % tab.cols) + 1;
    const preview = (tab.cells[focusedCell.index] || '').trim().replace(/\s+/g, ' ');
    const short = preview.length > 40 ? preview.slice(0, 37) + '…' : preview;
    return 'R' + row + 'C' + col + ' idx ' + focusedCell.index +
      (short ? ' — "' + short + '"' : ' — empty');
  }

  function buildUnderHoodSnapshot() {
    const document = activeDocument();
    const tab = activeTab();
    const docLabel = document
      ? (document.title || 'Untitled') + (document.filePath ? ' (' + document.filePath + ')' : ' (unsaved)')
      : 'none';
    const tabLabel = tab
      ? tab.title + ' [' + tab.id + ']' + (isMasterTab(tab) ? ' (Master)' : '')
      : 'none';
    const gridLabel = tab ? (tab.cols + '×' + tab.rows + ' (' + tab.cells.length + ' cells)') : 'n/a';
    const loc = autosaveLocations;
    const docsBackup = loc && loc.documentsBackupDir
      ? loc.documentsBackupDir + '/' + (loc.latestC2copyName || 'latest-autosave.c2copy')
      : '(unavailable)';
    const userBackup = loc && loc.userDataBackupDir
      ? loc.userDataBackupDir + '/' + (loc.latestC2copyName || 'latest-autosave.c2copy')
      : '(unavailable)';
    const saveTarget = document && document.filePath
      ? document.filePath
      : (docsBackup !== '(unavailable)' ? docsBackup : userBackup);
    return [
      { label: 'Active document', value: docLabel },
      { label: 'Autosave target', value: saveTarget },
      { label: 'Recoverable backup (Documents)', value: docsBackup },
      { label: 'Recoverable backup (userData)', value: userBackup },
      { label: 'Active tab', value: tabLabel },
      { label: 'Selected cell', value: describeSelectedCell() },
      { label: 'Grid size', value: gridLabel },
      { label: 'Part separator', value: JSON.stringify(state.separators.part) },
      { label: 'Column separator', value: JSON.stringify(state.separators.column) },
      { label: 'Row separator', value: JSON.stringify(state.separators.row) },
      { label: 'Global combined', value: state.globalCombined ? 'yes' : 'no' },
      { label: 'Confirmed links', value: String(state.confirmedLinks.length) },
      { label: 'Open documents', value: String(state.documents.length) },
      { label: 'Event log size', value: String(eventLog.length) }
    ];
  }

  function renderStatusPanel() {
    if (!el.statusPanel || !el.statusSnapshot || !el.statusEvents) return;
    const rows = buildUnderHoodSnapshot();
    el.statusSnapshot.innerHTML = rows.map(function (row) {
      return '<dt>' + escapeHtml(row.label) + '</dt><dd>' + escapeHtml(row.value) + '</dd>';
    }).join('');

    if (eventLog.length === 0) {
      el.statusEvents.innerHTML = '<li class="event-empty">No events yet — actions will appear here.</li>';
      return;
    }
    const newestFirst = eventLog.slice().reverse();
    el.statusEvents.innerHTML = newestFirst.map(function (entry) {
      const kindClass = entry.kind ? ' class="' + escapeHtml(entry.kind) + '"' : '';
      return '<li' + kindClass + '>' +
        '<span class="event-time">' + escapeHtml(formatClock(entry.at)) + '</span>' +
        '<span class="event-msg">' + escapeHtml(entry.msg) + '</span>' +
        '</li>';
    }).join('');
  }

  function setStatusPanelOpen(open) {
    statusPanelOpen = !!open;
    if (!el.statusPanel || !el.status) return;
    el.statusPanel.hidden = !statusPanelOpen;
    el.status.setAttribute('aria-expanded', statusPanelOpen ? 'true' : 'false');
    if (el.statusRow) el.statusRow.classList.toggle('is-panel-open', statusPanelOpen);
    if (statusPanelOpen) renderStatusPanel();
  }

  function pushEvent(msg, kind) {
    if (!msg) return;
    eventLog.push({
      at: Date.now(),
      msg: String(msg),
      kind: kind || ''
    });
    if (eventLog.length > EVENT_LOG_LIMIT) {
      eventLog.splice(0, eventLog.length - EVENT_LOG_LIMIT);
    }
    if (statusPanelOpen) renderStatusPanel();
  }

  function setStatus(msg, kind) {
    el.status.textContent = msg || '';
    el.status.className = 'status' + (kind ? ' ' + kind : '');
    if (msg) {
      pushEvent(msg, kind);
      clearTimeout(setStatus._t);
      setStatus._t = setTimeout(function () {
        if (!statusPanelOpen) {
          el.status.textContent = '';
          el.status.className = 'status';
        }
      }, 2200);
    }
  }

  /** Join non-empty cells of one row with the configured column separator. */
  function rowText(tab, rowIndex) {
    const parts = [];
    for (let c = 0; c < tab.cols; c++) {
      const v = (tab.cells[rowIndex * tab.cols + c] || '').trim();
      if (v) parts.push(v);
    }
    return parts.join(separatorValue(state.separators.column));
  }

  /** All non-empty rows, joined by the configured row separator. */
  function partText(tab) {
    const lines = [];
    for (let r = 0; r < tab.rows; r++) {
      const line = rowText(tab, r);
      if (line) lines.push(line);
    }
    return lines.join(separatorValue(state.separators.row));
  }

  function tabHasContent(tab) {
    return tab.cells.some(function (c) {
      return (c || '').trim().length > 0;
    }) || !!(state.partPrompts[tab.id] || '').trim();
  }

  function linkUid() {
    return 'link-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 7);
  }

  function currentPromptScope() {
    if (state.globalCombined) return 'global';
    const tab = activeTab();
    return tab ? tab.id : 'global';
  }

  function getPromptText(scope) {
    const target = scope || currentPromptScope();
    if (target === 'global') return state.combinedPrompt || '';
    return state.partPrompts[target] || '';
  }

  function setPromptText(scope, text) {
    if (scope === 'global') state.combinedPrompt = text;
    else state.partPrompts[scope] = text;
    // Every Combined edit copies the active prompt to the clipboard (Copy button kept).
    if (initialized && scope === currentPromptScope()) scheduleCombinedAutoCopy();
  }

  let combinedAutoCopyTimer = null;
  let lastCombinedAutoCopyStatusAt = 0;

  function scheduleCombinedAutoCopy() {
    if (combinedAutoCopyTimer) clearTimeout(combinedAutoCopyTimer);
    combinedAutoCopyTimer = setTimeout(function () {
      combinedAutoCopyTimer = null;
      autoCopyCombinedToClipboard();
    }, 120);
  }

  function autoCopyCombinedToClipboard() {
    const text = getCombinedPlainText();
    const now = Date.now();
    const showStatus = now - lastCombinedAutoCopyStatusAt > 700;
    writeTextToClipboard(text == null ? '' : String(text)).then(function () {
      if (showStatus) {
        lastCombinedAutoCopyStatusAt = Date.now();
        setStatus('Copied Combined', 'ok');
      }
    }).catch(function () {
      // Fall back through the shared helper (may surface Copy failed).
      if (showStatus) {
        lastCombinedAutoCopyStatusAt = Date.now();
        copyTextWithStatus(text, 'Copied Combined');
      }
    });
  }

  // Click/focus Combined copies current text (like cell click-copy). Edit + auto-copy-on-edit kept.
  let lastCombinedActivateCopyKey = '';
  let lastCombinedActivateCopyAt = 0;

  function copyCombinedOnActivate() {
    if (!initialized) return;
    const text = getCombinedPlainText();
    const value = text == null ? '' : String(text);
    if (!value) return;
    const key = currentPromptScope() + ':' + value;
    const now = Date.now();
    // Guard against click+focus double-fire (same pattern as cell auto-copy).
    if (key === lastCombinedActivateCopyKey && now - lastCombinedActivateCopyAt < 300) return;
    lastCombinedActivateCopyKey = key;
    lastCombinedActivateCopyAt = now;
    lastCombinedAutoCopyStatusAt = now;
    copyTextWithStatus(value, 'Copied Combined');
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

  function normalizeConfirmedLinks(raw) {
    if (!Array.isArray(raw)) return [];
    const links = [];
    for (let i = 0; i < raw.length; i++) {
      const item = raw[i];
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
        cellIndex: cellIndex,
        text: item.text,
        start: start,
        end: end,
        scope: scope
      });
    }
    return links;
  }

  /** Repair start/end when Combined text still contains link.text (survives load drift). */
  function repairConfirmedLinkOffset(link, text) {
    if (!link || typeof text !== 'string') return false;
    if (link.start >= 0 && link.end <= text.length && link.start <= link.end &&
        text.slice(link.start, link.end) === link.text) {
      return true;
    }
    if (!link.text) return false;
    var from = Math.max(0, link.start - 80);
    var idx = text.indexOf(link.text, from);
    if (idx === -1) idx = text.indexOf(link.text);
    if (idx === -1) return false;
    link.start = idx;
    link.end = idx + link.text.length;
    return true;
  }

  function repairConfirmedLinksForScope(scope) {
    var target = scope || currentPromptScope();
    var text = getPromptText(target);
    state.confirmedLinks = state.confirmedLinks.filter(function (link) {
      if (link.scope !== target) return true;
      if (!linkMatchesSource(link, link.text)) return false;
      return repairConfirmedLinkOffset(link, text);
    });
  }

  function repairAllConfirmedLinks() {
    var scopes = {};
    state.confirmedLinks.forEach(function (link) { scopes[link.scope] = true; });
    Object.keys(scopes).forEach(function (scope) {
      repairConfirmedLinksForScope(scope);
    });
  }

  function linksForScope(scope) {
    const target = scope || currentPromptScope();
    return state.confirmedLinks
      .filter(function (link) { return link.scope === target; })
      .sort(function (a, b) { return a.start - b.start; });
  }

  function sourceCellText(tabId, cellIndex) {
    const tab = state.tabs.find(function (t) { return t.id === tabId; });
    if (!tab || cellIndex < 0 || cellIndex >= tab.cells.length) return null;
    return (tab.cells[cellIndex] || '').trim();
  }

  function linkMatchesSource(link, segmentText) {
    const expected = sourceCellText(link.tabId, link.cellIndex);
    if (expected === null) return false;
    return String(segmentText) === expected;
  }

  function isCellConfirmed(tabId, cellIndex) {
    return state.confirmedLinks.some(function (link) {
      return link.tabId === tabId && link.cellIndex === cellIndex;
    });
  }

  function isCellConfirmedInScope(tabId, cellIndex, scope) {
    const target = scope || currentPromptScope();
    return state.confirmedLinks.some(function (link) {
      return link.tabId === tabId && link.cellIndex === cellIndex && link.scope === target;
    });
  }

  /** True when this Master cell (or matching text from any linked cell) is in Combined. */
  function isMasterCellRepresentedInCombined(master, masterIdx) {
    if (!master || masterIdx < 0 || masterIdx >= master.cells.length) return false;
    if (isCellConfirmed(master.id, masterIdx)) return true;
    const text = (master.cells[masterIdx] || '').trim();
    if (!text) return false;
    return state.confirmedLinks.some(function (link) {
      const expected = sourceCellText(link.tabId, link.cellIndex);
      return expected !== null && expected === text;
    });
  }

  function linksForCellsInScope(tabId, cellIndices, scope) {
    const target = scope || currentPromptScope();
    const wanted = {};
    for (let i = 0; i < cellIndices.length; i++) wanted[cellIndices[i]] = true;
    return state.confirmedLinks.filter(function (link) {
      return link.scope === target && link.tabId === tabId && wanted[link.cellIndex];
    });
  }

  function isRowIncluded(tab, rowIndex) {
    if (!tab || rowIndex < 0 || rowIndex >= tab.rows) return false;
    const scope = currentPromptScope();
    for (let c = 0; c < tab.cols; c++) {
      const idx = rowIndex * tab.cols + c;
      if (!(tab.cells[idx] || '').trim()) continue;
      if (isCellConfirmedInScope(tab.id, idx, scope)) return true;
    }
    return false;
  }

  function dropConfirmedLink(linkId) {
    state.confirmedLinks = state.confirmedLinks.filter(function (link) {
      return link.id !== linkId;
    });
  }

  function dropLinksForScope(scope) {
    state.confirmedLinks = state.confirmedLinks.filter(function (link) {
      return link.scope !== scope;
    });
  }

  function dropLinksForTab(tabId) {
    state.confirmedLinks = state.confirmedLinks.filter(function (link) {
      return link.tabId !== tabId && link.scope !== tabId;
    });
  }

  function remapConfirmedCellIndex(tabId, fromIndex, toIndex) {
    state.confirmedLinks.forEach(function (link) {
      if (link.tabId === tabId && link.cellIndex === fromIndex) {
        link.cellIndex = toIndex;
      }
    });
  }

  function remapConfirmedAfterColumnAdd(tabId, oldCols, newCols) {
    state.confirmedLinks.forEach(function (link) {
      if (link.tabId !== tabId) return;
      const row = Math.floor(link.cellIndex / oldCols);
      const col = link.cellIndex % oldCols;
      link.cellIndex = row * newCols + col;
    });
  }

  function swapConfirmedRowIndices(tabId, cols, rowA, rowB) {
    state.confirmedLinks.forEach(function (link) {
      if (link.tabId !== tabId) return;
      const row = Math.floor(link.cellIndex / cols);
      const col = link.cellIndex % cols;
      if (row === rowA) link.cellIndex = rowB * cols + col;
      else if (row === rowB) link.cellIndex = rowA * cols + col;
    });
  }

  function shiftConfirmedRowsDown(tabId, cols, fromRow, emptyRow) {
    // moveRow-down shifts occupied rows [fromRow, emptyRow) down by one.
    state.confirmedLinks.forEach(function (link) {
      if (link.tabId !== tabId) return;
      const row = Math.floor(link.cellIndex / cols);
      const col = link.cellIndex % cols;
      if (row >= fromRow && row < emptyRow) {
        link.cellIndex = (row + 1) * cols + col;
      }
    });
  }

  function ensureSleptCells(tab) {
    if (!tab) return;
    tab.sleptCells = normalizeSleptCells(tab.sleptCells, tab.cols * tab.rows);
  }

  function isCellSlept(tab, cellIndex) {
    if (!tab || cellIndex < 0) return false;
    ensureSleptCells(tab);
    return !!tab.sleptCells[cellIndex];
  }

  function setCellSlept(tab, cellIndex, slept) {
    if (!tab || cellIndex < 0 || cellIndex >= tab.cells.length) return;
    ensureSleptCells(tab);
    tab.sleptCells[cellIndex] = !!slept;
  }

  function remapSleptAfterColumnAdd(tab, oldCols, newCols) {
    const oldSlept = Array.isArray(tab.sleptCells) ? tab.sleptCells : [];
    const next = emptySleptCells(newCols, tab.rows);
    for (let r = 0; r < tab.rows; r++) {
      for (let c = 0; c < oldCols; c++) {
        const oldIdx = r * oldCols + c;
        next[r * newCols + c] = !!(oldIdx < oldSlept.length && oldSlept[oldIdx]);
      }
    }
    tab.sleptCells = next;
  }

  function swapSleptRowIndices(tab, rowA, rowB) {
    ensureSleptCells(tab);
    for (let col = 0; col < tab.cols; col++) {
      const a = rowA * tab.cols + col;
      const b = rowB * tab.cols + col;
      const tmp = tab.sleptCells[a];
      tab.sleptCells[a] = tab.sleptCells[b];
      tab.sleptCells[b] = tmp;
    }
  }

  function shiftSleptRowsDown(tab, fromRow, emptyRow) {
    ensureSleptCells(tab);
    for (let row = emptyRow; row > fromRow + 1; row--) {
      for (let col = 0; col < tab.cols; col++) {
        tab.sleptCells[row * tab.cols + col] = tab.sleptCells[(row - 1) * tab.cols + col];
      }
    }
    const moving = tab.sleptCells.slice(fromRow * tab.cols, (fromRow + 1) * tab.cols);
    for (let col = 0; col < tab.cols; col++) {
      tab.sleptCells[(fromRow + 1) * tab.cols + col] = moving[col];
      tab.sleptCells[fromRow * tab.cols + col] = false;
    }
  }

  function toggleCellSleep(cellIndex, slept) {
    const tab = activeTab();
    if (!tab || cellIndex < 0 || cellIndex >= tab.cells.length) return;
    pushHistory();
    setCellSlept(tab, cellIndex, slept);
    const wrap = el.cellGrid.querySelector('.cell-wrap[data-row="' + Math.floor(cellIndex / tab.cols) + '"][data-col="' + (cellIndex % tab.cols) + '"]');
    if (wrap) wrap.classList.toggle('is-slept', !!slept);
    scheduleSave();
    const row = Math.floor(cellIndex / tab.cols) + 1;
    const col = (cellIndex % tab.cols) + 1;
    setStatus(slept ? ('Slept R' + row + 'C' + col) : ('Woke R' + row + 'C' + col));
  }

  /**
   * Click+drag paint for Combined / sleep / row checkboxes.
   * Kinds never mix: a drag started on Combined only paints Combined, etc.
   * First control sets the target value; later controls of the same kind match it.
   */
  let checkboxDrag = null;

  function endCheckboxDragListeners() {
    document.removeEventListener('pointermove', onCheckboxDragMove);
    document.removeEventListener('pointerup', endCheckboxDrag);
    document.removeEventListener('pointercancel', endCheckboxDrag);
    document.body.classList.remove('painting-checkboxes');
  }

  function ensureCellConfirmedState(cellIndex, wantOn, opts) {
    const quiet = !!(opts && opts.quiet);
    const tab = activeTab();
    if (!tab || cellIndex < 0 || cellIndex >= tab.cells.length) return false;
    const scope = currentPromptScope();
    const isOn = isCellConfirmedInScope(tab.id, cellIndex, scope);
    if (isOn === !!wantOn) return false;
    if (wantOn) {
      const value = (tab.cells[cellIndex] || '').trim();
      if (!value) return false;
      appendPieces([{
        type: 'confirmed',
        text: value,
        tabId: tab.id,
        cellIndex: cellIndex
      }], quiet ? null : (tab.title + ' R' + (Math.floor(cellIndex / tab.cols) + 1) +
        'C' + ((cellIndex % tab.cols) + 1)), { quiet: quiet });
      return true;
    }
    const links = linksForCellsInScope(tab.id, [cellIndex], scope);
    if (!links.length) return false;
    removeLinksFromCombined(links);
    if (!quiet) refreshAfterConfirmedChange();
    return true;
  }

  function ensureRowIncludedState(rowIndex, wantOn, opts) {
    const quiet = !!(opts && opts.quiet);
    const tab = activeTab();
    if (!tab || rowIndex < 0 || rowIndex >= tab.rows) return false;
    const scope = currentPromptScope();
    const isOn = isRowIncluded(tab, rowIndex);
    if (isOn === !!wantOn) return false;
    if (wantOn) {
      const pieces = rowConfirmedPieces(tab, rowIndex);
      if (!pieces.length) return false;
      appendPieces(pieces, quiet ? null : (tab.title + ' row ' + (rowIndex + 1)), { quiet: quiet });
      return true;
    }
    const links = linksForCellsInScope(tab.id, rowCellIndices(tab, rowIndex), scope);
    if (!links.length) return false;
    removeLinksFromCombined(links);
    if (!quiet) refreshAfterConfirmedChange();
    return true;
  }

  function ensureCellSleepState(cellIndex, wantSlept, opts) {
    const quiet = !!(opts && opts.quiet);
    const tab = activeTab();
    if (!tab || cellIndex < 0 || cellIndex >= tab.cells.length) return false;
    if (!!isCellSlept(tab, cellIndex) === !!wantSlept) return false;
    setCellSlept(tab, cellIndex, wantSlept);
    const wrap = el.cellGrid.querySelector(
      '.cell-wrap[data-row="' + Math.floor(cellIndex / tab.cols) + '"][data-col="' + (cellIndex % tab.cols) + '"]'
    );
    if (wrap) wrap.classList.toggle('is-slept', !!wantSlept);
    const input = wrap ? wrap.querySelector('.cell-sleep-input') : null;
    if (input) input.checked = !!wantSlept;
    if (!quiet) {
      scheduleSave();
      const row = Math.floor(cellIndex / tab.cols) + 1;
      const col = (cellIndex % tab.cols) + 1;
      setStatus(wantSlept ? ('Slept R' + row + 'C' + col) : ('Woke R' + row + 'C' + col));
    }
    return true;
  }

  function applyCheckboxDragKey(kind, key) {
    if (!checkboxDrag || checkboxDrag.kind !== kind) return;
    if (checkboxDrag.visited[key]) return;
    checkboxDrag.visited[key] = true;
    const want = checkboxDrag.value;
    if (kind === 'combined') ensureCellConfirmedState(key, want, { quiet: true });
    else if (kind === 'row') ensureRowIncludedState(key, want, { quiet: true });
    else if (kind === 'sleep') ensureCellSleepState(key, want, { quiet: true });
    if (kind === 'combined' || kind === 'row') applyAppendCheckedState();
  }

  function checkboxDragHit(clientX, clientY) {
    const node = document.elementFromPoint(clientX, clientY);
    if (!node || typeof node.closest !== 'function') return null;
    if (checkboxDrag.kind === 'combined') {
      const btn = node.closest('.cell-append');
      if (!btn || !el.cellGrid.contains(btn)) return null;
      const idx = parseInt(btn.dataset.idx, 10);
      if (Number.isNaN(idx)) return null;
      return { kind: 'combined', key: idx };
    }
    if (checkboxDrag.kind === 'row') {
      const btn = node.closest('.row-append');
      if (!btn || !el.cellGrid.contains(btn)) return null;
      const row = parseInt(btn.dataset.row, 10);
      if (Number.isNaN(row)) return null;
      return { kind: 'row', key: row };
    }
    if (checkboxDrag.kind === 'sleep') {
      const input = node.closest('.cell-sleep-input') ||
        (node.closest('.cell-sleep') && node.closest('.cell-sleep').querySelector('.cell-sleep-input'));
      if (!input || !el.cellGrid.contains(input)) return null;
      const wrap = input.closest('.cell-wrap');
      if (!wrap) return null;
      const row = parseInt(wrap.dataset.row, 10);
      const col = parseInt(wrap.dataset.col, 10);
      const tab = activeTab();
      if (!tab || Number.isNaN(row) || Number.isNaN(col)) return null;
      return { kind: 'sleep', key: row * tab.cols + col };
    }
    return null;
  }

  function onCheckboxDragMove(e) {
    if (!checkboxDrag) return;
    const hit = checkboxDragHit(e.clientX, e.clientY);
    if (!hit || hit.kind !== checkboxDrag.kind) return;
    applyCheckboxDragKey(hit.kind, hit.key);
  }

  function endCheckboxDrag() {
    if (!checkboxDrag) {
      endCheckboxDragListeners();
      return;
    }
    const kind = checkboxDrag.kind;
    checkboxDrag = null;
    historySuspended = false;
    endCheckboxDragListeners();
    if (kind === 'combined' || kind === 'row') refreshAfterConfirmedChange();
    else scheduleSave();
  }

  function beginCheckboxDrag(kind, key, value, event) {
    if (checkboxDrag) endCheckboxDrag();
    if (cellRangeDrag) return;
    checkboxDrag = {
      kind: kind,
      value: !!value,
      visited: {},
      pointerId: event && event.pointerId
    };
    checkboxDrag.visited[key] = true;
    pushHistory();
    historySuspended = true;
    document.body.classList.add('painting-checkboxes');
    document.addEventListener('pointermove', onCheckboxDragMove);
    document.addEventListener('pointerup', endCheckboxDrag);
    document.addEventListener('pointercancel', endCheckboxDrag);
    if (kind === 'combined') ensureCellConfirmedState(key, value, { quiet: true });
    else if (kind === 'row') ensureRowIncludedState(key, value, { quiet: true });
    else if (kind === 'sleep') ensureCellSleepState(key, value, { quiet: true });
    if (kind === 'combined' || kind === 'row') applyAppendCheckedState();
  }

  function applyAppendCheckedState() {
    const tab = activeTab();
    if (!tab) return;
    const scope = currentPromptScope();

    const rowButtons = el.cellGrid.querySelectorAll('.row-append');
    for (let i = 0; i < rowButtons.length; i++) {
      const btn = rowButtons[i];
      const row = parseInt(btn.dataset.row, 10);
      if (Number.isNaN(row)) continue;
      const on = isRowIncluded(tab, row);
      btn.classList.toggle('is-checked', on);
      btn.setAttribute('aria-pressed', on ? 'true' : 'false');
      btn.title = on
        ? 'Remove row ' + (row + 1) + ' from combined prompt'
        : 'Add row ' + (row + 1) + ' to combined prompt';
      btn.setAttribute('aria-label', btn.title);
      // Keep In Combined filter in sync when checkbox state changes.
      const hidden = !rowMatchesFilter(tab, row);
      const controls = btn.closest('.row-controls');
      if (controls) controls.classList.toggle('is-row-filtered', hidden);
      const wraps = el.cellGrid.querySelectorAll('.cell-wrap[data-row="' + row + '"]');
      // Fallback: mark by dataset on wrap if present; else scan cells.
      if (wraps.length) {
        for (let w = 0; w < wraps.length; w++) wraps[w].classList.toggle('is-row-filtered', hidden);
      } else {
        const cells = el.cellGrid.querySelectorAll('.cell[data-row="' + row + '"]');
        for (let c = 0; c < cells.length; c++) {
          const wrap = cells[c].closest ? cells[c].closest('.cell-wrap') : null;
          if (wrap) wrap.classList.toggle('is-row-filtered', hidden);
        }
      }
    }

    const cellButtons = el.cellGrid.querySelectorAll('.cell-append');
    for (let i = 0; i < cellButtons.length; i++) {
      const btn = cellButtons[i];
      const idx = parseInt(btn.dataset.idx, 10);
      if (Number.isNaN(idx)) continue;
      const on = isCellConfirmedInScope(tab.id, idx, scope);
      btn.classList.toggle('is-checked', on);
      btn.setAttribute('aria-pressed', on ? 'true' : 'false');
      const row = Math.floor(idx / tab.cols) + 1;
      const col = (idx % tab.cols) + 1;
      btn.title = on
        ? 'Remove cell R' + row + 'C' + col + ' from combined prompt'
        : 'Add cell R' + row + 'C' + col + ' to combined prompt';
      btn.setAttribute('aria-label', btn.title);
    }
  }

  function applyConfirmedCellHighlights() {
    const tab = activeTab();
    if (!tab) return;
    const cells = el.cellGrid.querySelectorAll('.cell');
    for (let i = 0; i < cells.length; i++) {
      const node = cells[i];
      const idx = parseInt(node.dataset.idx, 10);
      const confirmed = !Number.isNaN(idx) && isCellConfirmed(tab.id, idx);
      node.classList.toggle('cell-confirmed', confirmed);
      const wrap = node.closest ? node.closest('.cell-wrap') : null;
      if (wrap) wrap.classList.toggle('cell-confirmed', confirmed);
    }
    applyAppendCheckedState();
    // Keep Master insert picker greens in sync with Combined confirmed links.
    if (!isMasterTab(tab)) syncMasterLibraryConfirmedState();
  }

  function syncMasterLibraryConfirmedState() {
    const master = state.tabs.find(isMasterTab);
    if (!master || !el.masterLibraryItems) return;
    const buttons = el.masterLibraryItems.querySelectorAll('.master-library-cell:not(.master-library-cell-empty)');
    for (let i = 0; i < buttons.length; i++) {
      const button = buttons[i];
      const row = parseInt(button.dataset.row, 10);
      const col = parseInt(button.dataset.col, 10);
      if (Number.isNaN(row) || Number.isNaN(col)) continue;
      const masterIdx = row * master.cols + col;
      button.classList.toggle('cell-confirmed', isMasterCellRepresentedInCombined(master, masterIdx));
    }
  }

  function readCombinedDomTextAndSpans() {
    const pieces = [];
    let text = '';

    function appendTextNode(value) {
      if (!value) return;
      text += value;
    }

    function walk(node) {
      if (!node) return;
      if (node.nodeType === Node.TEXT_NODE) {
        appendTextNode(node.nodeValue || '');
        return;
      }
      if (node.nodeType !== Node.ELEMENT_NODE) return;

      if (node.classList && node.classList.contains('confirmed-segment')) {
        const seg = node.textContent || '';
        const start = text.length;
        appendTextNode(seg);
        pieces.push({
          node: node,
          id: node.dataset.linkId || '',
          text: seg,
          start: start,
          end: text.length
        });
        return;
      }

      if (node.tagName === 'BR') {
        appendTextNode('\n');
        return;
      }

      const isBlock = /^(DIV|P|LI|H[1-6]|PRE)$/i.test(node.tagName);
      if (isBlock && text && !text.endsWith('\n')) appendTextNode('\n');
      const children = node.childNodes;
      for (let i = 0; i < children.length; i++) walk(children[i]);
      if (isBlock && text && !text.endsWith('\n')) appendTextNode('\n');
    }

    const kids = el.combined.childNodes;
    for (let i = 0; i < kids.length; i++) walk(kids[i]);
    // Contenteditable often leaves a trailing newline from a final empty block.
    if (text.endsWith('\n') && !getPromptText().endsWith('\n') && pieces.length === 0 && text.length === 1) {
      // keep as-is for empty-ish editors
    }
    return { text: text, spans: pieces };
  }

  function unwrapConfirmedNode(node) {
    if (!node || !node.parentNode) return;
    const textNode = document.createTextNode(node.textContent || '');
    node.parentNode.replaceChild(textNode, node);
  }

  function syncConfirmedFromCombinedDom() {
    const scope = currentPromptScope();
    const snapshot = readCombinedDomTextAndSpans();
    const keepIds = {};
    const seen = {};

    snapshot.spans.forEach(function (span) {
      const link = state.confirmedLinks.find(function (item) {
        return item.id === span.id && item.scope === scope;
      });
      const stillGood = link && !seen[span.id] && linkMatchesSource(link, span.text);
      if (!stillGood) {
        unwrapConfirmedNode(span.node);
        return;
      }
      seen[span.id] = true;
      keepIds[span.id] = true;
    });

    state.confirmedLinks = state.confirmedLinks.filter(function (link) {
      if (link.scope !== scope) return true;
      return !!keepIds[link.id];
    });

    const after = readCombinedDomTextAndSpans();
    setPromptText(scope, after.text);
    after.spans.forEach(function (span) {
      const link = state.confirmedLinks.find(function (item) {
        return item.id === span.id && item.scope === scope;
      });
      if (!link) {
        unwrapConfirmedNode(span.node);
        return;
      }
      link.start = span.start;
      link.end = span.end;
      link.text = span.text;
    });
    el.combined.classList.toggle('is-empty', !after.text);
    applyConfirmedCellHighlights();
  }

  function revalidateLinksForCell(tabId, cellIndex, opts) {
    const expected = sourceCellText(tabId, cellIndex);
    let removed = false;
    state.confirmedLinks = state.confirmedLinks.filter(function (link) {
      if (link.tabId !== tabId || link.cellIndex !== cellIndex) return true;
      if (expected !== null && link.text === expected) return true;
      removed = true;
      return false;
    });
    if (removed && !(opts && opts.silent)) {
      renderCombinedPrompt();
      applyConfirmedCellHighlights();
    }
    return removed;
  }

  function revalidateAllConfirmedLinks() {
    let removed = false;
    state.confirmedLinks = state.confirmedLinks.filter(function (link) {
      if (linkMatchesSource(link, link.text)) return true;
      removed = true;
      return false;
    });
    if (removed) {
      renderCombinedPrompt();
      applyConfirmedCellHighlights();
    }
    return removed;
  }

  /**
   * Remove confirmed link ranges from the prompt text, dropping matching
   * column/row separators between removed pieces and a leading part separator.
   */
  function removeLinksFromCombined(linksToRemove) {
    if (!linksToRemove || !linksToRemove.length) return false;
    const scope = linksToRemove[0].scope;
    const removeIds = {};
    for (let i = 0; i < linksToRemove.length; i++) {
      if (linksToRemove[i].scope !== scope) continue;
      removeIds[linksToRemove[i].id] = true;
    }
    if (!Object.keys(removeIds).length) return false;

    let text = getPromptText(scope);
    const all = linksForScope(scope);
    const partSep = separatorValue(state.separators.part);
    const colSep = separatorValue(state.separators.column);
    const rowSep = separatorValue(state.separators.row);

    const ranges = [];
    for (let i = 0; i < all.length; i++) {
      if (!removeIds[all[i].id]) continue;
      ranges.push({ start: all[i].start, end: all[i].end });
    }
    if (!ranges.length) return false;

    ranges.sort(function (a, b) { return a.start - b.start; });
    const merged = [];
    for (let i = 0; i < ranges.length; i++) {
      const range = ranges[i];
      if (!merged.length) {
        merged.push({ start: range.start, end: range.end });
        continue;
      }
      const last = merged[merged.length - 1];
      const gap = text.slice(last.end, range.start);
      if (gap === '' || gap === colSep || gap === rowSep) {
        last.end = range.end;
      } else {
        merged.push({ start: range.start, end: range.end });
      }
    }

    for (let i = 0; i < merged.length; i++) {
      const range = merged[i];
      if (!partSep || range.start < partSep.length) continue;
      const before = text.slice(range.start - partSep.length, range.start);
      if (before !== partSep) continue;
      const covered = all.some(function (link) {
        if (removeIds[link.id]) return false;
        return link.start < range.start && link.end > range.start - partSep.length;
      });
      if (!covered) range.start -= partSep.length;
    }

    merged.sort(function (a, b) { return b.start - a.start; });
    for (let i = 0; i < merged.length; i++) {
      const range = merged[i];
      const len = range.end - range.start;
      if (len <= 0) continue;
      text = text.slice(0, range.start) + text.slice(range.end);
      state.confirmedLinks.forEach(function (link) {
        if (link.scope !== scope || removeIds[link.id]) return;
        if (link.start >= range.end) {
          link.start -= len;
          link.end -= len;
        }
      });
    }

    state.confirmedLinks = state.confirmedLinks.filter(function (link) {
      return !removeIds[link.id];
    });
    setPromptText(scope, text);
    return true;
  }

  function refreshAfterConfirmedChange() {
    renderCombinedPrompt();
    applyConfirmedCellHighlights();
    renderMasterLibrary();
    scheduleSave();
  }

  function toggleCellConfirmed(cellIndex) {
    const tab = activeTab();
    if (!tab || cellIndex < 0 || cellIndex >= tab.cells.length) return;
    const scope = currentPromptScope();
    if (scope !== 'global' && !tab) return;

    if (isCellConfirmedInScope(tab.id, cellIndex, scope)) {
      const links = linksForCellsInScope(tab.id, [cellIndex], scope);
      if (!links.length) {
        applyAppendCheckedState();
        return;
      }
      pushHistory();
      removeLinksFromCombined(links);
      refreshAfterConfirmedChange();
      const row = Math.floor(cellIndex / tab.cols) + 1;
      const col = (cellIndex % tab.cols) + 1;
      setStatus('Removed R' + row + 'C' + col + ' from combined', 'ok');
      return;
    }

    const value = (tab.cells[cellIndex] || '').trim();
    if (!value) {
      setStatus('Nothing to append', 'err');
      return;
    }
    appendPieces([{
      type: 'confirmed',
      text: value,
      tabId: tab.id,
      cellIndex: cellIndex
    }], tab.title + ' R' + (Math.floor(cellIndex / tab.cols) + 1) +
      'C' + ((cellIndex % tab.cols) + 1));
  }

  function appendWithPartSeparator(current, text) {
    const partSeparator = separatorValue(state.separators.part);
    if (current && partSeparator && !current.endsWith(partSeparator)) {
      return current + partSeparator + text;
    }
    return current + text;
  }

  /**
   * Append plain and confirmed pieces into the active combined prompt.
   * pieces: [{ type:'plain'|'confirmed', text, tabId?, cellIndex? }]
   */
  function appendPieces(pieces, label, opts) {
    const quiet = !!(opts && opts.quiet);
    if (!pieces || !pieces.length) {
      if (!quiet) setStatus('Nothing to append', 'err');
      return;
    }
    const scope = currentPromptScope();
    if (scope !== 'global' && !activeTab()) return;

    let current = getPromptText(scope);
    const partSeparator = separatorValue(state.separators.part);
    const toAdd = [];
    if (current && partSeparator && !current.endsWith(partSeparator)) {
      toAdd.push({ type: 'plain', text: partSeparator });
    }
    for (let i = 0; i < pieces.length; i++) toAdd.push(pieces[i]);

    // Preview that we will actually add something before snapshotting undo.
    let willAdd = '';
    for (let i = 0; i < toAdd.length; i++) {
      const piece = toAdd[i];
      const text = piece.text || '';
      if (!text && piece.type !== 'plain') continue;
      willAdd += text;
    }
    if (!willAdd) {
      if (!quiet) setStatus('Nothing to append', 'err');
      return;
    }

    pushHistory();

    let offset = current.length;
    let added = '';
    for (let i = 0; i < toAdd.length; i++) {
      const piece = toAdd[i];
      const text = piece.text || '';
      if (!text && piece.type !== 'plain') continue;
      if (piece.type === 'confirmed') {
        const id = linkUid();
        state.confirmedLinks.push({
          id: id,
          tabId: piece.tabId,
          cellIndex: piece.cellIndex,
          text: text,
          start: offset,
          end: offset + text.length,
          scope: scope
        });
      }
      added += text;
      offset += text.length;
    }

    if (!added) {
      setStatus('Nothing to append', 'err');
      return;
    }

    setPromptText(scope, current + added);
    if (quiet) {
      scheduleSave();
      return;
    }
    renderCombinedPrompt();
    applyConfirmedCellHighlights();
    renderMasterLibrary();
    scheduleSave();
    setStatus('Appended' + (label ? ' "' + label + '"' : ''), 'ok');
  }

  function appendText(text, label) {
    if (!text) {
      setStatus('Nothing to append', 'err');
      return;
    }
    appendPieces([{ type: 'plain', text: text }], label);
  }

  function rowConfirmedPieces(tab, rowIndex) {
    const pieces = [];
    const colSep = separatorValue(state.separators.column);
    for (let c = 0; c < tab.cols; c++) {
      const idx = rowIndex * tab.cols + c;
      const value = (tab.cells[idx] || '').trim();
      if (!value) continue;
      if (pieces.length) pieces.push({ type: 'plain', text: colSep });
      pieces.push({
        type: 'confirmed',
        text: value,
        tabId: tab.id,
        cellIndex: idx
      });
    }
    return pieces;
  }

  function partConfirmedPieces(tab) {
    const pieces = [];
    const rowSep = separatorValue(state.separators.row);
    for (let r = 0; r < tab.rows; r++) {
      const rowPieces = rowConfirmedPieces(tab, r);
      if (!rowPieces.length) continue;
      if (pieces.length) pieces.push({ type: 'plain', text: rowSep });
      for (let i = 0; i < rowPieces.length; i++) pieces.push(rowPieces[i]);
    }
    return pieces;
  }

  function renderCombinedPrompt() {
    const tab = activeTab();
    const scope = currentPromptScope();
    // Repair offsets before painting so greens survive restart / load drift.
    repairConfirmedLinksForScope(scope);
    const text = getPromptText(scope);
    const links = linksForScope(scope).filter(function (link) {
      if (link.start < 0 || link.end > text.length || link.start > link.end) return false;
      return text.slice(link.start, link.end) === link.text;
    });

    // Drop only links that still fail after repair (edited away or source mismatch).
    const validIds = {};
    links.forEach(function (link) { validIds[link.id] = true; });
    state.confirmedLinks = state.confirmedLinks.filter(function (link) {
      if (link.scope !== scope) return true;
      return !!validIds[link.id];
    });

    el.combined.replaceChildren();
    let pos = 0;
    links.forEach(function (link) {
      if (link.start > pos) {
        el.combined.appendChild(document.createTextNode(text.slice(pos, link.start)));
      }
      const span = document.createElement('span');
      span.className = 'confirmed-segment';
      span.dataset.linkId = link.id;
      span.textContent = link.text;
      span.title = 'Confirmed from linked cell — edit to unlink';
      el.combined.appendChild(span);
      pos = link.end;
    });
    if (pos < text.length) {
      el.combined.appendChild(document.createTextNode(text.slice(pos)));
    }

    el.combined.classList.toggle('is-empty', !text);
    el.globalCombined.checked = state.globalCombined;
    applyConfirmedCellHighlights();
  }

  function mergePartPrompts() {
    state.tabs.forEach(function (tab) {
      const prompt = state.partPrompts[tab.id];
      if (!prompt) return;
      const base = state.combinedPrompt || '';
      const partSeparator = separatorValue(state.separators.part);
      let prefix = '';
      if (base && partSeparator && !base.endsWith(partSeparator)) prefix = partSeparator;
      const insertAt = base.length + prefix.length;
      state.combinedPrompt = base + prefix + prompt;
      state.confirmedLinks.forEach(function (link) {
        if (link.scope !== tab.id) return;
        link.scope = 'global';
        link.start += insertAt;
        link.end += insertAt;
      });
    });
    state.partPrompts = {};
  }

  function getCombinedPlainText() {
    return getPromptText(currentPromptScope());
  }

  function applyData(data) {
    focusedCell = null;
    const tabs = [];
    for (let i = 0; i < data.tabs.length; i++) {
      const n = normalizeTab(data.tabs[i]);
      if (n) tabs.push(n);
    }
    if (tabs.length === 0) {
      tabs.push(makeTab('tab-1', 'Part 1'));
    }
    const masterIndex = tabs.findIndex(function (tab) {
      return tab.id === 'master' || tab.title.trim().toLowerCase() === 'master';
    });
    if (masterIndex === -1) {
      tabs.unshift(makeTab('master', 'Master'));
    } else {
      const master = tabs.splice(masterIndex, 1)[0];
      master.title = 'Master';
      tabs.unshift(master);
    }
    state.tabs = tabs;
    state.activeTabId = data.activeTabId || defaultActiveTabId(tabs);
    if (!state.tabs.some(function (t) {
      return t.id === state.activeTabId;
    })) {
      state.activeTabId = defaultActiveTabId(tabs) || tabs[0].id;
    }
    state.combinedPrompt = typeof data.combinedPrompt === 'string' ? data.combinedPrompt : '';
    state.globalCombined = typeof data.globalCombined === 'boolean' ? data.globalCombined : true;
    state.partPrompts = data.partPrompts && typeof data.partPrompts === 'object'
      ? Object.keys(data.partPrompts).reduce(function (prompts, id) {
        if (typeof data.partPrompts[id] === 'string') prompts[id] = data.partPrompts[id];
        return prompts;
      }, {})
      : {};
    state.separators = normalizeSeparators(data.separators);
    state.confirmedLinks = normalizeConfirmedLinks(data.confirmedLinks);
    if (state.globalCombined) mergePartPrompts();
    // Re-anchor Combined ranges after merge so cell + Master-insert greens restore.
    repairAllConfirmedLinks();
    el.partSeparator.value = state.separators.part;
    el.columnSeparator.value = state.separators.column;
    el.rowSeparator.value = state.separators.row;

    renderCombinedPrompt();
    renderTabs();
    renderGrid();
    renderMasterLibrary();
  }

  function renderDocuments() {
    el.documentBar.innerHTML = '';
    state.documents.forEach(function (promptDocument) {
      const item = document.createElement('div');
      item.className = 'document-tab-item';

      const select = document.createElement('button');
      select.type = 'button';
      select.className = 'document-tab' + (promptDocument.id === state.activeDocumentId ? ' active' : '');
      select.title = promptDocument.filePath || promptDocument.title;
      select.setAttribute('role', 'tab');
      select.setAttribute('aria-selected', promptDocument.id === state.activeDocumentId ? 'true' : 'false');
      select.draggable = true;
      select.title += ' (drag to reorder)';

      const title = document.createElement('span');
      title.className = 'document-tab-title';
      title.textContent = promptDocument.title;
      select.appendChild(title);
      select.addEventListener('click', function () {
        selectDocument(promptDocument.id);
      });
      select.addEventListener('dragstart', function (event) {
        event.dataTransfer.setData('text/plain', promptDocument.id);
        event.dataTransfer.effectAllowed = 'move';
        item.classList.add('dragging');
      });
      select.addEventListener('dragend', function () {
        item.classList.remove('dragging');
        el.documentBar.querySelectorAll('.drop-before, .drop-after').forEach(function (target) {
          target.classList.remove('drop-before', 'drop-after');
        });
      });
      select.addEventListener('dragover', function (event) {
        event.preventDefault();
        event.dataTransfer.dropEffect = 'move';
        el.documentBar.querySelectorAll('.drop-before, .drop-after').forEach(function (target) {
          target.classList.remove('drop-before', 'drop-after');
        });
        const bounds = item.getBoundingClientRect();
        item.classList.add(event.clientX < bounds.left + bounds.width / 2 ? 'drop-before' : 'drop-after');
      });
      select.addEventListener('dragleave', function () {
        item.classList.remove('drop-before', 'drop-after');
      });
      select.addEventListener('drop', function (event) {
        event.preventDefault();
        const draggedId = event.dataTransfer.getData('text/plain');
        const bounds = item.getBoundingClientRect();
        reorderDocument(draggedId, promptDocument.id, event.clientX >= bounds.left + bounds.width / 2);
      });

      const close = document.createElement('button');
      close.type = 'button';
      close.className = 'document-tab-close';
      close.textContent = '×';
      close.title = 'Close ' + promptDocument.title;
      close.setAttribute('aria-label', close.title);
      close.addEventListener('click', function (event) {
        event.stopPropagation();
        closeDocument(promptDocument.id);
      });

      item.appendChild(select);
      item.appendChild(close);
      el.documentBar.appendChild(item);
    });

    const activePrompt = activeDocument();
    if (window.click2copy && window.click2copy.setWindowTitle) {
      window.click2copy.setWindowTitle(activePrompt ? activePrompt.title + ' - Click2Copy' : 'Click2Copy');
    }
  }

  function documentDataFromResult(result) {
    if (!result || !result.ok || !result.data) return null;
    return result.data;
  }

  function normalizedPath(filePath) {
    return String(filePath).replace(/\\/g, '/').toLowerCase();
  }

  function documentName(filePath) {
    return String(filePath).split(/[\\/]/).pop() || 'Prompt';
  }

  function isEmptyDocument(data) {
    return data &&
      !data.combinedPrompt &&
      data.tabs.every(function (tab) {
        return tab.cells.every(function (cell) { return !cell; });
      });
  }

  async function flushCurrentDocument() {
    clearTimeout(state.saveTimer);
    return persist();
  }

  function captureCurrentDocument() {
    clearTimeout(state.saveTimer);
    const document = activeDocument();
    if (document) document.data = snapshot();
  }

  async function selectDocument(id) {
    if (id === state.activeDocumentId) return;
    const next = state.documents.find(function (document) { return document.id === id; });
    if (!next) return;
    await flushCurrentDocument();
    state.activeDocumentId = id;
    clearHistory();
    applyData(next.data);
    renderDocuments();
    scheduleSave();
  }

  async function createDocument() {
    await flushCurrentDocument();
    const data = makeDefaultData();
    data.activeTabId = defaultActiveTabId(data.tabs) || data.tabs[0].id;
    const document = {
      id: uid(),
      title: 'Untitled',
      filePath: null,
      data: data
    };
    state.documents.push(document);
    state.activeDocumentId = document.id;
    clearHistory();
    applyData(document.data);
    renderDocuments();
    scheduleSave();
  }

  async function openDocumentPath(filePath) {
    const existing = state.documents.find(function (document) {
      return document.filePath && normalizedPath(document.filePath) === normalizedPath(filePath);
    });
    if (existing) {
      await selectDocument(existing.id);
      return;
    }

    const result = await window.click2copy.openDocument(filePath);
    const data = documentDataFromResult(result);
    if (!data) {
      setStatus((result && result.error) || 'Could not open prompt file', 'err');
      return;
    }
    captureCurrentDocument();
    if (state.documents.length === 1 &&
      state.documents[0].id === 'document-default' &&
      !state.documents[0].filePath &&
      isEmptyDocument(state.documents[0].data)) {
      state.documents = [];
    }
    const document = {
      id: uid(),
      title: documentName(result.filePath || filePath),
      filePath: result.filePath || filePath,
      data: data
    };
    state.documents.push(document);
    state.activeDocumentId = document.id;
    clearHistory();
    applyData(data);
    renderDocuments();
    scheduleSave();
  }

  async function openDocuments() {
    try {
      const result = await window.click2copy.openDocuments();
      if (!result || result.canceled) return;
      if (!result.ok || !Array.isArray(result.documents)) {
        setStatus((result && result.error) || 'Could not open prompt files', 'err');
        return;
      }
      const opened = [];
      let failed = 0;
      for (const documentResult of result.documents) {
        if (!documentResult.ok) {
          failed++;
          continue;
        }
        const existing = state.documents.find(function (document) {
          return document.filePath &&
            normalizedPath(document.filePath) === normalizedPath(documentResult.filePath);
        }) || opened.find(function (document) {
          return normalizedPath(document.filePath) === normalizedPath(documentResult.filePath);
        });
        if (existing) {
          opened.push(existing);
          continue;
        }
        const data = documentDataFromResult(documentResult);
        if (!data) {
          failed++;
          continue;
        }
        opened.push({
          id: uid(),
          title: documentName(documentResult.filePath),
          filePath: documentResult.filePath,
          data: data
        });
      }

      if (opened.length === 0) {
        if (failed) setStatus('Could not open ' + failed + ' prompt file' + (failed === 1 ? '' : 's'), 'err');
        return;
      }

      captureCurrentDocument();
      if (state.documents.length === 1 &&
        state.documents[0].id === 'document-default' &&
        !state.documents[0].filePath &&
        isEmptyDocument(state.documents[0].data)) {
        state.documents = [];
      }
      opened.forEach(function (document) {
        if (!state.documents.some(function (current) { return current.id === document.id; })) {
          state.documents.push(document);
        }
      });
      state.activeDocumentId = opened[opened.length - 1].id;
      clearHistory();
      applyData(activeDocument().data);
      renderDocuments();
      scheduleSave();
      setStatus(
        'Opened ' + opened.length + ' prompt' + (opened.length === 1 ? '' : 's') +
        ' as tabs' + (failed ? ' (' + failed + ' failed)' : ''),
        failed ? 'err' : 'ok'
      );
    } catch (err) {
      console.error(err);
      setStatus('Could not open prompt file: ' + (err.message || 'Unknown error'), 'err');
    }
  }

  async function saveActiveDocument() {
    const document = activeDocument();
    if (!document) return;
    if (!document.filePath) {
      await saveActiveDocumentAs();
      return;
    }
    const saved = await flushCurrentDocument();
    if (!saved) return;
    setStatus('Saved ' + document.title, 'ok');
  }

  async function saveActiveDocumentAs() {
    const document = activeDocument();
    if (!document) return;
    try {
      const suggestedName = /\.c2copy$/i.test(document.title)
        ? document.title
        : document.title + '.c2copy';
      const result = await window.click2copy.saveDocumentAs(snapshot(), suggestedName);
      if (!result || result.canceled) return;
      if (!result.ok) {
        setStatus(result.error || 'Save failed', 'err');
        return;
      }
      document.filePath = result.filePath;
      document.title = documentName(result.filePath);
      document.data = snapshot();
      renderDocuments();
      const saved = await flushCurrentDocument();
      if (!saved) return;
      setStatus('Saved ' + document.title, 'ok');
    } catch (err) {
      console.error(err);
      setStatus('Save failed', 'err');
    }
  }

  if (window.click2copy && window.click2copy.onMenuAction) {
    window.click2copy.onMenuAction(function (action) {
      if (action === 'new-document') createDocument();
      else if (action === 'open-documents') openDocuments();
      else if (action === 'save-document') saveActiveDocument();
      else if (action === 'save-document-as') saveActiveDocumentAs();
    });
  }

  async function closeDocument(id) {
    const document = state.documents.find(function (item) { return item.id === id; });
    if (!document) return;
    if (!document.filePath && !window.confirm('Close this unsaved prompt?')) return;
    await flushCurrentDocument();
    const index = state.documents.indexOf(document);
    state.documents.splice(index, 1);
    if (state.documents.length === 0) {
      const data = makeDefaultData();
      data.activeTabId = defaultActiveTabId(data.tabs) || data.tabs[0].id;
      const replacement = { id: uid(), title: 'Untitled', filePath: null, data: data };
      state.documents.push(replacement);
      state.activeDocumentId = replacement.id;
      clearHistory();
      applyData(replacement.data);
      renderDocuments();
      scheduleSave();
      return;
    }
    if (state.activeDocumentId === id) {
      const next = state.documents[Math.min(index, state.documents.length - 1)];
      state.activeDocumentId = next.id;
      clearHistory();
      applyData(next.data);
    }
    renderDocuments();
    scheduleSave();
  }

  function renderTabs() {
    el.tabBar.innerHTML = '';
    state.tabs.forEach(function (tab) {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'tab' + (tab.id === state.activeTabId ? ' active' : '');
      btn.textContent = tab.title;
      btn.setAttribute('role', 'tab');
      btn.setAttribute('aria-selected', tab.id === state.activeTabId ? 'true' : 'false');
      btn.dataset.id = tab.id;
      btn.draggable = !isMasterTab(tab);
      if (isMasterTab(tab)) btn.classList.add('master-tab');
      btn.title = isMasterTab(tab)
        ? 'Master parts — edit reusable text here'
        : tab.title + ' (drag to reorder; double-click to rename)';

      btn.addEventListener('click', function () {
        selectTab(tab.id);
      });
      btn.addEventListener('dragstart', function (event) {
        event.dataTransfer.setData('text/plain', tab.id);
        event.dataTransfer.effectAllowed = 'move';
        btn.classList.add('dragging');
      });
      btn.addEventListener('dragend', function () {
        btn.classList.remove('dragging');
        el.tabBar.querySelectorAll('.drop-before, .drop-after').forEach(function (item) {
          item.classList.remove('drop-before', 'drop-after');
        });
      });
      btn.addEventListener('dragover', function (event) {
        event.preventDefault();
        event.dataTransfer.dropEffect = 'move';
        el.tabBar.querySelectorAll('.drop-before, .drop-after').forEach(function (item) {
          item.classList.remove('drop-before', 'drop-after');
        });
        const bounds = btn.getBoundingClientRect();
        btn.classList.add(event.clientX < bounds.left + bounds.width / 2 ? 'drop-before' : 'drop-after');
      });
      btn.addEventListener('dragleave', function () {
        btn.classList.remove('drop-before', 'drop-after');
      });
      btn.addEventListener('drop', function (event) {
        event.preventDefault();
        const draggedId = event.dataTransfer.getData('text/plain');
        const bounds = btn.getBoundingClientRect();
        reorderTab(draggedId, tab.id, event.clientX >= bounds.left + bounds.width / 2);
      });
      btn.addEventListener('dblclick', function (e) {
        e.preventDefault();
        e.stopPropagation();
        if (isMasterTab(tab)) return;
        startInlineRename(btn, tab);
      });
      btn.addEventListener('contextmenu', function (e) {
        e.preventDefault();
        if (isMasterTab(tab)) return;
        const action = window.prompt('Tab actions — type: rename | delete', 'rename');
        if (!action) return;
        const a = action.trim().toLowerCase();
        if (a === 'rename') startInlineRename(btn, tab);
        else if (a === 'delete') deleteTab(tab.id);
      });

      el.tabBar.appendChild(btn);
    });

    // Keep the add-part control in the tab strip, immediately after the last part tab.
    el.tabBar.appendChild(el.btnAdd);

    const tab = activeTab();
    el.partLabel.textContent = tab
      ? tab.title + ' (' + tab.cols + '×' + tab.rows + ')'
      : 'Part content';
    el.btnDelete.disabled = isMasterTab(tab) || partTabs().length <= 1;
    el.btnRename.disabled = isMasterTab(tab);
    el.masterLibrary.hidden = isMasterTab(tab);
  }

  function renderMasterLibrary() {
    el.masterLibraryItems.innerHTML = '';
    const master = state.tabs.find(isMasterTab);
    const current = activeTab();
    if (!master || !current || isMasterTab(current)) return;

    let hasContent = false;
    for (let i = 0; i < master.cells.length; i++) {
      if (master.cells[i] && String(master.cells[i]).trim()) {
        hasContent = true;
        break;
      }
    }
    if (!hasContent) {
      const empty = document.createElement('span');
      empty.className = 'master-library-empty';
      empty.textContent = 'Add reusable text in the Master part first.';
      el.masterLibraryItems.appendChild(empty);
      return;
    }

    const grid = document.createElement('div');
    grid.className = 'master-library-grid';
    grid.setAttribute('role', 'grid');
    // Mirror Master tab column widths (shared tab.columnWidths state).
    grid.style.gridTemplateColumns = contentColumnTemplate(master, 0);

    for (let row = 0; row < master.rows; row++) {
      let rowHasContent = false;
      for (let col = 0; col < master.cols; col++) {
        const cellText = master.cells[row * master.cols + col];
        if (cellText && String(cellText).trim()) {
          rowHasContent = true;
          break;
        }
      }
      if (!rowHasContent) continue;

      for (let col = 0; col < master.cols; col++) {
        const text = master.cells[row * master.cols + col] || '';
        const trimmed = String(text).trim();
        if (trimmed) {
          const button = document.createElement('button');
          button.type = 'button';
          button.className = 'master-library-cell';
          const masterIdx = row * master.cols + col;
          if (isMasterCellRepresentedInCombined(master, masterIdx)) button.classList.add('cell-confirmed');
          button.textContent = text;
          button.title = 'Add to the selected cell, or next empty cell, in ' + current.title;
          button.setAttribute('role', 'gridcell');
          button.dataset.row = String(row);
          button.dataset.col = String(col);
          button.addEventListener('click', function () {
            insertMasterText(text);
          });
          grid.appendChild(button);
        } else {
          const blank = document.createElement('div');
          blank.className = 'master-library-cell master-library-cell-empty';
          blank.setAttribute('role', 'gridcell');
          blank.setAttribute('aria-hidden', 'true');
          blank.dataset.row = String(row);
          blank.dataset.col = String(col);
          grid.appendChild(blank);
        }
      }
    }

    el.masterLibraryItems.appendChild(grid);
  }

  function insertMasterText(text) {
    const tab = activeTab();
    if (!tab || isMasterTab(tab) || tab.cells.length === 0) return;

    let index = -1;
    let usedFirstCellFallback = false;
    if (focusedCell && focusedCell.tabId === tab.id &&
        Number.isInteger(focusedCell.index) &&
        focusedCell.index >= 0 && focusedCell.index < tab.cells.length) {
      index = focusedCell.index;
    } else {
      index = tab.cells.findIndex(function (cell) {
        return !cell || !cell.trim();
      });
      if (index === -1) {
        index = 0;
        usedFirstCellFallback = true;
      }
    }

    pushHistory();
    tab.cells[index] = text;
    focusedCell = { tabId: tab.id, index: index };
    revalidateLinksForCell(tab.id, index, { silent: true });
    renderTabs();
    renderGrid();
    renderMasterLibrary();
    renderCombinedPrompt();
    scheduleSave();
    const cell = el.cellGrid.querySelector('textarea.cell[data-idx="' + index + '"]');
    if (cell) cell.focus();
    setStatus(
      usedFirstCellFallback
        ? 'Added Master text to ' + tab.title + ' (replaced the first cell; no empty cells)'
        : 'Added Master text to ' + tab.title,
      'ok'
    );
  }

  function reorderTab(draggedId, targetId, afterTarget) {
    if (!draggedId || draggedId === targetId) return;
    const fromIndex = state.tabs.findIndex(function (tab) { return tab.id === draggedId; });
    const targetIndex = state.tabs.findIndex(function (tab) { return tab.id === targetId; });
    if (fromIndex <= 0 || targetIndex <= 0) return;

    pushHistory();
    const moved = state.tabs.splice(fromIndex, 1)[0];
    let insertAt = state.tabs.findIndex(function (tab) { return tab.id === targetId; });
    if (afterTarget) insertAt++;
    state.tabs.splice(insertAt, 0, moved);
    renderTabs();
    scheduleSave();
  }

  function reorderDocument(draggedId, targetId, afterTarget) {
    if (!draggedId || draggedId === targetId) return;
    const fromIndex = state.documents.findIndex(function (item) { return item.id === draggedId; });
    const targetIndex = state.documents.findIndex(function (item) { return item.id === targetId; });
    if (fromIndex < 0 || targetIndex < 0) return;

    const moved = state.documents.splice(fromIndex, 1)[0];
    let insertAt = state.documents.findIndex(function (item) { return item.id === targetId; });
    if (afterTarget) insertAt++;
    state.documents.splice(insertAt, 0, moved);
    renderDocuments();
    scheduleSave();
  }

  function currentColumnWidths(tab) {
    const headers = el.cellGrid.querySelectorAll('.column-header');
    if (headers.length !== tab.cols) return null;
    return Array.prototype.map.call(headers, function (header) {
      return Math.max(MIN_COLUMN_WIDTH, header.getBoundingClientRect().width);
    });
  }

  function applyGridColumns(tab) {
    // Row gutter (40px) + shared content widths from tab.columnWidths.
    el.cellGrid.style.gridTemplateColumns = '40px ' + contentColumnTemplate(tab, 120);
  }


  function uniqueCol1Values(tab) {
    const seen = Object.create(null);
    const values = [];
    if (!tab || !tab.cols) return values;
    for (let r = 0; r < tab.rows; r++) {
      const v = (tab.cells[r * tab.cols] || '').trim();
      if (seen[v]) continue;
      seen[v] = true;
      values.push(v);
    }
    values.sort(function (a, b) {
      if (!a && !b) return 0;
      if (!a) return 1;
      if (!b) return -1;
      return a.localeCompare(b, undefined, { sensitivity: 'base', numeric: true });
    });
    return values;
  }

  function rowMatchesCol1ValueFilter(tab, rowIndex) {
    if (col1ValueFilter === null) return true;
    const v = (tab.cells[rowIndex * tab.cols] || '').trim();
    return col1ValueFilter.has(v);
  }

  function rowMatchesFilter(tab, rowIndex) {
    let base = true;
    if (gridRowFilter === 'nonempty') base = !rowIsEmpty(tab, rowIndex);
    else if (gridRowFilter === 'included') base = isRowIncluded(tab, rowIndex);
    if (!base) return false;
    return rowMatchesCol1ValueFilter(tab, rowIndex);
  }

  function applyRowFilterVisibility() {
    const tab = activeTab();
    if (!tab || !el.cellGrid) return;
    const rowButtons = el.cellGrid.querySelectorAll('.row-append');
    for (let i = 0; i < rowButtons.length; i++) {
      const btn = rowButtons[i];
      const row = parseInt(btn.dataset.row, 10);
      if (Number.isNaN(row)) continue;
      const hidden = !rowMatchesFilter(tab, row);
      const controls = btn.closest('.row-controls');
      if (controls) controls.classList.toggle('is-row-filtered', hidden);
      const wraps = el.cellGrid.querySelectorAll('.cell-wrap[data-row="' + row + '"]');
      for (let w = 0; w < wraps.length; w++) wraps[w].classList.toggle('is-row-filtered', hidden);
    }
  }

  function syncRowFilterButtons() {
    const buttons = [el.btnFilterAll, el.btnFilterNonempty, el.btnFilterIncluded];
    for (let i = 0; i < buttons.length; i++) {
      const btn = buttons[i];
      if (!btn) continue;
      const on = btn.dataset.filter === gridRowFilter;
      btn.classList.toggle('is-active', on);
      btn.setAttribute('aria-pressed', on ? 'true' : 'false');
    }
  }

  function setGridRowFilter(mode) {
    if (mode !== 'all' && mode !== 'nonempty' && mode !== 'included') return;
    if (gridRowFilter === mode) return;
    gridRowFilter = mode;
    syncRowFilterButtons();
    renderGrid();
    const labels = { all: 'Showing all rows', nonempty: 'Showing non-empty rows', included: 'Showing rows in Combined' };
    setStatus(labels[mode] || 'Row filter updated', 'ok');
  }

  function isCol1ValueSelected(value) {
    if (col1ValueFilter === null) return true;
    return col1ValueFilter.has(value);
  }

  function setCol1ValueFilterSelection(selectedValues, allValues) {
    const selected = [];
    const seen = Object.create(null);
    for (let i = 0; i < selectedValues.length; i++) {
      const v = selectedValues[i];
      if (Object.prototype.hasOwnProperty.call(seen, v)) continue;
      seen[v] = true;
      selected.push(v);
    }
    let allOn = selected.length === allValues.length;
    if (allOn) {
      for (let i = 0; i < allValues.length; i++) {
        if (!Object.prototype.hasOwnProperty.call(seen, allValues[i])) {
          allOn = false;
          break;
        }
      }
    }
    col1ValueFilter = allOn ? null : new Set(selected);
    applyRowFilterVisibility();
    syncCol1FilterControls();
  }

  function positionCol1FilterMenu(btn, menu) {
    if (!btn || !menu) return;
    const rect = btn.getBoundingClientRect();
    const pad = 8;
    const width = Math.max(220, Math.min(320, window.innerWidth - pad * 2));
    let left = rect.left;
    if (left + width > window.innerWidth - pad) left = Math.max(pad, window.innerWidth - pad - width);
    if (left < pad) left = pad;
    menu.style.position = 'fixed';
    menu.style.top = Math.round(rect.bottom + 4) + 'px';
    menu.style.left = Math.round(left) + 'px';
    menu.style.width = width + 'px';
    menu.style.zIndex = '60';
  }

  function syncCol1FilterControls() {
    const btn = el.cellGrid && el.cellGrid.querySelector('.column-col1-filter-btn');
    if (!btn) return;
    const active = col1ValueFilter !== null;
    btn.classList.toggle('is-active', active);
    btn.setAttribute('aria-pressed', active ? 'true' : 'false');
    btn.setAttribute('aria-expanded', col1FilterMenuOpen ? 'true' : 'false');
    const count = col1ValueFilter ? col1ValueFilter.size : 0;
    btn.title = active
      ? ('Column 1 value filter on (' + count + ' selected) — click to change')
      : 'Filter rows by Column 1 values';
    btn.setAttribute('aria-label', btn.title);
    const menu = el.cellGrid.querySelector('.column-col1-filter-menu');
    if (menu) {
      menu.hidden = !col1FilterMenuOpen;
      if (col1FilterMenuOpen) positionCol1FilterMenu(btn, menu);
    }
  }

  function closeCol1FilterMenu() {
    if (!col1FilterMenuOpen) return;
    col1FilterMenuOpen = false;
    syncCol1FilterControls();
  }

  function buildCol1FilterMenu(tab, menu) {
    menu.innerHTML = '';
    const values = uniqueCol1Values(tab);

    const actions = document.createElement('div');
    actions.className = 'column-col1-filter-actions';

    const selectAll = document.createElement('button');
    selectAll.type = 'button';
    selectAll.className = 'column-col1-filter-action';
    selectAll.textContent = 'Select all';
    selectAll.addEventListener('click', function (e) {
      e.preventDefault();
      e.stopPropagation();
      col1ValueFilter = null;
      buildCol1FilterMenu(tab, menu);
      applyRowFilterVisibility();
      syncCol1FilterControls();
      setStatus('Showing all Column 1 values', 'ok');
    });

    const clearBtn = document.createElement('button');
    clearBtn.type = 'button';
    clearBtn.className = 'column-col1-filter-action';
    clearBtn.textContent = 'Clear';
    clearBtn.addEventListener('click', function (e) {
      e.preventDefault();
      e.stopPropagation();
      col1ValueFilter = new Set();
      buildCol1FilterMenu(tab, menu);
      applyRowFilterVisibility();
      syncCol1FilterControls();
      setStatus('Column 1 value filter cleared (no rows)', 'ok');
    });

    actions.appendChild(selectAll);
    actions.appendChild(clearBtn);
    menu.appendChild(actions);

    const list = document.createElement('div');
    list.className = 'column-col1-filter-list';
    list.setAttribute('role', 'group');
    list.setAttribute('aria-label', 'Column 1 values');

    if (!values.length) {
      const empty = document.createElement('div');
      empty.className = 'column-col1-filter-empty';
      empty.textContent = 'No values';
      list.appendChild(empty);
    } else {
      for (let i = 0; i < values.length; i++) {
        const value = values[i];
        const label = document.createElement('label');
        label.className = 'column-col1-filter-option';

        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.checked = isCol1ValueSelected(value);
        if (value) cb.dataset.value = value;
        else cb.dataset.blank = '1';
        cb.addEventListener('click', function (e) {
          e.stopPropagation();
        });
        cb.addEventListener('change', function () {
          const selected = [];
          const boxes = list.querySelectorAll('input[type="checkbox"]');
          for (let b = 0; b < boxes.length; b++) {
            if (!boxes[b].checked) continue;
            selected.push(boxes[b].dataset.blank === '1' ? '' : (boxes[b].dataset.value || ''));
          }
          setCol1ValueFilterSelection(selected, values);
          const n = col1ValueFilter === null ? values.length : col1ValueFilter.size;
          setStatus(
            col1ValueFilter === null
              ? 'Showing all Column 1 values'
              : ('Showing ' + n + ' Column 1 value' + (n === 1 ? '' : 's')),
            'ok'
          );
        });

        const text = document.createElement('span');
        text.className = 'column-col1-filter-option-text';
        text.textContent = value ? value : '(blank)';
        if (!value) text.classList.add('is-blank');

        label.appendChild(cb);
        label.appendChild(text);
        list.appendChild(label);
      }
    }

    menu.appendChild(list);
  }

  function toggleCol1FilterMenu(tab, btn, menu) {
    col1FilterMenuOpen = !col1FilterMenuOpen;
    if (col1FilterMenuOpen) {
      buildCol1FilterMenu(tab, menu);
      positionCol1FilterMenu(btn, menu);
    }
    syncCol1FilterControls();
  }

  function renderGrid() {
    const tab = activeTab();
    el.cellGrid.innerHTML = '';
    if (!tab) return;

    // The first grid row is a header row. Its handles resize the matching cell column.
    applyGridColumns(tab);

    const corner = document.createElement('div');
    corner.className = 'grid-corner';
    corner.setAttribute('aria-hidden', 'true');
    el.cellGrid.appendChild(corner);

    for (let c = 0; c < tab.cols; c++) {
      const header = document.createElement('div');
      header.className = 'column-header';
      header.dataset.col = String(c);
      header.setAttribute('role', 'columnheader');

      const label = document.createElement('span');
      label.className = 'column-header-label';
      label.textContent = 'Column ' + (c + 1);
      header.appendChild(label);

      if (c === 0) {
        header.classList.add('column-header-sortable');
        const sortControls = document.createElement('span');
        sortControls.className = 'column-sort-controls';
        sortControls.setAttribute('role', 'group');
        sortControls.setAttribute('aria-label', 'Sort by column 1');

        const sortAsc = document.createElement('button');
        sortAsc.type = 'button';
        sortAsc.className = 'column-sort-btn';
        sortAsc.textContent = 'A–Z';
        sortAsc.title = 'Sort rows by column 1 A–Z';
        sortAsc.setAttribute('aria-label', sortAsc.title);
        sortAsc.addEventListener('click', function (e) {
          e.preventDefault();
          e.stopPropagation();
          sortRowsByColumn(0, 'asc');
        });

        const sortDesc = document.createElement('button');
        sortDesc.type = 'button';
        sortDesc.className = 'column-sort-btn';
        sortDesc.textContent = 'Z–A';
        sortDesc.title = 'Sort rows by column 1 Z–A';
        sortDesc.setAttribute('aria-label', sortDesc.title);
        sortDesc.addEventListener('click', function (e) {
          e.preventDefault();
          e.stopPropagation();
          sortRowsByColumn(0, 'desc');
        });

        sortControls.appendChild(sortAsc);
        sortControls.appendChild(sortDesc);
        header.appendChild(sortControls);

        const filterWrap = document.createElement('span');
        filterWrap.className = 'column-col1-filter';

        const filterBtn = document.createElement('button');
        filterBtn.type = 'button';
        filterBtn.className = 'column-sort-btn column-col1-filter-btn';
        filterBtn.textContent = 'Values';
        filterBtn.title = 'Filter rows by Column 1 values';
        filterBtn.setAttribute('aria-label', filterBtn.title);
        filterBtn.setAttribute('aria-haspopup', 'true');
        filterBtn.setAttribute('aria-expanded', 'false');
        filterBtn.setAttribute('aria-pressed', 'false');

        const filterMenu = document.createElement('div');
        filterMenu.className = 'column-col1-filter-menu';
        filterMenu.hidden = true;
        filterMenu.setAttribute('role', 'dialog');
        filterMenu.setAttribute('aria-label', 'Column 1 value filter');

        filterBtn.addEventListener('click', function (e) {
          e.preventDefault();
          e.stopPropagation();
          toggleCol1FilterMenu(tab, filterBtn, filterMenu);
        });
        filterMenu.addEventListener('click', function (e) {
          e.stopPropagation();
        });

        filterWrap.appendChild(filterBtn);
        filterWrap.appendChild(filterMenu);
        header.appendChild(filterWrap);
        header.classList.add('column-header-filterable');
      }

      const handle = document.createElement('span');
      handle.className = 'column-resize-handle';
      handle.setAttribute('role', 'separator');
      handle.setAttribute('aria-orientation', 'vertical');
      handle.setAttribute('aria-label', 'Resize column ' + (c + 1));
      handle.title = 'Drag to resize column ' + (c + 1);
      handle.tabIndex = 0;
      handle.addEventListener('pointerdown', function (e) {
        beginColumnResize(e, c, handle);
      });
      handle.addEventListener('keydown', function (e) {
        if (e.key !== 'ArrowLeft' && e.key !== 'ArrowRight') return;
        e.preventDefault();
        resizeColumnByKeyboard(tab, c, e.key === 'ArrowRight' ? 16 : -16);
      });
      header.appendChild(handle);
      el.cellGrid.appendChild(header);
    }

    for (let r = 0; r < tab.rows; r++) {
      const rowControls = document.createElement('div');
      rowControls.className = 'row-controls';

      const rowBtn = document.createElement('button');
      rowBtn.type = 'button';
      rowBtn.className = 'row-append';
      rowBtn.dataset.row = String(r);
      rowBtn.setAttribute('aria-pressed', 'false');
      rowBtn.innerHTML = '<span class="row-append-mark" aria-hidden="true"></span>';
      rowBtn.addEventListener('pointerdown', function (ev) {
        if (ev.button != null && ev.button !== 0) return;
        ev.preventDefault();
        ev.stopPropagation();
        const wantOn = !isRowIncluded(tab, r);
        beginCheckboxDrag('row', r, wantOn, ev);
      });
      rowBtn.addEventListener('click', function (ev) {
        // Handled on pointerdown for click+drag paint; suppress leftover click.
        ev.preventDefault();
        ev.stopPropagation();
      });
      rowControls.appendChild(rowBtn);

      const moveControls = document.createElement('span');
      moveControls.className = 'row-move-controls';

      const moveUp = document.createElement('button');
      moveUp.type = 'button';
      moveUp.className = 'row-move';
      moveUp.textContent = '↑';
      moveUp.title = 'Move row ' + (r + 1) + ' up';
      moveUp.setAttribute('aria-label', moveUp.title);
      moveUp.disabled = r === 0;
      moveUp.addEventListener('click', function () {
        moveRow(r, -1);
      });
      moveControls.appendChild(moveUp);

      const moveDown = document.createElement('button');
      moveDown.type = 'button';
      moveDown.className = 'row-move';
      moveDown.textContent = '↓';
      moveDown.title = 'Move row ' + (r + 1) + ' down';
      moveDown.setAttribute('aria-label', moveDown.title);
      moveDown.addEventListener('click', function () {
        moveRow(r, 1);
      });
      moveControls.appendChild(moveDown);

      rowControls.appendChild(moveControls);
      el.cellGrid.appendChild(rowControls);

      for (let c = 0; c < tab.cols; c++) {
        const idx = r * tab.cols + c;
        const wrap = document.createElement('div');
        wrap.className = 'cell-wrap';
        wrap.dataset.row = String(r);
        wrap.dataset.col = String(c);
        if (isCellConfirmed(tab.id, idx)) wrap.classList.add('cell-confirmed');
        if (isCellSlept(tab, idx)) wrap.classList.add('is-slept');

        const gutter = document.createElement('div');
        gutter.className = 'cell-gutter';

        const sleepLabel = document.createElement('label');
        sleepLabel.className = 'cell-sleep';
        sleepLabel.title = 'Sleep cell — exclude from Append all active';
        const sleepCb = document.createElement('input');
        sleepCb.type = 'checkbox';
        sleepCb.className = 'cell-sleep-input';
        sleepCb.checked = isCellSlept(tab, idx);
        sleepCb.setAttribute('aria-label', 'Sleep row ' + (r + 1) + ' column ' + (c + 1));
        sleepCb.addEventListener('click', function (ev) {
          ev.stopPropagation();
        });
        sleepCb.addEventListener('pointerdown', function (ev) {
          if (ev.button != null && ev.button !== 0) return;
          ev.stopPropagation();
          const wantSlept = !sleepCb.checked;
          beginCheckboxDrag('sleep', idx, wantSlept, ev);
          // Prevent the native click toggle; drag paint owns the value.
          ev.preventDefault();
        });
        sleepCb.addEventListener('change', function (ev) {
          // Ignored during/after pointerdown paint; keep as keyboard fallback.
          if (checkboxDrag) {
            ev.preventDefault();
            return;
          }
          toggleCellSleep(idx, sleepCb.checked);
        });
        const sleepZzz = document.createElement('span');
        sleepZzz.className = 'cell-sleep-zzz';
        sleepZzz.setAttribute('aria-hidden', 'true');
        sleepZzz.textContent = 'Zzz';
        sleepLabel.appendChild(sleepCb);
        sleepLabel.appendChild(sleepZzz);
        gutter.appendChild(sleepLabel);

        const cellBtn = document.createElement('button');
        cellBtn.type = 'button';
        cellBtn.className = 'cell-append';
        cellBtn.dataset.idx = String(idx);
        cellBtn.dataset.row = String(r);
        cellBtn.dataset.col = String(c);
        cellBtn.setAttribute('aria-pressed', 'false');
        cellBtn.innerHTML = '<span class="cell-append-mark" aria-hidden="true"></span>';
        cellBtn.addEventListener('pointerdown', function (ev) {
          if (ev.button != null && ev.button !== 0) return;
          ev.preventDefault();
          ev.stopPropagation();
          const wantOn = !isCellConfirmedInScope(tab.id, idx, currentPromptScope());
          beginCheckboxDrag('combined', idx, wantOn, ev);
        });
        cellBtn.addEventListener('click', function (ev) {
          ev.preventDefault();
          ev.stopPropagation();
        });
        gutter.appendChild(cellBtn);
        wrap.appendChild(gutter);

        const ta = document.createElement('textarea');
        ta.className = 'cell';
        if (isCellConfirmed(tab.id, idx)) ta.classList.add('cell-confirmed');
        ta.rows = 2;
        ta.spellcheck = false;
        ta.placeholder = 'R' + (r + 1) + 'C' + (c + 1);
        ta.value = tab.cells[idx] || '';
        ta.dataset.row = String(r);
        ta.dataset.col = String(c);
        ta.dataset.idx = String(idx);
        ta.setAttribute('aria-label', 'Row ' + (r + 1) + ' column ' + (c + 1));
        ta.addEventListener('input', onCellInput);
        ta.addEventListener('keydown', onCellKeydown);
        ta.addEventListener('pointerdown', onCellPointerDownSelect);
        ta.addEventListener('focus', onCellFocusSelect);
        ta.addEventListener('click', onCellClickSelect);
        ta.addEventListener('paste', onCellPaste);
        ta.addEventListener('copy', onCellCopy);
        ta.addEventListener('cut', onCellCut);
        wrap.appendChild(ta);
        el.cellGrid.appendChild(wrap);
      }

      const rowHidden = !rowMatchesFilter(tab, r);
      rowControls.classList.toggle('is-row-filtered', rowHidden);
      const rowWraps = el.cellGrid.querySelectorAll('.cell-wrap[data-row="' + r + '"]');
      for (let w = 0; w < rowWraps.length; w++) {
        rowWraps[w].classList.toggle('is-row-filtered', rowHidden);
      }
    }

    syncRowFilterButtons();
    applyAppendCheckedState();
    if (col1FilterMenuOpen) {
      const menu = el.cellGrid.querySelector('.column-col1-filter-menu');
      if (menu) buildCol1FilterMenu(tab, menu);
    }
    syncCol1FilterControls();
  }

  function resizeColumnByKeyboard(tab, col, delta) {
    const widths = Array.isArray(tab.columnWidths) && tab.columnWidths.length === tab.cols
      ? tab.columnWidths.slice()
      : currentColumnWidths(tab);
    if (!widths) return;
    pushHistory({ coalesce: true });
    const next = Math.max(MIN_COLUMN_WIDTH, widths[col] + delta);
    if (col < widths.length - 1) {
      const available = widths[col] + widths[col + 1] - MIN_COLUMN_WIDTH;
      widths[col] = Math.min(next, available);
      widths[col + 1] = available - widths[col] + MIN_COLUMN_WIDTH;
    } else {
      widths[col] = next;
    }
    tab.columnWidths = widths;
    applyGridColumns(tab);
    scheduleSave();
  }

  function beginColumnResize(e, col, handle) {
    const tab = activeTab();
    if (!tab) return;
    e.preventDefault();
    e.stopPropagation();

    const widths = currentColumnWidths(tab);
    if (!widths) return;
    pushHistory();
    const startX = e.clientX;
    const startWidths = widths.slice();
    const hasNext = col < tab.cols - 1;
    let moved = false;
    document.body.classList.add('resizing-columns');
    el.cellGrid.classList.add('resizing');

    function onMove(moveEvent) {
      const delta = moveEvent.clientX - startX;
      let nextWidth = Math.max(MIN_COLUMN_WIDTH, startWidths[col] + delta);
      const nextWidths = startWidths.slice();
      if (hasNext) {
        const pairWidth = startWidths[col] + startWidths[col + 1];
        nextWidth = Math.min(nextWidth, pairWidth - MIN_COLUMN_WIDTH);
        nextWidths[col + 1] = pairWidth - nextWidth;
      }
      nextWidths[col] = nextWidth;
      tab.columnWidths = nextWidths;
      applyGridColumns(tab);
      moved = moved || Math.abs(delta) > 1;
      scheduleSave();
    }

    function finish() {
      document.removeEventListener('pointermove', onMove);
      document.removeEventListener('pointerup', finish);
      document.removeEventListener('pointercancel', finish);
      document.body.classList.remove('resizing-columns');
      el.cellGrid.classList.remove('resizing');
      if (moved) setStatus('Column resized', 'ok');
    }

    document.addEventListener('pointermove', onMove);
    document.addEventListener('pointerup', finish);
    document.addEventListener('pointercancel', finish);
    if (handle.setPointerCapture) {
      try { handle.setPointerCapture(e.pointerId); } catch (err) { /* no-op */ }
    }
  }

  function resizeCombinedSection(height) {
    const available = el.combinedSection.parentElement.clientHeight;
    const minimum = 150;
    const maximum = Math.max(minimum, available - 230);
    const next = Math.max(minimum, Math.min(maximum, height));
    el.combinedSection.style.flex = '0 1 ' + next + 'px';
    try {
      localStorage.setItem('click2copy-combined-height', String(next));
    } catch (err) {
      console.warn('Could not save combined prompt height:', err);
    }
  }

  function beginSectionResize(event) {
    event.preventDefault();
    const startY = event.clientY;
    const startHeight = el.combinedSection.getBoundingClientRect().height;
    document.body.classList.add('resizing-sections');

    function onMove(moveEvent) {
      resizeCombinedSection(startHeight - (moveEvent.clientY - startY));
    }

    function finish() {
      document.removeEventListener('pointermove', onMove);
      document.removeEventListener('pointerup', finish);
      document.removeEventListener('pointercancel', finish);
      document.body.classList.remove('resizing-sections');
    }

    document.addEventListener('pointermove', onMove);
    document.addEventListener('pointerup', finish);
    document.addEventListener('pointercancel', finish);
  }

  function resizeCombinedSectionByKeyboard(delta) {
    resizeCombinedSection(el.combinedSection.getBoundingClientRect().height + delta);
  }

  function rememberFocusedCell(e) {
    const tab = activeTab();
    const idx = parseInt(e.currentTarget.dataset.idx, 10);
    if (!tab || Number.isNaN(idx) || idx < 0 || idx >= tab.cells.length) return;
    focusedCell = { tabId: tab.id, index: idx };
  }

  // Click2Copy: selecting/focusing a cell with content writes it to the clipboard.
  // Prefer click/focus (not input/keystrokes while typing). Guard against click+focus double-fire.
  // Click+drag across cells selects a rectangle; mouseup copies TSV (tabs/newlines like sheets).
  let lastAutoCopyKey = '';
  let lastAutoCopyAt = 0;
  let suppressCellAutoCopy = false;
  let cellRangeDrag = null;

  function writeTextToClipboard(text) {
    if (navigator.clipboard && typeof navigator.clipboard.writeText === 'function') {
      return navigator.clipboard.writeText(text);
    }
    return Promise.reject(new Error('Clipboard API unavailable'));
  }

  function copyTextWithStatus(text, okStatus) {
    const value = text == null ? '' : String(text);
    const statusOk = okStatus || 'Copied cell';
    writeTextToClipboard(value).then(function () {
      setStatus(statusOk, 'ok');
    }).catch(function () {
      const prev = document.activeElement;
      try {
        const helper = document.createElement('textarea');
        helper.value = value;
        helper.setAttribute('readonly', '');
        helper.style.position = 'fixed';
        helper.style.left = '-9999px';
        document.body.appendChild(helper);
        helper.select();
        const ok = document.execCommand('copy');
        document.body.removeChild(helper);
        if (prev && typeof prev.focus === 'function') {
          try { prev.focus(); } catch (err) { /* no-op */ }
        }
        if (ok) setStatus(statusOk, 'ok');
        else setStatus('Copy failed', 'err');
      } catch (err) {
        if (prev && typeof prev.focus === 'function') {
          try { prev.focus(); } catch (e2) { /* no-op */ }
        }
        setStatus('Copy failed', 'err');
      }
    });
  }

  function autoCopyCellToClipboard(ta) {
    if (suppressCellAutoCopy) return;
    if (!ta || ta.tagName !== 'TEXTAREA') return;
    const value = ta.value == null ? '' : String(ta.value);
    if (!value) return;
    const tab = activeTab();
    const idx = parseInt(ta.dataset.idx, 10);
    const key = (tab ? tab.id : '') + ':' + (Number.isNaN(idx) ? '' : idx) + ':' + value;
    const now = Date.now();
    if (key === lastAutoCopyKey && now - lastAutoCopyAt < 300) return;
    lastAutoCopyKey = key;
    lastAutoCopyAt = now;
    copyTextWithStatus(value, 'Copied cell');
  }

  function clearCellRangeHighlight() {
    if (!el.cellGrid) return;
    const nodes = el.cellGrid.querySelectorAll(
      '.cell-wrap.cell-range-selected, .cell-wrap.selected, .cell.cell-range-selected, .cell.selected'
    );
    for (let i = 0; i < nodes.length; i++) {
      nodes[i].classList.remove('cell-range-selected', 'selected');
    }
  }

  function normalizeRangeBounds(r0, c0, r1, c1) {
    return {
      rMin: Math.min(r0, r1),
      rMax: Math.max(r0, r1),
      cMin: Math.min(c0, c1),
      cMax: Math.max(c0, c1)
    };
  }

  function applyCellRangeHighlight(r0, c0, r1, c1) {
    clearCellRangeHighlight();
    if (!el.cellGrid) return;
    const b = normalizeRangeBounds(r0, c0, r1, c1);
    for (let r = b.rMin; r <= b.rMax; r++) {
      for (let c = b.cMin; c <= b.cMax; c++) {
        const wrap = el.cellGrid.querySelector(
          '.cell-wrap[data-row="' + r + '"][data-col="' + c + '"]'
        );
        if (!wrap) continue;
        // Mark every cell in the rectangle — not only the one under the pointer.
        wrap.classList.add('cell-range-selected', 'selected');
        const ta = wrap.querySelector('textarea.cell');
        if (ta) ta.classList.add('cell-range-selected', 'selected');
      }
    }
  }

  function buildCellRangeTsv(tab, r0, c0, r1, c1) {
    const b = normalizeRangeBounds(r0, c0, r1, c1);
    const lines = [];
    for (let r = b.rMin; r <= b.rMax; r++) {
      const cells = [];
      for (let c = b.cMin; c <= b.cMax; c++) {
        cells.push(tab.cells[r * tab.cols + c] || '');
      }
      lines.push(cells.join('\t'));
    }
    return lines.join('\n');
  }

  function cellWrapFromPoint(clientX, clientY) {
    const hit = document.elementFromPoint(clientX, clientY);
    if (!hit || typeof hit.closest !== 'function') return null;
    const wrap = hit.closest('.cell-wrap');
    if (!wrap || !el.cellGrid.contains(wrap)) return null;
    if (wrap.classList.contains('is-row-filtered')) return null;
    return wrap;
  }

  function finishCellRangeListeners() {
    document.removeEventListener('pointermove', onCellRangePointerMove);
    document.removeEventListener('pointerup', onCellRangePointerUp);
    document.removeEventListener('pointercancel', onCellRangePointerUp);
    document.removeEventListener('dragstart', onCellRangeDragStartPrevent, true);
    document.body.classList.remove('selecting-cell-range');
    if (cellRangeDrag && cellRangeDrag.captureEl && cellRangeDrag.pointerId != null) {
      try {
        if (cellRangeDrag.captureEl.releasePointerCapture) {
          cellRangeDrag.captureEl.releasePointerCapture(cellRangeDrag.pointerId);
        }
      } catch (err) { /* no-op */ }
    }
  }

  function onCellRangeDragStartPrevent(e) {
    // Native textarea text-drag would steal the gesture and leave only one cell lit.
    if (cellRangeDrag && cellRangeDrag.active) e.preventDefault();
  }

  function collapseCellTextSelection() {
    try {
      const sel = window.getSelection();
      if (sel && sel.removeAllRanges) sel.removeAllRanges();
    } catch (err) { /* no-op */ }
    const active = document.activeElement;
    if (active && active.tagName === 'TEXTAREA' && active.classList.contains('cell')) {
      try {
        const pos = typeof active.selectionEnd === 'number' ? active.selectionEnd : 0;
        active.setSelectionRange(pos, pos);
      } catch (err2) { /* no-op */ }
    }
  }

  function beginMultiCellRangeDrag() {
    if (!cellRangeDrag || cellRangeDrag.multi) return;
    cellRangeDrag.multi = true;
    suppressCellAutoCopy = true;
    document.body.classList.add('selecting-cell-range');
    collapseCellTextSelection();
  }

  function onCellRangePointerMove(e) {
    if (!cellRangeDrag || !cellRangeDrag.active) return;
    const wrap = cellWrapFromPoint(e.clientX, e.clientY);
    if (!wrap) return;
    const row = parseInt(wrap.dataset.row, 10);
    const col = parseInt(wrap.dataset.col, 10);
    if (Number.isNaN(row) || Number.isNaN(col)) return;
    if (row === cellRangeDrag.endRow && col === cellRangeDrag.endCol) return;
    cellRangeDrag.endRow = row;
    cellRangeDrag.endCol = col;
    const multi = row !== cellRangeDrag.startRow || col !== cellRangeDrag.startCol;
    if (!multi) {
      if (!cellRangeDrag.multi) return;
      applyCellRangeHighlight(
        cellRangeDrag.startRow, cellRangeDrag.startCol,
        cellRangeDrag.endRow, cellRangeDrag.endCol
      );
      return;
    }
    beginMultiCellRangeDrag();
    if (e.cancelable) e.preventDefault();
    // Full rectangle from anchor → pointer, every move — not only the hovered cell.
    applyCellRangeHighlight(
      cellRangeDrag.startRow, cellRangeDrag.startCol,
      cellRangeDrag.endRow, cellRangeDrag.endCol
    );
  }

  function onCellRangePointerUp(e) {
    if (!cellRangeDrag || !cellRangeDrag.active) return;
    const drag = cellRangeDrag;
    drag.active = false;
    finishCellRangeListeners();

    // Refresh end cell from release point when possible.
    const wrap = cellWrapFromPoint(e.clientX, e.clientY);
    if (wrap) {
      const row = parseInt(wrap.dataset.row, 10);
      const col = parseInt(wrap.dataset.col, 10);
      if (!Number.isNaN(row) && !Number.isNaN(col)) {
        drag.endRow = row;
        drag.endCol = col;
      }
    }

    const b = normalizeRangeBounds(drag.startRow, drag.startCol, drag.endRow, drag.endCol);
    const rows = b.rMax - b.rMin + 1;
    const cols = b.cMax - b.cMin + 1;
    const isMulti = rows > 1 || cols > 1;

    if (isMulti) {
      suppressCellAutoCopy = true;
      const tab = activeTab();
      if (tab) {
        const tsv = buildCellRangeTsv(tab, b.rMin, b.cMin, b.rMax, b.cMax);
        const label = rows + '\u00d7' + cols + ' cells';
        lastAutoCopyKey = (tab.id || '') + ':range:' + b.rMin + ',' + b.cMin + ':' + b.rMax + ',' + b.cMax;
        lastAutoCopyAt = Date.now();
        copyTextWithStatus(tsv, 'Copied ' + label);
      }
      applyCellRangeHighlight(b.rMin, b.cMin, b.rMax, b.cMax);
      window.setTimeout(function () {
        clearCellRangeHighlight();
        suppressCellAutoCopy = false;
        cellRangeDrag = null;
      }, 120);
      return;
    }

    clearCellRangeHighlight();
    suppressCellAutoCopy = false;
    cellRangeDrag = null;
  }

  function onCellPointerDownSelect(e) {
    if (e.button != null && e.button !== 0) return;
    const ta = e.currentTarget;
    const row = parseInt(ta.dataset.row, 10);
    const col = parseInt(ta.dataset.col, 10);
    if (Number.isNaN(row) || Number.isNaN(col)) return;
    // Re-click / fresh drag; do not disturb Combined/sleep (those are outside the textarea).
    if (cellRangeDrag && cellRangeDrag.active) {
      finishCellRangeListeners();
    }
    cellRangeDrag = {
      active: true,
      multi: false,
      startRow: row,
      startCol: col,
      endRow: row,
      endCol: col,
      pointerId: e.pointerId,
      captureEl: ta
    };
    document.addEventListener('pointermove', onCellRangePointerMove);
    document.addEventListener('pointerup', onCellRangePointerUp);
    document.addEventListener('pointercancel', onCellRangePointerUp);
    document.addEventListener('dragstart', onCellRangeDragStartPrevent, true);
    if (ta.setPointerCapture && e.pointerId != null) {
      try { ta.setPointerCapture(e.pointerId); } catch (err) { /* no-op */ }
    }
  }

  function onCellFocusSelect(e) {
    rememberFocusedCell(e);
    autoCopyCellToClipboard(e.currentTarget);
  }

  function onCellClickSelect(e) {
    rememberFocusedCell(e);
    // Prefer click when selecting a cell (also covers re-click of an already-focused cell).
    autoCopyCellToClipboard(e.currentTarget);
  }

  function onCellInput(e) {
    const tab = activeTab();
    if (!tab) return;
    rememberFocusedCell(e);
    const idx = parseInt(e.target.dataset.idx, 10);
    if (Number.isNaN(idx) || idx < 0 || idx >= tab.cells.length) return;
    pushHistory({ coalesce: true });
    tab.cells[idx] = e.target.value;
    revalidateLinksForCell(tab.id, idx);
    const confirmed = isCellConfirmed(tab.id, idx);
    e.target.classList.toggle('cell-confirmed', confirmed);
    const wrap = e.target.closest ? e.target.closest('.cell-wrap') : null;
    if (wrap) wrap.classList.toggle('cell-confirmed', confirmed);
    applyAppendCheckedState();
    if (isMasterTab(tab)) renderMasterLibrary();
    scheduleSave();
  }

  function focusCell(index) {
    const cell = el.cellGrid.querySelector('textarea.cell[data-idx="' + index + '"]');
    if (cell) cell.focus();
  }

  function moveToNextCell(currentCell, index) {
    // Explicitly finish the current edit before moving, like a spreadsheet.
    if (currentCell && typeof currentCell.blur === 'function') currentCell.blur();
    focusCell(index);
  }

  function onCellKeydown(e) {
    if (e.key !== 'Enter' || e.shiftKey || e.isComposing) return;

    const tab = activeTab();
    const idx = parseInt(e.currentTarget.dataset.idx, 10);
    if (!tab || Number.isNaN(idx) || idx < 0 || idx >= tab.cells.length) return;

    e.preventDefault();
    if (idx < tab.cells.length - 1) {
      moveToNextCell(e.currentTarget, idx + 1);
      return;
    }

    // Blur before re-rendering so the current edit is finished cleanly.
    e.currentTarget.blur();
    // Keep Enter row-major: append a row after the final cell, then focus it.
    addRow();
    focusCell(idx + 1);
  }

  function selectTab(id) {
    if (id === state.activeTabId) return;
    const tab = state.tabs.find(function (t) {
      return t.id === id;
    });
    if (!tab) return;
    state.activeTabId = id;
    focusedCell = null;
    renderTabs();
    renderGrid();
    renderMasterLibrary();
    renderCombinedPrompt();
    scheduleSave();
  }

  function startInlineRename(btn, tab) {
    if (btn.classList.contains('editing')) return;
    btn.classList.add('editing');
    btn.textContent = '';

    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'tab-rename-input';
    input.value = tab.title;
    input.setAttribute('aria-label', 'Rename tab');
    btn.appendChild(input);
    input.focus();
    input.select();

    let done = false;
    function finish(commit) {
      if (done) return;
      done = true;
      const next = input.value.trim();
      if (commit && next && next !== tab.title) {
        if (next.toLowerCase() === 'master') {
          setStatus('Only the Master library tab may be named Master', 'err');
        } else {
          pushHistory();
          tab.title = next;
          scheduleSave();
        }
      }
      renderTabs();
    }

    input.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') {
        e.preventDefault();
        finish(true);
      } else if (e.key === 'Escape') {
        e.preventDefault();
        finish(false);
      }
    });
    input.addEventListener('blur', function () {
      finish(true);
    });
  }

  function addTab() {
    pushHistory();
    const tab = makeTab(uid(), nextPartTitle());
    state.tabs.push(tab);
    state.activeTabId = tab.id;
    renderTabs();
    renderGrid();
    renderMasterLibrary();
    scheduleSave();
    setStatus('Tab added');
  }

  function renameActiveTab() {
    const tab = activeTab();
    if (!tab || isMasterTab(tab)) return;
    const btn = el.tabBar.querySelector('.tab.active');
    if (btn) {
      startInlineRename(btn, tab);
      return;
    }
    const next = window.prompt('Rename tab', tab.title);
    if (next == null) return;
    const trimmed = next.trim();
    if (!trimmed) return;
    if (trimmed.toLowerCase() === 'master') {
      setStatus('Only the Master library tab may be named Master', 'err');
      return;
    }
    pushHistory();
    tab.title = trimmed;
    renderTabs();
    scheduleSave();
  }

  function deleteTab(id) {
    const tab = state.tabs.find(function (item) { return item.id === id; });
    if (!tab) return;
    if (isMasterTab(tab)) {
      setStatus('The Master part cannot be deleted', 'err');
      return;
    }
    if (partTabs().length <= 1) {
      setStatus('Keep at least one part tab alongside Master', 'err');
      return;
    }

    if (tabHasContent(tab)) {
      const ok = window.confirm(
        'Delete "' + tab.title + '"? Its cell content and part-specific appended prompt will be lost.'
      );
      if (!ok) return;
    }

    const idx = state.tabs.findIndex(function (t) {
      return t.id === id;
    });
    pushHistory();
    state.tabs.splice(idx, 1);
    delete state.partPrompts[id];
    dropLinksForTab(id);
    if (state.activeTabId === id) {
      const next = state.tabs[Math.min(idx, state.tabs.length - 1)];
      state.activeTabId = next.id;
    }
    renderTabs();
    renderGrid();
    renderMasterLibrary();
    renderCombinedPrompt();
    scheduleSave();
    setStatus('Tab deleted');
  }


  /**
   * Parse an HTML clipboard fragment into a row-major string matrix.
   * Returns null when no usable <table> is present.
   */
  function parseHtmlTable(html) {
    if (!html || typeof html !== 'string') return null;
    let doc;
    try {
      doc = new DOMParser().parseFromString(html, 'text/html');
    } catch (err) {
      return null;
    }
    const table = doc.querySelector('table');
    if (!table) return null;
    const rows = [];
    const trs = table.querySelectorAll('tr');
    for (let i = 0; i < trs.length; i++) {
      const cells = [];
      const tds = trs[i].querySelectorAll('td, th');
      for (let j = 0; j < tds.length; j++) {
        const td = tds[j];
        const raw = td.innerText != null ? td.innerText : (td.textContent || '');
        cells.push(String(raw).replace(/\u00a0/g, ' ').replace(/\r\n/g, '\n').replace(/\r/g, '\n'));
        const colspan = parseInt(td.getAttribute('colspan'), 10);
        if (Number.isInteger(colspan) && colspan > 1) {
          for (let extra = 1; extra < colspan; extra++) cells.push('');
        }
      }
      if (cells.length) rows.push(cells);
    }
    if (!rows.length) return null;
    let maxCols = 0;
    for (let i = 0; i < rows.length; i++) {
      if (rows[i].length > maxCols) maxCols = rows[i].length;
    }
    for (let i = 0; i < rows.length; i++) {
      while (rows[i].length < maxCols) rows[i].push('');
    }
    return rows;
  }

  /**
   * Parse tab-separated / newline-separated plain text (Excel / Sheets text/plain).
   * A trailing newline from spreadsheet apps is stripped so it does not create a blank row.
   */
  function parseTsv(text) {
    if (text == null) return null;
    let body = String(text).replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    if (body.endsWith('\n')) body = body.slice(0, -1);
    if (body === '') return [['']];
    const lines = body.split('\n');
    const rows = lines.map(function (line) {
      return line.split('\t');
    });
    let maxCols = 0;
    for (let i = 0; i < rows.length; i++) {
      if (rows[i].length > maxCols) maxCols = rows[i].length;
    }
    for (let i = 0; i < rows.length; i++) {
      while (rows[i].length < maxCols) rows[i].push('');
    }
    return rows;
  }

  function matrixCellCount(matrix) {
    if (!matrix || !matrix.length) return 0;
    let n = 0;
    for (let i = 0; i < matrix.length; i++) n += matrix[i].length;
    return n;
  }

  function isMultiCellMatrix(matrix) {
    return !!(matrix && (matrix.length > 1 || (matrix[0] && matrix[0].length > 1)));
  }

  /**
   * Prefer HTML table (preserves structure from Sheets/Excel); fall back to TSV text/plain.
   */
  function parseClipboardMatrix(clipboardData) {
    if (!clipboardData) return null;
    const html = clipboardData.getData('text/html');
    const fromHtml = parseHtmlTable(html);
    if (fromHtml && isMultiCellMatrix(fromHtml)) return fromHtml;
    const plain = clipboardData.getData('text/plain');
    const fromTsv = parseTsv(plain);
    if (fromTsv && isMultiCellMatrix(fromTsv)) return fromTsv;
    // Single-cell HTML table with richer text than plain
    if (fromHtml && matrixCellCount(fromHtml) === 1) return fromHtml;
    return null;
  }

  function ensureTabSize(tab, rows, cols) {
    ensureSleptCells(tab);
    while (tab.cols < cols) {
      const oldCols = tab.cols;
      const newCells = [];
      for (let r = 0; r < tab.rows; r++) {
        for (let c = 0; c < tab.cols; c++) {
          newCells.push(tab.cells[r * tab.cols + c] || '');
        }
        newCells.push('');
      }
      if (Array.isArray(tab.columnWidths) && tab.columnWidths.length === tab.cols) {
        const defaultWidth = tab.columnWidths[tab.columnWidths.length - 1] || DEFAULT_COLUMN_WIDTH;
        tab.columnWidths = tab.columnWidths.concat([defaultWidth]);
      }
      tab.cols += 1;
      tab.cells = newCells;
      remapSleptAfterColumnAdd(tab, oldCols, tab.cols);
      remapConfirmedAfterColumnAdd(tab.id, oldCols, tab.cols);
    }
    while (tab.rows < rows) {
      for (let c = 0; c < tab.cols; c++) {
        tab.cells.push('');
        tab.sleptCells.push(false);
      }
      tab.rows += 1;
    }
  }

  function pasteMatrixAt(tab, startRow, startCol, matrix) {
    if (!tab || !matrix || !matrix.length) return null;
    const pasteRows = matrix.length;
    const pasteCols = matrix[0].length;
    ensureTabSize(tab, startRow + pasteRows, startCol + pasteCols);
    for (let r = 0; r < pasteRows; r++) {
      for (let c = 0; c < pasteCols; c++) {
        const idx = (startRow + r) * tab.cols + (startCol + c);
        tab.cells[idx] = matrix[r][c] == null ? '' : String(matrix[r][c]);
        revalidateLinksForCell(tab.id, idx, { silent: true });
      }
    }
    renderCombinedPrompt();
    applyConfirmedCellHighlights();
    return {
      rows: pasteRows,
      cols: pasteCols,
      endRow: startRow + pasteRows - 1,
      endCol: startCol + pasteCols - 1
    };
  }

  function onCellPaste(e) {
    const tab = activeTab();
    if (!tab) return;
    const startIdx = parseInt(e.currentTarget.dataset.idx, 10);
    if (Number.isNaN(startIdx) || startIdx < 0 || startIdx >= tab.cells.length) return;

    const matrix = parseClipboardMatrix(e.clipboardData);
    if (!matrix || !isMultiCellMatrix(matrix)) {
      // Single-cell / plain text: let the textarea handle a normal paste.
      return;
    }

    e.preventDefault();
    pushHistory();
    const startRow = Math.floor(startIdx / tab.cols);
    const startCol = startIdx % tab.cols;
    const result = pasteMatrixAt(tab, startRow, startCol, matrix);
    if (!result) return;

    focusedCell = { tabId: tab.id, index: startIdx };
    renderTabs();
    renderGrid();
    renderMasterLibrary();
    scheduleSave();
    const cell = el.cellGrid.querySelector('textarea.cell[data-idx="' + startIdx + '"]');
    if (cell) cell.focus();
    setStatus(
      'Pasted ' + result.rows + '×' + result.cols + ' cells at R' + (startRow + 1) +
      'C' + (startCol + 1) + ' (' + tab.cols + '×' + tab.rows + ' grid)',
      'ok'
    );
  }

  function onCellCopy(e) {
    const ta = e.currentTarget;
    if (!ta || ta.tagName !== 'TEXTAREA') return;
    // If the user highlighted a substring, keep the browser's default copy.
    if (typeof ta.selectionStart === 'number' && typeof ta.selectionEnd === 'number' &&
        ta.selectionStart !== ta.selectionEnd) {
      return;
    }
    const value = ta.value == null ? '' : String(ta.value);
    if (!e.clipboardData) return;
    e.preventDefault();
    e.clipboardData.setData('text/plain', value);
    setStatus('Copied cell', 'ok');
  }

  function onCellCut(e) {
    const tab = activeTab();
    const ta = e.currentTarget;
    if (!tab || !ta || ta.tagName !== 'TEXTAREA') return;
    if (typeof ta.selectionStart === 'number' && typeof ta.selectionEnd === 'number' &&
        ta.selectionStart !== ta.selectionEnd) {
      return;
    }
    const value = ta.value == null ? '' : String(ta.value);
    if (!e.clipboardData) return;
    e.preventDefault();
    e.clipboardData.setData('text/plain', value);
    const idx = parseInt(ta.dataset.idx, 10);
    if (!Number.isNaN(idx) && idx >= 0 && idx < tab.cells.length) {
      pushHistory();
      tab.cells[idx] = '';
      ta.value = '';
      revalidateLinksForCell(tab.id, idx);
      const confirmed = isCellConfirmed(tab.id, idx);
      ta.classList.toggle('cell-confirmed', confirmed);
      const wrap = ta.closest ? ta.closest('.cell-wrap') : null;
      if (wrap) wrap.classList.toggle('cell-confirmed', confirmed);
      applyAppendCheckedState();
      if (isMasterTab(tab)) renderMasterLibrary();
      scheduleSave();
    }
    setStatus('Cut cell', 'ok');
  }

  function addRow() {
    const tab = activeTab();
    if (!tab) return;
    pushHistory();
    ensureSleptCells(tab);
    for (let c = 0; c < tab.cols; c++) {
      tab.cells.push('');
      tab.sleptCells.push(false);
    }
    tab.rows += 1;
    renderTabs();
    renderGrid();
    renderMasterLibrary();
    scheduleSave();
    setStatus('Row added (' + tab.cols + '×' + tab.rows + ')');
  }

  function rowIsEmpty(tab, rowIndex) {
    for (let col = 0; col < tab.cols; col++) {
      if ((tab.cells[rowIndex * tab.cols + col] || '').trim()) return false;
    }
    return true;
  }

  function sortRowsByColumn(colIndex, direction) {
    const tab = activeTab();
    if (!tab) return;
    if (colIndex < 0 || colIndex >= tab.cols) return;
    if (direction !== 'asc' && direction !== 'desc') return;

    pushHistory();
    ensureSleptCells(tab);

    const cols = tab.cols;
    const rows = tab.rows;
    const order = [];
    for (let r = 0; r < rows; r++) order.push(r);

    order.sort(function (a, b) {
      const va = (tab.cells[a * cols + colIndex] || '').trim();
      const vb = (tab.cells[b * cols + colIndex] || '').trim();
      // Keep blank column-1 cells at the bottom for both directions.
      if (!va && !vb) return a - b;
      if (!va) return 1;
      if (!vb) return -1;
      const cmp = va.localeCompare(vb, undefined, { sensitivity: 'base', numeric: true });
      if (cmp !== 0) return direction === 'desc' ? -cmp : cmp;
      return a - b;
    });

    const newCells = [];
    const newSlept = [];
    for (let i = 0; i < order.length; i++) {
      const src = order[i];
      for (let c = 0; c < cols; c++) {
        const idx = src * cols + c;
        newCells.push(tab.cells[idx] || '');
        newSlept.push(!!tab.sleptCells[idx]);
      }
    }

    const oldToNew = new Array(rows);
    for (let newR = 0; newR < order.length; newR++) {
      oldToNew[order[newR]] = newR;
    }

    tab.cells = newCells;
    tab.sleptCells = newSlept;

    state.confirmedLinks.forEach(function (link) {
      if (link.tabId !== tab.id) return;
      const row = Math.floor(link.cellIndex / cols);
      const col = link.cellIndex % cols;
      if (row < 0 || row >= rows) return;
      link.cellIndex = oldToNew[row] * cols + col;
    });

    renderTabs();
    renderGrid();
    renderMasterLibrary();
    scheduleSave();
    setStatus(direction === 'desc' ? 'Sorted by column 1 Z–A' : 'Sorted by column 1 A–Z');
  }

  function moveRow(rowIndex, direction) {
    const tab = activeTab();
    if (!tab || (direction !== -1 && direction !== 1)) return;
    const destination = rowIndex + direction;
    if (rowIndex < 0 || rowIndex >= tab.rows || destination < 0) return;

    pushHistory();
    ensureSleptCells(tab);
    if (direction < 0) {
      for (let col = 0; col < tab.cols; col++) {
        const sourceIndex = rowIndex * tab.cols + col;
        const destinationIndex = destination * tab.cols + col;
        const cell = tab.cells[destinationIndex];
        tab.cells[destinationIndex] = tab.cells[sourceIndex];
        tab.cells[sourceIndex] = cell;
      }
      swapSleptRowIndices(tab, rowIndex, destination);
      swapConfirmedRowIndices(tab.id, tab.cols, rowIndex, destination);
    } else {
      let emptyRow = rowIndex + 1;
      while (emptyRow < tab.rows && !rowIsEmpty(tab, emptyRow)) emptyRow++;

      if (emptyRow === tab.rows) {
        tab.cells.push.apply(tab.cells, emptyCells(tab.cols, 1));
        tab.sleptCells.push.apply(tab.sleptCells, emptySleptCells(tab.cols, 1));
        tab.rows++;
        emptyRow = tab.rows - 1;
      }

      const movingCells = tab.cells.slice(rowIndex * tab.cols, (rowIndex + 1) * tab.cols);
      for (let row = emptyRow; row > rowIndex + 1; row--) {
        for (let col = 0; col < tab.cols; col++) {
          tab.cells[row * tab.cols + col] = tab.cells[(row - 1) * tab.cols + col];
        }
      }
      for (let col = 0; col < tab.cols; col++) {
        tab.cells[(rowIndex + 1) * tab.cols + col] = movingCells[col];
        tab.cells[rowIndex * tab.cols + col] = '';
      }
      shiftSleptRowsDown(tab, rowIndex, emptyRow);
      shiftConfirmedRowsDown(tab.id, tab.cols, rowIndex, emptyRow);
    }

    renderTabs();
    renderGrid();
    scheduleSave();
    setStatus(direction < 0 ? 'Row moved up' : 'Row moved down');
  }

  function addColumn() {
    const tab = activeTab();
    if (!tab) return;
    pushHistory();
    ensureSleptCells(tab);
    const oldCols = tab.cols;
    const newCells = [];
    for (let r = 0; r < tab.rows; r++) {
      for (let c = 0; c < tab.cols; c++) {
        newCells.push(tab.cells[r * tab.cols + c] || '');
      }
      newCells.push('');
    }
    if (Array.isArray(tab.columnWidths) && tab.columnWidths.length === tab.cols) {
      const defaultWidth = tab.columnWidths[tab.columnWidths.length - 1] || DEFAULT_COLUMN_WIDTH;
      tab.columnWidths = tab.columnWidths.concat([defaultWidth]);
    }
    tab.cols += 1;
    tab.cells = newCells;
    remapSleptAfterColumnAdd(tab, oldCols, tab.cols);
    remapConfirmedAfterColumnAdd(tab.id, oldCols, tab.cols);
    renderTabs();
    renderGrid();
    renderMasterLibrary();
    scheduleSave();
    setStatus('Column added (' + tab.cols + '×' + tab.rows + ')');
  }

  function rowCellIndices(tab, rowIndex) {
    const indices = [];
    for (let c = 0; c < tab.cols; c++) indices.push(rowIndex * tab.cols + c);
    return indices;
  }

  function rowConfirmedPiecesMissing(tab, rowIndex) {
    const scope = currentPromptScope();
    const pieces = [];
    const colSep = separatorValue(state.separators.column);
    for (let c = 0; c < tab.cols; c++) {
      const idx = rowIndex * tab.cols + c;
      const value = (tab.cells[idx] || '').trim();
      if (!value) continue;
      if (isCellSlept(tab, idx)) continue;
      if (isCellConfirmedInScope(tab.id, idx, scope)) continue;
      if (pieces.length) pieces.push({ type: 'plain', text: colSep });
      pieces.push({
        type: 'confirmed',
        text: value,
        tabId: tab.id,
        cellIndex: idx
      });
    }
    return pieces;
  }

  function toggleRow(rowIndex) {
    const tab = activeTab();
    if (!tab) return;
    if (rowIndex < 0 || rowIndex >= tab.rows) return;
    const scope = currentPromptScope();

    if (isRowIncluded(tab, rowIndex)) {
      const links = linksForCellsInScope(tab.id, rowCellIndices(tab, rowIndex), scope);
      if (!links.length) {
        applyAppendCheckedState();
        return;
      }
      pushHistory();
      removeLinksFromCombined(links);
      refreshAfterConfirmedChange();
      setStatus('Removed "' + tab.title + ' row ' + (rowIndex + 1) + '"', 'ok');
      return;
    }

    const pieces = rowConfirmedPieces(tab, rowIndex);
    appendPieces(pieces, tab.title + ' row ' + (rowIndex + 1));
  }

  function appendRow(rowIndex) {
    toggleRow(rowIndex);
  }

  function appendAll() {
    const tab = activeTab();
    if (!tab) return;
    const pieces = [];
    const rowSep = separatorValue(state.separators.row);
    for (let r = 0; r < tab.rows; r++) {
      const rowPieces = rowConfirmedPiecesMissing(tab, r);
      if (!rowPieces.length) continue;
      if (pieces.length) pieces.push({ type: 'plain', text: rowSep });
      for (let i = 0; i < rowPieces.length; i++) pieces.push(rowPieces[i]);
    }
    appendPieces(pieces, tab.title);
  }

  function selectCombinedContents() {
    el.combined.focus();
    const range = document.createRange();
    range.selectNodeContents(el.combined);
    const selection = window.getSelection();
    if (!selection) return;
    selection.removeAllRanges();
    selection.addRange(range);
  }

  async function copyCombined() {
    const text = getCombinedPlainText();
    if (!text) {
      setStatus('Combined prompt is empty', 'err');
      return;
    }
    try {
      await navigator.clipboard.writeText(text);
      setStatus('Copied to clipboard', 'ok');
    } catch (err) {
      selectCombinedContents();
      try {
        document.execCommand('copy');
        setStatus('Copied to clipboard', 'ok');
      } catch (e2) {
        setStatus('Copy failed', 'err');
      }
    }
  }

  function showUndoClearButton(visible) {
    if (!el.btnUndoClear) return;
    el.btnUndoClear.hidden = !visible;
  }

  function dismissUndoClear() {
    lastClearSnapshot = null;
    showUndoClearButton(false);
  }

  function clearCombined() {
    const scope = currentPromptScope();
    // Snapshot before clear for the dedicated "Undo last clear" button (and undo stack).
    lastClearSnapshot = cloneCurrentDocument();
    pushHistory();
    setPromptText(scope, '');
    dropLinksForScope(scope);
    renderCombinedPrompt();
    applyConfirmedCellHighlights();
    scheduleSave();
    showUndoClearButton(true);
    setStatus('Cleared');
  }

  function undoLastClear() {
    if (!lastClearSnapshot) {
      setStatus('Nothing to undo');
      showUndoClearButton(false);
      return;
    }
    endHistoryCoalesce();
    redoStack.push(cloneCurrentDocument());
    if (redoStack.length > UNDO_LIMIT) redoStack.shift();
    // Drop the matching clear entry from the undo stack when it is still on top.
    if (undoStack.length) undoStack.pop();
    const entry = lastClearSnapshot;
    lastClearSnapshot = null;
    showUndoClearButton(false);
    restoreHistoryEntry(entry);
    setStatus('Undid clear', 'ok');
  }

  async function exportBackup() {
    if (!window.click2copy || typeof window.click2copy.exportData !== 'function') {
      setStatus('Export unavailable', 'err');
      return;
    }
    try {
      const result = await window.click2copy.exportData(snapshot());
      if (!result || result.canceled) {
        setStatus('Export canceled');
        return;
      }
      if (!result.ok) {
        setStatus(result.error || 'Export failed', 'err');
        return;
      }
      setStatus('Exported backup', 'ok');
    } catch (err) {
      console.error(err);
      setStatus('Export failed', 'err');
    }
  }

  async function importBackup() {
    if (!window.click2copy || typeof window.click2copy.importData !== 'function') {
      setStatus('Import unavailable', 'err');
      return;
    }
    try {
      const result = await window.click2copy.importData();
      if (!result || result.canceled) {
        setStatus('Import canceled');
        return;
      }
      if (!result.ok || !result.data) {
        setStatus(result.error || 'Import failed', 'err');
        return;
      }

      const ok = window.confirm(
        'Replace all tabs and combined prompt with this backup?'
      );
      if (!ok) {
        setStatus('Import canceled');
        return;
      }

      pushHistory();
      applyData(result.data);
      scheduleSave();
      setStatus('Imported backup (' + state.tabs.length + ' tabs)', 'ok');
    } catch (err) {
      console.error(err);
      setStatus('Import failed', 'err');
    }
  }

  el.btnAdd.addEventListener('click', addTab);
  el.btnRename.addEventListener('click', renameActiveTab);
  el.btnDelete.addEventListener('click', function () {
    if (state.activeTabId) deleteTab(state.activeTabId);
  });
  el.btnAddRow.addEventListener('click', addRow);
  el.btnAddCol.addEventListener('click', addColumn);
  if (el.btnFilterAll) el.btnFilterAll.addEventListener('click', function () { setGridRowFilter('all'); });
  if (el.btnFilterNonempty) el.btnFilterNonempty.addEventListener('click', function () { setGridRowFilter('nonempty'); });
  if (el.btnFilterIncluded) el.btnFilterIncluded.addEventListener('click', function () { setGridRowFilter('included'); });
  document.addEventListener('pointerdown', function (e) {
    if (!col1FilterMenuOpen) return;
    const wrap = el.cellGrid && el.cellGrid.querySelector('.column-col1-filter');
    if (wrap && wrap.contains(e.target)) return;
    closeCol1FilterMenu();
  });
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') closeCol1FilterMenu();
  });
  window.addEventListener('resize', closeCol1FilterMenu);
  if (el.cellGrid && el.cellGrid.parentElement) {
    el.cellGrid.parentElement.addEventListener('scroll', closeCol1FilterMenu, { passive: true });
  }
  el.btnAppend.addEventListener('click', appendAll);
  el.btnCopy.addEventListener('click', copyCombined);
  el.btnClear.addEventListener('click', clearCombined);
  el.btnUndoClear.addEventListener('click', undoLastClear);
  el.btnExport.addEventListener('click', exportBackup);
  el.btnImport.addEventListener('click', importBackup);
  el.combinedResizer.addEventListener('pointerdown', beginSectionResize);
  el.combinedResizer.addEventListener('keydown', function (event) {
    if (event.key !== 'ArrowUp' && event.key !== 'ArrowDown') return;
    event.preventDefault();
    resizeCombinedSectionByKeyboard(event.key === 'ArrowUp' ? 24 : -24);
  });

  function onSeparatorInput(key, input) {
    input.addEventListener('input', function () {
      pushHistory({ coalesce: true });
      state.separators[key] = input.value;
      scheduleSave();
    });
  }

  onSeparatorInput('part', el.partSeparator);
  onSeparatorInput('column', el.columnSeparator);
  onSeparatorInput('row', el.rowSeparator);

  if (window.click2copy && window.click2copy.onOpenDocument) {
    window.click2copy.onOpenDocument(function (filePath) {
      if (!initialized) {
        pendingDocumentPaths.push(filePath);
        return;
      }
      openDocumentPath(filePath).catch(function (err) {
        console.error(err);
        setStatus('Could not open prompt file: ' + (err.message || 'Unknown error'), 'err');
      });
    });
  }

  el.combined.addEventListener('focus', function () {
    copyCombinedOnActivate();
  });
  el.combined.addEventListener('click', function () {
    copyCombinedOnActivate();
  });

  el.combined.addEventListener('input', function () {
    dismissUndoClear();
    pushHistory({ coalesce: true });
    syncConfirmedFromCombinedDom();
    scheduleSave();
  });

  el.combined.addEventListener('keydown', function (event) {
    if (event.key !== 'Enter' || event.isComposing) return;
    event.preventDefault();
    if (document.queryCommandSupported && document.queryCommandSupported('insertText')) {
      document.execCommand('insertText', false, '\n');
    } else {
      const selection = window.getSelection();
      if (!selection || !selection.rangeCount) return;
      const range = selection.getRangeAt(0);
      range.deleteContents();
      const node = document.createTextNode('\n');
      range.insertNode(node);
      range.setStartAfter(node);
      range.collapse(true);
      selection.removeAllRanges();
      selection.addRange(range);
      dismissUndoClear();
      pushHistory({ coalesce: true });
      syncConfirmedFromCombinedDom();
      scheduleSave();
    }
  });

  el.combined.addEventListener('paste', function (event) {
    event.preventDefault();
    const pasted = (event.clipboardData || window.clipboardData).getData('text/plain');
    if (document.queryCommandSupported && document.queryCommandSupported('insertText')) {
      document.execCommand('insertText', false, pasted || '');
    } else {
      const selection = window.getSelection();
      if (!selection || !selection.rangeCount) return;
      const range = selection.getRangeAt(0);
      range.deleteContents();
      const node = document.createTextNode(pasted || '');
      range.insertNode(node);
      range.setStartAfter(node);
      range.collapse(true);
      selection.removeAllRanges();
      selection.addRange(range);
      dismissUndoClear();
      pushHistory({ coalesce: true });
      syncConfirmedFromCombinedDom();
      scheduleSave();
    }
  });

  el.globalCombined.addEventListener('change', function () {
    pushHistory();
    const next = el.globalCombined.checked;
    if (next && !state.globalCombined) mergePartPrompts();
    state.globalCombined = next;
    renderCombinedPrompt();
    scheduleSave();
  });

  async function init() {
    try {
      const savedHeight = Number(localStorage.getItem('click2copy-combined-height'));
      if (Number.isFinite(savedHeight) && savedHeight > 0) resizeCombinedSection(savedHeight);
    } catch (err) {
      console.warn('Could not load combined prompt height:', err);
    }

    if (window.click2copy && window.click2copy.getAppVersion) {
      try {
        el.appVersion.textContent = 'v' + await window.click2copy.getAppVersion();
      } catch (err) {
        console.error('Could not read app version:', err);
      }
    }

    if (window.click2copy && typeof window.click2copy.getAutosaveLocations === 'function') {
      try {
        autosaveLocations = await window.click2copy.getAutosaveLocations();
      } catch (err) {
        console.error('Could not read autosave locations:', err);
      }
    }

    let session;
    if (window.click2copy) {
      try {
        session = await window.click2copy.loadSession();
      } catch (err) {
        console.error(err);
      }
    }

    if (!session || !Array.isArray(session.documents) || session.documents.length === 0) {
      let data = null;
      if (window.click2copy) {
        try {
          data = await window.click2copy.loadData();
        } catch (err) {
          console.error(err);
        }
      }
      const fallback = data && Array.isArray(data.tabs) && data.tabs.length > 0 ? data : makeDefaultData();
      if (!fallback.activeTabId) fallback.activeTabId = defaultActiveTabId(fallback.tabs) || fallback.tabs[0].id;
      session = {
        documents: [{ id: 'document-default', title: 'Untitled', filePath: null, data: fallback }],
        activeDocumentId: 'document-default'
      };
    }

    state.documents = session.documents.filter(function (item) {
      return item && typeof item.id === 'string' && item.data && Array.isArray(item.data.tabs);
    });
    if (state.documents.length === 0) {
      const data = makeDefaultData();
      data.activeTabId = defaultActiveTabId(data.tabs) || data.tabs[0].id;
      state.documents = [{ id: uid(), title: 'Untitled', filePath: null, data: data }];
    }
    state.activeDocumentId = state.documents.some(function (item) {
      return item.id === session.activeDocumentId;
    }) ? session.activeDocumentId : state.documents[0].id;

    let failedFile = false;
    for (const promptDocument of state.documents) {
      if (!promptDocument.filePath) continue;
      try {
        const sessionData = promptDocument.data;
        const result = await window.click2copy.openDocument(promptDocument.filePath);
        if (result && result.ok) {
          const fileData = result.data;
          // Named-file reload must not wipe greens: legacy/.c2copy without confirmedLinks
          // keeps session links when Combined text still matches.
          const fileLinks = normalizeConfirmedLinks(fileData && fileData.confirmedLinks);
          const sessionLinks = normalizeConfirmedLinks(sessionData && sessionData.confirmedLinks);
          if (fileLinks.length === 0 && sessionLinks.length > 0) {
            const fileCombined = typeof fileData.combinedPrompt === 'string' ? fileData.combinedPrompt : '';
            const sessionCombined = typeof sessionData.combinedPrompt === 'string' ? sessionData.combinedPrompt : '';
            if (fileCombined === sessionCombined) {
              fileData.confirmedLinks = sessionLinks;
              if (sessionData.partPrompts && typeof sessionData.partPrompts === 'object') {
                fileData.partPrompts = sessionData.partPrompts;
              }
            }
          } else if (fileLinks.length > 0) {
            fileData.confirmedLinks = fileLinks;
          }
          promptDocument.data = fileData;
          promptDocument.title = documentName(promptDocument.filePath);
        } else {
          failedFile = true;
          promptDocument.filePath = null;
        }
      } catch (err) {
        console.error(err);
        failedFile = true;
        promptDocument.filePath = null;
      }
    }

    const current = activeDocument();
    clearHistory();
    applyData(current.data);
    // Capture repaired greens back onto the document before launch persist.
    if (current) current.data = snapshot();
    renderDocuments();
    initialized = true;
    await persist();
    if (failedFile) setStatus('A prompt file could not be read; restored its last saved session copy', 'err');
    while (pendingDocumentPaths.length > 0) {
      const filePath = pendingDocumentPaths.shift();
      await openDocumentPath(filePath);
    }
  }

  if (el.status) {
    el.status.addEventListener('click', function () {
      setStatusPanelOpen(!statusPanelOpen);
    });
  }
  if (el.statusPanelClose) {
    el.statusPanelClose.addEventListener('click', function (event) {
      event.stopPropagation();
      setStatusPanelOpen(false);
    });
  }
  document.addEventListener('keydown', function (event) {
    if (event.key === 'Escape' && statusPanelOpen) setStatusPanelOpen(false);

    const mod = event.ctrlKey || event.metaKey;
    if (!mod || event.altKey) return;
    // Skip when renaming a tab inline (native text field undo is fine there).
    const target = event.target;
    if (target && target.classList && target.classList.contains('tab-rename-input')) return;

    const key = event.key.length === 1 ? event.key.toLowerCase() : event.key;
    if (key === 'z' && !event.shiftKey) {
      event.preventDefault();
      undo();
      return;
    }
    if ((key === 'z' && event.shiftKey) || key === 'y') {
      event.preventDefault();
      redo();
    }
  });
  document.addEventListener('pointerdown', function (event) {
    if (!statusPanelOpen || !el.statusPanel) return;
    const target = event.target;
    if (el.statusPanel.contains(target) || el.status.contains(target)) return;
    setStatusPanelOpen(false);
  });

  init();
})();
