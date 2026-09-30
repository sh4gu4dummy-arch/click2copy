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
  let initialized = false;
  let focusedCell = null;
  /** UI-only row filter: 'all' | 'nonempty' | 'included' */
  let gridRowFilter = 'all';
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

      const document = activeDocument();
      if (document && document.filePath) {
        const result = await window.click2copy.saveDocument(document.filePath, snapshot());
        if (!result || !result.ok) {
          setStatus((result && result.error) || 'Failed to save prompt file', 'err');
          return false;
        }
      } else {
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
    return [
      { label: 'Active document', value: docLabel },
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
  }

  function normalizeConfirmedLinks(raw) {
    if (!Array.isArray(raw)) return [];
    const links = [];
    for (let i = 0; i < raw.length; i++) {
      const item = raw[i];
      if (!item || typeof item !== 'object') continue;
      if (typeof item.id !== 'string' || !item.id) continue;
      if (typeof item.tabId !== 'string' || !item.tabId) continue;
      if (!Number.isInteger(item.cellIndex) || item.cellIndex < 0) continue;
      if (typeof item.text !== 'string') continue;
      if (!Number.isInteger(item.start) || item.start < 0) continue;
      if (!Number.isInteger(item.end) || item.end < item.start) continue;
      const scope = typeof item.scope === 'string' && item.scope ? item.scope : 'global';
      links.push({
        id: item.id,
        tabId: item.tabId,
        cellIndex: item.cellIndex,
        text: item.text,
        start: item.start,
        end: item.end,
        scope: scope
      });
    }
    return links;
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
  function appendPieces(pieces, label) {
    if (!pieces || !pieces.length) {
      setStatus('Nothing to append', 'err');
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
      setStatus('Nothing to append', 'err');
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
    const text = getPromptText(scope);
    const links = linksForScope(scope).filter(function (link) {
      if (link.start < 0 || link.end > text.length || link.start > link.end) return false;
      return text.slice(link.start, link.end) === link.text;
    });

    // Drop stale ranges that no longer match the prompt string.
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


  function rowMatchesFilter(tab, rowIndex) {
    if (gridRowFilter === 'all') return true;
    if (gridRowFilter === 'nonempty') return !rowIsEmpty(tab, rowIndex);
    if (gridRowFilter === 'included') return isRowIncluded(tab, rowIndex);
    return true;
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
      header.textContent = 'Column ' + (c + 1);

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
      rowBtn.addEventListener('click', function () {
        toggleRow(r);
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
        sleepCb.addEventListener('change', function () {
          toggleCellSleep(idx, sleepCb.checked);
        });
        sleepLabel.appendChild(sleepCb);
        gutter.appendChild(sleepLabel);

        const cellBtn = document.createElement('button');
        cellBtn.type = 'button';
        cellBtn.className = 'cell-append';
        cellBtn.dataset.idx = String(idx);
        cellBtn.dataset.row = String(r);
        cellBtn.dataset.col = String(c);
        cellBtn.setAttribute('aria-pressed', 'false');
        cellBtn.innerHTML = '<span class="cell-append-mark" aria-hidden="true"></span>';
        cellBtn.addEventListener('click', function (ev) {
          ev.preventDefault();
          ev.stopPropagation();
          toggleCellConfirmed(idx);
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
        ta.addEventListener('focus', rememberFocusedCell);
        ta.addEventListener('click', rememberFocusedCell);
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
        const result = await window.click2copy.openDocument(promptDocument.filePath);
        if (result && result.ok) {
          promptDocument.data = result.data;
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
