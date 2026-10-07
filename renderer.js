(function () {
  'use strict';

  const DEFAULT_COLS = 3;
  const DEFAULT_ROWS = 8;
  const DEFAULT_COLUMN_WIDTH = 160;
  const MIN_COLUMN_WIDTH = 100;
  /** Match --cell-min-h; floor for Fit (checkbox + nest + ~40px). Was 64 — left huge dead space on short rows. */
  const MIN_ROW_HEIGHT = 44;
  const DEFAULT_SEPARATORS = {
    part: '\\n\\n',
    column: ' | ',
    row: '\\n'
  };

  /** Pickable tab icons (Master + part tabs). Tools uses fixed wrench only. */
  const TAB_ICON_IDS = ['text', 'video', 'img', 'sparkle', 'tag', 'folder', 'layers', 'hash'];
  const TAB_ICON_LABELS = {
    text: 'Text (Aa)',
    video: 'Video',
    img: 'Image',
    sparkle: 'Sparkle',
    tag: 'Tag',
    folder: 'Folder',
    layers: 'Layers',
    hash: 'Hash',
    wrench: 'Tools'
  };
  const DEFAULT_PART_ICON = 'text';
  const DEFAULT_MASTER_ICON = 'layers';
  const TOOLS_ICON = 'wrench';

  const TAB_ICON_SVG = {
    text: '<svg viewBox="0 0 16 16" aria-hidden="true"><text x="8" y="12" text-anchor="middle" font-size="10" font-weight="700" font-family="Segoe UI,system-ui,sans-serif" fill="currentColor">Aa</text></svg>',
    video: '<svg viewBox="0 0 16 16" aria-hidden="true"><rect x="1.5" y="3.5" width="9" height="9" rx="1.5" fill="none" stroke="currentColor" stroke-width="1.4"/><path d="M11.2 6.2 L14.5 4.2 V11.8 L11.2 9.8 Z" fill="currentColor"/></svg>',
    img: '<svg viewBox="0 0 16 16" aria-hidden="true"><rect x="1.5" y="2.5" width="13" height="11" rx="1.5" fill="none" stroke="currentColor" stroke-width="1.4"/><circle cx="5.5" cy="6" r="1.3" fill="currentColor"/><path d="M2.5 12.2 L6.2 8.2 L8.5 10.2 L11 7.5 L14.5 12.2 Z" fill="currentColor"/></svg>',
    sparkle: '<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M8 1.2 L9.1 6.1 L14 7.2 L9.1 8.3 L8 13.2 L6.9 8.3 L2 7.2 L6.9 6.1 Z" fill="currentColor"/><path d="M12.5 1.8 L13 3.5 L14.7 4 L13 4.5 L12.5 6.2 L12 4.5 L10.3 4 L12 3.5 Z" fill="currentColor"/></svg>',
    tag: '<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M2.2 8.8 V3.2 H7.8 L13.8 9.2 L9.2 13.8 Z" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linejoin="round"/><circle cx="5.2" cy="5.2" r="1.1" fill="currentColor"/></svg>',
    folder: '<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M1.5 4.2 H6 L7.4 5.6 H14.5 V12.5 H1.5 Z" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linejoin="round"/></svg>',
    layers: '<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M8 2.2 L14 5.2 L8 8.2 L2 5.2 Z" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linejoin="round"/><path d="M2 8 L8 11 L14 8" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linejoin="round"/><path d="M2 10.6 L8 13.6 L14 10.6" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linejoin="round"/></svg>',
    hash: '<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M6.2 2.5 L5.2 13.5 M10.8 2.5 L9.8 13.5 M2.5 6.2 H13.5 M2.5 10.2 H13.5" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>',
    wrench: '<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M10.8 2.2 a3.2 3.2 0 0 0-4.3 3.7 L2.2 10.2 l1.3 1.3 2.3-2.3 1.3 1.3-2.3 2.3 1.3 1.3 4.3-4.3 a3.2 3.2 0 0 0 3.7-4.3 l-2.2 2.2 -1.4-1.4 2.2-2.2 z" fill="currentColor"/></svg>'
  };

  function isTabIconId(id) {
    return typeof id === 'string' && TAB_ICON_IDS.indexOf(id) !== -1;
  }

  function defaultIconForTab(tab) {
    if (tab && (tab.id === 'master' || (tab.title && tab.title.trim().toLowerCase() === 'master'))) {
      return DEFAULT_MASTER_ICON;
    }
    return DEFAULT_PART_ICON;
  }

  function resolveTabIcon(tab) {
    if (tab && isTabIconId(tab.icon)) return tab.icon;
    return defaultIconForTab(tab);
  }

  function tabIconMarkup(iconId) {
    return TAB_ICON_SVG[iconId] || TAB_ICON_SVG[DEFAULT_PART_ICON];
  }

  /** Excel-style column letters: 0→A, 25→Z, 26→AA, … */
  function columnLetter(colIndex) {
    let n = (Number(colIndex) || 0) + 1;
    if (n < 1) n = 1;
    let s = '';
    while (n > 0) {
      n--;
      s = String.fromCharCode(65 + (n % 26)) + s;
      n = Math.floor(n / 26);
    }
    return s;
  }

  /** Excel-style cell address from 0-based row/col: (0,0)→A1 */
  function cellAddress(rowIndex, colIndex) {
    return columnLetter(colIndex) + String((Number(rowIndex) || 0) + 1);
  }

  /** Address from flat cell index on a tab (0-based). */
  function cellAddressFromIndex(tab, cellIndex) {
    if (!tab || !Number.isInteger(cellIndex) || cellIndex < 0) return '?';
    const cols = tab.cols || 1;
    return cellAddress(Math.floor(cellIndex / cols), cellIndex % cols);
  }

  const state = {
    documents: [],
    activeDocumentId: null,
    tabs: [],
    activeTabId: null,
    combinedPrompt: '',
    globalCombined: true,
    matchSourceOrder: false,
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
  /**
   * Per-tab grid filter prefs (UI-only).
   * Keyed by tab id → { rowFilter: 'all'|'nonempty'|'included', valueFilter: null|Set }.
   * Prevents editing Values / row filter on one tab from overwriting another tab's selections.
   */
  const gridFilterPrefsByTabId = Object.create(null);
  let col1FilterMenuOpen = false;
  /** UI-only Master insert filter menu open (ephemeral; not per-tab) */
  let masterLibFilterMenuOpen = false;
  /** Skip scroll-dismiss while Values filter refresh changes Master insert height. */
  let suppressMasterLibMenuScrollClose = false;
  /**
   * Per-part-tab Master insert prefs (UI-only, not persisted).
   * Keyed by part tab id → { valueFilter: null|Set, sortDir: null|'asc'|'desc' }.
   * Switching part tabs restores that tab's Values selection and A–Z sort.
   */
  const masterLibPrefsByPartId = Object.create(null);
  /** UI-only Tools tab — not stored in document tabs; acts on lastPartTabId */
  let toolsTabActive = false;
  let lastPartTabId = null;
  /** UI-only Find / Find All / Replace over the active part tab (cells + nest pages). */
  let partSearchHits = [];
  let partSearchHitIndex = -1;
  let partSearchQuery = '';
  let partSearchFindAll = false;
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
    partSection: document.querySelector('.part-section'),
    toolsPanel: document.getElementById('tools-panel'),
    toolsPanelHint: document.getElementById('tools-panel-hint'),
    combined: document.getElementById('combined-prompt'),
    masterSegmentDialog: document.getElementById('master-segment-dialog'),
    btnMasterOverwrite: document.getElementById('btn-master-overwrite'),
    btnMasterKeepLocal: document.getElementById('btn-master-keep-local'),
    btnMasterCancelEdit: document.getElementById('btn-master-cancel-edit'),
    globalCombined: document.getElementById('global-combined'),
    matchSourceOrder: document.getElementById('match-source-order'),
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
    btnFitRows: document.getElementById('btn-fit-rows'),
    btnAppend: document.getElementById('btn-append'),
    btnCopy: document.getElementById('btn-copy'),
    btnClear: document.getElementById('btn-clear'),
    btnUndoClear: document.getElementById('btn-undo-clear'),
    btnExport: document.getElementById('btn-export'),
    btnImport: document.getElementById('btn-import'),
    partSeparator: document.getElementById('part-separator'),
    columnSeparator: document.getElementById('column-separator'),
    rowSeparator: document.getElementById('row-separator'),
    partSearchBar: document.getElementById('part-search-bar'),
    partTabPageChrome: document.getElementById('part-tab-page-chrome'),
    partFindInput: document.getElementById('part-find-input'),
    partReplaceInput: document.getElementById('part-replace-input'),
    btnFind: document.getElementById('btn-find'),
    btnFindAll: document.getElementById('btn-find-all'),
    btnReplace: document.getElementById('btn-replace'),
    partSearchStatus: document.getElementById('part-search-status'),
    updateBanner: document.getElementById('update-banner'),
    updateBannerText: document.getElementById('update-banner-text'),
    btnUpdateRestart: document.getElementById('btn-update-restart'),
    btnUpdateDismiss: document.getElementById('btn-update-dismiss')
  };

  function uid() {
    return 'tab-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 7);
  }

  function emptyCells(cols, rows) {
    return Array(cols * rows).fill('');
  }


  /** Parallel to cells: each entry is an array of nest objects { pages, page }. */
  function emptyNestedCells(cols, rows) {
    const out = [];
    const n = cols * rows;
    for (let i = 0; i < n; i++) out.push([]);
    return out;
  }

  function makeEmptyNest() {
    return { pages: [''], page: 0 };
  }

  /** One nest: pages[] of strings + current page index. Legacy string → one page. */
  function normalizeNest(item) {
    if (typeof item === 'string') {
      return { pages: [item], page: 0 };
    }
    if (!item || typeof item !== 'object') return makeEmptyNest();
    let pages;
    if (Array.isArray(item.pages) && item.pages.length) {
      pages = item.pages.map(function (p) { return typeof p === 'string' ? p : ''; });
    } else if (typeof item.text === 'string') {
      pages = [item.text];
    } else {
      pages = [''];
    }
    let page = Number.isInteger(item.page) ? item.page : 0;
    if (page < 0) page = 0;
    if (page >= pages.length) page = pages.length - 1;
    return { pages: pages, page: page };
  }

  function normalizeNestList(list) {
    if (!Array.isArray(list)) return [];
    return list.map(normalizeNest);
  }

  /** Deep clone nest list (normalizeNest always returns fresh objects). */
  function cloneNestList(list) {
    return normalizeNestList(list);
  }

  /** Compare nest page contents (ignore current page index — view state only). */
  function nestsContentEqual(a, b) {
    const aa = normalizeNestList(a);
    const bb = normalizeNestList(b);
    if (aa.length !== bb.length) return false;
    for (let i = 0; i < aa.length; i++) {
      if (aa[i].pages.length !== bb[i].pages.length) return false;
      for (let p = 0; p < aa[i].pages.length; p++) {
        if (aa[i].pages[p] !== bb[i].pages[p]) return false;
      }
    }
    return true;
  }


  const CELL_SHADE_IDS = ['green', 'yellow', 'red'];

  function emptyCellShades(cols, rows) {
    return Array((cols || 0) * (rows || 0)).fill(null);
  }

  function normalizeCellShade(value) {
    if (value === 'green' || value === 'yellow' || value === 'red') return value;
    return null;
  }

  function normalizeCellShades(shades, length) {
    const needed = length > 0 ? length : 0;
    let out;
    if (Array.isArray(shades)) {
      out = shades.map(normalizeCellShade);
    } else {
      out = Array(needed).fill(null);
    }
    if (out.length < needed) {
      while (out.length < needed) out.push(null);
    } else if (out.length > needed) {
      out = out.slice(0, needed);
    }
    return out;
  }

  function makeEmptyCellPages(text) {
    return { pages: [typeof text === 'string' ? text : ''], page: 0 };
  }

  /** Parent-cell pages — same shape as a nest: { pages: string[], page }. */
  function normalizeCellPagesEntry(item, fallbackText) {
    if (typeof item === 'string') return { pages: [item], page: 0 };
    if (!item || typeof item !== 'object') {
      return makeEmptyCellPages(typeof fallbackText === 'string' ? fallbackText : '');
    }
    let pages;
    if (Array.isArray(item.pages) && item.pages.length) {
      pages = item.pages.map(function (p) { return typeof p === 'string' ? p : ''; });
    } else if (typeof item.text === 'string') {
      pages = [item.text];
    } else if (typeof fallbackText === 'string') {
      pages = [fallbackText];
    } else {
      pages = [''];
    }
    let page = Number.isInteger(item.page) ? item.page : 0;
    if (page < 0) page = 0;
    if (page >= pages.length) page = pages.length - 1;
    return { pages: pages, page: page };
  }

  function cloneCellPagesEntry(entry) {
    return normalizeCellPagesEntry(entry, '');
  }

  function normalizeCellPages(raw, cells, length) {
    const needed = length > 0 ? length : 0;
    const src = Array.isArray(raw) ? raw : [];
    const cellSrc = Array.isArray(cells) ? cells : [];
    const out = [];
    for (let i = 0; i < needed; i++) {
      const fallback = typeof cellSrc[i] === 'string' ? cellSrc[i] : '';
      // Pages are canonical; cells[i] only seeds a missing entry (legacy / first create).
      out.push(normalizeCellPagesEntry(src[i], fallback));
    }
    return out;
  }

  /** Compare page contents (ignore current page index — view state only). */
  function cellPagesContentEqual(a, b) {
    const aa = normalizeCellPagesEntry(a, '');
    const bb = normalizeCellPagesEntry(b, '');
    if (aa.pages.length !== bb.pages.length) return false;
    for (let p = 0; p < aa.pages.length; p++) {
      if (aa.pages[p] !== bb.pages[p]) return false;
    }
    return true;
  }

  function ensureCellShades(tab) {
    if (!tab) return;
    const needed = (tab.cols || 0) * (tab.rows || 0);
    tab.cellShades = normalizeCellShades(tab.cellShades, needed);
  }

  function ensureCellPages(tab) {
    if (!tab) return;
    const needed = (tab.cols || 0) * (tab.rows || 0);
    tab.cellPages = normalizeCellPages(tab.cellPages, tab.cells, needed);
    // Mirror visible page into cells[].
    for (let i = 0; i < needed; i++) {
      const entry = tab.cellPages[i];
      if (!entry || !entry.pages.length) continue;
      tab.cells[i] = entry.pages[entry.page] == null ? '' : String(entry.pages[entry.page]);
    }
  }

  function getCellPages(tab, cellIndex) {
    if (!tab || cellIndex < 0) return makeEmptyCellPages('');
    ensureCellPages(tab);
    if (cellIndex >= tab.cellPages.length) return makeEmptyCellPages('');
    return tab.cellPages[cellIndex];
  }

  function getCellShade(tab, cellIndex) {
    if (!tab || cellIndex < 0) return null;
    ensureCellShades(tab);
    if (cellIndex >= tab.cellShades.length) return null;
    return tab.cellShades[cellIndex];
  }

  function setCellShade(tab, cellIndex, shade) {
    if (!tab || cellIndex < 0 || cellIndex >= tab.cells.length) return;
    ensureCellShades(tab);
    tab.cellShades[cellIndex] = normalizeCellShade(shade);
  }

  /** Write textarea/live value into the current page + cells[i]. */
  function writeCellCurrentPage(tab, cellIndex, text) {
    if (!tab || cellIndex < 0 || cellIndex >= tab.cells.length) return;
    ensureCellPages(tab);
    const entry = tab.cellPages[cellIndex];
    const value = text == null ? '' : String(text);
    tab.cells[cellIndex] = value;
    if (entry && entry.pages && entry.pages.length) {
      const page = entry.page || 0;
      if (page >= 0 && page < entry.pages.length) entry.pages[page] = value;
    }
  }

  /** After flipping page index, mirror pages[page] → cells[i]. */
  function syncCellTextFromPages(tab, cellIndex) {
    if (!tab || cellIndex < 0 || cellIndex >= tab.cells.length) return;
    ensureCellPages(tab);
    const entry = tab.cellPages[cellIndex];
    if (!entry || !entry.pages.length) return;
    const page = entry.page || 0;
    tab.cells[cellIndex] = entry.pages[page] == null ? '' : String(entry.pages[page]);
  }

  /**
   * Grow/normalize cellPages length without mirroring pages→cells.
   * ensureCellPages() mirrors and can wipe a pending cells[] write (Master insert
   * on an empty part page is the common case).
   */
  function ensureCellPagesArray(tab) {
    if (!tab) return;
    const needed = (tab.cols || 0) * (tab.rows || 0);
    tab.cellPages = normalizeCellPages(tab.cellPages, tab.cells, needed);
  }

  function copyPagesFromMasterToCell(targetTab, targetIndex, masterIdx) {
    if (!targetTab || targetIndex < 0) return;
    // Do not ensureCellPages() here — it mirrors stale empty page snapshots onto
    // cells[] and can blank a Master insert before the clone lands.
    ensureCellPagesArray(targetTab);
    if (targetIndex >= targetTab.cellPages.length) return;
    const master = state.tabs.find(isMasterTab);
    if (!master || !Number.isInteger(masterIdx) || masterIdx < 0 ||
        masterIdx >= master.cells.length) {
      targetTab.cellPages[targetIndex] = makeEmptyCellPages(targetTab.cells[targetIndex] || '');
      return;
    }
    // Normalize Master page array without mirroring (avoids wiping Master.cells
    // when cellPages desynced from the visible library text).
    ensureCellPagesArray(master);
    const live = master.cells[masterIdx] == null ? '' : String(master.cells[masterIdx]);
    let entry = cloneCellPagesEntry(
      master.cellPages[masterIdx] || makeEmptyCellPages(live)
    );
    // Master library paints master.cells — trust that for the active page slot.
    if (live !== (entry.pages[entry.page] || '')) {
      entry.pages[entry.page] = live;
    }
    targetTab.cellPages[targetIndex] = entry;
    syncCellTextFromPages(targetTab, targetIndex);
  }

  /**
   * Push Master[masterIdx] pages onto every Master-origin part cell locked to that index.
   * Skips an in-progress Master cell edit session on that cell.
   */
  function syncLockedPartPagesFromMaster(masterIdx) {
    if (!Number.isInteger(masterIdx) || masterIdx < 0) return;
    const master = state.tabs.find(isMasterTab);
    if (!master || masterIdx >= master.cells.length) return;
    ensureCellPages(master);
    const source = cloneCellPagesEntry(getCellPages(master, masterIdx));
    state.tabs.forEach(function (tab) {
      if (!tab || isMasterTab(tab)) return;
      ensureCellLocks(tab);
      ensureCellPages(tab);
      for (let i = 0; i < tab.cellLocks.length; i++) {
        const lock = tab.cellLocks[i];
        if (!lock || lock.masterOrigin !== true) continue;
        if (!Number.isInteger(lock.masterCellIndex) || lock.masterCellIndex !== masterIdx) continue;
        if (isMasterCellEditFor(tab, i)) continue;
        if (cellPagesContentEqual(tab.cellPages[i], source) &&
            (tab.cells[i] || '') === (source.pages[source.page] || '')) {
          // Still sync page index (view state) if contents match.
          if (tab.cellPages[i].page !== source.page) {
            tab.cellPages[i].page = source.page;
            syncCellTextFromPages(tab, i);
            liveSyncConfirmedLinksForCell(tab.id, i, null, { silent: true });
          }
          continue;
        }
        tab.cellPages[i] = cloneCellPagesEntry(source);
        syncCellTextFromPages(tab, i);
        liveSyncConfirmedLinksForCell(tab.id, i, null, { silent: true });
      }
    });
  }

  function swapCellPagesIndices(tab, indexA, indexB) {
    if (!tab || indexA === indexB) return;
    ensureCellPages(tab);
    if (indexA < 0 || indexB < 0 || indexA >= tab.cellPages.length || indexB >= tab.cellPages.length) return;
    const tmp = tab.cellPages[indexA];
    tab.cellPages[indexA] = tab.cellPages[indexB];
    tab.cellPages[indexB] = tmp;
  }

  function swapCellShadeIndices(tab, indexA, indexB) {
    if (!tab || indexA === indexB) return;
    ensureCellShades(tab);
    if (indexA < 0 || indexB < 0 || indexA >= tab.cellShades.length || indexB >= tab.cellShades.length) return;
    const tmp = tab.cellShades[indexA];
    tab.cellShades[indexA] = tab.cellShades[indexB];
    tab.cellShades[indexB] = tmp;
  }

  function remapCellPagesAfterColumnAdd(tab, oldCols, newCols) {
    const oldPages = Array.isArray(tab.cellPages) ? tab.cellPages : [];
    const next = [];
    for (let r = 0; r < tab.rows; r++) {
      for (let c = 0; c < newCols; c++) {
        if (c < oldCols) {
          const oldIdx = r * oldCols + c;
          next.push(oldIdx < oldPages.length
            ? normalizeCellPagesEntry(oldPages[oldIdx], tab.cells[r * newCols + c] || '')
            : makeEmptyCellPages(tab.cells[r * newCols + c] || ''));
        } else {
          next.push(makeEmptyCellPages(''));
        }
      }
    }
    tab.cellPages = next;
  }

  function remapCellShadesAfterColumnAdd(tab, oldCols, newCols) {
    const oldShades = Array.isArray(tab.cellShades) ? tab.cellShades : [];
    const next = [];
    for (let r = 0; r < tab.rows; r++) {
      for (let c = 0; c < newCols; c++) {
        if (c < oldCols) {
          const oldIdx = r * oldCols + c;
          next.push(oldIdx < oldShades.length ? normalizeCellShade(oldShades[oldIdx]) : null);
        } else {
          next.push(null);
        }
      }
    }
    tab.cellShades = next;
  }

  function copyCellPagesRow(tab, fromRow, toRow) {
    ensureCellPages(tab);
    for (let col = 0; col < tab.cols; col++) {
      tab.cellPages[toRow * tab.cols + col] = cloneCellPagesEntry(
        tab.cellPages[fromRow * tab.cols + col]
      );
    }
  }

  function clearCellPagesRow(tab, rowIndex) {
    ensureCellPages(tab);
    for (let col = 0; col < tab.cols; col++) {
      tab.cellPages[rowIndex * tab.cols + col] = makeEmptyCellPages('');
    }
  }

  function copyCellShadeRow(tab, fromRow, toRow) {
    ensureCellShades(tab);
    for (let col = 0; col < tab.cols; col++) {
      tab.cellShades[toRow * tab.cols + col] = tab.cellShades[fromRow * tab.cols + col];
    }
  }

  function clearCellShadeRow(tab, rowIndex) {
    ensureCellShades(tab);
    for (let col = 0; col < tab.cols; col++) {
      tab.cellShades[rowIndex * tab.cols + col] = null;
    }
  }

  function isMasterCellEditFor(tab, cellIndex) {
    return !!(isMasterCellEditSession() && tab &&
      masterSegmentEdit.tabId === tab.id &&
      masterSegmentEdit.cellIndex === cellIndex);
  }

  /** Locked Master-inserted cell and not in an unlock edit session. */
  function isCellMasterEditBlocked(tab, cellIndex) {
    return !!(tab && !isMasterTab(tab) && isCellMasterLocked(tab, cellIndex) &&
      !isMasterCellEditFor(tab, cellIndex));
  }

  /** Replace target nests with a clone of Master[masterIdx] nests (or []). */
  function copyNestsFromMasterToCell(targetTab, targetIndex, masterIdx) {
    if (!targetTab || targetIndex < 0) return;
    ensureNestedCells(targetTab);
    if (targetIndex >= targetTab.nestedCells.length) return;
    const master = state.tabs.find(isMasterTab);
    if (!master || !Number.isInteger(masterIdx) || masterIdx < 0 ||
        masterIdx >= master.cells.length) {
      targetTab.nestedCells[targetIndex] = [];
      return;
    }
    ensureNestedCells(master);
    targetTab.nestedCells[targetIndex] = cloneNestList(getCellNests(master, masterIdx));
  }

  /**
   * Drop Combined nest links whose nestIndex is out of range after a nest sync,
   * then live-rewrite remaining links for the cell.
   */
  function syncNestLinksAfterNestReplace(tabId, cellIndex) {
    const tab = state.tabs.find(function (t) { return t.id === tabId; });
    if (!tab) return;
    ensureNestedCells(tab);
    const nestCount = getCellNests(tab, cellIndex).length;
    const toRemove = [];
    state.confirmedLinks.forEach(function (link) {
      if (link.tabId !== tabId || link.cellIndex !== cellIndex) return;
      const ni = linkNestIndex(link);
      if (ni === null) return;
      if (ni < 0 || ni >= nestCount) toRemove.push(link);
    });
    if (toRemove.length) {
      const byScope = {};
      toRemove.forEach(function (link) {
        if (!byScope[link.scope]) byScope[link.scope] = [];
        byScope[link.scope].push(link);
      });
      Object.keys(byScope).forEach(function (scope) {
        removeLinksFromCombined(byScope[scope]);
      });
      const removeIds = {};
      for (let i = 0; i < toRemove.length; i++) removeIds[toRemove[i].id] = true;
      state.confirmedLinks = state.confirmedLinks.filter(function (link) {
        return !removeIds[link.id];
      });
    }
    liveSyncConfirmedLinksForCell(tabId, cellIndex, undefined, { silent: true });
  }

  /**
   * Push Master[masterIdx] nests onto every Master-origin part cell locked to that index.
   * Skips an in-progress Master cell edit session on that cell.
   */
  function syncLockedPartNestsFromMaster(masterIdx) {
    if (!Number.isInteger(masterIdx) || masterIdx < 0) return;
    const master = state.tabs.find(isMasterTab);
    if (!master || masterIdx >= master.cells.length) return;
    ensureNestedCells(master);
    const source = cloneNestList(getCellNests(master, masterIdx));
    state.tabs.forEach(function (tab) {
      if (!tab || isMasterTab(tab)) return;
      ensureCellLocks(tab);
      ensureNestedCells(tab);
      for (let i = 0; i < tab.cellLocks.length; i++) {
        const lock = tab.cellLocks[i];
        if (!lock || lock.masterOrigin !== true) continue;
        if (!Number.isInteger(lock.masterCellIndex) || lock.masterCellIndex !== masterIdx) continue;
        if (isMasterCellEditFor(tab, i)) continue;
        if (nestsContentEqual(tab.nestedCells[i], source)) continue;
        tab.nestedCells[i] = cloneNestList(source);
        syncNestLinksAfterNestReplace(tab.id, i);
      }
    });
  }

  /** Nested-cell arrays parallel to cells; missing/short pad with []. */
  function normalizeNestedCells(nested, length) {
    const needed = length > 0 ? length : 0;
    let out;
    if (Array.isArray(nested)) {
      out = nested.map(normalizeNestList);
    } else {
      out = emptyNestedCells(1, needed);
    }
    if (out.length < needed) {
      while (out.length < needed) out.push([]);
    } else if (out.length > needed) {
      out = out.slice(0, needed);
    }
    return out;
  }

  function normalizeCellLock(item) {
    if (!item || typeof item !== 'object') return null;
    if (item.masterOrigin !== true) return null;
    const out = { masterOrigin: true, locked: item.locked === false ? false : true };
    if (Number.isInteger(item.masterCellIndex) && item.masterCellIndex >= 0) {
      out.masterCellIndex = item.masterCellIndex;
    }
    return out;
  }

  function normalizeCellLocks(locks, length) {
    const needed = length > 0 ? length : 0;
    let out;
    if (Array.isArray(locks)) {
      out = locks.map(normalizeCellLock);
    } else {
      out = Array(needed).fill(null);
    }
    if (out.length < needed) {
      while (out.length < needed) out.push(null);
    } else if (out.length > needed) {
      out = out.slice(0, needed);
    }
    return out;
  }

  function emptyCellLocks(cols, rows) {
    return Array(cols * rows).fill(null);
  }

  function ensureCellLocks(tab) {
    if (!tab) return;
    const needed = (tab.cols || 0) * (tab.rows || 0);
    tab.cellLocks = normalizeCellLocks(tab.cellLocks, needed);
  }

  function getCellLock(tab, cellIndex) {
    if (!tab || !Number.isInteger(cellIndex) || cellIndex < 0) return null;
    ensureCellLocks(tab);
    if (cellIndex >= tab.cellLocks.length) return null;
    return tab.cellLocks[cellIndex];
  }

  function isCellMasterLocked(tab, cellIndex) {
    const lock = getCellLock(tab, cellIndex);
    return !!(lock && lock.masterOrigin === true && lock.locked !== false);
  }

  function setCellMasterLock(tab, cellIndex, masterCellIndex) {
    if (!tab || isMasterTab(tab)) return;
    ensureCellLocks(tab);
    if (cellIndex < 0 || cellIndex >= tab.cellLocks.length) return;
    const lock = { masterOrigin: true, locked: true };
    if (Number.isInteger(masterCellIndex) && masterCellIndex >= 0) {
      lock.masterCellIndex = masterCellIndex;
    } else {
      const text = tab.cells[cellIndex] || '';
      const found = findMasterCellIndexByText(text);
      if (found !== null) lock.masterCellIndex = found;
    }
    tab.cellLocks[cellIndex] = lock;
  }

  function clearCellMasterLock(tab, cellIndex) {
    if (!tab) return;
    ensureCellLocks(tab);
    if (cellIndex < 0 || cellIndex >= tab.cellLocks.length) return;
    tab.cellLocks[cellIndex] = null;
  }

  function swapCellLockIndices(tab, indexA, indexB) {
    if (!tab || indexA === indexB) return;
    ensureCellLocks(tab);
    if (indexA < 0 || indexB < 0 || indexA >= tab.cellLocks.length || indexB >= tab.cellLocks.length) return;
    const tmp = tab.cellLocks[indexA];
    tab.cellLocks[indexA] = tab.cellLocks[indexB];
    tab.cellLocks[indexB] = tmp;
  }

  function remapCellLocksAfterColumnAdd(tab, oldCols, newCols) {
    ensureCellLocks(tab);
    const oldLocks = tab.cellLocks.slice();
    const next = emptyCellLocks(newCols, tab.rows);
    for (let r = 0; r < tab.rows; r++) {
      for (let c = 0; c < oldCols; c++) {
        const oldIdx = r * oldCols + c;
        next[r * newCols + c] = oldIdx < oldLocks.length ? oldLocks[oldIdx] : null;
      }
    }
    tab.cellLocks = next;
  }

  function copyCellLockRow(tab, fromRow, toRow) {
    ensureCellLocks(tab);
    for (let col = 0; col < tab.cols; col++) {
      tab.cellLocks[toRow * tab.cols + col] = tab.cellLocks[fromRow * tab.cols + col];
    }
  }

  function clearCellLockRow(tab, rowIndex) {
    ensureCellLocks(tab);
    for (let col = 0; col < tab.cols; col++) {
      tab.cellLocks[rowIndex * tab.cols + col] = null;
    }
  }

  function nestPageText(nest, pageIndex) {
    if (!nest || !Array.isArray(nest.pages)) return '';
    if (pageIndex < 0 || pageIndex >= nest.pages.length) return '';
    return nest.pages[pageIndex] == null ? '' : String(nest.pages[pageIndex]);
  }

  function nestExportText(nest) {
    if (!nest || !Array.isArray(nest.pages)) return '';
    const parts = [];
    for (let i = 0; i < nest.pages.length; i++) {
      const t = (nest.pages[i] || '').trim();
      if (t) parts.push(t);
    }
    return parts.join(separatorValue(state.separators.row));
  }

  function nestHasContent(nest) {
    return !!nestExportText(nest);
  }

  /** Boolean array parallel to cells; missing/short arrays pad with false. */

  function makeTab(id, title, icon) {
    const cells = emptyCells(DEFAULT_COLS, DEFAULT_ROWS);
    const tab = {
      id: id,
      title: title,
      cols: DEFAULT_COLS,
      rows: DEFAULT_ROWS,
      cells: cells,
      nestedCells: emptyNestedCells(DEFAULT_COLS, DEFAULT_ROWS),
      cellLocks: emptyCellLocks(DEFAULT_COLS, DEFAULT_ROWS),
      cellPages: normalizeCellPages(null, cells, DEFAULT_COLS * DEFAULT_ROWS),
      cellShades: emptyCellShades(DEFAULT_COLS, DEFAULT_ROWS),
      pages: null,
      page: 0
    };
    // Bootstrap default page from empty grid (Combined keys filled on first flush).
    tab.pages = [captureTabPageGridOnly(tab)];
    tab.page = 0;
    tab.icon = isTabIconId(icon) ? icon : defaultIconForTab(tab);
    return tab;
  }

  function makeDefaultData() {
    return {
      tabs: [
        makeTab('master', 'Master', DEFAULT_MASTER_ICON),
        makeTab('tab-1', 'Part 1', DEFAULT_PART_ICON),
        makeTab('tab-2', 'Part 2', DEFAULT_PART_ICON),
        makeTab('tab-3', 'Part 3', DEFAULT_PART_ICON)
      ],
      activeTabId: 'tab-1',
      combinedPrompt: '',
      globalCombined: true,
      matchSourceOrder: false,
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

  function normalizeRowHeights(heights, rows) {
    if (!Array.isArray(heights) || heights.length !== rows) return null;
    const normalized = heights.map(function (height) {
      const value = Number(height);
      return Number.isFinite(value) && value >= MIN_ROW_HEIGHT ? value : null;
    });
    return normalized.every(function (height) { return height !== null; })
      ? normalized
      : null;
  }

  /** Keep tab.rowHeights length in sync when rows grow/shrink (only if already set). */
  function ensureRowHeightsLength(tab) {
    if (!tab || !Array.isArray(tab.rowHeights)) return;
    while (tab.rowHeights.length < tab.rows) tab.rowHeights.push(MIN_ROW_HEIGHT);
    if (tab.rowHeights.length > tab.rows) tab.rowHeights = tab.rowHeights.slice(0, tab.rows);
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
      const legacy = {
        id: t.id,
        title: t.title,
        cols: DEFAULT_COLS,
        rows: DEFAULT_ROWS,
        cells: cells,
        nestedCells: emptyNestedCells(DEFAULT_COLS, DEFAULT_ROWS),
        cellLocks: emptyCellLocks(DEFAULT_COLS, DEFAULT_ROWS),
        cellPages: normalizeCellPages(null, cells, DEFAULT_COLS * DEFAULT_ROWS),
        cellShades: emptyCellShades(DEFAULT_COLS, DEFAULT_ROWS),
        pages: null,
        page: 0
      };
      ensureTabPages(legacy);
      legacy.icon = isTabIconId(t.icon) ? t.icon : defaultIconForTab(legacy);
      return legacy;
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
    const rowHeights = normalizeRowHeights(t.rowHeights, rows);
    // sleptCells from older docs are ignored (sleep UI dropped).
    const nestedCells = normalizeNestedCells(t.nestedCells, needed);
    const cellLocks = normalizeCellLocks(t.cellLocks, needed);
    const cellPages = normalizeCellPages(t.cellPages, cells, needed);
    const cellShades = normalizeCellShades(t.cellShades, needed);
    for (let i = 0; i < needed; i++) {
      const entry = cellPages[i];
      if (entry && Array.isArray(entry.pages)) {
        cells[i] = entry.pages[entry.page] == null ? '' : String(entry.pages[entry.page]);
      }
    }
    const normalized = {
      id: t.id,
      title: t.title,
      cols: cols,
      rows: rows,
      cells: cells,
      nestedCells: nestedCells,
      cellLocks: cellLocks,
      cellPages: cellPages,
      cellShades: cellShades,
      pages: Array.isArray(t.pages) ? t.pages : null,
      page: Number.isInteger(t.page) ? t.page : 0
    };
    normalized.icon = isTabIconId(t.icon) ? t.icon : defaultIconForTab(normalized);
    if (columnWidths) normalized.columnWidths = columnWidths;
    if (rowHeights) normalized.rowHeights = rowHeights;
    // Persist/migrate part-tab pages; hydrate working grid from active page.
    ensureTabPages(normalized);
    const masterLike = normalized.id === 'master' ||
      (typeof normalized.title === 'string' && normalized.title.trim().toLowerCase() === 'master');
    if (!masterLike && normalized.pages && normalized.pages[normalized.page]) {
      applyTabPageGrid(normalized, normalized.pages[normalized.page]);
    } else if (masterLike) {
      // Keep Master as a single normalized page mirror of root fields.
      normalized.pages = [captureTabPageGridOnly(normalized)];
      normalized.page = 0;
    }
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

  function resolveLastPartTabId() {
    if (lastPartTabId && state.tabs.some(function (tab) {
      return tab.id === lastPartTabId && !isMasterTab(tab);
    })) {
      return lastPartTabId;
    }
    const current = activeTab();
    if (current && !isMasterTab(current)) {
      lastPartTabId = current.id;
      return lastPartTabId;
    }
    const parts = partTabs();
    lastPartTabId = parts.length ? parts[0].id : null;
    return lastPartTabId;
  }

  /** Tab that Tools (+ Row / + Column / Rename) and grid edits should target. */
  function toolsTargetTab() {
    if (!toolsTabActive) return activeTab();
    const id = resolveLastPartTabId();
    return state.tabs.find(function (tab) { return tab.id === id; }) || null;
  }

  function updateToolsChrome() {
    if (el.partSection) {
      el.partSection.classList.toggle('is-tools-view', toolsTabActive);
    }
    if (el.toolsPanel) {
      el.toolsPanel.hidden = !toolsTabActive;
    }
    const target = toolsTargetTab();
    if (el.toolsPanelHint) {
      el.toolsPanelHint.textContent = target
        ? ('Grid tools for active part: ' + target.title + ' (' + target.cols + '×' + target.rows + ')')
        : 'Grid tools — select a part tab first';
    }
    if (el.btnAddRow) el.btnAddRow.disabled = !target;
    if (el.btnAddCol) el.btnAddCol.disabled = !target;
    if (el.btnRename) el.btnRename.disabled = !target || isMasterTab(target);
  }

  function selectToolsTab() {
    resolveLastPartTabId();
    if (!lastPartTabId) {
      setStatus('Add a part tab before opening Tools', 'err');
      return;
    }
    const switchedPart = state.activeTabId !== lastPartTabId;
    if (toolsTabActive && !switchedPart) return;
    // Keep document activeTabId on the part Tools will mutate (not Master).
    if (switchedPart) state.activeTabId = lastPartTabId;
    toolsTabActive = true;
    focusedCell = null;
    clearStickyCellRange();
    closeMasterLibFilterMenu();
    masterLibFilterMenuOpen = false;
    renderTabs();
    if (switchedPart) {
      renderGrid();
      renderMasterLibrary();
      renderCombinedPrompt();
      scheduleSave();
    } else {
      updateToolsChrome();
    }
    const target = toolsTargetTab();
    setStatus('Tools — acting on ' + (target ? target.title : 'part'));
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
    // Flush each part tab's working grid + Combined keys into pages[page].
    state.tabs.forEach(function (tab) {
      if (!isMasterTab(tab)) flushTabPage(tab);
      else {
        ensureTabPages(tab);
        tab.pages = [captureTabPageGridOnly(tab)];
        tab.page = 0;
      }
    });
    return {
      tabs: state.tabs,
      activeTabId: state.activeTabId,
      combinedPrompt: state.combinedPrompt,
      globalCombined: state.globalCombined,
      matchSourceOrder: state.matchSourceOrder,
      partPrompts: Object.assign({}, state.partPrompts),
      separators: state.separators,
      confirmedLinks: state.confirmedLinks.map(function (link) {
        const out = {
          id: link.id,
          tabId: link.tabId,
          cellIndex: link.cellIndex,
          text: link.text,
          start: link.start,
          end: link.end,
          scope: link.scope
        };
        if (Number.isInteger(link.nestIndex) && link.nestIndex >= 0) {
          out.nestIndex = link.nestIndex;
        }
        if (link.masterOrigin === true) out.masterOrigin = true;
        if (link.masterOrigin === false) out.masterOrigin = false;
        if (link.locked === true) out.locked = true;
        if (link.locked === false) out.locked = false;
        if (Number.isInteger(link.masterCellIndex) && link.masterCellIndex >= 0) {
          out.masterCellIndex = link.masterCellIndex;
        }
        return out;
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
      matchSourceOrder: typeof data.matchSourceOrder === 'boolean' ? data.matchSourceOrder : false,
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
    const addr = cellAddressFromIndex(tab, focusedCell.index);
    const preview = (tab.cells[focusedCell.index] || '').trim().replace(/\s+/g, ' ');
    const short = preview.length > 40 ? preview.slice(0, 37) + '…' : preview;
    return addr + ' idx ' + focusedCell.index +
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
      { label: 'Match source order', value: state.matchSourceOrder ? 'yes' : 'no' },
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
    ensureNestedCells(tab);
    const hasCell = tab.cells.some(function (c, i) {
      if ((c || '').trim().length > 0) return true;
      const nests = tab.nestedCells[i] || [];
      return nests.some(function (n) { return nestHasContent(n); });
    });
    return hasCell || !!(state.partPrompts[tab.id] || '').trim();
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

  // Click/focus Combined is the only auto clipboard path (Copy button + Ctrl/Cmd+C kept).
  let lastCombinedAutoCopyStatusAt = 0;
  let lastCombinedActivateCopyKey = '';
  let lastCombinedActivateCopyAt = 0;
  let combinedActivateSkipCopy = false;

  /**
   * Last known Combined caret/selection as plain-text offsets into the active
   * Combined prompt. Used when checkbox / Append / Master Combined-add runs
   * after focus has left the Combined editor. null start ⇒ append at end.
   */
  let lastCombinedCaret = { start: null, end: null, scope: null };

  function readCombinedSelectionOffsets() {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return null;
    const range = sel.getRangeAt(0);
    if (!el.combined.contains(range.commonAncestorContainer) &&
        range.commonAncestorContainer !== el.combined) {
      return null;
    }
    function offsetAt(node, offset) {
      try {
        const pre = document.createRange();
        pre.selectNodeContents(el.combined);
        pre.setEnd(node, offset);
        return pre.toString().length;
      } catch (err) {
        return null;
      }
    }
    const start = offsetAt(range.startContainer, range.startOffset);
    const end = offsetAt(range.endContainer, range.endOffset);
    if (start == null || end == null) return null;
    return {
      start: Math.min(start, end),
      end: Math.max(start, end)
    };
  }

  function rememberCombinedCaretFromDom() {
    const offsets = readCombinedSelectionOffsets();
    if (!offsets) return;
    const scope = currentPromptScope();
    const textLen = (getPromptText(scope) || '').length;
    lastCombinedCaret = {
      start: Math.max(0, Math.min(offsets.start, textLen)),
      end: Math.max(0, Math.min(offsets.end, textLen)),
      scope: scope
    };
  }

  function snapCaretOutOfConfirmedLinks(scope, pos) {
    const links = linksForScope(scope);
    for (let i = 0; i < links.length; i++) {
      const link = links[i];
      if (pos > link.start && pos < link.end) return link.end;
    }
    return pos;
  }

  /**
   * Where Combined-add paths should insert: live Combined selection if focused,
   * else last known caret for this scope, else end (legacy append).
   * Non-collapsed selection is replaced. Collapsed caret inside a confirmed
   * segment snaps to after that segment so greens/links stay intact.
   */
  function resolveCombinedInsertRange(scope, textLen) {
    let start = null;
    let end = null;
    if (document.activeElement === el.combined) {
      const live = readCombinedSelectionOffsets();
      if (live) {
        start = live.start;
        end = live.end;
        lastCombinedCaret = { start: start, end: end, scope: scope };
      }
    }
    if (start == null && lastCombinedCaret.scope === scope &&
        typeof lastCombinedCaret.start === 'number') {
      start = lastCombinedCaret.start;
      end = typeof lastCombinedCaret.end === 'number'
        ? lastCombinedCaret.end
        : start;
    }
    if (start == null) {
      return { start: textLen, end: textLen };
    }
    start = Math.max(0, Math.min(start, textLen));
    end = Math.max(0, Math.min(end == null ? start : end, textLen));
    if (end < start) {
      const tmp = start;
      start = end;
      end = tmp;
    }
    if (start === end) {
      const snapped = snapCaretOutOfConfirmedLinks(scope, start);
      start = end = snapped;
    }
    return { start: start, end: end };
  }

  function setCombinedCaretOffset(offset) {
    if (!el.combined) return;
    const textLen = (getCombinedPlainText() || '').length;
    offset = Math.max(0, Math.min(typeof offset === 'number' ? offset : textLen, textLen));
    const sel = window.getSelection();
    if (!sel) return;

    function placeAt(node, pos) {
      const range = document.createRange();
      if (node.nodeType === Node.TEXT_NODE) {
        range.setStart(node, Math.max(0, Math.min(pos, (node.nodeValue || '').length)));
      } else {
        range.selectNodeContents(node);
        range.collapse(pos <= 0);
      }
      range.collapse(true);
      sel.removeAllRanges();
      sel.addRange(range);
    }

    let remaining = offset;
    const kids = el.combined.childNodes;
    if (!kids.length) {
      const range = document.createRange();
      range.selectNodeContents(el.combined);
      range.collapse(true);
      sel.removeAllRanges();
      sel.addRange(range);
      return;
    }
    for (let i = 0; i < kids.length; i++) {
      const node = kids[i];
      const len = (node.textContent || '').length;
      if (remaining > len) {
        remaining -= len;
        continue;
      }
      if (node.nodeType === Node.TEXT_NODE) {
        placeAt(node, remaining);
        return;
      }
      if (node.classList && node.classList.contains('confirmed-segment')) {
        const tn = node.firstChild;
        if (tn && tn.nodeType === Node.TEXT_NODE) placeAt(tn, remaining);
        else placeAt(node, remaining <= 0 ? 0 : 1);
        return;
      }
      const tn = node.firstChild;
      if (tn && tn.nodeType === Node.TEXT_NODE) placeAt(tn, remaining);
      else placeAt(el.combined, 1);
      return;
    }
    const range = document.createRange();
    range.selectNodeContents(el.combined);
    range.collapse(false);
    sel.removeAllRanges();
    sel.addRange(range);
  }

  function copyCombinedOnActivate(e) {
    if (!initialized) return;
    if (combinedActivateSkipCopy || (e && (e.ctrlKey || e.metaKey))) return;
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
      const nestIndex = toNonNegInt(item.nestIndex);
      const link = {
        id: item.id,
        tabId: item.tabId,
        cellIndex: cellIndex,
        text: item.text,
        start: start,
        end: end,
        scope: scope
      };
      if (nestIndex !== null) link.nestIndex = nestIndex;
      if (item.masterOrigin === true) link.masterOrigin = true;
      if (item.masterOrigin === false) link.masterOrigin = false;
      if (item.locked === true) link.locked = true;
      if (item.locked === false) link.locked = false;
      const masterCellIndex = toNonNegInt(item.masterCellIndex);
      if (masterCellIndex !== null) link.masterCellIndex = masterCellIndex;
      // Master-origin segments default to locked when flag omitted.
      if (link.masterOrigin === true && item.locked === undefined) link.locked = true;
      links.push(link);
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
      // Col1 filter-only: keep for checkbox UX; never re-anchor into Combined text.
      if (isCol1FilterOnlyLink(link)) {
        var expectedCol1 = sourceLinkText(link);
        if (expectedCol1 === null || expectedCol1 === '') return false;
        link.text = expectedCol1;
        link.start = 0;
        link.end = 0;
        return true;
      }
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

  /** nestIndex on a link: integer >= 0 means nest; otherwise parent cell. */
  function linkNestIndex(link) {
    if (!link) return null;
    return Number.isInteger(link.nestIndex) && link.nestIndex >= 0 ? link.nestIndex : null;
  }

  function sourceCellText(tabId, cellIndex) {
    const tab = state.tabs.find(function (t) { return t.id === tabId; });
    if (!tab || cellIndex < 0 || cellIndex >= tab.cells.length) return null;
    return cellParentText(tab, cellIndex);
  }

  function sourceNestText(tabId, cellIndex, nestIndex) {
    const tab = state.tabs.find(function (t) { return t.id === tabId; });
    if (!tab || cellIndex < 0 || cellIndex >= tab.cells.length) return null;
    if (!Number.isInteger(nestIndex) || nestIndex < 0) return null;
    const nests = getCellNests(tab, cellIndex);
    if (nestIndex >= nests.length) return null;
    return nestExportText(nests[nestIndex]);
  }

  function sourceLinkText(link) {
    if (!link) return null;
    const nestIdx = linkNestIndex(link);
    if (nestIdx === null) return sourceCellText(link.tabId, link.cellIndex);
    return sourceNestText(link.tabId, link.cellIndex, nestIdx);
  }

  function linkMatchesSource(link, segmentText) {
    const expected = sourceLinkText(link);
    if (expected === null) return false;
    return String(segmentText) === expected;
  }

  function isCellConfirmed(tabId, cellIndex) {
    // Greens / confirmed UI follow the *current* Combined scope (active tab or global).
    return isCellConfirmedInScope(tabId, cellIndex, currentPromptScope());
  }

  function isCellConfirmedInScope(tabId, cellIndex, scope) {
    const target = scope || currentPromptScope();
    return state.confirmedLinks.some(function (link) {
      return link.tabId === tabId && link.cellIndex === cellIndex &&
        link.scope === target && linkNestIndex(link) === null;
    });
  }

  function isNestConfirmed(tabId, cellIndex, nestIndex) {
    return isNestConfirmedInScope(tabId, cellIndex, nestIndex, currentPromptScope());
  }

  function isNestConfirmedInScope(tabId, cellIndex, nestIndex, scope) {
    if (!Number.isInteger(nestIndex) || nestIndex < 0) return false;
    const target = scope || currentPromptScope();
    return state.confirmedLinks.some(function (link) {
      return link.tabId === tabId && link.cellIndex === cellIndex &&
        link.scope === target && linkNestIndex(link) === nestIndex;
    });
  }

  function linksForNestInScope(tabId, cellIndex, nestIndex, scope) {
    const target = scope || currentPromptScope();
    return state.confirmedLinks.filter(function (link) {
      return link.scope === target && link.tabId === tabId &&
        link.cellIndex === cellIndex && linkNestIndex(link) === nestIndex;
    });
  }

  /**
   * True when this Master cell’s text is represented as a confirmed Combined
   * segment in the *current* Combined scope (active part / Global prompt).
   * Uses the same bar as cell greens: a live confirmed link whose segment
   * text exactly equals the Master cell (linkMatchesSource + link.text).
   * Avoids false greens from: other tabs’ Combined scopes, other parts’
   * segments in Global Combined, nest links whose parent happens to match,
   * stale links that no longer match their source, or substring coincidence.
   */
  function isMasterCellRepresentedInCombined(master, masterIdx) {
    if (!master || masterIdx < 0 || masterIdx >= master.cells.length) return false;
    const scope = currentPromptScope();
    const text = cellParentText(master, masterIdx);
    if (!text) return false;
    // Direct Master-cell confirmed link in this Combined scope.
    if (isCellConfirmedInScope(master.id, masterIdx, scope)) return true;

    const current = activeTab();
    const activePartId = current && !isMasterTab(current) ? current.id : null;

    return state.confirmedLinks.some(function (link) {
      if (link.scope !== scope) return false;
      // Col1 filter-only links are not Combined segments.
      if (isCol1FilterOnlyLink(link)) return false;
      // When Global Combined is shared, only count segments from the active
      // part (or Master itself) — not other tabs’ contributions.
      if (activePartId && link.tabId !== activePartId && link.tabId !== master.id) {
        return false;
      }
      if (!linkMatchesSource(link, link.text)) return false;
      return String(link.text) === text;
    });
  }

  /**
   * Active Master unlock/edit session (null when idle).
   * Combined: { kind:'combined', linkId, originalText, scope }
   * Part cell: { kind:'cell', tabId, cellIndex, originalText }
   */
  let masterSegmentEdit = null;
  let masterSegmentDialogOpen = false;

  function isMasterCellEditSession() {
    return !!(masterSegmentEdit && masterSegmentEdit.kind === 'cell');
  }

  function isMasterCombinedEditSession() {
    return !!(masterSegmentEdit && masterSegmentEdit.kind === 'combined');
  }

  function updateMasterEditDialogCopy() {
    const title = document.getElementById('master-segment-dialog-title');
    const body = el.masterSegmentDialog
      ? el.masterSegmentDialog.querySelector('.modal-body')
      : null;
    if (isMasterCellEditSession()) {
      if (title) title.textContent = 'Master cell edited';
      if (body) {
        body.textContent =
          'Overwrite the Master library cell (parent text + nests) and other tabs using that Master cell, keep this as local cell text/nests only, or cancel and discard the edit?';
      }
    } else {
      if (title) title.textContent = 'Master segment edited';
      if (body) {
        body.textContent =
          'Overwrite the Master library (and other tabs using that Master text), keep this as local Combined text only, or cancel and discard the edit?';
      }
    }
  }

  /** True when a live selection / target range intersects a locked Master Combined span. */
  function rangeTouchesLockedMasterSegment(range) {
    if (!range || !el.combined) return null;
    const lockedNodes = el.combined.querySelectorAll('.confirmed-segment.master-segment-locked');
    for (let i = 0; i < lockedNodes.length; i++) {
      const node = lockedNodes[i];
      try {
        if (range.intersectsNode(node)) return node;
      } catch (err) { /* detached */ }
    }
    return null;
  }

  function selectionTouchesLockedMasterSegment() {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return null;
    const hit = rangeTouchesLockedMasterSegment(sel.getRangeAt(0));
    if (hit) return hit;
    const nodes = [sel.anchorNode, sel.focusNode];
    for (let i = 0; i < nodes.length; i++) {
      const node = nodes[i];
      const elNode = node && node.nodeType === Node.TEXT_NODE ? node.parentElement : node;
      const locked = elNode && elNode.closest
        ? elNode.closest('.confirmed-segment.master-segment-locked')
        : null;
      if (locked && el.combined.contains(locked)) return locked;
    }
    return null;
  }

  /** Locked span that Backspace/Delete at a collapsed caret would remove or bite into. */
  function lockedSegmentAdjacentToCaret(key) {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount || !sel.isCollapsed) return null;
    const range = sel.getRangeAt(0);
    let node = range.startContainer;
    let offset = range.startOffset;

    function lockedFromNode(n) {
      if (!n) return null;
      if (n.nodeType === Node.ELEMENT_NODE && n.classList &&
          n.classList.contains('confirmed-segment') &&
          n.classList.contains('master-segment-locked') &&
          el.combined.contains(n)) {
        return n;
      }
      if (n.nodeType === Node.TEXT_NODE && n.parentElement) {
        return lockedFromNode(n.parentElement);
      }
      return null;
    }

    if (key === 'Backspace') {
      if (node.nodeType === Node.TEXT_NODE && offset === 0) {
        return lockedFromNode(node.previousSibling) ||
          lockedFromNode(node.parentElement && node.parentElement.previousSibling);
      }
      if (node.nodeType === Node.ELEMENT_NODE && offset > 0) {
        return lockedFromNode(node.childNodes[offset - 1]);
      }
    }
    if (key === 'Delete') {
      if (node.nodeType === Node.TEXT_NODE && offset === (node.textContent || '').length) {
        return lockedFromNode(node.nextSibling) ||
          lockedFromNode(node.parentElement && node.parentElement.nextSibling);
      }
      if (node.nodeType === Node.ELEMENT_NODE && offset < node.childNodes.length) {
        return lockedFromNode(node.childNodes[offset]);
      }
    }
    return null;
  }

  function findMasterCellIndexByText(text) {
    const master = state.tabs.find(isMasterTab);
    if (!master || text === null || text === undefined) return null;
    const want = String(text);
    for (let i = 0; i < master.cells.length; i++) {
      if (cellParentText(master, i) === want) return i;
    }
    return null;
  }

  function findMasterOriginForText(text) {
    const idx = findMasterCellIndexByText(text);
    if (idx === null) return null;
    return { masterCellIndex: idx };
  }

  function isMasterOriginLink(link) {
    return !!(link && link.masterOrigin === true);
  }

  function isMasterLockedLink(link) {
    if (!isMasterOriginLink(link)) return false;
    return link.locked !== false;
  }

  function copyLinkMasterFields(from, to) {
    if (!from || !to) return to;
    if (from.masterOrigin === true) to.masterOrigin = true;
    if (from.masterOrigin === false) to.masterOrigin = false;
    if (from.locked === true) to.locked = true;
    if (from.locked === false) to.locked = false;
    if (Number.isInteger(from.masterCellIndex) && from.masterCellIndex >= 0) {
      to.masterCellIndex = from.masterCellIndex;
    }
    return to;
  }

  function applyMasterOriginToNewLink(link, piece) {
    if (!link) return link;
    const master = state.tabs.find(isMasterTab);
    if (master && piece && piece.tabId === master.id) {
      link.masterOrigin = true;
      link.locked = true;
      if (Number.isInteger(piece.cellIndex) && piece.cellIndex >= 0) {
        link.masterCellIndex = piece.cellIndex;
      }
      return link;
    }
    // Prefer explicit Master-lock metadata on the part cell (insert-from-Master).
    if (piece && piece.tabId && Number.isInteger(piece.cellIndex)) {
      const srcTab = state.tabs.find(function (t) { return t.id === piece.tabId; });
      const cellLock = srcTab && !isMasterTab(srcTab) ? getCellLock(srcTab, piece.cellIndex) : null;
      if (cellLock && cellLock.masterOrigin === true) {
        link.masterOrigin = true;
        link.locked = true;
        if (Number.isInteger(cellLock.masterCellIndex) && cellLock.masterCellIndex >= 0) {
          link.masterCellIndex = cellLock.masterCellIndex;
        }
        return link;
      }
    }
    const origin = findMasterOriginForText(link.text);
    if (origin) {
      link.masterOrigin = true;
      link.locked = true;
      link.masterCellIndex = origin.masterCellIndex;
    }
    return link;
  }

  function linkById(linkId, scope) {
    const target = scope || currentPromptScope();
    return state.confirmedLinks.find(function (link) {
      return link.id === linkId && link.scope === target;
    }) || null;
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
      if (isCellConfirmedInScope(tab.id, idx, scope)) return true;
      const nests = getCellNests(tab, idx);
      for (let n = 0; n < nests.length; n++) {
        if (isNestConfirmedInScope(tab.id, idx, n, scope)) return true;
      }
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

  /** Swap Combined-link cell indices for two cells (content moves with the link). */
  function swapConfirmedCellIndices(tabId, indexA, indexB) {
    if (indexA === indexB) return;
    state.confirmedLinks.forEach(function (link) {
      if (link.tabId !== tabId) return;
      if (link.cellIndex === indexA) link.cellIndex = indexB;
      else if (link.cellIndex === indexB) link.cellIndex = indexA;
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

  /** Bump Combined links on rows >= fromRow down by one (after a push-down grow). */
  function shiftConfirmedRowsFrom(tabId, cols, fromRow) {
    state.confirmedLinks.forEach(function (link) {
      if (link.tabId !== tabId) return;
      const row = Math.floor(link.cellIndex / cols);
      const col = link.cellIndex % cols;
      if (row >= fromRow) {
        link.cellIndex = (row + 1) * cols + col;
      }
    });
  }

  /** Remap Combined links when rows are reordered (↑ extract+insert). */
  function remapConfirmedRowsByOrder(tabId, cols, oldToNew) {
    state.confirmedLinks.forEach(function (link) {
      if (link.tabId !== tabId) return;
      const row = Math.floor(link.cellIndex / cols);
      const col = link.cellIndex % cols;
      if (row < 0 || row >= oldToNew.length) return;
      link.cellIndex = oldToNew[row] * cols + col;
    });
  }

  /** When Master rows reorder, keep part-tab locks + Combined masterCellIndex in sync. */
  function remapMasterCellIndices(oldToNew, cols) {
    if (!oldToNew || !cols) return;
    function mapIndex(idx) {
      if (!Number.isInteger(idx) || idx < 0) return idx;
      const row = Math.floor(idx / cols);
      const col = idx % cols;
      if (row < 0 || row >= oldToNew.length) return idx;
      const next = oldToNew[row];
      if (!Number.isInteger(next) || next < 0) return idx;
      return next * cols + col;
    }
    state.tabs.forEach(function (tab) {
      if (isMasterTab(tab)) return;
      ensureCellLocks(tab);
      for (let i = 0; i < tab.cellLocks.length; i++) {
        const lock = tab.cellLocks[i];
        if (lock && Number.isInteger(lock.masterCellIndex) && lock.masterCellIndex >= 0) {
          lock.masterCellIndex = mapIndex(lock.masterCellIndex);
        }
      }
    });
    state.confirmedLinks.forEach(function (link) {
      if (Number.isInteger(link.masterCellIndex) && link.masterCellIndex >= 0) {
        link.masterCellIndex = mapIndex(link.masterCellIndex);
      }
    });
  }










  function ensureNestedCells(tab) {
    if (!tab) return;
    tab.nestedCells = normalizeNestedCells(tab.nestedCells, tab.cols * tab.rows);
  }

  function getCellNests(tab, cellIndex) {
    if (!tab || cellIndex < 0) return [];
    ensureNestedCells(tab);
    if (cellIndex >= tab.nestedCells.length) return [];
    return tab.nestedCells[cellIndex];
  }

  function setCellNests(tab, cellIndex, nests) {
    if (!tab || cellIndex < 0 || cellIndex >= tab.cells.length) return;
    ensureNestedCells(tab);
    tab.nestedCells[cellIndex] = normalizeNestList(nests);
  }

  /** Parent cell text only (nests have their own Combined checkboxes). */
  function cellParentText(tab, cellIndex) {
    if (!tab || cellIndex < 0 || cellIndex >= tab.cells.length) return '';
    return (tab.cells[cellIndex] || '').trim();
  }

  /**
   * Legacy helper name: parent text for Combined. Nests append separately via
   * their own confirmed links / nestIndex.
   */
  function cellCombinedText(tab, cellIndex) {
    return cellParentText(tab, cellIndex);
  }

  function cellHasExportableContent(tab, cellIndex) {
    if (cellParentText(tab, cellIndex)) return true;
    const nests = getCellNests(tab, cellIndex);
    for (let i = 0; i < nests.length; i++) {
      if (nestHasContent(nests[i])) return true;
    }
    return false;
  }

  /**
   * Confirmed pieces for one cell stack: parent (nestIndex omitted) then each
   * non-empty nest, separated by the row separator.
   */
  function cellConfirmedPieces(tab, cellIndex, opts) {
    const pieces = [];
    if (!tab || cellIndex < 0 || cellIndex >= tab.cells.length) return pieces;
    const scope = (opts && opts.scope) || currentPromptScope();
    const skipConfirmed = !!(opts && opts.skipConfirmed);
    const rowSep = separatorValue(state.separators.row);

    const parent = cellParentText(tab, cellIndex);
    // Part-tab Column A is filter-only — never emit parent text into Combined.
    if (parent && !isPartTabCol1Cell(tab, cellIndex) &&
        !(skipConfirmed && isCellConfirmedInScope(tab.id, cellIndex, scope))) {
      pieces.push({
        type: 'confirmed',
        text: parent,
        tabId: tab.id,
        cellIndex: cellIndex
      });
    }

    const nests = getCellNests(tab, cellIndex);
    for (let n = 0; n < nests.length; n++) {
      const nestText = nestExportText(nests[n]);
      if (!nestText) continue;
      if (skipConfirmed && isNestConfirmedInScope(tab.id, cellIndex, n, scope)) continue;
      if (pieces.length) pieces.push({ type: 'plain', text: rowSep });
      pieces.push({
        type: 'confirmed',
        text: nestText,
        tabId: tab.id,
        cellIndex: cellIndex,
        nestIndex: n
      });
    }
    return pieces;
  }

  function swapNestedCellIndices(tab, indexA, indexB) {
    if (!tab || indexA === indexB) return;
    ensureNestedCells(tab);
    if (indexA < 0 || indexB < 0 || indexA >= tab.nestedCells.length || indexB >= tab.nestedCells.length) return;
    const tmp = tab.nestedCells[indexA];
    tab.nestedCells[indexA] = tab.nestedCells[indexB];
    tab.nestedCells[indexB] = tmp;
  }

  function remapNestedAfterColumnAdd(tab, oldCols, newCols) {
    const oldNested = Array.isArray(tab.nestedCells) ? tab.nestedCells : [];
    const next = emptyNestedCells(newCols, tab.rows);
    for (let r = 0; r < tab.rows; r++) {
      for (let c = 0; c < oldCols; c++) {
        const oldIdx = r * oldCols + c;
        next[r * newCols + c] = oldIdx < oldNested.length
          ? normalizeNestList(oldNested[oldIdx])
          : [];
      }
    }
    tab.nestedCells = next;
  }

  function copyNestedRow(tab, fromRow, toRow) {
    ensureNestedCells(tab);
    for (let col = 0; col < tab.cols; col++) {
      tab.nestedCells[toRow * tab.cols + col] = normalizeNestList(tab.nestedCells[fromRow * tab.cols + col]);
    }
  }

  function clearNestedRow(tab, rowIndex) {
    ensureNestedCells(tab);
    for (let col = 0; col < tab.cols; col++) {
      tab.nestedCells[rowIndex * tab.cols + col] = [];
    }
  }

  /**
   * Resolve cellIndex for nest + from the clicked control's cell-wrap (row/col),
   * then data-idx, then closure fallback. Avoids wrong-cell attaches when Fit
   * row heights stretched the gutter and shoved + toward the next row visually.
   */
  function cellIndexFromNestAddEvent(ev, fallbackIdx) {
    const tab = activeTab();
    const btn = ev && ev.currentTarget;
    if (btn && typeof btn.closest === 'function') {
      const wrap = btn.closest('.cell-wrap');
      if (wrap && tab) {
        const row = parseInt(wrap.dataset.row, 10);
        const col = parseInt(wrap.dataset.col, 10);
        if (!Number.isNaN(row) && !Number.isNaN(col) &&
            row >= 0 && col >= 0 && row < tab.rows && col < tab.cols) {
          return row * tab.cols + col;
        }
        const ta = wrap.querySelector('textarea.cell');
        if (ta) {
          const fromTa = parseInt(ta.dataset.idx, 10);
          if (!Number.isNaN(fromTa) && fromTa >= 0) return fromTa;
        }
      }
    }
    if (btn && btn.dataset) {
      const fromBtn = parseInt(btn.dataset.idx, 10);
      if (!Number.isNaN(fromBtn) && fromBtn >= 0) return fromBtn;
      if (tab) {
        const row = parseInt(btn.dataset.row, 10);
        const col = parseInt(btn.dataset.col, 10);
        if (!Number.isNaN(row) && !Number.isNaN(col) &&
            row >= 0 && col >= 0 && row < tab.rows && col < tab.cols) {
          return row * tab.cols + col;
        }
      }
    }
    return Number.isInteger(fallbackIdx) ? fallbackIdx : -1;
  }

  /** Re-measure and persist one row after nest add/remove so nests stay in-cell. */
  function refitRowHeightAt(row) {
    const tab = activeTab();
    if (!tab || !el.cellGrid || row < 0 || row >= tab.rows) return;
    const wraps = el.cellGrid.querySelectorAll('.cell-wrap[data-row="' + row + '"]');
    // Clear pinned heights first so empty/short rows can shrink, not only grow.
    for (let i = 0; i < wraps.length; i++) clearPinnedHeightsForFitMeasure(wraps[i]);
    const controls = el.cellGrid.querySelector('.row-controls[data-row="' + row + '"]');
    if (controls) clearPinnedHeightsForFitMeasure(controls);
    for (let i = 0; i < wraps.length; i++) prepareWrapWidthsForFitMeasure(wraps[i]);
    let maxH = MIN_ROW_HEIGHT;
    for (let i = 0; i < wraps.length; i++) {
      const stackH = measureCellStackHeight(wraps[i], tab);
      if (stackH > maxH) maxH = stackH;
    }
    applyHeightToGridRow(row, maxH);
    // Second pass after apply (wrap widths final); may grow OR shrink slightly.
    let maxH2 = MIN_ROW_HEIGHT;
    for (let i = 0; i < wraps.length; i++) {
      const stackH = measureCellStackHeight(wraps[i], tab);
      if (stackH > maxH2) maxH2 = stackH;
    }
    if (maxH2 !== maxH) {
      maxH = maxH2;
      applyHeightToGridRow(row, maxH);
    }
    if (Array.isArray(tab.rowHeights)) {
      ensureRowHeightsLength(tab);
      tab.rowHeights[row] = maxH;
    }
  }

  function addNestedCell(cellIndex) {
    const tab = activeTab();
    if (!tab || cellIndex < 0 || cellIndex >= tab.cells.length) return;
    if (isCellMasterEditBlocked(tab, cellIndex)) {
      setStatus('Locked Master cell — double-click to unlock before editing', 'err');
      return;
    }
    pushHistory();
    ensureNestedCells(tab);
    tab.nestedCells[cellIndex] = tab.nestedCells[cellIndex].concat([makeEmptyNest()]);
    const nestIndex = tab.nestedCells[cellIndex].length - 1;
    const row = Math.floor(cellIndex / tab.cols);
    if (isMasterTab(tab)) syncLockedPartNestsFromMaster(cellIndex);
    renderGrid();
    // Persisted rowHeights from before the nest would clip / spill; grow this row.
    refitRowHeightAt(row);
    scheduleSave();
    const nestTa = el.cellGrid.querySelector(
      'textarea.cell-nest-input[data-idx="' + cellIndex + '"][data-nest="' + nestIndex + '"]'
    );
    if (nestTa) nestTa.focus();
    setStatus('Added nest under ' + cellAddressFromIndex(tab, cellIndex));
  }

  function remapNestLinksAfterRemove(tabId, cellIndex, removedNestIndex) {
    const byScope = {};
    const toRemove = [];
    state.confirmedLinks.forEach(function (link) {
      if (link.tabId !== tabId || link.cellIndex !== cellIndex) return;
      const ni = linkNestIndex(link);
      if (ni === null) return;
      if (ni === removedNestIndex) toRemove.push(link);
    });
    toRemove.forEach(function (link) {
      if (!byScope[link.scope]) byScope[link.scope] = [];
      byScope[link.scope].push(link);
    });
    Object.keys(byScope).forEach(function (scope) {
      removeLinksFromCombined(byScope[scope]);
    });
    const removeIds = {};
    for (let i = 0; i < toRemove.length; i++) removeIds[toRemove[i].id] = true;
    state.confirmedLinks = state.confirmedLinks.filter(function (link) {
      return !removeIds[link.id];
    });
    state.confirmedLinks.forEach(function (link) {
      if (link.tabId !== tabId || link.cellIndex !== cellIndex) return;
      const ni = linkNestIndex(link);
      if (ni !== null && ni > removedNestIndex) link.nestIndex = ni - 1;
    });
  }

  function removeNestedCell(cellIndex, nestIndex) {
    const tab = activeTab();
    if (!tab || cellIndex < 0 || cellIndex >= tab.cells.length) return;
    if (isCellMasterEditBlocked(tab, cellIndex)) {
      setStatus('Locked Master cell — double-click to unlock before editing', 'err');
      return;
    }
    ensureNestedCells(tab);
    const nests = tab.nestedCells[cellIndex];
    if (nestIndex < 0 || nestIndex >= nests.length) return;
    pushHistory();
    remapNestLinksAfterRemove(tab.id, cellIndex, nestIndex);
    nests.splice(nestIndex, 1);
    revalidateLinksForCell(tab.id, cellIndex, { silent: true });
    if (isMasterTab(tab)) syncLockedPartNestsFromMaster(cellIndex);
    const row = Math.floor(cellIndex / tab.cols);
    renderGrid();
    refitRowHeightAt(row);
    renderCombinedPrompt();
    applyConfirmedCellHighlights();
    applyAppendCheckedState();
    scheduleSave();
    setStatus('Removed nest from ' + cellAddressFromIndex(tab, cellIndex));
  }

  function refreshNestConfirmedUi(idx, nestIdx, nestEl) {
    const tab = activeTab();
    if (!tab) return;
    const confirmed = isNestConfirmed(tab.id, idx, nestIdx);
    const nestRow = nestEl && nestEl.closest ? nestEl.closest('.cell-nest') : null;
    if (nestRow) nestRow.classList.toggle('cell-confirmed', confirmed);
    if (nestEl && nestEl.classList) nestEl.classList.toggle('cell-confirmed', confirmed);
    applyAppendCheckedState();
    if (isMasterTab(tab)) renderMasterLibrary();
  }

  function onNestInput(e) {
    const tab = activeTab();
    if (!tab) return;
    const idx = parseInt(e.target.dataset.idx, 10);
    const nestIdx = parseInt(e.target.dataset.nest, 10);
    if (Number.isNaN(idx) || idx < 0 || idx >= tab.cells.length) return;
    if (Number.isNaN(nestIdx) || nestIdx < 0) return;
    // Locked Master-inserted cell nests: read-only until double-click unlock.
    if (isCellMasterEditBlocked(tab, idx)) {
      ensureNestedCells(tab);
      const nestLocked = tab.nestedCells[idx] && tab.nestedCells[idx][nestIdx];
      const pageLocked = nestLocked ? (nestLocked.page || 0) : 0;
      e.target.value = nestLocked && nestLocked.pages
        ? (nestLocked.pages[pageLocked] || '')
        : '';
      setStatus('Locked Master cell — double-click to unlock before editing', 'err');
      return;
    }
    ensureNestedCells(tab);
    const nest = tab.nestedCells[idx][nestIdx];
    if (!nest) return;
    const page = nest.page || 0;
    if (page < 0 || page >= nest.pages.length) return;
    if (stickyCellRange) clearStickyCellRange();
    clearKeyboardCellRange();
    pushHistory({ coalesce: true });
    nest.pages[page] = e.target.value;
    liveSyncConfirmedLinksForCell(tab.id, idx, nestIdx);
    refreshNestConfirmedUi(idx, nestIdx, e.target);
    // Master tab nest edits push to locked part cells pointing at this Master index.
    if (isMasterTab(tab)) syncLockedPartNestsFromMaster(idx);
    maybeLiveRefitRowFromInput(Math.floor(idx / tab.cols));
    scheduleSave();
  }

  function setNestPage(cellIndex, nestIndex, pageIndex, opts) {
    const tab = activeTab();
    if (!tab || cellIndex < 0 || cellIndex >= tab.cells.length) return;
    ensureNestedCells(tab);
    const nest = tab.nestedCells[cellIndex][nestIndex];
    if (!nest || !nest.pages.length) return;
    let next = pageIndex;
    if (next < 0) next = 0;
    if (next >= nest.pages.length) next = nest.pages.length - 1;
    if (nest.page === next && !(opts && opts.forceRender)) return;
    // Page flip is view state — still persist so reopen lands on same page.
    nest.page = next;
    if (opts && opts.skipRender) {
      scheduleSave();
      return;
    }
    renderGrid();
    scheduleSave();
    const nestTa = el.cellGrid.querySelector(
      'textarea.cell-nest-input[data-idx="' + cellIndex + '"][data-nest="' + nestIndex + '"]'
    );
    if (nestTa) nestTa.focus();
  }

  function addNestPage(cellIndex, nestIndex) {
    const tab = activeTab();
    if (!tab || cellIndex < 0 || cellIndex >= tab.cells.length) return;
    if (isCellMasterEditBlocked(tab, cellIndex)) {
      setStatus('Locked Master cell — double-click to unlock before editing', 'err');
      return;
    }
    ensureNestedCells(tab);
    const nest = tab.nestedCells[cellIndex][nestIndex];
    if (!nest) return;
    pushHistory();
    nest.pages.push('');
    nest.page = nest.pages.length - 1;
    if (isMasterTab(tab)) syncLockedPartNestsFromMaster(cellIndex);
    renderGrid();
    refitRowHeightAt(Math.floor(cellIndex / tab.cols));
    scheduleSave();
    const nestTa = el.cellGrid.querySelector(
      'textarea.cell-nest-input[data-idx="' + cellIndex + '"][data-nest="' + nestIndex + '"]'
    );
    if (nestTa) nestTa.focus();
    setStatus('Added nest page (' + nest.pages.length + ')');
  }

  function removeNestPage(cellIndex, nestIndex) {
    const tab = activeTab();
    if (!tab || cellIndex < 0 || cellIndex >= tab.cells.length) return;
    if (isCellMasterEditBlocked(tab, cellIndex)) {
      setStatus('Locked Master cell — double-click to unlock before editing', 'err');
      return;
    }
    ensureNestedCells(tab);
    const nest = tab.nestedCells[cellIndex][nestIndex];
    if (!nest || nest.pages.length <= 1) {
      removeNestedCell(cellIndex, nestIndex);
      return;
    }
    pushHistory();
    const removeAt = nest.page || 0;
    nest.pages.splice(removeAt, 1);
    if (nest.page >= nest.pages.length) nest.page = nest.pages.length - 1;
    liveSyncConfirmedLinksForCell(tab.id, cellIndex, nestIndex, { silent: true });
    if (isMasterTab(tab)) syncLockedPartNestsFromMaster(cellIndex);
    renderGrid();
    renderCombinedPrompt();
    applyConfirmedCellHighlights();
    applyAppendCheckedState();
    scheduleSave();
    setStatus('Removed nest page (' + nest.pages.length + ' left)');
  }

  function stepNestPage(cellIndex, nestIndex, delta) {
    const tab = activeTab();
    if (!tab) return;
    ensureNestedCells(tab);
    const nest = tab.nestedCells[cellIndex] && tab.nestedCells[cellIndex][nestIndex];
    if (!nest) return;
    const cur = nest.page || 0;
    const next = cur + delta;
    if (next < 0) return;
    if (next >= nest.pages.length) {
      // Past the end → create a new page (clear UX for multi-page nests).
      addNestPage(cellIndex, nestIndex);
      return;
    }
    setNestPage(cellIndex, nestIndex, next);
  }



  function setCellPage(cellIndex, pageIndex, opts) {
    const tab = activeTab();
    if (!tab || cellIndex < 0 || cellIndex >= tab.cells.length) return;
    ensureCellPages(tab);
    if (!tab.cellPages[cellIndex] || !tab.cellPages[cellIndex].pages.length) return;
    // Flush live textarea into the page we are leaving.
    const liveTa = el.cellGrid
      ? el.cellGrid.querySelector('textarea.cell[data-idx="' + cellIndex + '"]')
      : null;
    if (liveTa && document.activeElement === liveTa) {
      writeCellCurrentPage(tab, cellIndex, liveTa.value);
    } else {
      // Keep cells[i] mirrored into current page before flip.
      writeCellCurrentPage(tab, cellIndex, tab.cells[cellIndex] || '');
    }
    // writeCellCurrentPage → ensureCellPages replaces cellPages[]; re-fetch entry.
    const entry = tab.cellPages[cellIndex];
    if (!entry || !entry.pages.length) return;
    let next = pageIndex;
    if (next < 0) next = 0;
    if (next >= entry.pages.length) next = entry.pages.length - 1;
    if (entry.page === next && !(opts && opts.forceRender)) return;
    entry.page = next;
    syncCellTextFromPages(tab, cellIndex);
    liveSyncConfirmedLinksForCell(tab.id, cellIndex, null, { silent: true });
    if (isMasterTab(tab)) syncLockedPartPagesFromMaster(cellIndex);
    if (opts && opts.skipRender) {
      scheduleSave();
      return;
    }
    renderGrid();
    renderCombinedPrompt();
    applyConfirmedCellHighlights();
    applyAppendCheckedState();
    scheduleSave();
    const ta = el.cellGrid.querySelector('textarea.cell[data-idx="' + cellIndex + '"]');
    if (ta) ta.focus();
  }

  function addCellPage(cellIndex, mode) {
    const tab = activeTab();
    if (!tab || cellIndex < 0 || cellIndex >= tab.cells.length) return;
    if (isCellMasterEditBlocked(tab, cellIndex)) {
      setStatus('Locked Master cell — double-click to unlock before editing', 'err');
      return;
    }
    const copyMode = mode === 'copy';
    ensureCellPages(tab);
    if (!tab.cellPages[cellIndex]) return;
    pushHistory();
    writeCellCurrentPage(tab, cellIndex, tab.cells[cellIndex] || '');
    // writeCellCurrentPage → ensureCellPages replaces cellPages[]; re-fetch before mutate.
    const entry = tab.cellPages[cellIndex];
    if (!entry) return;
    const cur = entry.pages[entry.page] == null ? '' : String(entry.pages[entry.page]);
    entry.pages.push(copyMode ? cur : '');
    entry.page = entry.pages.length - 1;
    syncCellTextFromPages(tab, cellIndex);
    liveSyncConfirmedLinksForCell(tab.id, cellIndex, null, { silent: true });
    if (isMasterTab(tab)) syncLockedPartPagesFromMaster(cellIndex);
    renderGrid();
    refitRowHeightAt(Math.floor(cellIndex / tab.cols));
    scheduleSave();
    const ta = el.cellGrid.querySelector('textarea.cell[data-idx="' + cellIndex + '"]');
    if (ta) ta.focus();
    setStatus(copyMode
      ? ('Added cell page — copy of current (' + entry.pages.length + ')')
      : ('Added cell page — blank (' + entry.pages.length + ')'));
  }

  function removeCellPage(cellIndex) {
    const tab = activeTab();
    if (!tab || cellIndex < 0 || cellIndex >= tab.cells.length) return;
    if (isCellMasterEditBlocked(tab, cellIndex)) {
      setStatus('Locked Master cell — double-click to unlock before editing', 'err');
      return;
    }
    ensureCellPages(tab);
    const entry = tab.cellPages[cellIndex];
    if (!entry || entry.pages.length <= 1) {
      setStatus('Cell already has only one page', 'err');
      return;
    }
    pushHistory();
    const removeAt = entry.page || 0;
    entry.pages.splice(removeAt, 1);
    if (entry.page >= entry.pages.length) entry.page = entry.pages.length - 1;
    syncCellTextFromPages(tab, cellIndex);
    liveSyncConfirmedLinksForCell(tab.id, cellIndex, null, { silent: true });
    if (isMasterTab(tab)) syncLockedPartPagesFromMaster(cellIndex);
    renderGrid();
    renderCombinedPrompt();
    applyConfirmedCellHighlights();
    applyAppendCheckedState();
    refitRowHeightAt(Math.floor(cellIndex / tab.cols));
    scheduleSave();
    setStatus('Removed cell page (' + entry.pages.length + ' left)');
  }

  function stepCellPage(cellIndex, delta) {
    const tab = activeTab();
    if (!tab) return;
    ensureCellPages(tab);
    const entry = tab.cellPages[cellIndex];
    if (!entry) return;
    const cur = entry.page || 0;
    const next = cur + delta;
    if (next < 0 || next >= entry.pages.length) return;
    setCellPage(cellIndex, next);
  }

  /* ── Part-tab pages (full grid snapshots per part) ───────────────────── */

  function captureTabConfirmedKeys(tabId) {
    const keys = [];
    const seen = Object.create(null);
    state.confirmedLinks.forEach(function (link) {
      if (!link || link.tabId !== tabId) return;
      const nest = Number.isInteger(link.nestIndex) && link.nestIndex >= 0
        ? link.nestIndex
        : null;
      const id = nest === null ? String(link.cellIndex) : (link.cellIndex + ':' + nest);
      if (seen[id]) return;
      seen[id] = true;
      const item = { cellIndex: link.cellIndex };
      if (nest !== null) item.nestIndex = nest;
      keys.push(item);
    });
    return keys;
  }

  function normalizeTabConfirmedKeys(raw) {
    if (!Array.isArray(raw)) return [];
    const out = [];
    const seen = Object.create(null);
    for (let i = 0; i < raw.length; i++) {
      const item = raw[i];
      if (!item || typeof item !== 'object') continue;
      if (!Number.isInteger(item.cellIndex) || item.cellIndex < 0) continue;
      const nest = Number.isInteger(item.nestIndex) && item.nestIndex >= 0
        ? item.nestIndex
        : null;
      const id = nest === null ? String(item.cellIndex) : (item.cellIndex + ':' + nest);
      if (seen[id]) continue;
      seen[id] = true;
      const entry = { cellIndex: item.cellIndex };
      if (nest !== null) entry.nestIndex = nest;
      out.push(entry);
    }
    return out;
  }

  /** Snapshot of the part tab's editable grid + Combined membership for one page. */
  function captureTabPage(tab) {
    if (!tab) return null;
    ensureNestedCells(tab);
    ensureCellLocks(tab);
    ensureCellPages(tab);
    ensureCellShades(tab);
    const needed = (tab.cols || 0) * (tab.rows || 0);
    const prev = (Array.isArray(tab.pages) && tab.pages[tab.page]) ? tab.pages[tab.page] : null;
    const snap = {
      cols: tab.cols,
      rows: tab.rows,
      cells: tab.cells.map(function (c) { return typeof c === 'string' ? c : ''; }),
      nestedCells: tab.nestedCells.map(function (list) { return cloneNestList(list); }),
      cellLocks: normalizeCellLocks(tab.cellLocks, needed),
      cellPages: tab.cellPages.map(function (entry) { return cloneCellPagesEntry(entry); }),
      cellShades: normalizeCellShades(tab.cellShades, needed),
      partPrompt: typeof state.partPrompts[tab.id] === 'string' ? state.partPrompts[tab.id] : '',
      confirmed: captureTabConfirmedKeys(tab.id),
      name: (prev && typeof prev.name === 'string') ? prev.name : ''
    };
    const cw = normalizeColumnWidths(tab.columnWidths, tab.cols);
    const rh = normalizeRowHeights(tab.rowHeights, tab.rows);
    if (cw) snap.columnWidths = cw.slice();
    if (rh) snap.rowHeights = rh.slice();
    return snap;
  }

  function normalizeTabPage(raw) {
    if (!raw || typeof raw !== 'object') return null;
    let cols = Number.isInteger(raw.cols) && raw.cols > 0 ? raw.cols : DEFAULT_COLS;
    let rows = Number.isInteger(raw.rows) && raw.rows > 0 ? raw.rows : DEFAULT_ROWS;
    let cells;
    if (Array.isArray(raw.cells)) {
      cells = raw.cells.map(function (c) { return typeof c === 'string' ? c : ''; });
    } else {
      cells = emptyCells(cols, rows);
    }
    const needed = cols * rows;
    if (cells.length < needed) cells = cells.concat(Array(needed - cells.length).fill(''));
    else if (cells.length > needed) cells = cells.slice(0, needed);
    const nestedCells = normalizeNestedCells(raw.nestedCells, needed);
    const cellLocks = normalizeCellLocks(raw.cellLocks, needed);
    const cellPages = normalizeCellPages(raw.cellPages, cells, needed);
    const cellShades = normalizeCellShades(raw.cellShades, needed);
    for (let i = 0; i < needed; i++) {
      const entry = cellPages[i];
      if (entry && Array.isArray(entry.pages)) {
        cells[i] = entry.pages[entry.page] == null ? '' : String(entry.pages[entry.page]);
      }
    }
    const snap = {
      cols: cols,
      rows: rows,
      cells: cells,
      nestedCells: nestedCells,
      cellLocks: cellLocks,
      cellPages: cellPages,
      cellShades: cellShades,
      partPrompt: typeof raw.partPrompt === 'string' ? raw.partPrompt : '',
      confirmed: normalizeTabConfirmedKeys(raw.confirmed),
      name: typeof raw.name === 'string' ? raw.name : ''
    };
    const cw = normalizeColumnWidths(raw.columnWidths, cols);
    const rh = normalizeRowHeights(raw.rowHeights, rows);
    if (cw) snap.columnWidths = cw;
    if (rh) snap.rowHeights = rh;
    return snap;
  }

  /** Deep copy a page snapshot. opts.unlock clears Master cell locks on the copy. */
  function cloneTabPage(snap, opts) {
    const base = normalizeTabPage(snap);
    if (!base) return null;
    if (opts && opts.unlock) {
      base.cellLocks = emptyCellLocks(base.cols, base.rows);
    }
    return base;
  }

  function applyTabPageGrid(tab, snap) {
    if (!tab || !snap) return;
    const page = normalizeTabPage(snap);
    if (!page) return;
    tab.cols = page.cols;
    tab.rows = page.rows;
    tab.cells = page.cells.slice();
    tab.nestedCells = page.nestedCells.map(function (list) { return cloneNestList(list); });
    tab.cellLocks = normalizeCellLocks(page.cellLocks, page.cols * page.rows);
    tab.cellPages = page.cellPages.map(function (entry) { return cloneCellPagesEntry(entry); });
    tab.cellShades = normalizeCellShades(page.cellShades, page.cols * page.rows);
    if (page.columnWidths) tab.columnWidths = page.columnWidths.slice();
    else delete tab.columnWidths;
    if (page.rowHeights) tab.rowHeights = page.rowHeights.slice();
    else delete tab.rowHeights;
    ensureCellPages(tab);
  }

  function removeTabCombinedContribution(tabId) {
    const doomed = state.confirmedLinks.filter(function (link) {
      return link && link.tabId === tabId;
    });
    if (!doomed.length) return;
    const byScope = {};
    doomed.forEach(function (link) {
      if (!byScope[link.scope]) byScope[link.scope] = [];
      byScope[link.scope].push(link);
    });
    Object.keys(byScope).forEach(function (scope) {
      removeLinksFromCombined(byScope[scope]);
    });
  }

  function restoreTabCombinedFromPage(tab, snap) {
    if (!tab || !snap) return;
    removeTabCombinedContribution(tab.id);
    if (typeof snap.partPrompt === 'string' && snap.partPrompt) {
      state.partPrompts[tab.id] = snap.partPrompt;
    } else {
      delete state.partPrompts[tab.id];
    }
    const keys = normalizeTabConfirmedKeys(snap.confirmed);
    // Parents first, then nests — preserves a sensible Combined order.
    keys.sort(function (a, b) {
      if (a.cellIndex !== b.cellIndex) return a.cellIndex - b.cellIndex;
      const an = Number.isInteger(a.nestIndex) ? a.nestIndex : -1;
      const bn = Number.isInteger(b.nestIndex) ? b.nestIndex : -1;
      return an - bn;
    });
    for (let i = 0; i < keys.length; i++) {
      const item = keys[i];
      if (Number.isInteger(item.nestIndex)) {
        ensureNestConfirmedState(item.cellIndex, item.nestIndex, true, { quiet: true });
      } else {
        ensureCellConfirmedState(item.cellIndex, true, { quiet: true, col1Cascade: false });
      }
    }
  }

  function flushLiveCellInputs(tab) {
    if (!tab || !el.cellGrid) return;
    const tas = el.cellGrid.querySelectorAll('textarea.cell[data-idx]');
    for (let i = 0; i < tas.length; i++) {
      const ta = tas[i];
      const idx = parseInt(ta.dataset.idx, 10);
      if (Number.isNaN(idx) || idx < 0 || idx >= tab.cells.length) continue;
      if (document.activeElement === ta || (ta.value || '') !== (tab.cells[idx] || '')) {
        writeCellCurrentPage(tab, idx, ta.value);
      }
    }
  }

  function flushTabPage(tab) {
    if (!tab || isMasterTab(tab)) return;
    ensureTabPages(tab);
    if (tab === activeTab()) flushLiveCellInputs(tab);
    tab.pages[tab.page] = captureTabPage(tab);
  }

  /** Grid-only page snapshot (no Combined keys) — safe during normalize/load. */
  function captureTabPageGridOnly(tab) {
    if (!tab) return null;
    ensureNestedCells(tab);
    ensureCellLocks(tab);
    ensureCellPages(tab);
    ensureCellShades(tab);
    const needed = (tab.cols || 0) * (tab.rows || 0);
    const prev = (Array.isArray(tab.pages) && tab.pages[tab.page]) ? tab.pages[tab.page] : null;
    const snap = {
      cols: tab.cols,
      rows: tab.rows,
      cells: tab.cells.map(function (c) { return typeof c === 'string' ? c : ''; }),
      nestedCells: tab.nestedCells.map(function (list) { return cloneNestList(list); }),
      cellLocks: normalizeCellLocks(tab.cellLocks, needed),
      cellPages: tab.cellPages.map(function (entry) { return cloneCellPagesEntry(entry); }),
      cellShades: normalizeCellShades(tab.cellShades, needed),
      partPrompt: '',
      confirmed: [],
      name: (prev && typeof prev.name === 'string') ? prev.name : ''
    };
    const cw = normalizeColumnWidths(tab.columnWidths, tab.cols);
    const rh = normalizeRowHeights(tab.rowHeights, tab.rows);
    if (cw) snap.columnWidths = cw.slice();
    if (rh) snap.rowHeights = rh.slice();
    return snap;
  }

  function ensureTabPages(tab) {
    if (!tab) return;
    // Master stays single-page (no chrome); still normalize shape for persist.
    if (!Array.isArray(tab.pages) || !tab.pages.length) {
      tab.pages = [captureTabPageGridOnly(tab)];
      tab.page = 0;
      return;
    }
    const normalized = [];
    for (let i = 0; i < tab.pages.length; i++) {
      const page = normalizeTabPage(tab.pages[i]);
      if (page) normalized.push(page);
    }
    if (!normalized.length) normalized.push(captureTabPageGridOnly(tab));
    tab.pages = normalized;
    let page = Number.isInteger(tab.page) ? tab.page : 0;
    if (page < 0) page = 0;
    if (page >= tab.pages.length) page = tab.pages.length - 1;
    tab.page = page;
  }

  function hydrateTabFromActivePage(tab) {
    if (!tab || isMasterTab(tab)) return;
    ensureTabPages(tab);
    applyTabPageGrid(tab, tab.pages[tab.page]);
  }

  function setTabPage(pageIndex, opts) {
    const tab = activeTab();
    if (!tab || isMasterTab(tab) || toolsTabActive) return;
    ensureTabPages(tab);
    flushLiveCellInputs(tab);
    tab.pages[tab.page] = captureTabPage(tab);
    let next = pageIndex;
    if (next < 0) next = 0;
    if (next >= tab.pages.length) next = tab.pages.length - 1;
    if (tab.page === next && !(opts && opts.forceRender)) {
      renderPartTabPageChrome();
      return;
    }
    tab.page = next;
    applyTabPageGrid(tab, tab.pages[tab.page]);
    restoreTabCombinedFromPage(tab, tab.pages[tab.page]);
    focusedCell = null;
    clearStickyCellRange();
    clearKeyboardCellRange();
    clearPartSearch({ keepInputs: true, keepQuery: true });
    renderTabs();
    renderGrid();
    renderCombinedPrompt();
    applyConfirmedCellHighlights();
    applyAppendCheckedState();
    renderMasterLibrary();
    renderPartTabPageChrome();
    scheduleSave();
  }

  function makeBlankTabPage(tab) {
    if (!tab) return null;
    const cols = tab.cols > 0 ? tab.cols : DEFAULT_COLS;
    const rows = tab.rows > 0 ? tab.rows : DEFAULT_ROWS;
    const cells = emptyCells(cols, rows);
    const snap = {
      cols: cols,
      rows: rows,
      cells: cells,
      nestedCells: emptyNestedCells(cols, rows),
      cellLocks: emptyCellLocks(cols, rows),
      cellPages: normalizeCellPages(null, cells, cols * rows),
      cellShades: emptyCellShades(cols, rows),
      partPrompt: '',
      confirmed: [],
      name: ''
    };
    // Keep column widths so blank page matches the current grid layout.
    const cw = normalizeColumnWidths(tab.columnWidths, cols);
    if (cw) snap.columnWidths = cw.slice();
    return normalizeTabPage(snap);
  }

  function addTabPage(mode) {
    const tab = activeTab();
    if (!tab || isMasterTab(tab) || toolsTabActive) return;
    const copyMode = mode === 'copy';
    ensureTabPages(tab);
    pushHistory();
    flushLiveCellInputs(tab);
    tab.pages[tab.page] = captureTabPage(tab);
    let nextPage;
    if (copyMode) {
      // Ash: copy current but not locked — unlock Master cell locks on the copy.
      nextPage = cloneTabPage(tab.pages[tab.page], { unlock: true });
    } else {
      nextPage = makeBlankTabPage(tab);
    }
    if (!nextPage) return;
    // Fresh page name — do not reuse the source page's label.
    nextPage.name = 'Page ' + (tab.pages.length + 1);
    tab.pages.push(nextPage);
    tab.page = tab.pages.length - 1;
    applyTabPageGrid(tab, nextPage);
    restoreTabCombinedFromPage(tab, nextPage);
    focusedCell = null;
    clearStickyCellRange();
    clearKeyboardCellRange();
    renderTabs();
    renderGrid();
    renderCombinedPrompt();
    applyConfirmedCellHighlights();
    applyAppendCheckedState();
    renderMasterLibrary();
    renderPartTabPageChrome();
    scheduleSave();
    setStatus(copyMode
      ? ('Added part page (' + tab.pages.length + ') — unlocked copy of previous')
      : ('Added part page (' + tab.pages.length + ') — blank'));
  }

  function removeTabPage() {
    const tab = activeTab();
    if (!tab || isMasterTab(tab) || toolsTabActive) return;
    ensureTabPages(tab);
    if (tab.pages.length <= 1) {
      setStatus('Part already has only one page', 'err');
      return;
    }
    pushHistory();
    flushLiveCellInputs(tab);
    const removeAt = tab.page || 0;
    tab.pages.splice(removeAt, 1);
    if (tab.page >= tab.pages.length) tab.page = tab.pages.length - 1;
    applyTabPageGrid(tab, tab.pages[tab.page]);
    restoreTabCombinedFromPage(tab, tab.pages[tab.page]);
    focusedCell = null;
    clearStickyCellRange();
    clearKeyboardCellRange();
    renderTabs();
    renderGrid();
    renderCombinedPrompt();
    applyConfirmedCellHighlights();
    applyAppendCheckedState();
    renderMasterLibrary();
    renderPartTabPageChrome();
    scheduleSave();
    setStatus('Removed part page (' + tab.pages.length + ' left)');
  }

  function stepTabPage(delta) {
    const tab = activeTab();
    if (!tab || isMasterTab(tab)) return;
    ensureTabPages(tab);
    const cur = tab.page || 0;
    const next = cur + delta;
    if (next < 0 || next >= tab.pages.length) return;
    setTabPage(next);
  }

  function defaultTabPageName(pageIndex) {
    return 'Page ' + (pageIndex + 1);
  }

  function tabPageDisplayName(snap, pageIndex) {
    if (snap && typeof snap.name === 'string' && snap.name.trim()) return snap.name.trim();
    return defaultTabPageName(pageIndex);
  }

  function setTabPageName(name, opts) {
    const tab = activeTab();
    if (!tab || isMasterTab(tab) || toolsTabActive) return;
    ensureTabPages(tab);
    const page = tab.pages[tab.page];
    if (!page) return;
    const next = typeof name === 'string' ? name.trim() : '';
    if ((page.name || '') === next) {
      if (opts && opts.forceRender) renderPartTabPageChrome();
      return;
    }
    pushHistory();
    page.name = next;
    renderPartTabPageChrome();
    scheduleSave();
  }

  function renderPartTabPageChrome() {
    const host = el.partTabPageChrome;
    if (!host) return;
    const tab = activeTab();
    const show = !!(tab && !isMasterTab(tab) && !toolsTabActive);
    host.hidden = !show;
    // Keep focus in the name field across re-renders when the user is typing.
    const active = document.activeElement;
    const keepNameFocus = !!(active && active.classList &&
      active.classList.contains('part-tab-page-name') && host.contains(active));
    const keepSelStart = keepNameFocus ? active.selectionStart : null;
    const keepSelEnd = keepNameFocus ? active.selectionEnd : null;
    const keepValue = keepNameFocus ? active.value : null;
    host.textContent = '';
    if (!show) return;
    ensureTabPages(tab);
    const pageCount = tab.pages.length;
    const page = Math.min(Math.max(tab.page || 0, 0), pageCount - 1);
    tab.page = page;
    const snap = tab.pages[page];

    const chrome = document.createElement('div');
    chrome.className = 'part-tab-page-chrome-inner';
    chrome.setAttribute('role', 'group');
    chrome.setAttribute('aria-label', 'Part pages');

    const prevBtn = document.createElement('button');
    prevBtn.type = 'button';
    prevBtn.className = 'part-tab-page-btn';
    prevBtn.textContent = '◀';
    prevBtn.title = 'Previous part page';
    prevBtn.setAttribute('aria-label', 'Previous part page');
    prevBtn.disabled = page <= 0;
    prevBtn.addEventListener('click', function (ev) {
      ev.preventDefault();
      stepTabPage(-1);
    });

    const label = document.createElement('span');
    label.className = 'part-tab-page-label';
    label.textContent = (page + 1) + '/' + pageCount;
    label.title = tabPageDisplayName(snap, page) + ' — page ' + (page + 1) + ' of ' + pageCount;

    const nextBtn = document.createElement('button');
    nextBtn.type = 'button';
    nextBtn.className = 'part-tab-page-btn';
    nextBtn.textContent = '▶';
    nextBtn.title = 'Next part page';
    nextBtn.setAttribute('aria-label', 'Next part page');
    nextBtn.disabled = page >= pageCount - 1;
    nextBtn.addEventListener('click', function (ev) {
      ev.preventDefault();
      stepTabPage(1);
    });

    const addBtn = document.createElement('button');
    addBtn.type = 'button';
    addBtn.className = 'part-tab-page-btn part-tab-page-add';
    addBtn.textContent = '+';
    addBtn.title = 'Add part page';
    addBtn.setAttribute('aria-label', 'Add part page');
    addBtn.addEventListener('click', function (ev) {
      ev.preventDefault();
      openAddPageChoiceMenu(addBtn, { kind: 'part' });
    });

    chrome.appendChild(prevBtn);
    chrome.appendChild(label);
    chrome.appendChild(nextBtn);
    chrome.appendChild(addBtn);

    if (pageCount > 1) {
      const removeBtn = document.createElement('button');
      removeBtn.type = 'button';
      removeBtn.className = 'part-tab-page-btn part-tab-page-remove';
      removeBtn.textContent = '×';
      removeBtn.title = 'Remove current part page';
      removeBtn.setAttribute('aria-label', 'Remove current part page');
      removeBtn.addEventListener('click', function (ev) {
        ev.preventDefault();
        removeTabPage();
      });
      chrome.appendChild(removeBtn);
    }

    const nameInput = document.createElement('input');
    nameInput.type = 'text';
    nameInput.className = 'part-tab-page-name';
    nameInput.value = keepNameFocus && keepValue != null
      ? keepValue
      : (snap && typeof snap.name === 'string' ? snap.name : '');
    nameInput.placeholder = defaultTabPageName(page);
    nameInput.title = 'Name this part page';
    nameInput.setAttribute('aria-label', 'Part page name');
    nameInput.spellcheck = false;
    nameInput.addEventListener('keydown', function (ev) {
      if (ev.key === 'Enter') {
        ev.preventDefault();
        nameInput.blur();
      } else if (ev.key === 'Escape') {
        ev.preventDefault();
        nameInput.value = snap && typeof snap.name === 'string' ? snap.name : '';
        nameInput.blur();
      }
    });
    nameInput.addEventListener('blur', function () {
      setTabPageName(nameInput.value);
    });
    // Typing should not steal focus via other key handlers / save churn.
    nameInput.addEventListener('mousedown', function (ev) { ev.stopPropagation(); });
    nameInput.addEventListener('click', function (ev) { ev.stopPropagation(); });

    host.appendChild(chrome);
    host.appendChild(nameInput);

    if (keepNameFocus) {
      nameInput.focus();
      try {
        if (keepSelStart != null && keepSelEnd != null) {
          nameInput.setSelectionRange(keepSelStart, keepSelEnd);
        }
      } catch (err) { /* ignore */ }
    }
  }



  function closeCellShadePicker() {
    const open = document.querySelectorAll('.cell-shade-picker');
    for (let i = 0; i < open.length; i++) open[i].remove();
    document.removeEventListener('pointerdown', onShadePickerOutside, true);
  }

  function onShadePickerOutside(ev) {
    const t = ev.target;
    if (t && typeof t.closest === 'function' &&
        (t.closest('.cell-shade-picker') || t.closest('.cell-shade-btn'))) {
      return;
    }
    closeCellShadePicker();
  }

  function openCellShadePicker(anchorBtn, cellIndex) {
    const tab = activeTab();
    if (!tab || cellIndex < 0 || cellIndex >= tab.cells.length) return;
    closeCellShadePicker();
    closeAddPageChoiceMenu();
    const picker = document.createElement('div');
    picker.className = 'cell-shade-picker';
    picker.setAttribute('role', 'listbox');
    picker.setAttribute('aria-label', 'Cell shade');
    const options = [
      { id: null, label: 'None', swatch: 'none' },
      { id: 'green', label: 'Light green', swatch: 'green' },
      { id: 'yellow', label: 'Light yellow', swatch: 'yellow' },
      { id: 'red', label: 'Light red', swatch: 'red' }
    ];
    const current = getCellShade(tab, cellIndex);
    options.forEach(function (opt) {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'cell-shade-option cell-shade-option-' + opt.swatch;
      btn.title = opt.label;
      btn.setAttribute('aria-label', opt.label);
      btn.setAttribute('role', 'option');
      if ((opt.id || null) === (current || null)) btn.classList.add('is-selected');
      btn.addEventListener('click', function (ev) {
        ev.preventDefault();
        ev.stopPropagation();
        pushHistory();
        setCellShade(tab, cellIndex, opt.id);
        closeCellShadePicker();
        const wrap = el.cellGrid
          ? el.cellGrid.querySelector('.cell-wrap textarea.cell[data-idx="' + cellIndex + '"]')
          : null;
        const wrapEl = wrap && wrap.closest ? wrap.closest('.cell-wrap') : null;
        if (wrapEl) {
          wrapEl.classList.remove('cell-shade-green', 'cell-shade-yellow', 'cell-shade-red');
          if (opt.id) wrapEl.classList.add('cell-shade-' + opt.id);
        } else {
          renderGrid();
        }
        scheduleSave();
        setStatus(opt.id ? ('Cell shade: ' + opt.label) : 'Cell shade cleared', 'ok');
      });
      picker.appendChild(btn);
    });
    const host = anchorBtn.closest('.cell-corner-tools') || anchorBtn.parentNode;
    host.appendChild(picker);
    document.addEventListener('pointerdown', onShadePickerOutside, true);
  }

  function closeAddPageChoiceMenu() {
    const open = document.querySelectorAll('.add-page-choice-menu');
    for (let i = 0; i < open.length; i++) open[i].remove();
    document.removeEventListener('pointerdown', onAddPageChoiceOutside, true);
  }

  function onAddPageChoiceOutside(ev) {
    const t = ev.target;
    if (t && typeof t.closest === 'function' &&
        (t.closest('.add-page-choice-menu') ||
         t.closest('.cell-page-add') ||
         t.closest('.part-tab-page-add'))) {
      return;
    }
    closeAddPageChoiceMenu();
  }

  /**
   * Shared "Add blank" / "Copy current" menu for cell-page + and part-tab page +.
   * Matches the compact picker UX used elsewhere (shade / tab actions).
   */
  function openAddPageChoiceMenu(anchorBtn, opts) {
    if (!anchorBtn) return;
    opts = opts || {};
    const kind = opts.kind === 'part' ? 'part' : 'cell';
    const cellIndex = Number.isInteger(opts.cellIndex) ? opts.cellIndex : -1;
    if (kind === 'cell') {
      const tab = activeTab();
      if (!tab || cellIndex < 0 || cellIndex >= tab.cells.length) return;
      if (isCellMasterEditBlocked(tab, cellIndex)) {
        setStatus('Locked Master cell — double-click to unlock before editing', 'err');
        return;
      }
    } else {
      const tab = activeTab();
      if (!tab || isMasterTab(tab) || toolsTabActive) return;
    }
    closeAddPageChoiceMenu();
    closeCellShadePicker();
    const menu = document.createElement('div');
    menu.className = 'add-page-choice-menu';
    menu.setAttribute('role', 'menu');
    menu.setAttribute('aria-label', kind === 'part' ? 'Add part page' : 'Add cell page');

    function addChoice(label, mode) {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'add-page-choice-option';
      btn.setAttribute('role', 'menuitem');
      btn.textContent = label;
      btn.addEventListener('click', function (ev) {
        ev.preventDefault();
        ev.stopPropagation();
        closeAddPageChoiceMenu();
        if (kind === 'part') addTabPage(mode);
        else addCellPage(cellIndex, mode);
      });
      menu.appendChild(btn);
    }
    addChoice('Add blank', 'blank');
    addChoice('Copy current', 'copy');

    if (kind === 'cell') {
      const host = anchorBtn.closest('.cell-corner-tools') ||
        anchorBtn.closest('.cell-page-chrome') ||
        anchorBtn.parentNode;
      host.appendChild(menu);
    } else {
      document.body.appendChild(menu);
      const rect = anchorBtn.getBoundingClientRect();
      const pad = 6;
      menu.style.position = 'fixed';
      menu.style.left = '0px';
      menu.style.top = '0px';
      const size = menu.getBoundingClientRect();
      let left = rect.left;
      let top = rect.bottom + 4;
      left = Math.max(pad, Math.min(left, window.innerWidth - size.width - pad));
      top = Math.max(pad, Math.min(top, window.innerHeight - size.height - pad));
      menu.style.left = left + 'px';
      menu.style.top = top + 'px';
      menu.style.zIndex = '1200';
    }
    document.addEventListener('pointerdown', onAddPageChoiceOutside, true);
  }

  /**
   * Click+drag paint for Combined / nest checkboxes.
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

  /**
   * Ash Col1 Combined cascade (part tabs only): when checking or unchecking
   * Column A, also (un)check same-row Col B and every other Col A cell with
   * the same trimmed value V (exact string match, same as Values filter) plus
   * each of those rows' Col B. Column A is filter-only — its text never enters
   * Combined (zero-length confirmed links for checkbox UX); Col B text still
   * appends. Nest Combined checkboxes are independent. Master skipped.
   *
   * Exclusive filter (Ash): checking a Col1 option also unchecks every other
   * Col1 value (≠ V) and those rows' Col B — radio-style across categories.
   * Same-value ON cascade and exclusive OFF run together on check; uncheck
   * only reverses the same-value set.
   *
   * Cascade is the DEFAULT for part-tab Column A Combined toggles; pass
   * opts.col1Cascade === false to opt out (Master insert).
   */
  function collectCol1CombinedCascadeIndices(tab, cellIndex) {
    const cols = tab.cols || 0;
    const cellCount = tab.cells ? tab.cells.length : 0;
    const rowsFromCells = cols > 0 ? Math.ceil(cellCount / cols) : 0;
    const rows = Math.max(tab.rows || 0, rowsFromCells);
    const value = cellParentText(tab, cellIndex);
    const out = [];
    const seen = Object.create(null);
    function add(idx) {
      if (!Number.isInteger(idx) || idx < 0 || idx >= cellCount) return;
      if (seen[idx]) return;
      seen[idx] = true;
      out.push(idx);
    }
    // Always keep the origin cell even when value/cols short-circuit.
    add(cellIndex);
    if (!value || cols < 2) return out;
    for (let r = 0; r < rows; r++) {
      const col1 = r * cols;
      if (cellParentText(tab, col1) !== value) continue;
      add(col1);
      add(col1 + 1);
    }
    return out;
  }

  /**
   * Exclusive Col1 filter OFF set: every Column A cell whose trimmed value
   * differs from the origin option, plus those rows' Column B. Same-value
   * rows are left for the ON cascade. Empty Col1 rows skipped.
   */
  function collectCol1ExclusiveOffIndices(tab, cellIndex) {
    const cols = tab.cols || 0;
    const cellCount = tab.cells ? tab.cells.length : 0;
    const rowsFromCells = cols > 0 ? Math.ceil(cellCount / cols) : 0;
    const rows = Math.max(tab.rows || 0, rowsFromCells);
    const value = cellParentText(tab, cellIndex);
    const out = [];
    const seen = Object.create(null);
    function add(idx) {
      if (!Number.isInteger(idx) || idx < 0 || idx >= cellCount) return;
      if (seen[idx]) return;
      seen[idx] = true;
      out.push(idx);
    }
    if (!value || cols < 1) return out;
    for (let r = 0; r < rows; r++) {
      const col1 = r * cols;
      const other = cellParentText(tab, col1);
      if (!other || other === value) continue;
      add(col1);
      if (cols >= 2) add(col1 + 1);
    }
    return out;
  }

  function shouldCol1CombinedCascade(tab, cellIndex) {
    if (!tab || isMasterTab(tab)) return false;
    const cols = tab.cols || 0;
    if (cols < 2) return false;
    if (!Number.isInteger(cellIndex) || cellIndex < 0 || cellIndex >= tab.cells.length) return false;
    return (cellIndex % cols) === 0;
  }


  /** Part-tab Column A (filter-only): never contributes text to Combined. */
  function isPartTabCol1Cell(tab, cellIndex) {
    if (!tab || isMasterTab(tab)) return false;
    const cols = tab.cols || 0;
    if (cols < 1) return false;
    if (!Number.isInteger(cellIndex) || cellIndex < 0 || cellIndex >= (tab.cells ? tab.cells.length : 0)) {
      return false;
    }
    return (cellIndex % cols) === 0;
  }

  /**
   * Confirmed parent link on part-tab Column A — checkbox / cascade UX only.
   * Nest links on Col1 are normal Combined content (not filter-only).
   */
  function isCol1FilterOnlyLink(link) {
    if (!link || linkNestIndex(link) !== null) return false;
    const tab = state.tabs.find(function (t) { return t.id === link.tabId; });
    return isPartTabCol1Cell(tab, link.cellIndex);
  }

  /** Mark Col1 confirmed for filter UX without inserting cell text into Combined. */
  function ensureCol1FilterLink(tab, cellIndex, scope) {
    const target = scope || currentPromptScope();
    if (!tab || !isPartTabCol1Cell(tab, cellIndex)) return false;
    if (isCellConfirmedInScope(tab.id, cellIndex, target)) return false;
    const text = cellParentText(tab, cellIndex);
    if (!text) return false;
    state.confirmedLinks.push({
      id: linkUid(),
      tabId: tab.id,
      cellIndex: cellIndex,
      text: text,
      start: 0,
      end: 0,
      scope: target
    });
    return true;
  }

  /**
   * Rebuild Combined scopes so legacy Col1 confirmed segments leave the prompt
   * text while keeping zero-length filter links for checkbox state.
   */
  function stripCol1TextFromAllCombinedScopes() {
    const seen = Object.create(null);
    const scopes = [];
    for (let i = 0; i < state.confirmedLinks.length; i++) {
      const link = state.confirmedLinks[i];
      if (!isCol1FilterOnlyLink(link)) continue;
      if (!link.scope || seen[link.scope]) continue;
      seen[link.scope] = true;
      scopes.push(link.scope);
    }
    let changed = false;
    for (let i = 0; i < scopes.length; i++) {
      const target = scopes[i];
      const links = state.confirmedLinks.filter(function (link) {
        return link.scope === target;
      });
      if (!links.length) continue;
      const ordered = state.matchSourceOrder
        ? links.slice().sort(compareLinksBySourceOrder)
        : links.slice().sort(function (a, b) { return a.start - b.start; });
      if (rejoinConfirmedLinks(target, ordered)) changed = true;
      else {
        // Even when text unchanged, pin Col1 filter links to zero-length.
        for (let j = 0; j < state.confirmedLinks.length; j++) {
          const link = state.confirmedLinks[j];
          if (link.scope !== target || !isCol1FilterOnlyLink(link)) continue;
          if (link.start !== 0 || link.end !== 0) {
            link.start = 0;
            link.end = 0;
            changed = true;
          }
        }
      }
    }
    return changed;
  }


  /** Resolve whether Col1 cascade applies. Default ON for Col1; false opts out. */
  function resolveCol1Cascade(tab, cellIndex, opts) {
    if (!shouldCol1CombinedCascade(tab, cellIndex)) return false;
    if (opts && opts.col1Cascade === false) return false;
    // Explicit true, or omitted (default on for Col1 Combined toggles).
    return true;
  }

  function ensureCellConfirmedState(cellIndex, wantOn, opts) {
    const quiet = !!(opts && opts.quiet);
    const tab = activeTab();
    if (!tab || cellIndex < 0 || cellIndex >= tab.cells.length) return false;
    const col1Cascade = resolveCol1Cascade(tab, cellIndex, opts);
    const scope = currentPromptScope();
    const isOn = isCellConfirmedInScope(tab.id, cellIndex, scope);
    if (isOn === !!wantOn) {
      // Already desired — still exclusive-clear rival Col1 categories on check.
      if (wantOn && col1Cascade) {
        const offIndices = collectCol1ExclusiveOffIndices(tab, cellIndex);
        if (offIndices.length) {
          const offLinks = linksForCellsInScope(tab.id, offIndices, scope).filter(function (link) {
            return linkNestIndex(link) === null;
          });
          if (offLinks.length) {
            removeLinksFromCombined(offLinks);
            if (!quiet) refreshAfterConfirmedChange();
            else scheduleSave();
            return true;
          }
        }
      }
      return false;
    }
    if (wantOn) {
      const value = cellParentText(tab, cellIndex);
      if (!value) return false;
      let indices = [cellIndex];
      // Auto-batch on Combined checkbox check/uncheck for Col1 — Master insert
      // passes col1Cascade:false to stay single-cell.
      if (col1Cascade) {
        // Exclusive filter: clear other Col1 options (+ their Col B) first so
        // cascade ON for this value cannot leave rival categories checked.
        const offIndices = collectCol1ExclusiveOffIndices(tab, cellIndex);
        if (offIndices.length) {
          const offLinks = linksForCellsInScope(tab.id, offIndices, scope).filter(function (link) {
            return linkNestIndex(link) === null;
          });
          if (offLinks.length) removeLinksFromCombined(offLinks);
        }
        indices = collectCol1CombinedCascadeIndices(tab, cellIndex);
      }
      const pieces = [];
      const colSep = separatorValue(state.separators.column);
      const rowSep = separatorValue(state.separators.row);
      let lastRow = -1;
      let lastHadPiece = false;
      let addedFilter = false;
      for (let i = 0; i < indices.length; i++) {
        const idx = indices[i];
        if (isCellConfirmedInScope(tab.id, idx, scope)) continue;
        const text = cellParentText(tab, idx);
        if (!text) continue;
        // Part-tab Col1: filter checkbox only — never append cell text to Combined.
        if (isPartTabCol1Cell(tab, idx)) {
          if (ensureCol1FilterLink(tab, idx, scope)) addedFilter = true;
          continue;
        }
        const row = Math.floor(idx / (tab.cols || 1));
        if (lastHadPiece) {
          if (row !== lastRow) {
            if (rowSep) pieces.push({ type: 'plain', text: rowSep });
          } else if (colSep) {
            pieces.push({ type: 'plain', text: colSep });
          }
        }
        pieces.push({
          type: 'confirmed',
          text: text,
          tabId: tab.id,
          cellIndex: idx
        });
        lastRow = row;
        lastHadPiece = true;
      }
      if (pieces.length) {
        appendPieces(pieces, quiet ? null : (tab.title + ' ' + cellAddressFromIndex(tab, cellIndex)), { quiet: quiet });
        return true;
      }
      if (!addedFilter) return false;
      if (!quiet) refreshAfterConfirmedChange();
      else scheduleSave();
      return true;
    }
    // Uncheck: reverse Col1 cascade (same category + Col2s) unless opted out.
    let indices = [cellIndex];
    if (col1Cascade) {
      indices = collectCol1CombinedCascadeIndices(tab, cellIndex);
    }
    const links = linksForCellsInScope(tab.id, indices, scope).filter(function (link) {
      return linkNestIndex(link) === null;
    });
    if (!links.length) return false;
    removeLinksFromCombined(links);
    if (!quiet) refreshAfterConfirmedChange();
    return true;
  }

  function markCol1CascadeVisited(cellIndex) {
    if (!checkboxDrag) return;
    const tab = activeTab();
    if (!tab || !shouldCol1CombinedCascade(tab, cellIndex)) return;
    const indices = collectCol1CombinedCascadeIndices(tab, cellIndex);
    for (let i = 0; i < indices.length; i++) {
      checkboxDrag.visited[indices[i]] = true;
    }
    // Exclusive OFF targets: mark visited so drag-paint does not re-check rivals
    // after this Col1 option cleared them.
    if (checkboxDrag.value) {
      const off = collectCol1ExclusiveOffIndices(tab, cellIndex);
      for (let i = 0; i < off.length; i++) {
        checkboxDrag.visited[off[i]] = true;
      }
    }
  }

  function applyCheckboxDragKey(kind, key) {
    if (!checkboxDrag || checkboxDrag.kind !== kind) return;
    if (checkboxDrag.visited[key]) return;
    checkboxDrag.visited[key] = true;
    const want = checkboxDrag.value;
    if (kind === 'combined') {
      ensureCellConfirmedState(key, want, { quiet: true, col1Cascade: true });
      markCol1CascadeVisited(key);
    } else if (kind === 'nest') {
      const parts = String(key).split(':');
      const cellIndex = parseInt(parts[0], 10);
      const nestIndex = parseInt(parts[1], 10);
      if (!Number.isNaN(cellIndex) && !Number.isNaN(nestIndex)) {
        ensureNestConfirmedState(cellIndex, nestIndex, want, { quiet: true });
      }
    }
    // Greens + checkbox marks — do not wait for pointerup refresh.
    applyConfirmedCellHighlights();
  }

  function checkboxDragHit(clientX, clientY) {
    const node = document.elementFromPoint(clientX, clientY);
    if (!node || typeof node.closest !== 'function') return null;
    if (checkboxDrag.kind === 'combined') {
      const btn = node.closest('.cell-append');
      if (!btn || btn.classList.contains('cell-nest-append') || !el.cellGrid.contains(btn)) return null;
      const idx = parseInt(btn.dataset.idx, 10);
      if (Number.isNaN(idx)) return null;
      return { kind: 'combined', key: idx };
    }
    if (checkboxDrag.kind === 'nest') {
      const btn = node.closest('.cell-nest-append');
      if (!btn || !el.cellGrid.contains(btn)) return null;
      const idx = parseInt(btn.dataset.idx, 10);
      const nest = parseInt(btn.dataset.nest, 10);
      if (Number.isNaN(idx) || Number.isNaN(nest)) return null;
      return { kind: 'nest', key: idx + ':' + nest };
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
    if (kind === 'combined' || kind === 'nest') refreshAfterConfirmedChange();
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
    if (kind === 'combined') {
      ensureCellConfirmedState(key, value, { quiet: true, col1Cascade: true });
      markCol1CascadeVisited(key);
    } else if (kind === 'nest') {
      const parts = String(key).split(':');
      const cellIndex = parseInt(parts[0], 10);
      const nestIndex = parseInt(parts[1], 10);
      if (!Number.isNaN(cellIndex) && !Number.isNaN(nestIndex)) {
        ensureNestConfirmedState(cellIndex, nestIndex, value, { quiet: true });
      }
    }
    // Greens + checkbox marks immediately (cascade uncheck must clear siblings now).
    applyConfirmedCellHighlights();
  }

  function applyAppendCheckedState() {
    const tab = activeTab();
    if (!tab) return;
    const scope = currentPromptScope();

    applyRowFilterVisibility();

    const cellButtons = el.cellGrid.querySelectorAll('.cell-append:not(.cell-nest-append)');
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
        ? 'Remove cell ' + cellAddress(row - 1, col - 1) + ' from combined prompt'
        : 'Add cell ' + cellAddress(row - 1, col - 1) + ' to combined prompt';
      btn.setAttribute('aria-label', btn.title);
    }

    const nestButtons = el.cellGrid.querySelectorAll('.cell-nest-append');
    for (let i = 0; i < nestButtons.length; i++) {
      const btn = nestButtons[i];
      const idx = parseInt(btn.dataset.idx, 10);
      const nest = parseInt(btn.dataset.nest, 10);
      if (Number.isNaN(idx) || Number.isNaN(nest)) continue;
      const on = isNestConfirmedInScope(tab.id, idx, nest, scope);
      btn.classList.toggle('is-checked', on);
      btn.setAttribute('aria-pressed', on ? 'true' : 'false');
      const row = Math.floor(idx / tab.cols) + 1;
      const col = (idx % tab.cols) + 1;
      btn.title = on
        ? 'Remove nest ' + (nest + 1) + ' of ' + cellAddress(row - 1, col - 1) + ' from combined prompt'
        : 'Add nest ' + (nest + 1) + ' of ' + cellAddress(row - 1, col - 1) + ' to combined prompt';
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
    const nestNodes = el.cellGrid.querySelectorAll('.cell-nest');
    for (let i = 0; i < nestNodes.length; i++) {
      const nestRow = nestNodes[i];
      const idx = parseInt(nestRow.dataset.idx, 10);
      const nest = parseInt(nestRow.dataset.nest, 10);
      const confirmed = !Number.isNaN(idx) && !Number.isNaN(nest) &&
        isNestConfirmed(tab.id, idx, nest);
      nestRow.classList.toggle('cell-confirmed', confirmed);
      const nestTa = nestRow.querySelector('textarea.cell-nest-input');
      if (nestTa) nestTa.classList.toggle('cell-confirmed', confirmed);
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
      const editingMaster = isMasterCombinedEditSession() && masterSegmentEdit.linkId === span.id;
      const stillGood = link && !seen[span.id] &&
        (editingMaster || linkMatchesSource(link, span.text));
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

  /**
   * Rewrite one confirmed segment in Combined to newText (in place), shifting
   * later links in the same scope. Returns false if the old span cannot be found.
   */
  function rewriteConfirmedLinkSegment(link, newText) {
    if (!link || typeof newText !== 'string') return false;
    const scope = link.scope;
    let text = getPromptText(scope);
    if (!repairConfirmedLinkOffset(link, text)) return false;
    text = getPromptText(scope);
    const oldStart = link.start;
    const oldEnd = link.end;
    const delta = newText.length - (oldEnd - oldStart);
    if (delta === 0 && text.slice(oldStart, oldEnd) === newText) {
      link.text = newText;
      return true;
    }
    text = text.slice(0, oldStart) + newText + text.slice(oldEnd);
    link.text = newText;
    link.start = oldStart;
    link.end = oldStart + newText.length;
    state.confirmedLinks.forEach(function (other) {
      if (other.scope !== scope || other.id === link.id) return;
      if (other.start >= oldEnd) {
        other.start += delta;
        other.end += delta;
      }
    });
    setPromptText(scope, text);
    return true;
  }

  /**
   * Ash: editing a green/confirmed cell or nest in the table also edits the
   * matching Combined segment live (keeps the link). Empty source removes it.
   * nestIndexFilter: undefined = all links for the cell; null = parent only;
   * integer = that nest only.
   */
  function liveSyncConfirmedLinksForCell(tabId, cellIndex, nestIndexFilter, opts) {
    const silent = !!(opts && opts.silent);
    const candidates = state.confirmedLinks.filter(function (link) {
      if (link.tabId !== tabId || link.cellIndex !== cellIndex) return false;
      if (nestIndexFilter === undefined) return true;
      return linkNestIndex(link) === nestIndexFilter;
    });
    if (!candidates.length) return false;

    const toRemove = [];
    const toRewrite = [];
    let filterMetaChanged = false;
    candidates.forEach(function (link) {
      const expected = sourceLinkText(link);
      // Col1 filter-only: sync metadata / drop when empty — never rewrite Combined.
      if (isCol1FilterOnlyLink(link)) {
        if (expected === null || expected === '') {
          toRemove.push(link);
          return;
        }
        if (link.text !== expected) {
          link.text = expected;
          filterMetaChanged = true;
        }
        link.start = 0;
        link.end = 0;
        return;
      }
      if (expected === null || expected === '') {
        toRemove.push(link);
        return;
      }
      if (link.text === expected) {
        const text = getPromptText(link.scope);
        if (repairConfirmedLinkOffset(link, text) &&
            text.slice(link.start, link.end) === expected) {
          return;
        }
        // Offset lost but source still has text — drop rather than guess.
        toRemove.push(link);
        return;
      }
      toRewrite.push(link);
    });

    let changed = false;

    // Rewrite before remove so removeLinksFromCombined sees up-to-date offsets.
    if (toRewrite.length) {
      const byScope = {};
      toRewrite.forEach(function (link) {
        if (!byScope[link.scope]) byScope[link.scope] = [];
        byScope[link.scope].push(link);
      });
      Object.keys(byScope).forEach(function (scope) {
        const links = byScope[scope].slice().sort(function (a, b) {
          return b.start - a.start;
        });
        for (let i = 0; i < links.length; i++) {
          const link = links[i];
          if (!state.confirmedLinks.some(function (l) { return l.id === link.id; })) continue;
          const expected = sourceLinkText(link);
          if (expected === null || expected === '') {
            toRemove.push(link);
            continue;
          }
          if (rewriteConfirmedLinkSegment(link, expected)) changed = true;
          else {
            // Cannot locate old segment — drop the link record (Combined text kept).
            state.confirmedLinks = state.confirmedLinks.filter(function (l) {
              return l.id !== link.id;
            });
            changed = true;
          }
        }
      });
    }

    if (toRemove.length) {
      const byScope = {};
      toRemove.forEach(function (link) {
        if (!byScope[link.scope]) byScope[link.scope] = [];
        byScope[link.scope].push(link);
      });
      Object.keys(byScope).forEach(function (scope) {
        // Deduplicate ids in case rewrite path also queued a remove.
        const seen = {};
        const list = byScope[scope].filter(function (link) {
          if (seen[link.id]) return false;
          seen[link.id] = true;
          return state.confirmedLinks.some(function (l) { return l.id === link.id; });
        });
        if (list.length && removeLinksFromCombined(list)) changed = true;
      });
    }

    if ((changed || filterMetaChanged) && !silent) {
      renderCombinedPrompt();
      applyConfirmedCellHighlights();
    }
    return changed || filterMetaChanged;
  }

  function revalidateLinksForCell(tabId, cellIndex, opts) {
    let removed = false;
    state.confirmedLinks = state.confirmedLinks.filter(function (link) {
      if (link.tabId !== tabId || link.cellIndex !== cellIndex) return true;
      if (isCol1FilterOnlyLink(link)) {
        const expected = sourceLinkText(link);
        if (expected === null || expected === '') {
          removed = true;
          return false;
        }
        link.text = expected;
        link.start = 0;
        link.end = 0;
        return true;
      }
      if (linkMatchesSource(link, link.text)) return true;
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
      if (isCol1FilterOnlyLink(link)) {
        const expected = sourceLinkText(link);
        if (expected === null || expected === '') {
          removed = true;
          return false;
        }
        link.text = expected;
        link.start = 0;
        link.end = 0;
        return true;
      }
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

  function tabSourceOrderIndex(tabId) {
    const idx = state.tabs.findIndex(function (tab) { return tab.id === tabId; });
    return idx < 0 ? 9999 : idx;
  }

  /** Sort key: tab bar order → cellIndex (row-major) → parent before nests → nestIndex. */
  function compareLinksBySourceOrder(a, b) {
    const ta = tabSourceOrderIndex(a.tabId);
    const tb = tabSourceOrderIndex(b.tabId);
    if (ta !== tb) return ta - tb;
    if (a.cellIndex !== b.cellIndex) return a.cellIndex - b.cellIndex;
    const na = linkNestIndex(a);
    const nb = linkNestIndex(b);
    if (na === null && nb === null) return 0;
    if (na === null) return -1;
    if (nb === null) return 1;
    return na - nb;
  }

  function separatorBetweenSourceLinks(prev, next) {
    const partSep = separatorValue(state.separators.part);
    const colSep = separatorValue(state.separators.column);
    const rowSep = separatorValue(state.separators.row);
    if (prev.tabId !== next.tabId) return partSep;
    const tab = state.tabs.find(function (t) { return t.id === next.tabId; });
    const cols = tab && tab.cols > 0 ? tab.cols : 1;
    if (prev.cellIndex === next.cellIndex) return rowSep;
    const prevRow = Math.floor(prev.cellIndex / cols);
    const nextRow = Math.floor(next.cellIndex / cols);
    if (prevRow === nextRow) return colSep;
    return rowSep;
  }

  /**
   * Rebuild one scope's confirmed span by rejoining `ordered` links with the
   * current part/column/row separators. Same plain-text rule as Match source
   * order: keep text before the first and after the last confirmed segment;
   * replace the middle. Link ids, text, and Master lock fields stay; ranges move.
   */
  function rejoinConfirmedLinks(scope, ordered) {
    if (!ordered || !ordered.length) return false;
    const target = scope || currentPromptScope();
    // Col1 filter-only links stay confirmed for checkbox UX but never enter Combined text.
    const contentOrdered = [];
    const filterLinks = [];
    for (let i = 0; i < ordered.length; i++) {
      if (isCol1FilterOnlyLink(ordered[i])) filterLinks.push(ordered[i]);
      else contentOrdered.push(ordered[i]);
    }
    let middle = '';
    const updated = [];
    for (let i = 0; i < contentOrdered.length; i++) {
      const link = contentOrdered[i];
      if (i > 0) middle += separatorBetweenSourceLinks(contentOrdered[i - 1], link);
      const start = middle.length;
      middle += link.text || '';
      const end = middle.length;
      const next = {
        id: link.id,
        tabId: link.tabId,
        cellIndex: link.cellIndex,
        text: link.text,
        start: start,
        end: end,
        scope: target
      };
      if (Number.isInteger(link.nestIndex) && link.nestIndex >= 0) {
        next.nestIndex = link.nestIndex;
      }
      copyLinkMasterFields(link, next);
      updated.push(next);
    }

    // Prefix/suffix from the full ordered set so legacy Col1 spans are sliced out
    // of the replaced middle when content is rebuilt without them.
    const byStart = ordered.slice().sort(function (a, b) { return a.start - b.start; });
    const current = getPromptText(target);
    const prefix = byStart.length ? current.slice(0, byStart[0].start) : '';
    const suffix = byStart.length ? current.slice(byStart[byStart.length - 1].end) : '';
    const shift = prefix.length;
    for (let i = 0; i < updated.length; i++) {
      updated[i].start += shift;
      updated[i].end += shift;
    }
    // Pin filter-only Col1 links at zero-length (not painted in Combined).
    const filterAnchor = shift + middle.length;
    for (let i = 0; i < filterLinks.length; i++) {
      const link = filterLinks[i];
      const expected = sourceLinkText(link);
      if (expected === null || expected === '') continue;
      const next = {
        id: link.id,
        tabId: link.tabId,
        cellIndex: link.cellIndex,
        text: expected,
        start: filterAnchor,
        end: filterAnchor,
        scope: target
      };
      copyLinkMasterFields(link, next);
      updated.push(next);
    }
    const nextText = prefix + middle + suffix;

    let unchanged = nextText === current && updated.length === ordered.length;
    if (unchanged) {
      const prevById = Object.create(null);
      for (let i = 0; i < ordered.length; i++) prevById[ordered[i].id] = ordered[i];
      for (let i = 0; i < updated.length; i++) {
        const prev = prevById[updated[i].id];
        if (!prev ||
            updated[i].start !== prev.start ||
            updated[i].end !== prev.end ||
            updated[i].text !== prev.text) {
          unchanged = false;
          break;
        }
      }
    }
    if (unchanged) return false;

    state.confirmedLinks = state.confirmedLinks.filter(function (link) {
      return link.scope !== target;
    }).concat(updated);
    setPromptText(target, nextText);
    if (target === currentPromptScope()) {
      const caret = shift + middle.length;
      lastCombinedCaret = { start: caret, end: caret, scope: target };
    }
    return true;
  }

  /**
   * When Match source order is ON, rebuild Combined confirmed segments in
   * grid/tab source order (row-major). Preserves plain text before the first
   * and after the last confirmed segment; replaces the middle.
   */
  function reorderCombinedToSourceOrder(scope) {
    if (!state.matchSourceOrder) return false;
    const target = scope || currentPromptScope();
    const links = state.confirmedLinks.filter(function (link) {
      return link.scope === target;
    });
    if (links.length === 0) return false;
    return rejoinConfirmedLinks(target, links.slice().sort(compareLinksBySourceOrder));
  }

  /**
   * After a Tools separator is committed, rejoin every scope that has confirmed
   * links using the new part/column/row separators. Match source order also
   * re-sorts; otherwise document order is kept so caret/append order is not wiped.
   */
  function rewriteCombinedForSeparators() {
    const seen = Object.create(null);
    const scopes = [];
    for (let i = 0; i < state.confirmedLinks.length; i++) {
      const scope = state.confirmedLinks[i].scope;
      if (!scope || seen[scope]) continue;
      seen[scope] = true;
      scopes.push(scope);
    }
    let changed = false;
    for (let i = 0; i < scopes.length; i++) {
      const target = scopes[i];
      if (state.matchSourceOrder) {
        if (reorderCombinedToSourceOrder(target)) changed = true;
        continue;
      }
      const links = state.confirmedLinks.filter(function (link) {
        return link.scope === target;
      });
      if (!links.length) continue;
      const ordered = links.slice().sort(function (a, b) { return a.start - b.start; });
      if (rejoinConfirmedLinks(target, ordered)) changed = true;
    }
    if (!changed) return false;
    renderCombinedPrompt();
    scheduleSave();
    return true;
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
    reorderCombinedToSourceOrder(currentPromptScope());
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
    const wantOn = !isCellConfirmedInScope(tab.id, cellIndex, scope);
    // Route through ensureCellConfirmedState so Col1 Combined cascades on both
    // check and uncheck (nests stay independent inside that helper).
    if (!wantOn) {
      pushHistory();
      const changed = ensureCellConfirmedState(cellIndex, false, { col1Cascade: true });
      if (!changed) {
        applyConfirmedCellHighlights();
        return;
      }
      const row = Math.floor(cellIndex / tab.cols) + 1;
      const col = (cellIndex % tab.cols) + 1;
      setStatus('Removed ' + cellAddress(row - 1, col - 1) + ' from combined', 'ok');
      return;
    }
    const value = cellParentText(tab, cellIndex);
    if (!value) {
      setStatus('Nothing to append', 'err');
      return;
    }
    ensureCellConfirmedState(cellIndex, true, { col1Cascade: true });
  }

  function toggleNestConfirmed(cellIndex, nestIndex) {
    const tab = activeTab();
    if (!tab || cellIndex < 0 || cellIndex >= tab.cells.length) return;
    if (!Number.isInteger(nestIndex) || nestIndex < 0) return;
    const scope = currentPromptScope();
    const nests = getCellNests(tab, cellIndex);
    if (nestIndex >= nests.length) return;

    if (isNestConfirmedInScope(tab.id, cellIndex, nestIndex, scope)) {
      const links = linksForNestInScope(tab.id, cellIndex, nestIndex, scope);
      if (!links.length) {
        applyAppendCheckedState();
        return;
      }
      pushHistory();
      removeLinksFromCombined(links);
      refreshAfterConfirmedChange();
      const row = Math.floor(cellIndex / tab.cols) + 1;
      const col = (cellIndex % tab.cols) + 1;
      setStatus('Removed nest ' + (nestIndex + 1) + ' of ' + cellAddress(row - 1, col - 1) + ' from combined', 'ok');
      return;
    }

    const value = nestExportText(nests[nestIndex]);
    if (!value) {
      setStatus('Nothing to append', 'err');
      return;
    }
    appendPieces([{
      type: 'confirmed',
      text: value,
      tabId: tab.id,
      cellIndex: cellIndex,
      nestIndex: nestIndex
    }], tab.title + ' ' + cellAddressFromIndex(tab, cellIndex) + ' nest ' + (nestIndex + 1));
  }

  function ensureNestConfirmedState(cellIndex, nestIndex, wantOn, opts) {
    const quiet = !!(opts && opts.quiet);
    const tab = activeTab();
    if (!tab || cellIndex < 0 || cellIndex >= tab.cells.length) return false;
    if (!Number.isInteger(nestIndex) || nestIndex < 0) return false;
    const scope = currentPromptScope();
    const isOn = isNestConfirmedInScope(tab.id, cellIndex, nestIndex, scope);
    if (isOn === !!wantOn) return false;
    if (wantOn) {
      const nests = getCellNests(tab, cellIndex);
      if (nestIndex >= nests.length) return false;
      const value = nestExportText(nests[nestIndex]);
      if (!value) return false;
      appendPieces([{
        type: 'confirmed',
        text: value,
        tabId: tab.id,
        cellIndex: cellIndex,
        nestIndex: nestIndex
      }], quiet ? null : (tab.title + ' ' + cellAddressFromIndex(tab, cellIndex) + ' nest ' + (nestIndex + 1)), { quiet: quiet });
      return true;
    }
    const links = linksForNestInScope(tab.id, cellIndex, nestIndex, scope);
    if (!links.length) return false;
    removeLinksFromCombined(links);
    if (!quiet) refreshAfterConfirmedChange();
    return true;
  }

  function appendWithPartSeparator(current, text) {
    const partSeparator = separatorValue(state.separators.part);
    if (current && partSeparator && !current.endsWith(partSeparator)) {
      return current + partSeparator + text;
    }
    return current + text;
  }

  /**
   * Insert plain and confirmed pieces into the active combined prompt at the
   * Combined caret/selection (or last known caret). Falls back to end append.
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
    const insertRange = resolveCombinedInsertRange(scope, current.length);
    const insertAt = insertRange.start;
    const insertEnd = insertRange.end;
    const before = current.slice(0, insertAt);
    const after = current.slice(insertEnd);
    const removedLen = insertEnd - insertAt;

    const partSeparator = separatorValue(state.separators.part);
    const toAdd = [];
    if (before && partSeparator && !before.endsWith(partSeparator)) {
      toAdd.push({ type: 'plain', text: partSeparator });
    }
    for (let i = 0; i < pieces.length; i++) toAdd.push(pieces[i]);

    // Preview that we will actually add something before snapshotting undo.
    // Col1 filter-only parents contribute no Combined text (handled below).
    let willAdd = '';
    let willAddFilterOnly = false;
    for (let i = 0; i < toAdd.length; i++) {
      const piece = toAdd[i];
      const text = piece.text || '';
      if (!text && piece.type !== 'plain') continue;
      if (piece.type === 'confirmed') {
        const pieceTab = state.tabs.find(function (t) { return t.id === piece.tabId; });
        const nestIdx = Number.isInteger(piece.nestIndex) && piece.nestIndex >= 0
          ? piece.nestIndex
          : null;
        if (nestIdx === null && isPartTabCol1Cell(pieceTab, piece.cellIndex)) {
          willAddFilterOnly = true;
          continue;
        }
      }
      willAdd += text;
    }
    if (!willAdd) {
      if (willAddFilterOnly) {
        for (let i = 0; i < toAdd.length; i++) {
          const piece = toAdd[i];
          if (piece.type !== 'confirmed') continue;
          const pieceTab = state.tabs.find(function (t) { return t.id === piece.tabId; });
          const nestIdx = Number.isInteger(piece.nestIndex) && piece.nestIndex >= 0
            ? piece.nestIndex
            : null;
          if (nestIdx === null && isPartTabCol1Cell(pieceTab, piece.cellIndex)) {
            ensureCol1FilterLink(pieceTab, piece.cellIndex, scope);
          }
        }
        if (!quiet) refreshAfterConfirmedChange();
        else scheduleSave();
        return;
      }
      if (!quiet) setStatus('Nothing to append', 'err');
      return;
    }

    pushHistory();

    // Drop links overlapping a replaced selection; shift survivors after the hole
    // into post-deletion coordinates (still relative to `before + after`).
    if (removedLen > 0) {
      state.confirmedLinks = state.confirmedLinks.filter(function (link) {
        if (link.scope !== scope) return true;
        if (link.end <= insertAt || link.start >= insertEnd) return true;
        return false;
      });
      state.confirmedLinks.forEach(function (link) {
        if (link.scope !== scope) return;
        if (link.start >= insertEnd) {
          link.start -= removedLen;
          link.end -= removedLen;
        }
      });
    }

    // Shift surviving links that begin at/after the insert point to make room.
    state.confirmedLinks.forEach(function (link) {
      if (link.scope !== scope) return;
      if (link.start >= insertAt) {
        link.start += willAdd.length;
        link.end += willAdd.length;
      }
    });

    let offset = insertAt;
    let added = '';
    for (let i = 0; i < toAdd.length; i++) {
      const piece = toAdd[i];
      const text = piece.text || '';
      if (!text && piece.type !== 'plain') continue;
      if (piece.type === 'confirmed') {
        const pieceTab = state.tabs.find(function (t) { return t.id === piece.tabId; });
        const nestIdx = Number.isInteger(piece.nestIndex) && piece.nestIndex >= 0
          ? piece.nestIndex
          : null;
        // Part-tab Col1 parent: filter-only link, no Combined text.
        if (nestIdx === null && isPartTabCol1Cell(pieceTab, piece.cellIndex)) {
          if (!isCellConfirmedInScope(piece.tabId, piece.cellIndex, scope)) {
            ensureCol1FilterLink(pieceTab, piece.cellIndex, scope);
          }
          continue;
        }
        const id = linkUid();
        const link = {
          id: id,
          tabId: piece.tabId,
          cellIndex: piece.cellIndex,
          text: text,
          start: offset,
          end: offset + text.length,
          scope: scope
        };
        if (nestIdx !== null) link.nestIndex = nestIdx;
        applyMasterOriginToNewLink(link, piece);
        state.confirmedLinks.push(link);
      }
      added += text;
      offset += text.length;
    }

    if (!added) {
      setStatus('Nothing to append', 'err');
      return;
    }

    setPromptText(scope, before + added + after);
    lastCombinedCaret = {
      start: insertAt + added.length,
      end: insertAt + added.length,
      scope: scope
    };

    let addedConfirmed = false;
    for (let i = 0; i < toAdd.length; i++) {
      if (toAdd[i].type === 'confirmed') { addedConfirmed = true; break; }
    }
    if (addedConfirmed) reorderCombinedToSourceOrder(scope);

    if (quiet) {
      scheduleSave();
      return;
    }
    renderCombinedPrompt();
    applyConfirmedCellHighlights();
    renderMasterLibrary();
    if (document.activeElement === el.combined) {
      setCombinedCaretOffset(lastCombinedCaret.start);
    }
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
      const cellPieces = cellConfirmedPieces(tab, idx);
      if (!cellPieces.length) continue;
      if (pieces.length) pieces.push({ type: 'plain', text: colSep });
      for (let i = 0; i < cellPieces.length; i++) pieces.push(cellPieces[i]);
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

  function refreshAfterMasterSegmentChange() {
    renderCombinedPrompt();
    applyConfirmedCellHighlights();
    renderMasterLibrary();
    renderGrid();
    scheduleSave();
  }

  function escapeCssIdent(value) {
    if (window.CSS && typeof CSS.escape === 'function') return CSS.escape(value);
    return String(value).replace(/[^a-zA-Z0-9_-]/g, '\\$&');
  }

  function findCombinedSegmentSpan(linkId) {
    if (!el.combined || !linkId) return null;
    return el.combined.querySelector('.confirmed-segment[data-link-id="' + escapeCssIdent(linkId) + '"]');
  }

  function beginMasterSegmentEdit(span) {
    if (!span || masterSegmentDialogOpen) return;
    const linkId = span.dataset.linkId || '';
    const link = linkById(linkId);
    if (!link || !isMasterLockedLink(link)) return;
    const originalText = link.text;
    masterSegmentEdit = {
      kind: 'combined',
      linkId: linkId,
      originalText: originalText,
      scope: link.scope
    };
    // Keep link.locked true in data until Overwrite/Keep/Cancel resolves so
    // autosave mid-edit does not persist an unlocked Master segment.
    span.classList.remove('master-segment-locked');
    span.classList.add('master-segment-editing');
    span.contentEditable = 'true';
    span.title = 'Editing Master segment — blur or Enter to confirm; Esc to cancel';
    // Focus inside the segment for editing.
    const selection = window.getSelection();
    const range = document.createRange();
    range.selectNodeContents(span);
    range.collapse(false);
    if (selection) {
      selection.removeAllRanges();
      selection.addRange(range);
    }
    span.focus();
    setStatus('Master segment unlocked — edit, then choose Overwrite / Keep local / Cancel', 'ok');
  }

  function beginMasterCellEdit(ta, tab, cellIndex, focusEl) {
    if (!tab || masterSegmentDialogOpen) return;
    if (isMasterTab(tab)) return;
    if (!isCellMasterLocked(tab, cellIndex)) return;
    const originalText = tab.cells[cellIndex] == null ? '' : String(tab.cells[cellIndex]);
    ensureNestedCells(tab);
    ensureCellPages(tab);
    masterSegmentEdit = {
      kind: 'cell',
      tabId: tab.id,
      cellIndex: cellIndex,
      originalText: originalText,
      originalNests: cloneNestList(getCellNests(tab, cellIndex)),
      originalPages: cloneCellPagesEntry(getCellPages(tab, cellIndex))
    };
    // Keep cellLocks.locked true in data until dialog resolves.
    const parentTa = el.cellGrid
      ? el.cellGrid.querySelector('textarea.cell[data-idx="' + cellIndex + '"]')
      : null;
    const unlockTa = parentTa || ta;
    const wrapNode = unlockTa && unlockTa.closest
      ? unlockTa.closest('.cell-wrap')
      : (ta && ta.closest ? ta.closest('.cell-wrap') : null);
    if (unlockTa) {
      unlockTa.readOnly = false;
      unlockTa.classList.remove('master-cell-locked');
      unlockTa.classList.add('master-cell-editing');
      unlockTa.title = 'Editing Master cell — blur or Enter to confirm; Esc to cancel';
    }
    if (wrapNode) {
      wrapNode.classList.remove('master-cell-locked');
      wrapNode.classList.add('master-cell-editing');
      const nestTas = wrapNode.querySelectorAll('textarea.cell-nest-input');
      for (let i = 0; i < nestTas.length; i++) {
        nestTas[i].readOnly = false;
        nestTas[i].classList.remove('master-cell-locked');
        nestTas[i].classList.add('master-cell-editing');
        nestTas[i].title = 'Editing Master nest — blur or Enter to confirm; Esc to cancel';
      }
    }
    const focusTarget = focusEl || unlockTa;
    if (focusTarget && typeof focusTarget.focus === 'function') {
      focusTarget.focus();
      try {
        if (typeof focusTarget.setSelectionRange === 'function') {
          const len = (focusTarget.value || '').length;
          focusTarget.setSelectionRange(len, len);
        }
      } catch (err) { /* no-op */ }
    }
    setStatus('Master cell unlocked — edit, then choose Overwrite / Keep local / Cancel', 'ok');
  }

  function showMasterSegmentDialog() {
    if (!el.masterSegmentDialog) return;
    updateMasterEditDialogCopy();
    masterSegmentDialogOpen = true;
    el.masterSegmentDialog.hidden = false;
    if (el.btnMasterOverwrite) el.btnMasterOverwrite.focus();
  }

  function hideMasterSegmentDialog() {
    masterSegmentDialogOpen = false;
    if (el.masterSegmentDialog) el.masterSegmentDialog.hidden = true;
  }

  function requestMasterSegmentFinish(opts) {
    opts = opts || {};
    if (!masterSegmentEdit || masterSegmentDialogOpen) return;

    if (masterSegmentEdit.kind === 'cell') {
      const tab = state.tabs.find(function (t) { return t.id === masterSegmentEdit.tabId; });
      if (!tab) {
        masterSegmentEdit = null;
        return;
      }
      const idx = masterSegmentEdit.cellIndex;
      const liveText = idx >= 0 && idx < tab.cells.length
        ? (tab.cells[idx] == null ? '' : String(tab.cells[idx]))
        : masterSegmentEdit.originalText;
      const liveNests = idx >= 0 ? getCellNests(tab, idx) : [];
      const nestsUnchanged = nestsContentEqual(
        liveNests,
        masterSegmentEdit.originalNests || []
      );
      const livePages = idx >= 0 ? getCellPages(tab, idx) : makeEmptyCellPages('');
      const pagesUnchanged = cellPagesContentEqual(
        livePages,
        masterSegmentEdit.originalPages || makeEmptyCellPages('')
      );
      if (!opts.forceDialog && liveText === masterSegmentEdit.originalText &&
          nestsUnchanged && pagesUnchanged) {
        ensureCellLocks(tab);
        if (tab.cellLocks[idx]) tab.cellLocks[idx].locked = true;
        masterSegmentEdit = null;
        refreshAfterMasterSegmentChange();
        return;
      }
      showMasterSegmentDialog();
      return;
    }

    const link = linkById(masterSegmentEdit.linkId, masterSegmentEdit.scope);
    if (!link) {
      masterSegmentEdit = null;
      return;
    }
    // Read live DOM text for the segment if present.
    const span = findCombinedSegmentSpan(masterSegmentEdit.linkId);
    const liveText = span ? (span.textContent || '') : link.text;
    if (!opts.forceDialog && liveText === masterSegmentEdit.originalText) {
      // No change — silently re-lock (data flag stayed locked).
      link.text = masterSegmentEdit.originalText;
      link.locked = true;
      masterSegmentEdit = null;
      refreshAfterMasterSegmentChange();
      return;
    }
    // Persist current edit into link/prompt before dialog choice.
    if (span) {
      syncConfirmedFromCombinedDom();
    }
    showMasterSegmentDialog();
  }

  function cancelMasterSegmentEdit() {
    if (!masterSegmentEdit) {
      hideMasterSegmentDialog();
      return;
    }
    const edit = masterSegmentEdit;
    hideMasterSegmentDialog();

    if (edit.kind === 'cell') {
      const tab = state.tabs.find(function (t) { return t.id === edit.tabId; });
      if (tab && edit.cellIndex >= 0 && edit.cellIndex < tab.cells.length) {
        pushHistory();
        tab.cells[edit.cellIndex] = edit.originalText;
        ensureNestedCells(tab);
        tab.nestedCells[edit.cellIndex] = cloneNestList(edit.originalNests || []);
        ensureCellPages(tab);
        tab.cellPages[edit.cellIndex] = cloneCellPagesEntry(
          edit.originalPages || makeEmptyCellPages(edit.originalText)
        );
        syncCellTextFromPages(tab, edit.cellIndex);
        ensureCellLocks(tab);
        if (tab.cellLocks[edit.cellIndex]) {
          tab.cellLocks[edit.cellIndex].locked = true;
          tab.cellLocks[edit.cellIndex].masterOrigin = true;
        } else {
          setCellMasterLock(tab, edit.cellIndex, null);
        }
        syncNestLinksAfterNestReplace(tab.id, edit.cellIndex);
        liveSyncConfirmedLinksForCell(tab.id, edit.cellIndex, null, { silent: true });
      }
      masterSegmentEdit = null;
      refreshAfterMasterSegmentChange();
      setStatus('Master cell edit cancelled — locked again', 'ok');
      return;
    }

    const link = linkById(edit.linkId, edit.scope);
    if (link) {
      pushHistory();
      rewriteConfirmedLinkSegment(link, edit.originalText);
      link.locked = true;
      link.masterOrigin = true;
    }
    masterSegmentEdit = null;
    refreshAfterMasterSegmentChange();
    setStatus('Master segment edit cancelled — locked again', 'ok');
  }

  function applyMasterOverwriteText(oldText, newText, preferredMasterIdx, nestsFromEdit, pagesFromEdit) {
    const master = state.tabs.find(isMasterTab);
    let masterIdx = Number.isInteger(preferredMasterIdx) ? preferredMasterIdx : -1;
    if (master) {
      if (masterIdx < 0 || masterIdx >= master.cells.length ||
          cellParentText(master, masterIdx) !== oldText) {
        const found = findMasterCellIndexByText(oldText);
        masterIdx = found === null ? -1 : found;
      }
      if (masterIdx >= 0) {
        // Optional nest/pages payload from a part-cell Overwrite (parent + nests + pages).
        if (pagesFromEdit !== undefined) {
          ensureCellPages(master);
          master.cellPages[masterIdx] = cloneCellPagesEntry(pagesFromEdit);
          // Prefer edited current-page text as the visible page.
          writeCellCurrentPage(master, masterIdx, newText);
        } else {
          writeCellCurrentPage(master, masterIdx, newText);
        }
        if (nestsFromEdit !== undefined) {
          ensureNestedCells(master);
          master.nestedCells[masterIdx] = cloneNestList(nestsFromEdit);
        }
      }
    }

    // Sync every tab cell that still holds the old Master text (library + parts).
    state.tabs.forEach(function (tab) {
      if (!tab || !Array.isArray(tab.cells)) return;
      for (let i = 0; i < tab.cells.length; i++) {
        if (master && tab.id === master.id && i === masterIdx) continue;
        if (cellParentText(tab, i) === oldText) {
          writeCellCurrentPage(tab, i, newText);
          if (!isMasterTab(tab)) {
            setCellMasterLock(tab, i, masterIdx >= 0 ? masterIdx : null);
          }
        }
      }
    });

    // Rewrite all Master-origin Combined segments that tracked this Master text.
    state.confirmedLinks.forEach(function (other) {
      if (!isMasterOriginLink(other)) return;
      const sameMaster = Number.isInteger(masterIdx) && masterIdx >= 0 &&
        other.masterCellIndex === masterIdx;
      const sameText = other.text === oldText;
      if (!sameMaster && !sameText) return;
      if (other.text !== newText) rewriteConfirmedLinkSegment(other, newText);
      other.masterOrigin = true;
      other.locked = true;
      if (masterIdx >= 0) other.masterCellIndex = masterIdx;
    });

    // Locked inserted cells: sync nests + pages from Master (same index).
    if (masterIdx >= 0) {
      syncLockedPartNestsFromMaster(masterIdx);
      syncLockedPartPagesFromMaster(masterIdx);
    }

    return masterIdx;
  }

  function overwriteMasterFromSegmentEdit() {
    if (!masterSegmentEdit) {
      hideMasterSegmentDialog();
      return;
    }
    const edit = masterSegmentEdit;
    hideMasterSegmentDialog();

    if (edit.kind === 'cell') {
      const tab = state.tabs.find(function (t) { return t.id === edit.tabId; });
      if (!tab || edit.cellIndex < 0 || edit.cellIndex >= tab.cells.length) {
        masterSegmentEdit = null;
        refreshAfterMasterSegmentChange();
        return;
      }
      const newText = tab.cells[edit.cellIndex] == null ? '' : String(tab.cells[edit.cellIndex]);
      const oldText = edit.originalText;
      const lock = getCellLock(tab, edit.cellIndex);
      const preferred = lock && Number.isInteger(lock.masterCellIndex) ? lock.masterCellIndex : -1;
      const nestsFromEdit = cloneNestList(getCellNests(tab, edit.cellIndex));
      const pagesFromEdit = cloneCellPagesEntry(getCellPages(tab, edit.cellIndex));
      pushHistory();
      const masterIdx = applyMasterOverwriteText(oldText, newText, preferred, nestsFromEdit, pagesFromEdit);
      setCellMasterLock(tab, edit.cellIndex, masterIdx >= 0 ? masterIdx : null);
      masterSegmentEdit = null;
      refreshAfterMasterSegmentChange();
      setStatus('Overwrote Master (text + nests + pages) and synced matching cells/segments', 'ok');
      return;
    }

    const link = linkById(edit.linkId, edit.scope);
    if (!link) {
      masterSegmentEdit = null;
      refreshAfterMasterSegmentChange();
      return;
    }
    const span = findCombinedSegmentSpan(edit.linkId);
    if (span) syncConfirmedFromCombinedDom();
    const newText = link.text;
    const oldText = edit.originalText;
    pushHistory();
    const preferred = Number.isInteger(link.masterCellIndex) ? link.masterCellIndex : -1;
    const masterIdx = applyMasterOverwriteText(oldText, newText, preferred);
    link.masterOrigin = true;
    link.locked = true;
    if (masterIdx >= 0) link.masterCellIndex = masterIdx;
    // Also lock the linked part cell when it holds this Master text.
    const srcTab = state.tabs.find(function (t) { return t.id === link.tabId; });
    if (srcTab && !isMasterTab(srcTab) && link.cellIndex >= 0 &&
        link.cellIndex < srcTab.cells.length &&
        cellParentText(srcTab, link.cellIndex) === newText) {
      setCellMasterLock(srcTab, link.cellIndex, masterIdx >= 0 ? masterIdx : null);
    }
    masterSegmentEdit = null;
    refreshAfterMasterSegmentChange();
    setStatus('Overwrote Master (text + nests) and synced matching cells/segments', 'ok');
  }

  function keepMasterSegmentLocalOnly() {
    if (!masterSegmentEdit) {
      hideMasterSegmentDialog();
      return;
    }
    const edit = masterSegmentEdit;
    hideMasterSegmentDialog();

    if (edit.kind === 'cell') {
      const tab = state.tabs.find(function (t) { return t.id === edit.tabId; });
      if (!tab || edit.cellIndex < 0 || edit.cellIndex >= tab.cells.length) {
        masterSegmentEdit = null;
        refreshAfterMasterSegmentChange();
        return;
      }
      const newText = tab.cells[edit.cellIndex] == null ? '' : String(tab.cells[edit.cellIndex]);
      pushHistory();
      clearCellMasterLock(tab, edit.cellIndex);
      // Drop Master lock/origin on Combined links for this cell (parent + nests).
      state.confirmedLinks.forEach(function (other) {
        if (other.tabId !== tab.id || other.cellIndex !== edit.cellIndex) return;
        if (!isMasterOriginLink(other)) return;
        other.masterOrigin = false;
        other.locked = false;
        if ('masterCellIndex' in other) delete other.masterCellIndex;
      });
      liveSyncConfirmedLinksForCell(tab.id, edit.cellIndex, undefined, { silent: true });
      masterSegmentEdit = null;
      refreshAfterMasterSegmentChange();
      setStatus('Kept as local cell text — unlinked from Master lock', 'ok');
      return;
    }

    const link = linkById(edit.linkId, edit.scope);
    if (!link) {
      masterSegmentEdit = null;
      refreshAfterMasterSegmentChange();
      return;
    }
    const span = findCombinedSegmentSpan(edit.linkId);
    if (span) syncConfirmedFromCombinedDom();
    const newText = link.text;
    pushHistory();

    // Update the linked part/Master cell (or nest) so the confirmed link stays valid.
    const nestIdx = linkNestIndex(link);
    const tab = state.tabs.find(function (t) { return t.id === link.tabId; });
    if (tab) {
      if (nestIdx === null) {
        if (link.cellIndex >= 0 && link.cellIndex < tab.cells.length) {
          writeCellCurrentPage(tab, link.cellIndex, newText);
          if (!isMasterTab(tab)) clearCellMasterLock(tab, link.cellIndex);
        }
      } else {
        ensureNestedCells(tab);
        const nests = getCellNests(tab, link.cellIndex);
        if (nestIdx < nests.length) {
          const nest = nests[nestIdx];
          // Replace nest export with a single page of the new text.
          nest.pages = [newText];
          nest.page = 0;
        }
      }
    }

    link.masterOrigin = false;
    link.locked = false;
    if ('masterCellIndex' in link) delete link.masterCellIndex;
    masterSegmentEdit = null;
    refreshAfterMasterSegmentChange();
    setStatus('Kept as local Combined text — unlinked from Master lock', 'ok');
  }

  function renderCombinedPrompt() {
    const tab = activeTab();
    const scope = currentPromptScope();
    // Repair offsets before painting so greens survive restart / load drift.
    repairConfirmedLinksForScope(scope);
    const text = getPromptText(scope);
    const links = linksForScope(scope).filter(function (link) {
      // Col1 filter-only: keep for checkbox state; never paint into Combined.
      if (isCol1FilterOnlyLink(link)) return true;
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
      if (isCol1FilterOnlyLink(link)) return;
      if (link.start > pos) {
        el.combined.appendChild(document.createTextNode(text.slice(pos, link.start)));
      }
      const span = document.createElement('span');
      span.className = 'confirmed-segment';
      span.dataset.linkId = link.id;
      span.textContent = link.text;
      if (isMasterOriginLink(link)) {
        span.classList.add('master-segment');
        const editing = isMasterCombinedEditSession() && masterSegmentEdit.linkId === link.id;
        if (editing) {
          span.classList.add('master-segment-editing');
          span.contentEditable = 'true';
          span.title = 'Editing Master segment — blur or Enter to confirm; Esc to cancel';
        } else if (isMasterLockedLink(link)) {
          span.classList.add('master-segment-locked');
          span.contentEditable = 'false';
          span.title = 'Locked Master segment — double-click to unlock and edit';
        } else {
          span.title = 'Master-origin segment (local) — edit Combined to unlink; edit cell/nest to update live';
        }
      } else {
        span.title = 'Confirmed from linked cell — edit Combined to unlink; edit cell/nest to update live';
      }
      el.combined.appendChild(span);
      pos = link.end;
    });
    if (pos < text.length) {
      el.combined.appendChild(document.createTextNode(text.slice(pos)));
    }

    el.combined.classList.toggle('is-empty', !text);
    el.globalCombined.checked = state.globalCombined;
    if (el.matchSourceOrder) el.matchSourceOrder.checked = !!state.matchSourceOrder;
    applyConfirmedCellHighlights();
    syncCombinedOverflowY();
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
    toolsTabActive = false;
    state.activeTabId = data.activeTabId || defaultActiveTabId(tabs);
    if (!state.tabs.some(function (t) {
      return t.id === state.activeTabId;
    })) {
      state.activeTabId = defaultActiveTabId(tabs) || tabs[0].id;
    }
    const active = state.tabs.find(function (t) { return t.id === state.activeTabId; });
    lastPartTabId = active && !isMasterTab(active)
      ? active.id
      : (partTabs()[0] ? partTabs()[0].id : null);
    // Keep per-part Master insert Values/sort prefs for tabs that still exist
    // (undo/redo must not wipe other tabs' remembered picker selections).
    pruneMasterLibPrefs();
    pruneGridFilterPrefs();
    state.combinedPrompt = typeof data.combinedPrompt === 'string' ? data.combinedPrompt : '';
    state.globalCombined = typeof data.globalCombined === 'boolean' ? data.globalCombined : true;
    state.matchSourceOrder = typeof data.matchSourceOrder === 'boolean' ? data.matchSourceOrder : false;
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
    // Col1 is filter-only: drop any legacy Column A text from Combined prompts.
    stripCol1TextFromAllCombinedScopes();
    if (el.partSeparator) el.partSeparator.value = state.separators.part;
    if (el.columnSeparator) el.columnSeparator.value = state.separators.column;
    if (el.rowSeparator) el.rowSeparator.value = state.separators.row;

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

  function fillTabButtonContent(btn, tab, options) {
    const opts = options || {};
    const iconId = opts.fixedIcon || resolveTabIcon(tab);
    btn.textContent = '';
    const icon = document.createElement('span');
    icon.className = 'tab-icon';
    icon.setAttribute('aria-hidden', 'true');
    icon.dataset.icon = iconId;
    icon.innerHTML = tabIconMarkup(iconId);
    const title = document.createElement('span');
    title.className = 'tab-title';
    title.textContent = opts.titleText != null ? opts.titleText : (tab ? tab.title : '');
    btn.appendChild(icon);
    btn.appendChild(title);
    return { icon: icon, title: title };
  }

  function renderTabs() {
    closeTabIconPicker();
    el.tabBar.innerHTML = '';
    const targetPartId = toolsTabActive ? resolveLastPartTabId() : null;
    state.tabs.forEach(function (tab) {
      const btn = document.createElement('button');
      btn.type = 'button';
      const isActive = !toolsTabActive && tab.id === state.activeTabId;
      const master = isMasterTab(tab);
      btn.className = 'tab' + (isActive ? ' active' : '') + (master ? ' master-tab' : ' part-tab');
      fillTabButtonContent(btn, tab);
      btn.setAttribute('role', 'tab');
      btn.setAttribute('aria-selected', isActive ? 'true' : 'false');
      btn.dataset.id = tab.id;
      btn.draggable = !master;
      if (toolsTabActive && tab.id === targetPartId) btn.classList.add('tools-target');
      btn.title = master
        ? 'Master parts — edit reusable text here (right-click to change icon)'
        : tab.title + ' (drag to reorder; double-click to rename; right-click to change icon)';

      btn.addEventListener('click', function () {
        selectTab(tab.id);
      });
      btn.addEventListener('dragstart', function (event) {
        closeTabIconPicker();
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
        if (master) return;
        if (e.target && e.target.closest && e.target.closest('.tab-icon')) return;
        startInlineRename(btn, tab);
      });
      btn.addEventListener('contextmenu', function (e) {
        e.preventDefault();
        openTabIconPicker(btn, tab, { x: e.clientX, y: e.clientY, showActions: !master });
      });

      el.tabBar.appendChild(btn);
    });

    // Keep the add-part control in the tab strip, immediately after the last part tab.
    el.tabBar.appendChild(el.btnAdd);

    // Shared Tools tab (UI-only) — parks + Row / + Column / Rename; acts on active part.
    // Fixed wrench icon — not customizable / not in the icon picker.
    // Appended last with CSS margin-left:auto so it sits at the far right of the tab bar.
    const toolsBtn = document.createElement('button');
    toolsBtn.type = 'button';
    toolsBtn.className = 'tab tools-tab' + (toolsTabActive ? ' active' : '');
    fillTabButtonContent(toolsBtn, null, { fixedIcon: TOOLS_ICON, titleText: 'Tools' });
    toolsBtn.setAttribute('role', 'tab');
    toolsBtn.setAttribute('aria-selected', toolsTabActive ? 'true' : 'false');
    toolsBtn.dataset.id = '__tools__';
    toolsBtn.draggable = false;
    toolsBtn.title = 'Shared tools — + Row, + Column, Rename for the active part tab (wrench icon fixed)';
    toolsBtn.addEventListener('click', function () {
      closeTabIconPicker();
      selectToolsTab();
    });
    el.tabBar.appendChild(toolsBtn);

    const tab = activeTab();
    el.btnDelete.disabled = !tab || isMasterTab(tab) || partTabs().length <= 1 || toolsTabActive;
    el.masterLibrary.hidden = toolsTabActive || isMasterTab(tab);
    updateToolsChrome();
    renderPartTabPageChrome();
  }

  function masterRowHasContent(master, row) {
    if (!master || row < 0 || row >= master.rows) return false;
    for (let col = 0; col < master.cols; col++) {
      const cellText = master.cells[row * master.cols + col];
      if (cellText && String(cellText).trim()) return true;
    }
    return false;
  }


  function masterLibPrefsPartId() {
    const tab = activeTab();
    if (tab && !isMasterTab(tab)) return tab.id;
    return lastPartTabId;
  }

  function getMasterLibPrefs(partId) {
    const id = partId || masterLibPrefsPartId();
    if (!id) return { valueFilter: null, sortDir: null };
    let prefs = masterLibPrefsByPartId[id];
    if (!prefs) {
      prefs = { valueFilter: null, sortDir: null };
      masterLibPrefsByPartId[id] = prefs;
    }
    return prefs;
  }

  function clearMasterLibPrefs() {
    Object.keys(masterLibPrefsByPartId).forEach(function (key) {
      delete masterLibPrefsByPartId[key];
    });
  }

  /** Drop Master-insert prefs for part ids that no longer exist; keep the rest. */
  function pruneMasterLibPrefs() {
    const live = Object.create(null);
    partTabs().forEach(function (tab) {
      live[tab.id] = true;
    });
    Object.keys(masterLibPrefsByPartId).forEach(function (key) {
      if (!live[key]) delete masterLibPrefsByPartId[key];
    });
  }

  function clearMasterLibPrefsForPart(partId) {
    if (partId) delete masterLibPrefsByPartId[partId];
  }

  function gridFilterPrefsTabId() {
    const tab = activeTab();
    return tab ? tab.id : null;
  }

  function getGridFilterPrefs(tabId) {
    const id = tabId || gridFilterPrefsTabId();
    if (!id) return { rowFilter: 'all', valueFilter: null };
    let prefs = gridFilterPrefsByTabId[id];
    if (!prefs) {
      prefs = { rowFilter: 'all', valueFilter: null };
      gridFilterPrefsByTabId[id] = prefs;
    }
    return prefs;
  }

  function pruneGridFilterPrefs() {
    const live = Object.create(null);
    state.tabs.forEach(function (tab) { live[tab.id] = true; });
    Object.keys(gridFilterPrefsByTabId).forEach(function (key) {
      if (!live[key]) delete gridFilterPrefsByTabId[key];
    });
  }


  function uniqueMasterLibCol1Values(master) {
    const seen = Object.create(null);
    const values = [];
    if (!master || !master.cols) return values;
    for (let r = 0; r < master.rows; r++) {
      if (!masterRowHasContent(master, r)) continue;
      const v = (master.cells[r * master.cols] || '').trim();
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

  function isMasterLibValueSelected(value) {
    const filter = getMasterLibPrefs().valueFilter;
    if (filter === null) return true;
    return filter.has(value);
  }

  function setMasterLibValueFilterSelection(selectedValues, allValues) {
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
    getMasterLibPrefs(masterLibPrefsPartId()).valueFilter = allOn ? null : new Set(selected);
    // Keep Values menu open and in place (match Column A Values): refresh results only.
    masterLibFilterMenuOpen = true;
    const master = state.tabs.find(isMasterTab);
    const current = activeTab();
    refreshMasterLibraryResults(master, current);
    syncMasterLibFilterControls();
  }

  function closeMasterLibFilterMenu() {
    if (!masterLibFilterMenuOpen) return;
    masterLibFilterMenuOpen = false;
    const menu = el.masterLibraryItems && el.masterLibraryItems.querySelector('.master-library-filter-menu');
    if (menu) menu.hidden = true;
    const btn = el.masterLibraryItems && el.masterLibraryItems.querySelector('.master-library-filter-btn');
    if (btn) btn.setAttribute('aria-expanded', 'false');
  }

  function syncMasterLibFilterControls() {
    const btn = el.masterLibraryItems && el.masterLibraryItems.querySelector('.master-library-filter-btn');
    if (!btn) return;
    const prefs = getMasterLibPrefs();
    const active = prefs.valueFilter !== null;
    btn.classList.toggle('is-active', active);
    btn.setAttribute('aria-pressed', active ? 'true' : 'false');
    btn.setAttribute('aria-expanded', masterLibFilterMenuOpen ? 'true' : 'false');
    const count = prefs.valueFilter ? prefs.valueFilter.size : 0;
    btn.title = active
      ? ('Master Column A value filter on (' + count + ' selected) — click to change')
      : 'Filter Master parts by Column A values';
    btn.setAttribute('aria-label', btn.title);
    const menu = el.masterLibraryItems.querySelector('.master-library-filter-menu');
    if (menu) {
      menu.hidden = !masterLibFilterMenuOpen;
      if (masterLibFilterMenuOpen) positionCol1FilterMenu(btn, menu);
    }
    const asc = el.masterLibraryItems.querySelector('.master-library-sort-asc');
    const desc = el.masterLibraryItems.querySelector('.master-library-sort-desc');
    if (asc) asc.classList.toggle('is-active', prefs.sortDir === 'asc');
    if (desc) desc.classList.toggle('is-active', prefs.sortDir === 'desc');
  }

  function buildMasterLibFilterMenu(master, menu) {
    menu.innerHTML = '';
    const values = uniqueMasterLibCol1Values(master);

    const actions = document.createElement('div');
    actions.className = 'column-col1-filter-actions';

    const selectAll = document.createElement('button');
    selectAll.type = 'button';
    selectAll.className = 'column-col1-filter-action';
    selectAll.textContent = 'Select all';
    selectAll.addEventListener('click', function (e) {
      e.preventDefault();
      e.stopPropagation();
      getMasterLibPrefs(masterLibPrefsPartId()).valueFilter = null;
      masterLibFilterMenuOpen = true;
      buildMasterLibFilterMenu(master, menu);
      const current = activeTab();
      refreshMasterLibraryResults(master, current);
      syncMasterLibFilterControls();
      setStatus('Showing all Master Column A values', 'ok');
    });

    const clearBtn = document.createElement('button');
    clearBtn.type = 'button';
    clearBtn.className = 'column-col1-filter-action';
    clearBtn.textContent = 'Clear';
    clearBtn.addEventListener('click', function (e) {
      e.preventDefault();
      e.stopPropagation();
      getMasterLibPrefs(masterLibPrefsPartId()).valueFilter = new Set();
      masterLibFilterMenuOpen = true;
      buildMasterLibFilterMenu(master, menu);
      const current = activeTab();
      refreshMasterLibraryResults(master, current);
      syncMasterLibFilterControls();
      setStatus('Master Column A filter cleared (none selected)', 'ok');
    });

    actions.appendChild(selectAll);
    actions.appendChild(clearBtn);
    menu.appendChild(actions);

    const list = document.createElement('div');
    list.className = 'column-col1-filter-list';

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
        cb.checked = isMasterLibValueSelected(value);
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
          masterLibFilterMenuOpen = true;
          setMasterLibValueFilterSelection(selected, values);
          const filter = getMasterLibPrefs().valueFilter;
          const n = filter === null ? values.length : filter.size;
          setStatus(
            filter === null
              ? 'Showing all Master Column A values'
              : ('Showing ' + n + ' Master Column A value' + (n === 1 ? '' : 's')),
            'ok'
          );
        });

        const textEl = document.createElement('span');
        textEl.className = 'column-col1-filter-option-text';
        textEl.textContent = value ? value : '(blank)';
        if (!value) textEl.classList.add('is-blank');

        label.appendChild(cb);
        label.appendChild(textEl);
        list.appendChild(label);
      }
    }

    menu.appendChild(list);
  }

  function toggleMasterLibFilterMenu(master, btn, menu) {
    masterLibFilterMenuOpen = !masterLibFilterMenuOpen;
    if (masterLibFilterMenuOpen) {
      buildMasterLibFilterMenu(master, menu);
      positionCol1FilterMenu(btn, menu);
    }
    syncMasterLibFilterControls();
  }

  function setMasterLibSortDir(dir) {
    if (dir !== 'asc' && dir !== 'desc') return;
    // Always key off the active part tab so sort never writes a shared/global slot.
    const partId = masterLibPrefsPartId();
    const prefs = getMasterLibPrefs(partId);
    prefs.sortDir = prefs.sortDir === dir ? null : dir;
    masterLibFilterMenuOpen = false;
    renderMasterLibrary();
    setStatus(
      prefs.sortDir === 'desc'
        ? 'Master insert sorted Z–A by Column A'
        : prefs.sortDir === 'asc'
          ? 'Master insert sorted A–Z by Column A'
          : 'Master insert sort cleared (Master order)',
      'ok'
    );
  }

  /** Scrollable Master-insert results pane (toolbar stays outside it). */
  function masterLibraryResultsEl() {
    if (!el.masterLibraryItems) return null;
    let results = el.masterLibraryItems.querySelector('.master-library-results');
    if (!results) {
      results = document.createElement('div');
      results.className = 'master-library-results';
      el.masterLibraryItems.appendChild(results);
    }
    return results;
  }

  /** Replace Master-insert result grid/empty only — leave Values toolbar/menu intact. */
  function refreshMasterLibraryResults(master, current) {
    if (!el.masterLibraryItems) return;
    if (!master || !current || isMasterTab(current)) return;
    const results = masterLibraryResultsEl();
    if (!results) return;

    suppressMasterLibMenuScrollClose = true;
    try {
      results.innerHTML = '';

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
        results.appendChild(empty);
        return;
      }

      let rows = [];
      for (let row = 0; row < master.rows; row++) {
        if (masterRowHasContent(master, row)) rows.push(row);
      }

      const prefs = getMasterLibPrefs(current.id);
      if (prefs.valueFilter !== null) {
        rows = rows.filter(function (row) {
          const v = (master.cells[row * master.cols] || '').trim();
          return prefs.valueFilter.has(v);
        });
      }

      if (prefs.sortDir === 'asc' || prefs.sortDir === 'desc') {
        const dir = prefs.sortDir;
        rows = rows.slice().sort(function (a, b) {
          const va = (master.cells[a * master.cols] || '').trim();
          const vb = (master.cells[b * master.cols] || '').trim();
          if (!va && !vb) return a - b;
          if (!va) return 1;
          if (!vb) return -1;
          const cmp = va.localeCompare(vb, undefined, { sensitivity: 'base', numeric: true });
          if (cmp !== 0) return dir === 'desc' ? -cmp : cmp;
          return a - b;
        });
      }

      if (!rows.length) {
        const empty = document.createElement('span');
        empty.className = 'master-library-empty';
        empty.textContent = prefs.valueFilter !== null
          ? 'No Master parts match the Column A Values filter.'
          : 'Add reusable text in the Master part first.';
        results.appendChild(empty);
        return;
      }

      const grid = document.createElement('div');
      grid.className = 'master-library-grid';
      grid.setAttribute('role', 'grid');
      grid.style.gridTemplateColumns = contentColumnTemplate(master, 0);

      for (let ri = 0; ri < rows.length; ri++) {
        const row = rows[ri];
        for (let col = 0; col < master.cols; col++) {
          const cellText = master.cells[row * master.cols + col] || '';
          const trimmed = String(cellText).trim();
          if (trimmed) {
            const button = document.createElement('button');
            button.type = 'button';
            button.className = 'master-library-cell';
            const masterIdx = row * master.cols + col;
            if (isMasterCellRepresentedInCombined(master, masterIdx)) button.classList.add('cell-confirmed');
            button.textContent = cellText;
            button.title = 'Add to the selected cell, or next empty cell, in ' + current.title;
            button.setAttribute('role', 'gridcell');
            button.dataset.row = String(row);
            button.dataset.col = String(col);
            button.addEventListener('click', function () {
              insertMasterText(cellText, masterIdx);
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
      results.appendChild(grid);
    } finally {
      requestAnimationFrame(function () {
        suppressMasterLibMenuScrollClose = false;
      });
    }
  }

  function renderMasterLibrary() {
    const keepMenuOpen = masterLibFilterMenuOpen;
    el.masterLibraryItems.innerHTML = '';
    const master = state.tabs.find(isMasterTab);
    const current = activeTab();
    if (!master || !current || isMasterTab(current)) {
      masterLibFilterMenuOpen = false;
      return;
    }

    let hasContent = false;
    for (let i = 0; i < master.cells.length; i++) {
      if (master.cells[i] && String(master.cells[i]).trim()) {
        hasContent = true;
        break;
      }
    }
    if (!hasContent) {
      masterLibFilterMenuOpen = false;
      const empty = document.createElement('span');
      empty.className = 'master-library-empty';
      empty.textContent = 'Add reusable text in the Master part first.';
      el.masterLibraryItems.appendChild(empty);
      return;
    }

    const toolbar = document.createElement('div');
    toolbar.className = 'master-library-toolbar';
    toolbar.setAttribute('role', 'group');
    toolbar.setAttribute('aria-label', 'Master insert sort and filter');

    const sortControls = document.createElement('span');
    sortControls.className = 'column-sort-controls';
    sortControls.setAttribute('role', 'group');
    sortControls.setAttribute('aria-label', 'Sort Master parts by column A');

    const sortAsc = document.createElement('button');
    sortAsc.type = 'button';
    sortAsc.className = 'column-sort-btn master-library-sort-asc';
    sortAsc.textContent = 'A–Z';
    sortAsc.title = 'Sort Master insert by column A A–Z (display only)';
    sortAsc.setAttribute('aria-label', sortAsc.title);
    sortAsc.addEventListener('click', function (e) {
      e.preventDefault();
      e.stopPropagation();
      setMasterLibSortDir('asc');
    });

    const sortDesc = document.createElement('button');
    sortDesc.type = 'button';
    sortDesc.className = 'column-sort-btn master-library-sort-desc';
    sortDesc.textContent = 'Z–A';
    sortDesc.title = 'Sort Master insert by column A Z–A (display only)';
    sortDesc.setAttribute('aria-label', sortDesc.title);
    sortDesc.addEventListener('click', function (e) {
      e.preventDefault();
      e.stopPropagation();
      setMasterLibSortDir('desc');
    });

    sortControls.appendChild(sortAsc);
    sortControls.appendChild(sortDesc);
    toolbar.appendChild(sortControls);

    const filterWrap = document.createElement('span');
    filterWrap.className = 'column-col1-filter master-library-filter';

    const filterBtn = document.createElement('button');
    filterBtn.type = 'button';
    filterBtn.className = 'column-sort-btn column-col1-filter-btn master-library-filter-btn';
    filterBtn.textContent = 'Values';
    filterBtn.title = 'Filter Master parts by Column A values';
    filterBtn.setAttribute('aria-label', filterBtn.title);
    filterBtn.setAttribute('aria-haspopup', 'true');
    filterBtn.setAttribute('aria-expanded', 'false');
    filterBtn.setAttribute('aria-pressed', 'false');

    const filterMenu = document.createElement('div');
    filterMenu.className = 'column-col1-filter-menu master-library-filter-menu';
    filterMenu.hidden = true;
    filterMenu.setAttribute('role', 'dialog');
    filterMenu.setAttribute('aria-label', 'Master Column A value filter');

    filterBtn.addEventListener('click', function (e) {
      e.preventDefault();
      e.stopPropagation();
      toggleMasterLibFilterMenu(master, filterBtn, filterMenu);
    });
    filterMenu.addEventListener('click', function (e) {
      e.stopPropagation();
    });

    filterWrap.appendChild(filterBtn);
    filterWrap.appendChild(filterMenu);
    toolbar.appendChild(filterWrap);
    el.masterLibraryItems.appendChild(toolbar);
    // Results scroll independently so Values / A–Z never leave the viewport.
    const results = document.createElement('div');
    results.className = 'master-library-results';
    el.masterLibraryItems.appendChild(results);

    refreshMasterLibraryResults(master, current);

    masterLibFilterMenuOpen = keepMenuOpen;
    if (keepMenuOpen) {
      buildMasterLibFilterMenu(master, filterMenu);
    }
    syncMasterLibFilterControls();
  }

  function insertMasterText(text, masterCellIndex) {
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

    const insertText = text == null ? '' : String(text);
    pushHistory();
    let masterIdx = Number.isInteger(masterCellIndex) && masterCellIndex >= 0
      ? masterCellIndex
      : findMasterCellIndexByText(insertText);
    if (masterIdx === null) masterIdx = -1;
    setCellMasterLock(tab, index, masterIdx >= 0 ? masterIdx : null);
    // Replace/sync nests + parent pages from the Master cell (Ash).
    copyNestsFromMasterToCell(tab, index, masterIdx);
    copyPagesFromMasterToCell(tab, index, masterIdx);
    // Clicked Master-library text is authoritative for the active page slot.
    // (copyPages alone can leave the cell blank when Master cellPages desyncs
    // from master.cells, especially after part-page switches / empty page 2.)
    writeCellCurrentPage(tab, index, insertText);
    syncNestLinksAfterNestReplace(tab.id, index);
    focusedCell = { tabId: tab.id, index: index };
    revalidateLinksForCell(tab.id, index, { silent: true });
    // Auto-check Combined for the inserted cell (no-op if already included).
    historySuspended = true;
    try {
      ensureCellConfirmedState(index, true, { quiet: true, col1Cascade: false });
    } finally {
      historySuspended = false;
    }
    renderTabs();
    renderGrid();
    // Flush AFTER renderGrid — flushLiveCellInputs would otherwise re-read the
    // pre-insert empty textarea and wipe the Master text on page 2+.
    flushTabPage(tab);
    renderCombinedPrompt();
    applyConfirmedCellHighlights();
    renderMasterLibrary();
    scheduleSave();
    const cell = el.cellGrid.querySelector('textarea.cell[data-idx="' + index + '"]');
    if (cell) cell.focus();
    if (!(tab.cells[index] || '').trim() && !insertText.trim()) {
      setStatus('Master insert skipped — that Master cell has no text', 'err');
      return;
    }
    if (!(tab.cells[index] || '').trim()) {
      setStatus('Master insert failed — cell stayed empty (page/cellPages)', 'err');
      return;
    }
    setStatus(
      usedFirstCellFallback
        ? 'Added Master text (+ nests) to ' + tab.title + ' (replaced the first cell; Combined checked; locked until double-click unlock)'
        : 'Added Master text (+ nests) to ' + tab.title + ' (Combined checked; locked until double-click unlock)',
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
    // Row gutter (52px: Excel row # + move controls) + content widths.
    el.cellGrid.style.gridTemplateColumns = '52px ' + contentColumnTemplate(tab, 120);
  }

  function columnWidthForMeasure(tab, col) {
    const widths = normalizeColumnWidths(tab && tab.columnWidths, tab && tab.cols);
    if (widths && col >= 0 && col < widths.length) return widths[col];
    return DEFAULT_COLUMN_WIDTH;
  }

  /** Match .cell-wrap gutter column (26px). */
  const CELL_WRAP_GUTTER_W = 26;
  /** Match .cell-wrap border-right (border-box shrinks the textarea track). */
  const CELL_WRAP_BORDER_X = 1;
  /** Match .cell-nest-input min-height. */
  const NEST_INPUT_MIN_H = 36;
  /**
   * Fallback nest chrome height when off-DOM (single horizontal control row + pad).
   * Chrome sits BESIDE the nest textarea, so nest row height is max(chrome, textarea).
   */
  const NEST_CHROME_FALLBACK_H = 24;
  /** .cell-nest border top+bottom (border-box); must be in stack math or overflow:hidden clips. */
  const NEST_BORDER_Y = 2;
  /**
   * Nest textarea is inset by margin-left 12 + margin-right 4 + nest borders 2 +
   * horizontal chrome row (~130: pad + checkbox/◀/n/n/▶/+/× + border). Conservative vs live width.
   */
  const NEST_WIDTH_INSET = 148;
  const NESTS_TOP_BORDER = 1;
  const NESTS_BOTTOM_PAD = 4;
  const NESTS_GAP = 2;
  /** Subpixel / rounding pad so fitted rows do not clip the last line. */
  const FIT_HEIGHT_PAD = 2;

  /**
   * Width used for wrap measure. Prefer live border-box width after overflow-y:hidden
   * (Fit clears scrollbars so text wraps at the full cell width). clientWidth while a
   * scrollbar is visible would over-size; a too-wide fallback under-sizes and clips.
   */
  function textareaWrapWidth(ta, fallbackWidth) {
    if (ta) {
      const ow = ta.offsetWidth || (ta.getBoundingClientRect && ta.getBoundingClientRect().width) || 0;
      if (ow > 0) return Math.floor(ow);
    }
    return Math.max(40, Math.floor(fallbackWidth || DEFAULT_COLUMN_WIDTH));
  }

  function columnCellWrapWidth(tab, col) {
    const colW = columnWidthForMeasure(tab, col);
    // Textarea track inside cell-wrap: column width − gutter − wrap border-right.
    return Math.max(40, colW - CELL_WRAP_GUTTER_W - CELL_WRAP_BORDER_X);
  }

  function nestTextareaFallbackWidth(parentWrapW) {
    return Math.max(40, parentWrapW - NEST_WIDTH_INSET);
  }

  /**
   * Before measuring, hide overflow on live textareas so offsetWidth matches the
   * post-fit wrap width (no scrollbar gutter). Does not change heights yet.
   */
  function prepareWrapWidthsForFitMeasure(scopeEl) {
    const root = scopeEl || el.cellGrid;
    if (!root) return;
    const nodes = root.querySelectorAll('textarea.cell, textarea.cell-nest-input');
    for (let i = 0; i < nodes.length; i++) {
      nodes[i].style.overflowY = 'hidden';
    }
  }

  /**
   * Measure wrapped content height for a cell/nest textarea (or raw text).
   * Off-DOM clone so filtered (display:none) rows still measure; width must be the
   * border-box wrap width (same font/padding via className).
   */
  function measureTextareaContentHeight(ta, className, widthPx, options) {
    options = options || {};
    const minH = options.minHeight != null ? options.minHeight : MIN_ROW_HEIGHT;
    const text = options.text != null
      ? String(options.text)
      : (ta && ta.value != null ? String(ta.value) : '');
    const width = Math.floor(widthPx);
    if (!(width > 0)) return minH;
    const helper = document.createElement('textarea');
    helper.className = className || 'cell';
    helper.setAttribute('aria-hidden', 'true');
    helper.tabIndex = -1;
    helper.rows = 1;
    helper.value = text;
    helper.style.cssText =
      'position:absolute;left:-99999px;top:0;height:0;min-height:0;max-height:none;' +
      'overflow:hidden;visibility:hidden;resize:none;box-sizing:border-box;' +
      'flex:none;white-space:pre-wrap;overflow-wrap:break-word;';
    helper.style.width = width + 'px';
    // Mirror live padding/font when available so class-only measure cannot drift.
    if (ta && window.getComputedStyle) {
      try {
        const cs = window.getComputedStyle(ta);
        if (cs) {
          helper.style.font = cs.font;
          helper.style.letterSpacing = cs.letterSpacing;
          helper.style.padding = cs.padding;
          helper.style.border = cs.border;
          helper.style.lineHeight = cs.lineHeight;
          helper.style.boxSizing = cs.boxSizing || 'border-box';
        }
      } catch (err) { /* ignore */ }
    }
    document.body.appendChild(helper);
    // Force layout; height:0 + scrollHeight is the standard wrap measure.
    const needed = Math.max(minH, Math.ceil(helper.scrollHeight) + FIT_HEIGHT_PAD);
    document.body.removeChild(helper);
    return needed;
  }

  function measureCellContentHeight(ta, tab) {
    if (!ta) return MIN_ROW_HEIGHT;
    const col = parseInt(ta.dataset.col, 10);
    const width = textareaWrapWidth(ta, columnCellWrapWidth(tab, col));
    return measureTextareaContentHeight(ta, 'cell', width, { minHeight: MIN_ROW_HEIGHT });
  }

  function nestChromeHeight(nestEl) {
    if (!nestEl) return NEST_CHROME_FALLBACK_H;
    const chrome = nestEl.querySelector('.cell-nest-chrome');
    if (chrome) {
      const h = chrome.getBoundingClientRect().height || chrome.offsetHeight;
      if (h > 0) return Math.ceil(h);
    }
    return NEST_CHROME_FALLBACK_H;
  }

  /** Tallest page content height for a nest (current DOM page + other pages from model). */
  function measureNestTallestPageHeight(nestEl, nestModel, nestWidth) {
    let maxH = NEST_INPUT_MIN_H;
    const nestTa = nestEl ? nestEl.querySelector('textarea.cell-nest-input') : null;
    if (nestTa) {
      maxH = Math.max(
        maxH,
        measureTextareaContentHeight(nestTa, 'cell-nest-input', nestWidth, {
          minHeight: NEST_INPUT_MIN_H
        })
      );
    }
    if (nestModel && Array.isArray(nestModel.pages)) {
      const currentPage = nestTa ? parseInt(nestTa.dataset.page, 10) : -1;
      for (let p = 0; p < nestModel.pages.length; p++) {
        if (p === currentPage && nestTa) continue;
        const pageH = measureTextareaContentHeight(null, 'cell-nest-input', nestWidth, {
          text: nestPageText(nestModel, p),
          minHeight: NEST_INPUT_MIN_H
        });
        if (pageH > maxH) maxH = pageH;
      }
    }
    return maxH;
  }

  /** Outer nest row height: max(chrome, tallest page) + top/bottom border (border-box). */
  function measureNestOuterHeight(nestEl, nestModel, nestWidth) {
    const chromeH = nestChromeHeight(nestEl);
    const pageH = measureNestTallestPageHeight(nestEl, nestModel, nestWidth);
    return Math.max(chromeH, pageH) + NEST_BORDER_Y;
  }

  /**
   * Parent content height + nest block (chrome beside textarea → max, not sum;
   * includes nests container border/padding/gaps, per-nest borders, tallest page).
   */
  function measureCellStackHeight(wrap, tab) {
    if (!wrap) return MIN_ROW_HEIGHT;
    const ta = wrap.querySelector('textarea.cell');
    let total = measureCellContentHeight(ta, tab);
    const nestsEl = wrap.querySelector('.cell-nests');
    if (!nestsEl) return total;

    const nestNodes = nestsEl.querySelectorAll('.cell-nest');
    if (!nestNodes.length) return total;

    const col = ta ? parseInt(ta.dataset.col, 10) : 0;
    const cellIdx = ta ? parseInt(ta.dataset.idx, 10) : -1;
    const nestModels = cellIdx >= 0 ? getCellNests(tab, cellIdx) : [];

    let nestsBlock = NESTS_TOP_BORDER + NESTS_BOTTOM_PAD;
    if (nestNodes.length > 1) nestsBlock += (nestNodes.length - 1) * NESTS_GAP;

    for (let i = 0; i < nestNodes.length; i++) {
      const nestEl = nestNodes[i];
      const nestTa = nestEl.querySelector('textarea.cell-nest-input');
      const parentWrapW = textareaWrapWidth(ta, columnCellWrapWidth(tab, col));
      const nestWidth = textareaWrapWidth(nestTa, nestTextareaFallbackWidth(parentWrapW));
      nestsBlock += measureNestOuterHeight(nestEl, nestModels[i], nestWidth);
    }
    return total + nestsBlock;
  }

  function setTextareaFittedHeight(ta, heightPx) {
    if (!ta) return;
    const h = Math.max(1, heightPx);
    ta.style.height = h + 'px';
    ta.style.minHeight = h + 'px';
    ta.style.maxHeight = h + 'px';
    // Prevent flex shrink from fighting pinned heights inside .cell-stack.
    ta.style.flex = '0 0 auto';
    // Avoid a residual scrollbar re-narrowing wrap width after fit.
    ta.style.overflowY = 'hidden';
  }

  /**
   * Drop previously pinned wrap/textarea/nest heights so Fit can SHRINK tall
   * empty/short rows. Old rowHeights / manual resize:vertical pins must not
   * floor the next measure/apply.
   */
  function clearPinnedHeightsForFitMeasure(scopeEl) {
    const root = scopeEl || el.cellGrid;
    if (!root) return;
    function clearBox(node) {
      if (!node || !node.style) return;
      node.style.minHeight = '';
      node.style.height = '';
      node.style.maxHeight = '';
    }
    // Include root when scope is a single wrap / row-controls / nest.
    if (root.classList) {
      if (root.classList.contains('cell-wrap') || root.classList.contains('row-controls') ||
          root.classList.contains('cell-nest')) {
        clearBox(root);
      }
    }
    const wraps = root.querySelectorAll('.cell-wrap');
    for (let i = 0; i < wraps.length; i++) clearBox(wraps[i]);
    const textareas = root.querySelectorAll('textarea.cell, textarea.cell-nest-input');
    for (let i = 0; i < textareas.length; i++) {
      const ta = textareas[i];
      ta.style.height = '';
      ta.style.minHeight = '';
      ta.style.maxHeight = '';
      ta.style.flex = '';
      ta.style.overflowY = 'hidden';
    }
    if (root.matches && root.matches('textarea.cell, textarea.cell-nest-input')) {
      root.style.height = '';
      root.style.minHeight = '';
      root.style.maxHeight = '';
      root.style.flex = '';
      root.style.overflowY = 'hidden';
    }
    const nests = root.querySelectorAll('.cell-nest');
    for (let i = 0; i < nests.length; i++) clearBox(nests[i]);
    const controls = root.querySelectorAll('.row-controls');
    for (let i = 0; i < controls.length; i++) clearBox(controls[i]);
  }

  function applyHeightToGridRow(row, heightPx) {
    if (!el.cellGrid) return;
    const tab = activeTab();
    const h = Math.max(MIN_ROW_HEIGHT, heightPx);
    const wraps = el.cellGrid.querySelectorAll('.cell-wrap[data-row="' + row + '"]');
    for (let w = 0; w < wraps.length; w++) {
      const wrap = wraps[w];
      wrap.style.minHeight = h + 'px';
      wrap.style.height = h + 'px';
      wrap.style.maxHeight = h + 'px';
      const ta = wrap.querySelector('textarea.cell');
      const nestNodes = wrap.querySelectorAll('.cell-nest');
      const cellIdx = ta ? parseInt(ta.dataset.idx, 10) : -1;
      const nestModels = cellIdx >= 0 && tab ? getCellNests(tab, cellIdx) : [];
      const col = ta ? parseInt(ta.dataset.col, 10) : 0;

      for (let n = 0; n < nestNodes.length; n++) {
        const nestEl = nestNodes[n];
        const nt = nestEl.querySelector('textarea.cell-nest-input');
        const parentWrapW = textareaWrapWidth(ta, columnCellWrapWidth(tab, col));
        const nestWidth = textareaWrapWidth(nt, nestTextareaFallbackWidth(parentWrapW));
        const chromeH = nestChromeHeight(nestEl);
        const pageH = measureNestTallestPageHeight(nestEl, nestModels[n], nestWidth);
        const nestInnerH = Math.max(chromeH, pageH);
        const nestOuterH = nestInnerH + NEST_BORDER_Y;
        if (nt) setTextareaFittedHeight(nt, pageH);
        nestEl.style.minHeight = nestOuterH + 'px';
        nestEl.style.height = nestOuterH + 'px';
        nestEl.style.maxHeight = nestOuterH + 'px';
      }

      if (ta) {
        const parentContent = measureCellContentHeight(ta, tab);
        // When nests exist, keep parent at content size; otherwise stretch to row height.
        setTextareaFittedHeight(
          ta,
          nestNodes.length ? Math.max(MIN_ROW_HEIGHT, parentContent) : h
        );
      }
    }
    const controls = el.cellGrid.querySelector('.row-controls[data-row="' + row + '"]');
    if (controls) {
      controls.style.minHeight = h + 'px';
      controls.style.height = h + 'px';
    }
  }

  function applyPersistedRowHeights(tab) {
    const heights = normalizeRowHeights(tab && tab.rowHeights, tab && tab.rows);
    if (!heights) return;
    prepareWrapWidthsForFitMeasure(el.cellGrid);
    for (let r = 0; r < heights.length; r++) {
      applyHeightToGridRow(r, heights[r]);
    }
  }

  /**
   * Re-measure one row and persist when Fit has been used (rowHeights set).
   * Safe/cheap live path after typing so content does not clip again.
   */
  function maybeLiveRefitRowFromInput(row) {
    const tab = activeTab();
    if (!tab || !Array.isArray(tab.rowHeights)) return;
    if (row < 0 || row >= tab.rows) return;
    refitRowHeightAt(row);
  }

  /**
   * One-click: size every row to its tallest cell content on the active tab grid.
   * Shrinks tall empty/short rows as well as growing clipped ones — replaces any
   * prior persisted rowHeights / manual resize pins (not grow-only).
   */
  function autoFitAllRowHeights() {
    const tab = activeTab();
    if (!tab || !el.cellGrid) return;
    pushHistory();
    // Drop old pinned heights so measure/apply can shrink empty rows.
    clearPinnedHeightsForFitMeasure(el.cellGrid);
    // Match post-fit wrap widths before measuring (no scrollbar gutter).
    prepareWrapWidthsForFitMeasure(el.cellGrid);
    const heights = [];
    for (let r = 0; r < tab.rows; r++) {
      // Floor is chrome min only — never the previous tall rowHeights[r].
      let maxH = MIN_ROW_HEIGHT;
      const wraps = el.cellGrid.querySelectorAll('.cell-wrap[data-row="' + r + '"]');
      for (let i = 0; i < wraps.length; i++) {
        const stackH = measureCellStackHeight(wraps[i], tab);
        if (stackH > maxH) maxH = stackH;
      }
      heights.push(maxH);
      applyHeightToGridRow(r, maxH);
    }
    // Second pass: re-measure at final wrap widths; allow shrink OR grow.
    prepareWrapWidthsForFitMeasure(el.cellGrid);
    for (let r = 0; r < tab.rows; r++) {
      let maxH = MIN_ROW_HEIGHT;
      const wraps = el.cellGrid.querySelectorAll('.cell-wrap[data-row="' + r + '"]');
      for (let i = 0; i < wraps.length; i++) {
        const stackH = measureCellStackHeight(wraps[i], tab);
        if (stackH > maxH) maxH = stackH;
      }
      if (maxH !== heights[r]) {
        heights[r] = maxH;
        applyHeightToGridRow(r, maxH);
      }
    }
    // Full replace so tall persisted heights cannot stick around.
    tab.rowHeights = heights;
    scheduleSave();
    setStatus('Auto-fitted row heights (' + tab.rows + ' rows)', 'ok');
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
    const filter = getGridFilterPrefs(tab && tab.id).valueFilter;
    if (filter === null) return true;
    const v = (tab.cells[rowIndex * tab.cols] || '').trim();
    return filter.has(v);
  }

  function rowMatchesFilter(tab, rowIndex) {
    let base = true;
    const rowFilter = getGridFilterPrefs(tab && tab.id).rowFilter;
    if (rowFilter === 'nonempty') base = !rowIsEmpty(tab, rowIndex);
    else if (rowFilter === 'included') base = isRowIncluded(tab, rowIndex);
    if (!base) return false;
    return rowMatchesCol1ValueFilter(tab, rowIndex);
  }

  function applyRowFilterVisibility() {
    const tab = activeTab();
    if (!tab || !el.cellGrid) return;
    const rowControls = el.cellGrid.querySelectorAll('.row-controls');
    for (let i = 0; i < rowControls.length; i++) {
      const controls = rowControls[i];
      const row = parseInt(controls.dataset.row, 10);
      if (Number.isNaN(row)) continue;
      const hidden = !rowMatchesFilter(tab, row);
      controls.classList.toggle('is-row-filtered', hidden);
      const wraps = el.cellGrid.querySelectorAll('.cell-wrap[data-row="' + row + '"]');
      for (let w = 0; w < wraps.length; w++) wraps[w].classList.toggle('is-row-filtered', hidden);
    }
  }

  function syncRowFilterButtons() {
    const buttons = [el.btnFilterAll, el.btnFilterNonempty, el.btnFilterIncluded];
    for (let i = 0; i < buttons.length; i++) {
      const btn = buttons[i];
      if (!btn) continue;
      const on = btn.dataset.filter === getGridFilterPrefs().rowFilter;
      btn.classList.toggle('is-active', on);
      btn.setAttribute('aria-pressed', on ? 'true' : 'false');
    }
  }

  function setGridRowFilter(mode) {
    if (mode !== 'all' && mode !== 'nonempty' && mode !== 'included') return;
    const prefs = getGridFilterPrefs();
    if (prefs.rowFilter === mode) return;
    prefs.rowFilter = mode;
    syncRowFilterButtons();
    renderGrid();
    const labels = { all: 'Showing all rows', nonempty: 'Showing non-empty rows', included: 'Showing rows in Combined' };
    setStatus(labels[mode] || 'Row filter updated', 'ok');
  }

  function isCol1ValueSelected(value) {
    const filter = getGridFilterPrefs().valueFilter;
    if (filter === null) return true;
    return filter.has(value);
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
    getGridFilterPrefs().valueFilter = allOn ? null : new Set(selected);
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
    const prefs = getGridFilterPrefs();
    const active = prefs.valueFilter !== null;
    btn.classList.toggle('is-active', active);
    btn.setAttribute('aria-pressed', active ? 'true' : 'false');
    btn.setAttribute('aria-expanded', col1FilterMenuOpen ? 'true' : 'false');
    const count = prefs.valueFilter ? prefs.valueFilter.size : 0;
    btn.title = active
      ? ('Column A value filter on (' + count + ' selected) — click to change')
      : 'Filter rows by Column A values';
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
      getGridFilterPrefs().valueFilter = null;
      buildCol1FilterMenu(tab, menu);
      applyRowFilterVisibility();
      syncCol1FilterControls();
      setStatus('Showing all Column A values', 'ok');
    });

    const clearBtn = document.createElement('button');
    clearBtn.type = 'button';
    clearBtn.className = 'column-col1-filter-action';
    clearBtn.textContent = 'Clear';
    clearBtn.addEventListener('click', function (e) {
      e.preventDefault();
      e.stopPropagation();
      getGridFilterPrefs().valueFilter = new Set();
      buildCol1FilterMenu(tab, menu);
      applyRowFilterVisibility();
      syncCol1FilterControls();
      setStatus('Column A value filter cleared (no rows)', 'ok');
    });

    actions.appendChild(selectAll);
    actions.appendChild(clearBtn);
    menu.appendChild(actions);

    const list = document.createElement('div');
    list.className = 'column-col1-filter-list';
    list.setAttribute('role', 'group');
    list.setAttribute('aria-label', 'Column A values');

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
          const filter = getGridFilterPrefs().valueFilter;
          const n = filter === null ? values.length : filter.size;
          setStatus(
            filter === null
              ? 'Showing all Column A values'
              : ('Showing ' + n + ' Column A value' + (n === 1 ? '' : 's')),
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
      label.textContent = columnLetter(c);
      label.title = 'Column ' + columnLetter(c);
      header.appendChild(label);

      if (c === 0) {
        header.classList.add('column-header-sortable');
        const sortControls = document.createElement('span');
        sortControls.className = 'column-sort-controls';
        sortControls.setAttribute('role', 'group');
        sortControls.setAttribute('aria-label', 'Sort by column A');

        const sortAsc = document.createElement('button');
        sortAsc.type = 'button';
        sortAsc.className = 'column-sort-btn';
        sortAsc.textContent = 'A–Z';
        sortAsc.title = 'Sort rows by column A A–Z';
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
        sortDesc.title = 'Sort rows by column A Z–A';
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
        filterBtn.title = 'Filter rows by Column A values';
        filterBtn.setAttribute('aria-label', filterBtn.title);
        filterBtn.setAttribute('aria-haspopup', 'true');
        filterBtn.setAttribute('aria-expanded', 'false');
        filterBtn.setAttribute('aria-pressed', 'false');

        const filterMenu = document.createElement('div');
        filterMenu.className = 'column-col1-filter-menu';
        filterMenu.hidden = true;
        filterMenu.setAttribute('role', 'dialog');
        filterMenu.setAttribute('aria-label', 'Column A value filter');

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
      handle.setAttribute('aria-label', 'Resize column ' + columnLetter(c));
      handle.title = 'Drag to resize column ' + columnLetter(c);
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
      rowControls.dataset.row = String(r);

      const rowNum = document.createElement('span');
      rowNum.className = 'row-header-label';
      rowNum.textContent = String(r + 1);
      rowNum.title = 'Row ' + (r + 1);
      rowNum.setAttribute('aria-hidden', 'true');
      rowControls.appendChild(rowNum);

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
        const cellLocked = !isMasterTab(tab) && isCellMasterLocked(tab, idx);
        const cellEditing = isMasterCellEditSession() &&
          masterSegmentEdit.tabId === tab.id && masterSegmentEdit.cellIndex === idx;
        if (cellEditing) wrap.classList.add('master-cell-editing');
        else if (cellLocked) wrap.classList.add('master-cell-locked');
        const shade = getCellShade(tab, idx);
        if (shade) wrap.classList.add('cell-shade-' + shade);
        const gutter = document.createElement('div');
        gutter.className = 'cell-gutter';

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

        const nestAddBtn = document.createElement('button');
        nestAddBtn.type = 'button';
        nestAddBtn.className = 'cell-nest-add';
        nestAddBtn.dataset.idx = String(idx);
        nestAddBtn.dataset.row = String(r);
        nestAddBtn.dataset.col = String(c);
        nestAddBtn.title = 'Add nested cell under this cell';
        nestAddBtn.setAttribute('aria-label', 'Add nest under ' + cellAddress(r, c));
        nestAddBtn.innerHTML = '<span class="cell-nest-add-mark" aria-hidden="true">+</span>';
        nestAddBtn.addEventListener('pointerdown', function (ev) {
          ev.preventDefault();
          ev.stopPropagation();
        });
        nestAddBtn.addEventListener('click', function (ev) {
          ev.preventDefault();
          ev.stopPropagation();
          const targetIdx = cellIndexFromNestAddEvent(ev, idx);
          if (targetIdx < 0) return;
          addNestedCell(targetIdx);
        });
        gutter.appendChild(nestAddBtn);
        wrap.appendChild(gutter);

        const stack = document.createElement('div');
        stack.className = 'cell-stack';

        // Top-right corner: shade picker + parent-cell page chrome (compact).
        ensureCellPages(tab);
        const pagesEntry = getCellPages(tab, idx);
        const pageCount = pagesEntry.pages.length;
        const page = Math.min(Math.max(pagesEntry.page || 0, 0), pageCount - 1);
        pagesEntry.page = page;

        const corner = document.createElement('div');
        corner.className = 'cell-corner-tools';
        corner.dataset.idx = String(idx);

        const shadeBtn = document.createElement('button');
        shadeBtn.type = 'button';
        shadeBtn.className = 'cell-shade-btn' + (shade ? ' cell-shade-btn-' + shade : '');
        shadeBtn.title = 'Cell shade';
        shadeBtn.setAttribute('aria-label', 'Cell shade');
        shadeBtn.innerHTML = '<span class="cell-shade-btn-mark" aria-hidden="true"></span>';
        shadeBtn.addEventListener('pointerdown', function (ev) {
          ev.preventDefault();
          ev.stopPropagation();
        });
        shadeBtn.addEventListener('click', function (ev) {
          ev.preventDefault();
          ev.stopPropagation();
          openCellShadePicker(shadeBtn, idx);
        });
        corner.appendChild(shadeBtn);

        const pageChrome = document.createElement('div');
        pageChrome.className = 'cell-page-chrome';

        const prevBtn = document.createElement('button');
        prevBtn.type = 'button';
        prevBtn.className = 'cell-page-btn';
        prevBtn.textContent = '◀';
        prevBtn.title = 'Previous cell page';
        prevBtn.setAttribute('aria-label', 'Previous cell page');
        prevBtn.disabled = page <= 0;
        prevBtn.addEventListener('click', function (ev) {
          ev.preventDefault();
          ev.stopPropagation();
          stepCellPage(idx, -1);
        });

        const pageLabel = document.createElement('span');
        pageLabel.className = 'cell-page-label';
        pageLabel.textContent = (page + 1) + '/' + pageCount;
        pageLabel.title = 'Cell page ' + (page + 1) + ' of ' + pageCount;

        const nextBtn = document.createElement('button');
        nextBtn.type = 'button';
        nextBtn.className = 'cell-page-btn';
        nextBtn.textContent = '▶';
        nextBtn.title = 'Next cell page';
        nextBtn.setAttribute('aria-label', 'Next cell page');
        nextBtn.disabled = page >= pageCount - 1;
        nextBtn.addEventListener('click', function (ev) {
          ev.preventDefault();
          ev.stopPropagation();
          stepCellPage(idx, 1);
        });

        const addPageBtn = document.createElement('button');
        addPageBtn.type = 'button';
        addPageBtn.className = 'cell-page-btn cell-page-add';
        addPageBtn.textContent = '+';
        addPageBtn.title = 'Add cell page';
        addPageBtn.setAttribute('aria-label', 'Add cell page');
        addPageBtn.disabled = cellLocked && !cellEditing;
        addPageBtn.addEventListener('click', function (ev) {
          ev.preventDefault();
          ev.stopPropagation();
          openAddPageChoiceMenu(addPageBtn, { kind: 'cell', cellIndex: idx });
        });

        pageChrome.appendChild(prevBtn);
        pageChrome.appendChild(pageLabel);
        pageChrome.appendChild(nextBtn);
        pageChrome.appendChild(addPageBtn);

        if (pageCount > 1) {
          const removePageBtn = document.createElement('button');
          removePageBtn.type = 'button';
          removePageBtn.className = 'cell-page-btn cell-page-remove';
          removePageBtn.textContent = '×';
          removePageBtn.title = 'Remove current cell page';
          removePageBtn.setAttribute('aria-label', 'Remove current cell page');
          removePageBtn.disabled = cellLocked && !cellEditing;
          removePageBtn.addEventListener('click', function (ev) {
            ev.preventDefault();
            ev.stopPropagation();
            removeCellPage(idx);
          });
          pageChrome.appendChild(removePageBtn);
        }

        corner.appendChild(pageChrome);
        stack.appendChild(corner);

        const ta = document.createElement('textarea');
        ta.className = 'cell';
        if (isCellConfirmed(tab.id, idx)) ta.classList.add('cell-confirmed');
        ta.rows = 2;
        ta.spellcheck = false;
        ta.placeholder = cellAddress(r, c);
        ta.value = tab.cells[idx] || '';
        ta.dataset.row = String(r);
        ta.dataset.col = String(c);
        ta.dataset.idx = String(idx);
        ta.setAttribute('aria-label', 'Cell ' + cellAddress(r, c));
        if (cellEditing) {
          ta.classList.add('master-cell-editing');
          ta.readOnly = false;
          ta.title = 'Editing Master cell — blur or Enter to confirm; Esc to cancel';
        } else if (cellLocked) {
          ta.classList.add('master-cell-locked');
          ta.readOnly = true;
          ta.title = 'Locked Master cell — double-click to unlock and edit';
        }
        ta.addEventListener('input', onCellInput);
        ta.addEventListener('keydown', onCellKeydown);
        ta.addEventListener('pointerdown', onCellPointerDownSelect);
        ta.addEventListener('focus', onCellFocusSelect);
        ta.addEventListener('click', onCellClickSelect);
        ta.addEventListener('dblclick', onCellDblClickMasterUnlock);
        ta.addEventListener('blur', onCellBlurMasterFinish);
        ta.addEventListener('paste', onCellPaste);
        ta.addEventListener('copy', onCellCopy);
        ta.addEventListener('cut', onCellCut);
        ta.addEventListener('beforeinput', onCellBeforeInputMasterLock);
        stack.appendChild(ta);

        const nests = getCellNests(tab, idx);
        if (nests.length) {
          const nestsEl = document.createElement('div');
          nestsEl.className = 'cell-nests';
          for (let ni = 0; ni < nests.length; ni++) {
            (function (nestIndex) {
              const nest = nests[nestIndex];
              const pageCount = nest.pages.length;
              const page = Math.min(Math.max(nest.page || 0, 0), pageCount - 1);
              nest.page = page;

              const nestRow = document.createElement('div');
              nestRow.className = 'cell-nest';
              nestRow.dataset.idx = String(idx);
              nestRow.dataset.nest = String(nestIndex);
              if (isNestConfirmed(tab.id, idx, nestIndex)) nestRow.classList.add('cell-confirmed');

              const nestChrome = document.createElement('div');
              nestChrome.className = 'cell-nest-chrome';

              const nestAppendBtn = document.createElement('button');
              nestAppendBtn.type = 'button';
              nestAppendBtn.className = 'cell-append cell-nest-append';
              nestAppendBtn.dataset.idx = String(idx);
              nestAppendBtn.dataset.nest = String(nestIndex);
              nestAppendBtn.setAttribute('aria-pressed', 'false');
              nestAppendBtn.innerHTML = '<span class="cell-append-mark" aria-hidden="true"></span>';
              if (isNestConfirmed(tab.id, idx, nestIndex)) {
                nestAppendBtn.classList.add('is-checked');
                nestAppendBtn.setAttribute('aria-pressed', 'true');
              }
              nestAppendBtn.addEventListener('pointerdown', function (ev) {
                if (ev.button != null && ev.button !== 0) return;
                ev.preventDefault();
                ev.stopPropagation();
                const wantOn = !isNestConfirmedInScope(tab.id, idx, nestIndex, currentPromptScope());
                beginCheckboxDrag('nest', idx + ':' + nestIndex, wantOn, ev);
              });
              nestAppendBtn.addEventListener('click', function (ev) {
                ev.preventDefault();
                ev.stopPropagation();
              });
              nestChrome.appendChild(nestAppendBtn);

              const prevBtn = document.createElement('button');
              prevBtn.type = 'button';
              prevBtn.className = 'cell-nest-page-btn';
              prevBtn.textContent = '\u25c0';
              prevBtn.title = 'Previous nest page';
              prevBtn.setAttribute('aria-label', 'Previous nest page');
              prevBtn.disabled = page <= 0;
              prevBtn.addEventListener('click', function (ev) {
                ev.preventDefault();
                ev.stopPropagation();
                stepNestPage(idx, nestIndex, -1);
              });

              const pageLabel = document.createElement('span');
              pageLabel.className = 'cell-nest-page-label';
              pageLabel.textContent = (page + 1) + '/' + pageCount;
              pageLabel.title = 'Nest page ' + (page + 1) + ' of ' + pageCount;

              const nextBtn = document.createElement('button');
              nextBtn.type = 'button';
              nextBtn.className = 'cell-nest-page-btn';
              nextBtn.textContent = '\u25b6';
              nextBtn.title = 'Next nest page (adds a page at the end)';
              nextBtn.setAttribute('aria-label', 'Next nest page');
              nextBtn.addEventListener('click', function (ev) {
                ev.preventDefault();
                ev.stopPropagation();
                stepNestPage(idx, nestIndex, 1);
              });

              const addPageBtn = document.createElement('button');
              addPageBtn.type = 'button';
              addPageBtn.className = 'cell-nest-page-btn cell-nest-page-add';
              addPageBtn.textContent = '+';
              addPageBtn.title = 'Add nest page';
              addPageBtn.setAttribute('aria-label', 'Add nest page');
              addPageBtn.addEventListener('click', function (ev) {
                ev.preventDefault();
                ev.stopPropagation();
                addNestPage(idx, nestIndex);
              });

              const removeBtn = document.createElement('button');
              removeBtn.type = 'button';
              removeBtn.className = 'cell-nest-remove';
              removeBtn.textContent = '\u00d7';
              removeBtn.title = pageCount > 1
                ? 'Remove current nest page'
                : 'Remove nested cell';
              removeBtn.setAttribute('aria-label', removeBtn.title);
              removeBtn.addEventListener('click', function (ev) {
                ev.preventDefault();
                ev.stopPropagation();
                removeNestPage(idx, nestIndex);
              });

              nestChrome.appendChild(prevBtn);
              nestChrome.appendChild(pageLabel);
              nestChrome.appendChild(nextBtn);
              nestChrome.appendChild(addPageBtn);
              nestChrome.appendChild(removeBtn);

              const nestTa = document.createElement('textarea');
              nestTa.className = 'cell-nest-input';
              if (isNestConfirmed(tab.id, idx, nestIndex)) nestTa.classList.add('cell-confirmed');
              nestTa.rows = 2;
              nestTa.spellcheck = false;
              nestTa.placeholder = 'Nest ' + (nestIndex + 1) + ' p' + (page + 1);
              nestTa.value = nestPageText(nest, page);
              nestTa.dataset.idx = String(idx);
              nestTa.dataset.nest = String(nestIndex);
              nestTa.dataset.page = String(page);
              nestTa.setAttribute(
                'aria-label',
                'Nest ' + (nestIndex + 1) + ' page ' + (page + 1) + ' of ' + pageCount +
                  ' under ' + cellAddress(r, c)
              );
              if (cellEditing) {
                nestTa.classList.add('master-cell-editing');
                nestTa.title = 'Editing Master nest — blur or Enter to confirm; Esc to cancel';
              } else if (cellLocked) {
                nestTa.readOnly = true;
                nestTa.classList.add('master-cell-locked');
                nestTa.title = 'Locked Master nest — double-click to unlock before editing';
              }
              nestTa.addEventListener('input', onNestInput);
              nestTa.addEventListener('beforeinput', onNestBeforeInputMasterLock);
              nestTa.addEventListener('dblclick', onNestDblClickMasterUnlock);
              nestTa.addEventListener('blur', onNestBlurMasterFinish);
              nestTa.addEventListener('keydown', function (ev) {
                if (ev.isComposing) return;
                if (isMasterCellEditFor(tab, idx) && !masterSegmentDialogOpen) {
                  if (ev.key === 'Escape') {
                    ev.preventDefault();
                    cancelMasterSegmentEdit();
                    return;
                  }
                  if (ev.key === 'Enter' && !ev.shiftKey) {
                    ev.preventDefault();
                    requestMasterSegmentFinish();
                    return;
                  }
                }
                if (isCellMasterEditBlocked(tab, idx)) {
                  const nav = ev.key === 'Tab' || ev.key === 'Escape' ||
                    ev.key === 'ArrowUp' || ev.key === 'ArrowDown' ||
                    ev.key === 'ArrowLeft' || ev.key === 'ArrowRight' ||
                    ev.key === 'Home' || ev.key === 'End' || ev.key === 'PageUp' || ev.key === 'PageDown';
                  const modNav = (ev.ctrlKey || ev.metaKey) && (ev.key === 'c' || ev.key === 'C' ||
                    ev.key === 'a' || ev.key === 'A');
                  // Allow page-flip arrows / Alt+arrows; block editing keys.
                  if (ev.altKey && (ev.key === 'ArrowLeft' || ev.key === 'ArrowRight')) {
                    /* page flip allowed below */
                  } else if (!nav && !modNav) {
                    if (!(ev.key === 'ArrowLeft' || ev.key === 'ArrowRight')) {
                      ev.preventDefault();
                      setStatus('Locked Master cell — double-click to unlock before editing', 'err');
                      return;
                    }
                  }
                }
                // Enter moves to next parent cell (no auto-copy); Shift+Enter inserts newline.
                if (ev.key === 'Enter' && !ev.shiftKey) {
                  ev.preventDefault();
                  if (stickyCellRange) clearStickyCellRange();
                  const active = activeTab();
                  if (!active || idx < 0 || idx >= active.cells.length) return;
                  if (idx < active.cells.length - 1) {
                    moveToNextCell(ev.currentTarget, idx + 1);
                    return;
                  }
                  ev.currentTarget.blur();
                  addRow();
                  focusCell(idx + 1);
                  return;
                }
                if (ev.altKey && (ev.key === 'ArrowLeft' || ev.key === 'ArrowRight')) {
                  ev.preventDefault();
                  stepNestPage(idx, nestIndex, ev.key === 'ArrowRight' ? 1 : -1);
                  return;
                }
                // Plain arrows at caret edge flip pages when not editing mid-text selection.
                if (ev.key === 'ArrowLeft' || ev.key === 'ArrowRight') {
                  const taEl = ev.currentTarget;
                  if (typeof taEl.selectionStart === 'number' && taEl.selectionStart === taEl.selectionEnd) {
                    const atStart = taEl.selectionStart === 0;
                    const atEnd = taEl.selectionStart === (taEl.value || '').length;
                    if (ev.key === 'ArrowLeft' && atStart && page > 0) {
                      ev.preventDefault();
                      stepNestPage(idx, nestIndex, -1);
                      return;
                    }
                    if (ev.key === 'ArrowRight' && atEnd) {
                      ev.preventDefault();
                      stepNestPage(idx, nestIndex, 1);
                      return;
                    }
                  }
                }
              });
              nestTa.addEventListener('pointerdown', function (ev) {
                ev.stopPropagation();
              });
              nestTa.addEventListener('click', function (ev) {
                ev.stopPropagation();
                if (ev.detail !== 3) return;
                const nestEl = ev.currentTarget;
                window.setTimeout(function () {
                  copyTripleClickSelection(nestEl);
                }, 0);
              });

              nestRow.appendChild(nestChrome);
              nestRow.appendChild(nestTa);
              nestsEl.appendChild(nestRow);
            })(ni);
          }
          stack.appendChild(nestsEl);
        }

        wrap.appendChild(stack);
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
    applyPersistedRowHeights(tab);
    restoreStickyCellRangeHighlight();
    applyPartSearchHighlights();
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

  /**
   * Combined native scrollbar only when text overflows the box.
   * Measure with overflow hidden so a classic Windows gutter cannot create
   * false overflow (Ash: bar showed even when prompt was shorter than the box).
   * CSS uses overflow-y: auto (never scroll); this toggles hidden when content fits.
   */
  function syncCombinedOverflowY() {
    const node = el.combined;
    if (!node) return;
    node.style.overflowY = 'hidden';
    const overflows = node.scrollHeight > node.clientHeight + 1;
    node.style.overflowY = overflows ? 'auto' : 'hidden';
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
    syncCombinedOverflowY();
  }

  function beginSectionResize(event) {
    event.preventDefault();
    const startY = event.clientY;
    const startHeight = el.combinedSection.getBoundingClientRect().height;
    document.body.classList.add('resizing-sections');

    function onMove(moveEvent) {
      resizeCombinedSection(startHeight + (moveEvent.clientY - startY));
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

  // Click2Copy cells: short click/focus leaves the caret (no auto-copy). Triple-click
  // selects + copies the entire cell (including paragraph breaks) + status toast.
  // Enter moves to the next cell without copying (Shift+Enter = newline). Combined
  // click/focus copies Combined; Ctrl/Cmd+C copies cell or sticky multi-cell TSV. A
  // fresh drag selects a rectangle; only a later drag started inside that sticky
  // selection moves/swaps the block. Shift+Arrow extends the sticky multi-cell highlight.
  // Delete/Backspace clears every cell in the sticky multi-cell range (current page).
  let lastAutoCopyKey = '';
  let lastAutoCopyAt = 0;
  let suppressCellAutoCopy = false;
  let cellRangeDrag = null;
  /** Sticky multi-cell range after drag-copy (Excel-ish). Cleared on Esc / outside click / edit / nav. */
  let stickyCellRange = null;
  /** Keyboard range anchor/focus for Excel-like Shift+Arrow extension. */
  let keyboardRangeAnchor = null;
  let keyboardRangeFocus = null;
  const CELL_DRAG_MOVE_PX = 8;

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

  /**
   * Triple-click helper: select + copy the ENTIRE cell value (including newlines /
   * paragraph breaks). Native textarea triple-click only selects one paragraph, so we
   * expand selection to the full value before copying. Same status toast as other copies.
   * Does not re-introduce single/double-click or Enter auto-copy.
   */
  function copyTripleClickSelection(ta) {
    if (suppressCellAutoCopy) return;
    if (!ta || ta.tagName !== 'TEXTAREA') return;
    const value = ta.value == null ? '' : String(ta.value);
    if (!value) return;
    // Override native paragraph-bounded triple-click selection with the whole cell.
    selectWholeCellContents(ta);
    const tab = activeTab();
    const idx = parseInt(ta.dataset.idx, 10);
    const nest = ta.dataset.nest != null ? String(ta.dataset.nest) : '';
    const key = (tab ? tab.id : '') + ':' + (Number.isNaN(idx) ? '' : idx) +
      (nest ? ':n' + nest : '') + ':' + value;
    const now = Date.now();
    if (key === lastAutoCopyKey && now - lastAutoCopyAt < 300) return;
    lastAutoCopyKey = key;
    lastAutoCopyAt = now;
    copyTextWithStatus(value, 'Copied cell');
  }

  /** Snap selection to the whole cell (used after relocate; clicks leave the caret). */
  function selectWholeCellContents(ta) {
    if (!ta || ta.tagName !== 'TEXTAREA') return;
    try {
      const len = (ta.value == null ? '' : String(ta.value)).length;
      if (typeof ta.select === 'function') ta.select();
      if (typeof ta.setSelectionRange === 'function') ta.setSelectionRange(0, len);
    } catch (err) { /* no-op */ }
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

  function clearCellRelocateHighlight() {
    if (!el.cellGrid) return;
    const nodes = el.cellGrid.querySelectorAll(
      '.cell-wrap.cell-relocate-source, .cell-wrap.cell-relocate-target, ' +
      '.cell.cell-relocate-source, .cell.cell-relocate-target'
    );
    for (let i = 0; i < nodes.length; i++) {
      nodes[i].classList.remove('cell-relocate-source', 'cell-relocate-target');
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

  function isMultiCellBounds(b) {
    return !!(b && (b.rMax > b.rMin || b.cMax > b.cMin));
  }

  function getStickyCellRange() {
    if (!stickyCellRange || !isMultiCellBounds(stickyCellRange)) return null;
    const tab = activeTab();
    if (!tab || stickyCellRange.tabId !== tab.id) return null;
    return stickyCellRange;
  }

  function cellInStickyRange(row, col) {
    const b = getStickyCellRange();
    if (!b) return false;
    return row >= b.rMin && row <= b.rMax && col >= b.cMin && col <= b.cMax;
  }

  function setStickyCellRange(bounds, tabId) {
    const tab = activeTab();
    const id = tabId || (tab && tab.id) || null;
    if (!bounds || !isMultiCellBounds(bounds) || !id) {
      stickyCellRange = null;
      clearCellRangeHighlight();
      return;
    }
    stickyCellRange = {
      rMin: bounds.rMin,
      rMax: bounds.rMax,
      cMin: bounds.cMin,
      cMax: bounds.cMax,
      tabId: id
    };
    applyCellRangeHighlight(bounds.rMin, bounds.cMin, bounds.rMax, bounds.cMax);
  }

  function clearStickyCellRange() {
    stickyCellRange = null;
    clearCellRangeHighlight();
  }

  function restoreStickyCellRangeHighlight() {
    const b = getStickyCellRange();
    if (!b) return;
    applyCellRangeHighlight(b.rMin, b.cMin, b.rMax, b.cMax);
  }

  /**
   * Delete/Backspace on a sticky multi-cell selection: clear current-page text in
   * every cell in the rectangle (same scope as Cut). Other pages, nests, and shades
   * stay. Locked Master cells block the whole clear (same as multi-cell paste).
   * Selection highlight stays so the user can keep working the block.
   */
  function clearStickyMultiCellContents() {
    const tab = activeTab();
    const b = getStickyCellRange();
    if (!tab || !b) return false;

    const indices = [];
    for (let r = b.rMin; r <= b.rMax; r++) {
      for (let c = b.cMin; c <= b.cMax; c++) {
        indices.push(r * tab.cols + c);
      }
    }

    if (!isMasterTab(tab)) {
      for (let i = 0; i < indices.length; i++) {
        const idx = indices[i];
        if (isCellMasterLocked(tab, idx) &&
            !(isMasterCellEditSession() &&
              masterSegmentEdit.tabId === tab.id &&
              masterSegmentEdit.cellIndex === idx)) {
          setStatus('Locked Master cell in selection — double-click to unlock first', 'err');
          return false;
        }
      }
    }

    let anyContent = false;
    for (let i = 0; i < indices.length; i++) {
      if ((tab.cells[indices[i]] || '') !== '') {
        anyContent = true;
        break;
      }
    }
    if (!anyContent) return true;

    pushHistory();
    ensureCellPages(tab);
    for (let i = 0; i < indices.length; i++) {
      const idx = indices[i];
      writeCellCurrentPage(tab, idx, '');
      if (!isMasterTab(tab)) clearCellMasterLock(tab, idx);
      revalidateLinksForCell(tab.id, idx, { silent: true });
    }
    renderCombinedPrompt();
    applyConfirmedCellHighlights();
    applyAppendCheckedState();
    if (el.cellGrid) {
      for (let i = 0; i < indices.length; i++) {
        const idx = indices[i];
        const ta = el.cellGrid.querySelector('textarea.cell[data-idx="' + idx + '"]');
        if (!ta) continue;
        ta.value = '';
        const confirmed = isCellConfirmed(tab.id, idx);
        ta.classList.toggle('cell-confirmed', confirmed);
        const wrap = ta.closest ? ta.closest('.cell-wrap') : null;
        if (wrap) wrap.classList.toggle('cell-confirmed', confirmed);
      }
    }
    restoreStickyCellRangeHighlight();
    if (isMasterTab(tab)) renderMasterLibrary();
    scheduleSave();
    const rows = b.rMax - b.rMin + 1;
    const cols = b.cMax - b.cMin + 1;
    setStatus('Cleared ' + rows + '\u00d7' + cols + ' cells', 'ok');
    return true;
  }

  function clearKeyboardCellRange() {
    keyboardRangeAnchor = null;
    keyboardRangeFocus = null;
  }

  /** Extend the active cell with Shift+Arrow, like a spreadsheet range selection. */
  function extendKeyboardCellRange(ta, key) {
    const tab = activeTab();
    if (!tab || !ta) return false;
    const row = parseInt(ta.dataset.row, 10);
    const col = parseInt(ta.dataset.col, 10);
    if (Number.isNaN(row) || Number.isNaN(col)) return false;

    if (!keyboardRangeAnchor || keyboardRangeAnchor.tabId !== tab.id) {
      keyboardRangeAnchor = { tabId: tab.id, row: row, col: col };
      keyboardRangeFocus = { row: row, col: col };
    }
    const current = keyboardRangeFocus || { row: row, col: col };
    const currentIdx = current.row * tab.cols + current.col;
    const nextIdx = adjacentCellIndex(currentIdx, tab.cols, tab.rows, key);
    if (nextIdx < 0) return true;

    keyboardRangeFocus = {
      row: Math.floor(nextIdx / tab.cols),
      col: nextIdx % tab.cols
    };
    const b = normalizeRangeBounds(
      keyboardRangeAnchor.row, keyboardRangeAnchor.col,
      keyboardRangeFocus.row, keyboardRangeFocus.col
    );
    if (isMultiCellBounds(b)) setStickyCellRange(b, tab.id);
    else clearStickyCellRange();
    return true;
  }

  function applyCellRangeHighlight(r0, c0, r1, c1) {
    clearCellRangeHighlight();
    clearCellRelocateHighlight();
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

  /**
   * Move/swap a rectangular cell block so its top-left lands at destRMin/destCMin.
   * Non-overlapping + dest has content ⇒ rectangle swap; otherwise move (overwrite dest,
   * clear vacated source). Combined-link indices follow content.
   * Returns { swapped, bounds } or null on failure.
   */
  function relocateOrSwapCellBlock(tab, srcBounds, destRMin, destCMin) {
    if (!tab || !srcBounds || !isMultiCellBounds(srcBounds)) return null;
    const nRows = srcBounds.rMax - srcBounds.rMin + 1;
    const nCols = srcBounds.cMax - srcBounds.cMin + 1;
    if (destRMin < 0 || destCMin < 0) return null;
    if (destRMin + nRows - 1 >= tab.rows || destCMin + nCols - 1 >= tab.cols) return null;
    if (destRMin === srcBounds.rMin && destCMin === srcBounds.cMin) return null;

    const srcIndices = [];
    const dstIndices = [];
    for (let r = 0; r < nRows; r++) {
      for (let c = 0; c < nCols; c++) {
        srcIndices.push((srcBounds.rMin + r) * tab.cols + (srcBounds.cMin + c));
        dstIndices.push((destRMin + r) * tab.cols + (destCMin + c));
      }
    }

    const overlap = !(
      srcBounds.rMax < destRMin ||
      destRMin + nRows - 1 < srcBounds.rMin ||
      srcBounds.cMax < destCMin ||
      destCMin + nCols - 1 < srcBounds.cMin
    );

    ensureNestedCells(tab);
    ensureCellLocks(tab);
    ensureCellPages(tab);
    ensureCellShades(tab);
    const srcCells = srcIndices.map(function (i) { return tab.cells[i]; });
    const srcNested = srcIndices.map(function (i) { return normalizeNestList(tab.nestedCells[i]); });
    const srcLocks = srcIndices.map(function (i) { return tab.cellLocks[i]; });
    const srcPages = srcIndices.map(function (i) { return cloneCellPagesEntry(tab.cellPages[i]); });
    const srcShades = srcIndices.map(function (i) { return tab.cellShades[i] || null; });
    const dstCells = dstIndices.map(function (i) { return tab.cells[i]; });
    const dstNested = dstIndices.map(function (i) { return normalizeNestList(tab.nestedCells[i]); });
    const dstLocks = dstIndices.map(function (i) { return tab.cellLocks[i]; });
    const dstPages = dstIndices.map(function (i) { return cloneCellPagesEntry(tab.cellPages[i]); });
    const dstShades = dstIndices.map(function (i) { return tab.cellShades[i] || null; });

    let destHadContent = false;
    if (!overlap) {
      for (let i = 0; i < dstCells.length; i++) {
        if ((dstCells[i] || '').trim()) {
          destHadContent = true;
          break;
        }
      }
    }

    pushHistory();

    const indexMap = {};
    const destroyed = {};

    if (!overlap && destHadContent) {
      for (let i = 0; i < srcIndices.length; i++) {
        tab.cells[dstIndices[i]] = srcCells[i];
        tab.nestedCells[dstIndices[i]] = srcNested[i];
        tab.cellLocks[dstIndices[i]] = srcLocks[i];
        tab.cellPages[dstIndices[i]] = srcPages[i];
        tab.cellShades[dstIndices[i]] = srcShades[i];
        tab.cells[srcIndices[i]] = dstCells[i];
        tab.nestedCells[srcIndices[i]] = dstNested[i];
        tab.cellLocks[srcIndices[i]] = dstLocks[i];
        tab.cellPages[srcIndices[i]] = dstPages[i];
        tab.cellShades[srcIndices[i]] = dstShades[i];
        indexMap[srcIndices[i]] = dstIndices[i];
        indexMap[dstIndices[i]] = srcIndices[i];
      }
    } else {
      const srcSet = {};
      for (let i = 0; i < srcIndices.length; i++) srcSet[srcIndices[i]] = true;
      for (let i = 0; i < dstIndices.length; i++) {
        if (!srcSet[dstIndices[i]]) destroyed[dstIndices[i]] = true;
      }
      for (let i = 0; i < srcIndices.length; i++) {
        tab.cells[srcIndices[i]] = '';
        tab.nestedCells[srcIndices[i]] = [];
        tab.cellLocks[srcIndices[i]] = null;
        tab.cellPages[srcIndices[i]] = makeEmptyCellPages('');
        tab.cellShades[srcIndices[i]] = null;
      }
      for (let i = 0; i < srcIndices.length; i++) {
        tab.cells[dstIndices[i]] = srcCells[i];
        tab.nestedCells[dstIndices[i]] = srcNested[i];
        tab.cellLocks[dstIndices[i]] = srcLocks[i];
        tab.cellPages[dstIndices[i]] = srcPages[i];
        tab.cellShades[dstIndices[i]] = srcShades[i];
        indexMap[srcIndices[i]] = dstIndices[i];
      }
    }

    state.confirmedLinks = state.confirmedLinks.filter(function (link) {
      if (link.tabId !== tab.id) return true;
      return !destroyed[link.cellIndex];
    });
    // Remap in two passes so swap cycles do not collide.
    const pending = [];
    state.confirmedLinks.forEach(function (link) {
      if (link.tabId !== tab.id) return;
      if (Object.prototype.hasOwnProperty.call(indexMap, link.cellIndex)) {
        pending.push({ link: link, to: indexMap[link.cellIndex] });
      }
    });
    pending.forEach(function (item) {
      item.link.cellIndex = item.to;
    });

    return {
      swapped: !overlap && destHadContent,
      bounds: {
        rMin: destRMin,
        rMax: destRMin + nRows - 1,
        cMin: destCMin,
        cMax: destCMin + nCols - 1
      }
    };
  }

  function clampBlockDest(tab, srcBounds, destRMin, destCMin) {
    const nRows = srcBounds.rMax - srcBounds.rMin + 1;
    const nCols = srcBounds.cMax - srcBounds.cMin + 1;
    const maxR = Math.max(0, tab.rows - nRows);
    const maxC = Math.max(0, tab.cols - nCols);
    return {
      destRMin: Math.max(0, Math.min(destRMin, maxR)),
      destCMin: Math.max(0, Math.min(destCMin, maxC))
    };
  }

  function applyBlockRelocateHighlight(srcBounds, destBounds) {
    clearCellRangeHighlight();
    clearCellRelocateHighlight();
    if (!el.cellGrid || !srcBounds) return;
    function markRect(b, cls) {
      for (let r = b.rMin; r <= b.rMax; r++) {
        for (let c = b.cMin; c <= b.cMax; c++) {
          const wrap = el.cellGrid.querySelector(
            '.cell-wrap[data-row="' + r + '"][data-col="' + c + '"]'
          );
          if (!wrap) continue;
          wrap.classList.add(cls);
          const ta = wrap.querySelector('textarea.cell');
          if (ta) ta.classList.add(cls);
        }
      }
    }
    markRect(srcBounds, 'cell-relocate-source');
    if (destBounds &&
        (destBounds.rMin !== srcBounds.rMin || destBounds.cMin !== srcBounds.cMin ||
         destBounds.rMax !== srcBounds.rMax || destBounds.cMax !== srcBounds.cMax)) {
      markRect(destBounds, 'cell-relocate-target');
    }
  }

  function finishCellRangeListeners() {
    document.removeEventListener('pointermove', onCellRangePointerMove);
    document.removeEventListener('pointerup', onCellRangePointerUp);
    document.removeEventListener('pointercancel', onCellRangePointerUp);
    document.removeEventListener('dragstart', onCellRangeDragStartPrevent, true);
    document.body.classList.remove('selecting-cell-range');
    document.body.classList.remove('relocating-cell');
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
    if (cellRangeDrag && cellRangeDrag.active && cellRangeDrag.mode) e.preventDefault();
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
    if (!cellRangeDrag || cellRangeDrag.mode === 'multi') return;
    if (cellRangeDrag.mode === 'block-relocate') return;
    cellRangeDrag.mode = 'multi';
    cellRangeDrag.multi = true;
    suppressCellAutoCopy = true;
    document.body.classList.remove('relocating-cell');
    document.body.classList.add('selecting-cell-range');
    clearStickyCellRange();
    clearCellRelocateHighlight();
    collapseCellTextSelection();
  }

  function beginBlockRelocateDrag() {
    if (!cellRangeDrag || cellRangeDrag.mode === 'block-relocate') return;
    if (cellRangeDrag.mode === 'multi') return;
    if (!cellRangeDrag.blockBounds || !isMultiCellBounds(cellRangeDrag.blockBounds)) return;
    cellRangeDrag.mode = 'block-relocate';
    cellRangeDrag.multi = false;
    suppressCellAutoCopy = true;
    document.body.classList.remove('selecting-cell-range');
    document.body.classList.add('relocating-cell');
    collapseCellTextSelection();
    const tab = activeTab();
    const b = cellRangeDrag.blockBounds;
    const deltaR = cellRangeDrag.endRow - cellRangeDrag.grabRow;
    const deltaC = cellRangeDrag.endCol - cellRangeDrag.grabCol;
    let destRMin = b.rMin + deltaR;
    let destCMin = b.cMin + deltaC;
    if (tab) {
      const clamped = clampBlockDest(tab, b, destRMin, destCMin);
      destRMin = clamped.destRMin;
      destCMin = clamped.destCMin;
    }
    const destBounds = {
      rMin: destRMin,
      rMax: destRMin + (b.rMax - b.rMin),
      cMin: destCMin,
      cMax: destCMin + (b.cMax - b.cMin)
    };
    cellRangeDrag.destRMin = destRMin;
    cellRangeDrag.destCMin = destCMin;
    applyBlockRelocateHighlight(b, destBounds);
  }

  function onCellRangePointerMove(e) {
    if (!cellRangeDrag || !cellRangeDrag.active) return;
    const dx = e.clientX - cellRangeDrag.startX;
    const dy = e.clientY - cellRangeDrag.startY;
    const distSq = dx * dx + dy * dy;
    const moved = distSq >= (CELL_DRAG_MOVE_PX * CELL_DRAG_MOVE_PX);

    const wrap = cellWrapFromPoint(e.clientX, e.clientY);
    let row = cellRangeDrag.endRow;
    let col = cellRangeDrag.endCol;
    if (wrap) {
      const r = parseInt(wrap.dataset.row, 10);
      const c = parseInt(wrap.dataset.col, 10);
      if (!Number.isNaN(r) && !Number.isNaN(c)) {
        row = r;
        col = c;
      }
    }

    // Stay in native text-select (I-beam) while the pointer remains in the
    // starting cell. Multi-cell range / sticky block-relocate only begin once
    // the drag actually crosses into another cell — that is what was hijacking
    // in-cell character selection and flipping the cursor via body classes.
    const leftStartCell =
      row !== cellRangeDrag.startRow || col !== cellRangeDrag.startCol;
    if (!cellRangeDrag.mode && moved && leftStartCell) {
      if (cellRangeDrag.fromStickyBlock) beginBlockRelocateDrag();
      else beginMultiCellRangeDrag();
    }

    if (row === cellRangeDrag.endRow && col === cellRangeDrag.endCol) {
      return;
    }
    cellRangeDrag.endRow = row;
    cellRangeDrag.endCol = col;

    if (cellRangeDrag.mode === 'multi') {
      if (e.cancelable) e.preventDefault();
      applyCellRangeHighlight(
        cellRangeDrag.startRow, cellRangeDrag.startCol,
        cellRangeDrag.endRow, cellRangeDrag.endCol
      );
      return;
    }

    if (cellRangeDrag.mode === 'block-relocate') {
      if (e.cancelable) e.preventDefault();
      const tab = activeTab();
      const b = cellRangeDrag.blockBounds;
      if (!tab || !b) return;
      const deltaR = cellRangeDrag.endRow - cellRangeDrag.grabRow;
      const deltaC = cellRangeDrag.endCol - cellRangeDrag.grabCol;
      const clamped = clampBlockDest(tab, b, b.rMin + deltaR, b.cMin + deltaC);
      cellRangeDrag.destRMin = clamped.destRMin;
      cellRangeDrag.destCMin = clamped.destCMin;
      applyBlockRelocateHighlight(b, {
        rMin: clamped.destRMin,
        rMax: clamped.destRMin + (b.rMax - b.rMin),
        cMin: clamped.destCMin,
        cMax: clamped.destCMin + (b.cMax - b.cMin)
      });
    }
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

    if (drag.mode === 'multi') {
      const b = normalizeRangeBounds(drag.startRow, drag.startCol, drag.endRow, drag.endCol);
      const rows = b.rMax - b.rMin + 1;
      const cols = b.cMax - b.cMin + 1;
      const isMulti = rows > 1 || cols > 1;
      if (isMulti) {
        suppressCellAutoCopy = true;
        const tab = activeTab();
        // Sticky multi-cell selection only — copy via Ctrl/Cmd+C (no auto-copy on release).
        setStickyCellRange(b, tab && tab.id);
        window.setTimeout(function () {
          suppressCellAutoCopy = false;
          cellRangeDrag = null;
          restoreStickyCellRangeHighlight();
        }, 120);
        return;
      }
      clearCellRangeHighlight();
      suppressCellAutoCopy = false;
      cellRangeDrag = null;
      // Jittered drag that never left the cell — leave caret (no auto-copy).
      return;
    }

    if (drag.mode === 'block-relocate') {
      suppressCellAutoCopy = true;
      clearCellRelocateHighlight();
      const tab = activeTab();
      const b = drag.blockBounds;
      if (tab && b && isMultiCellBounds(b)) {
        const deltaR = drag.endRow - drag.grabRow;
        const deltaC = drag.endCol - drag.grabCol;
        const clamped = clampBlockDest(tab, b, b.rMin + deltaR, b.cMin + deltaC);
        const same = clamped.destRMin === b.rMin && clamped.destCMin === b.cMin;
        if (!same) {
          const result = relocateOrSwapCellBlock(tab, b, clamped.destRMin, clamped.destCMin);
          if (result) {
            stickyCellRange = null;
            renderTabs();
            renderGrid();
            renderMasterLibrary();
            renderCombinedPrompt();
            scheduleSave();
            setStickyCellRange(result.bounds, tab.id);
            const rows = result.bounds.rMax - result.bounds.rMin + 1;
            const cols = result.bounds.cMax - result.bounds.cMin + 1;
            const sizeLabel = rows + '\u00d7' + cols;
            const toLabel = cellAddress(result.bounds.rMin, result.bounds.cMin);
            setStatus(
              result.swapped
                ? ('Swapped ' + sizeLabel + ' block \u2194 ' + toLabel)
                : ('Moved ' + sizeLabel + ' block \u2192 ' + toLabel),
              'ok'
            );
            window.setTimeout(function () {
              suppressCellAutoCopy = false;
              cellRangeDrag = null;
              restoreStickyCellRangeHighlight();
              const focusIdx = result.bounds.rMin * tab.cols + result.bounds.cMin;
              focusCell(focusIdx);
              const ta = el.cellGrid.querySelector('textarea.cell[data-idx="' + focusIdx + '"]');
              if (ta) selectWholeCellContents(ta);
            }, 0);
            return;
          }
        }
      }
      // Cancelled / same place: keep sticky selection highlight.
      restoreStickyCellRangeHighlight();
      suppressCellAutoCopy = false;
      cellRangeDrag = null;
      return;
    }

    // Short click / hold-without-drag: place the caret and clear any sticky
    // multi-cell selection, including when the clicked cell is inside it.
    // A real drag from inside the range was handled above as block relocation.
    clearStickyCellRange();
    clearCellRelocateHighlight();
    suppressCellAutoCopy = false;
    cellRangeDrag = null;
  }

  function onCellPointerDownSelect(e) {
    if (e.button != null && e.button !== 0) return;
    const ta = e.currentTarget;
    clearKeyboardCellRange();
    const row = parseInt(ta.dataset.row, 10);
    const col = parseInt(ta.dataset.col, 10);
    if (Number.isNaN(row) || Number.isNaN(col)) return;
    // Re-click / fresh drag; do not disturb Combined toggles (those are outside the textarea).
    if (cellRangeDrag && cellRangeDrag.active) {
      finishCellRangeListeners();
    }

    const insideSticky = cellInStickyRange(row, col);
    if (!insideSticky && stickyCellRange) {
      // Click outside sticky multi-cell selection clears it (Excel-ish).
      clearStickyCellRange();
    }

    const sticky = insideSticky ? getStickyCellRange() : null;
    cellRangeDrag = {
      active: true,
      mode: null,
      multi: false,
      startRow: row,
      startCol: col,
      endRow: row,
      endCol: col,
      startX: e.clientX,
      startY: e.clientY,
      pointerId: e.pointerId,
      captureEl: ta,
      fromStickyBlock: !!sticky,
      blockBounds: sticky ? {
        rMin: sticky.rMin,
        rMax: sticky.rMax,
        cMin: sticky.cMin,
        cMax: sticky.cMax,
        tabId: sticky.tabId
      } : null,
      grabRow: row,
      grabCol: col,
      destRMin: sticky ? sticky.rMin : row,
      destCMin: sticky ? sticky.cMin : col
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
    // Leave caret where the user clicked / Tab landed; no activate auto-copy.
    rememberFocusedCell(e);
  }

  function onCellClickSelect(e) {
    // Leave caret at click position (or word select on double-click). Triple-click
    // selects + copies the entire cell (incl. paragraph breaks) + status toast.
    rememberFocusedCell(e);
    if (e.detail !== 3) return;
    const ta = e.currentTarget;
    // Defer past the browser's native paragraph selection, then expand to whole cell.
    window.setTimeout(function () {
      copyTripleClickSelection(ta);
    }, 0);
  }

  function onCellDblClickMasterUnlock(e) {
    const tab = activeTab();
    if (!tab || isMasterTab(tab) || masterSegmentDialogOpen) return;
    const ta = e.currentTarget;
    const idx = parseInt(ta.dataset.idx, 10);
    if (Number.isNaN(idx) || idx < 0 || idx >= tab.cells.length) return;
    if (!isCellMasterLocked(tab, idx)) return;
    e.preventDefault();
    e.stopPropagation();
    beginMasterCellEdit(ta, tab, idx);
  }

  function onNestDblClickMasterUnlock(e) {
    const tab = activeTab();
    if (!tab || isMasterTab(tab) || masterSegmentDialogOpen) return;
    const nestTa = e.currentTarget;
    const idx = parseInt(nestTa.dataset.idx, 10);
    if (Number.isNaN(idx) || idx < 0 || idx >= tab.cells.length) return;
    if (!isCellMasterLocked(tab, idx)) return;
    e.preventDefault();
    e.stopPropagation();
    const parentTa = el.cellGrid
      ? el.cellGrid.querySelector('textarea.cell[data-idx="' + idx + '"]')
      : null;
    beginMasterCellEdit(parentTa || nestTa, tab, idx, nestTa);
  }

  function onNestBlurMasterFinish(e) {
    if (!isMasterCellEditSession() || masterSegmentDialogOpen) return;
    const tab = activeTab();
    const nestTa = e.currentTarget;
    if (!tab || !nestTa) return;
    const idx = parseInt(nestTa.dataset.idx, 10);
    if (masterSegmentEdit.tabId !== tab.id || masterSegmentEdit.cellIndex !== idx) return;
    const next = e.relatedTarget;
    if (next && el.masterSegmentDialog && el.masterSegmentDialog.contains(next)) return;
    // Stay in-session when focus moves to parent/sibling nest of the same cell.
    if (next && next.closest) {
      const wrap = nestTa.closest('.cell-wrap');
      if (wrap && wrap.contains(next)) return;
    }
    window.setTimeout(function () {
      if (!isMasterCellEditSession() || masterSegmentDialogOpen) return;
      const active = document.activeElement;
      if (el.masterSegmentDialog && el.masterSegmentDialog.contains(active)) return;
      if (active === nestTa) return;
      const wrap = nestTa.closest ? nestTa.closest('.cell-wrap') : null;
      if (wrap && wrap.contains(active)) return;
      requestMasterSegmentFinish();
    }, 0);
  }

  function onNestBeforeInputMasterLock(e) {
    if (masterSegmentDialogOpen) {
      e.preventDefault();
      return;
    }
    const tab = activeTab();
    if (!tab || isMasterTab(tab)) return;
    const nestTa = e.currentTarget;
    const idx = parseInt(nestTa.dataset.idx, 10);
    if (Number.isNaN(idx) || idx < 0 || idx >= tab.cells.length) return;
    if (isMasterCellEditFor(tab, idx)) return;
    if (isCellMasterLocked(tab, idx) || nestTa.readOnly) {
      e.preventDefault();
      setStatus('Locked Master cell — double-click to unlock before editing', 'err');
    }
  }

  function onCellBlurMasterFinish(e) {
    if (!isMasterCellEditSession() || masterSegmentDialogOpen) return;
    const tab = activeTab();
    const ta = e.currentTarget;
    if (!tab || !ta) return;
    const idx = parseInt(ta.dataset.idx, 10);
    if (masterSegmentEdit.tabId !== tab.id || masterSegmentEdit.cellIndex !== idx) return;
    const next = e.relatedTarget;
    if (next && el.masterSegmentDialog && el.masterSegmentDialog.contains(next)) return;
    if (next && next.closest) {
      const wrap = ta.closest('.cell-wrap');
      if (wrap && wrap.contains(next)) return;
    }
    window.setTimeout(function () {
      if (!isMasterCellEditSession() || masterSegmentDialogOpen) return;
      const active = document.activeElement;
      if (el.masterSegmentDialog && el.masterSegmentDialog.contains(active)) return;
      if (active === ta) return;
      const wrap = ta.closest ? ta.closest('.cell-wrap') : null;
      if (wrap && wrap.contains(active)) return;
      requestMasterSegmentFinish();
    }, 0);
  }

  function onCellBeforeInputMasterLock(e) {
    if (masterSegmentDialogOpen) {
      e.preventDefault();
      return;
    }
    const tab = activeTab();
    if (!tab || isMasterTab(tab)) return;
    const ta = e.currentTarget;
    const idx = parseInt(ta.dataset.idx, 10);
    if (Number.isNaN(idx) || idx < 0 || idx >= tab.cells.length) return;
    // Allow typing only while this cell is the active unlocked Master edit session.
    if (isMasterCellEditSession() &&
        masterSegmentEdit.tabId === tab.id &&
        masterSegmentEdit.cellIndex === idx) {
      return;
    }
    if (isCellMasterLocked(tab, idx) || ta.readOnly) {
      e.preventDefault();
      setStatus('Locked Master cell — double-click to unlock before editing', 'err');
    }
  }

  function onCellInput(e) {
    const tab = activeTab();
    if (!tab) return;
    rememberFocusedCell(e);
    const idx = parseInt(e.target.dataset.idx, 10);
    if (Number.isNaN(idx) || idx < 0 || idx >= tab.cells.length) return;
    // Hard block: locked Master cells are read-only until double-click unlock.
    if (!isMasterTab(tab) && isCellMasterLocked(tab, idx) &&
        !(isMasterCellEditSession() &&
          masterSegmentEdit.tabId === tab.id &&
          masterSegmentEdit.cellIndex === idx)) {
      e.target.value = tab.cells[idx] || '';
      setStatus('Locked Master cell — double-click to unlock before editing', 'err');
      return;
    }
    if (stickyCellRange) clearStickyCellRange();
    clearKeyboardCellRange();
    pushHistory({ coalesce: true });
    writeCellCurrentPage(tab, idx, e.target.value);
    liveSyncConfirmedLinksForCell(tab.id, idx, null);
    const confirmed = isCellConfirmed(tab.id, idx);
    e.target.classList.toggle('cell-confirmed', confirmed);
    const wrap = e.target.closest ? e.target.closest('.cell-wrap') : null;
    if (wrap) wrap.classList.toggle('cell-confirmed', confirmed);
    applyAppendCheckedState();
    // Master parent-page edits push pages (all pages + current text) to locked parts.
    if (isMasterTab(tab)) {
      syncLockedPartPagesFromMaster(idx);
      renderMasterLibrary();
    }
    const row = Math.floor(idx / tab.cols);
    maybeLiveRefitRowFromInput(row);
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

  function isWholeCellSelected(ta) {
    if (!ta || ta.tagName !== 'TEXTAREA') return false;
    const len = (ta.value == null ? '' : String(ta.value)).length;
    return ta.selectionStart === 0 && ta.selectionEnd === len;
  }

  /** True when caret is collapsed at the text edge matching the arrow direction. */
  function caretAtArrowBoundary(ta, key) {
    if (!ta || ta.tagName !== 'TEXTAREA') return false;
    const val = ta.value == null ? '' : String(ta.value);
    const start = ta.selectionStart;
    const end = ta.selectionEnd;
    if (start !== end) return false;
    if (key === 'ArrowLeft') return start === 0;
    if (key === 'ArrowRight') return start === val.length;
    // Up/Down: first/last visual line (no newline before/after caret).
    if (key === 'ArrowUp') return val.slice(0, start).indexOf('\n') === -1;
    if (key === 'ArrowDown') return val.slice(end).indexOf('\n') === -1;
    return false;
  }

  function adjacentCellIndex(idx, cols, rows, key) {
    const row = Math.floor(idx / cols);
    const col = idx % cols;
    if (key === 'ArrowLeft') {
      if (col <= 0) return -1;
      return idx - 1;
    }
    if (key === 'ArrowRight') {
      if (col >= cols - 1) return -1;
      return idx + 1;
    }
    if (key === 'ArrowUp') {
      if (row <= 0) return -1;
      return idx - cols;
    }
    if (key === 'ArrowDown') {
      if (row >= rows - 1) return -1;
      return idx + cols;
    }
    return -1;
  }

  function onCellKeydown(e) {
    if (e.isComposing) return;
    const tab = activeTab();
    const idx = parseInt(e.currentTarget.dataset.idx, 10);
    if (!tab || Number.isNaN(idx) || idx < 0 || idx >= tab.cells.length) return;

    if (isMasterCellEditSession() &&
        masterSegmentEdit.tabId === tab.id &&
        masterSegmentEdit.cellIndex === idx &&
        !masterSegmentDialogOpen) {
      if (e.key === 'Escape') {
        e.preventDefault();
        cancelMasterSegmentEdit();
        return;
      }
      if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        requestMasterSegmentFinish();
        return;
      }
    }

    // Sticky multi-cell Delete/Backspace clears the whole rectangle. Must run
    // before the locked-cell guard (focus may be on a locked cell in the range)
    // and before the Enter/Tab/Arrow-only early return (that return was why
    // Delete previously did nothing for a multi-cell selection).
    const isClearKey = (e.key === 'Delete' || e.key === 'Backspace') &&
      !e.ctrlKey && !e.metaKey && !e.altKey;
    if (isClearKey && getStickyCellRange()) {
      e.preventDefault();
      clearStickyMultiCellContents();
      return;
    }

    if (!isMasterTab(tab) && isCellMasterLocked(tab, idx) &&
        !(isMasterCellEditSession() &&
          masterSegmentEdit.tabId === tab.id &&
          masterSegmentEdit.cellIndex === idx)) {
      // Allow navigation; block editing keys.
      const nav = e.key === 'Tab' || e.key === 'Escape' ||
        e.key === 'ArrowUp' || e.key === 'ArrowDown' ||
        e.key === 'ArrowLeft' || e.key === 'ArrowRight' ||
        e.key === 'Home' || e.key === 'End' || e.key === 'PageUp' || e.key === 'PageDown';
      const modNav = (e.ctrlKey || e.metaKey) && (e.key === 'c' || e.key === 'C' ||
        e.key === 'a' || e.key === 'A');
      if (!nav && !modNav) {
        e.preventDefault();
        if (e.key === 'Enter' && !e.shiftKey) {
          // Still allow Enter navigation on locked cells (no edit / no auto-copy).
        } else {
          setStatus('Locked Master cell — double-click to unlock before editing', 'err');
          return;
        }
      }
    }

    const isEnter = e.key === 'Enter' && !e.shiftKey;
    const isTab = e.key === 'Tab';
    const isArrow = e.key === 'ArrowUp' || e.key === 'ArrowDown' ||
      e.key === 'ArrowLeft' || e.key === 'ArrowRight';
    if (!isEnter && !isTab && !isArrow) return;

    if (isArrow) {
      if (e.shiftKey && !e.altKey && !e.ctrlKey && !e.metaKey) {
        // Shift+Arrow selects cells, not text, when the caret is at a grid boundary.
        if (!isWholeCellSelected(e.currentTarget) && !caretAtArrowBoundary(e.currentTarget, e.key)) return;
        e.preventDefault();
        extendKeyboardCellRange(e.currentTarget, e.key);
        return;
      }
      // Alt+Left/Right flips parent-cell pages (same idea as nest pages).
      if (e.altKey && (e.key === 'ArrowLeft' || e.key === 'ArrowRight') &&
          !e.ctrlKey && !e.metaKey) {
        e.preventDefault();
        stepCellPage(idx, e.key === 'ArrowRight' ? 1 : -1);
        return;
      }
      // Leave Ctrl/Meta+arrow for text selection / OS shortcuts.
      if (e.altKey || e.ctrlKey || e.metaKey) return;
      const ta = e.currentTarget;
      // Whole-cell select (e.g. after relocate), or caret at edge → move cells.
      // Mid-text caret (normal click) keeps normal caret movement.
      if (!isWholeCellSelected(ta) && !caretAtArrowBoundary(ta, e.key)) return;
      const next = adjacentCellIndex(idx, tab.cols, tab.rows, e.key);
      if (next < 0) return;
      e.preventDefault();
      if (stickyCellRange) clearStickyCellRange();
      clearKeyboardCellRange();
      moveToNextCell(ta, next);
      return;
    }

    // Keep Tab/Shift+Tab inside the cell grid (skip Combined controls).
    e.preventDefault();
    if (stickyCellRange) clearStickyCellRange();
    clearKeyboardCellRange();

    // Enter navigates to the next cell (no auto-copy). Shift+Enter is newline.
    // Tab navigates without copying. Use Ctrl/Cmd+C or Combined click to copy.
    if (isTab && e.shiftKey) {
      if (idx > 0) moveToNextCell(e.currentTarget, idx - 1);
      return;
    }

    // Enter or Tab: next cell (row-major).
    if (idx < tab.cells.length - 1) {
      moveToNextCell(e.currentTarget, idx + 1);
      return;
    }

    if (!isEnter) return;

    // Blur before re-rendering so the current edit is finished cleanly.
    e.currentTarget.blur();
    // Keep Enter row-major: append a row after the final cell, then focus it.
    addRow();
    focusCell(idx + 1);
  }

  function selectTab(id) {
    const tab = state.tabs.find(function (t) {
      return t.id === id;
    });
    if (!tab) return;
    const sameTab = id === state.activeTabId;
    if (sameTab && !toolsTabActive) return;
    toolsTabActive = false;
    state.activeTabId = id;
    if (!isMasterTab(tab)) lastPartTabId = id;
    focusedCell = null;
    clearStickyCellRange();
    clearKeyboardCellRange();
    clearPartSearch({ keepInputs: true, keepQuery: true });
    if (el.partFindInput) partSearchQuery = el.partFindInput.value;
    closeMasterLibFilterMenu();
    masterLibFilterMenuOpen = false;
    renderTabs();
    renderGrid();
    // Repair/drop Combined links for this tab before Master-insert greens paint.
    renderCombinedPrompt();
    renderMasterLibrary();
    scheduleSave();
  }

  function startInlineRename(btn, tab) {
    if (btn.classList.contains('editing')) return;
    closeTabIconPicker();
    btn.classList.add('editing');
    const parts = fillTabButtonContent(btn, tab);
    if (parts.title) parts.title.remove();

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
        } else if (next.toLowerCase() === 'tools') {
          setStatus('Tools is reserved for the shared tools tab', 'err');
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

  let tabIconPickerEl = null;
  let tabIconPickerTabId = null;

  function closeTabIconPicker() {
    if (tabIconPickerEl && tabIconPickerEl.parentNode) {
      tabIconPickerEl.parentNode.removeChild(tabIconPickerEl);
    }
    tabIconPickerEl = null;
    tabIconPickerTabId = null;
  }

  function setTabIcon(tab, iconId) {
    if (!tab || !isTabIconId(iconId)) return;
    if (resolveTabIcon(tab) === iconId) {
      closeTabIconPicker();
      return;
    }
    pushHistory();
    tab.icon = iconId;
    closeTabIconPicker();
    renderTabs();
    scheduleSave();
    setStatus('Tab icon → ' + (TAB_ICON_LABELS[iconId] || iconId), 'ok');
  }

  function openTabIconPicker(anchorBtn, tab, opts) {
    opts = opts || {};
    if (!tab || !anchorBtn) return;
    closeTabIconPicker();
    tabIconPickerTabId = tab.id;

    const pop = document.createElement('div');
    pop.className = 'tab-icon-picker';
    pop.setAttribute('role', 'dialog');
    pop.setAttribute('aria-label', 'Choose tab icon');

    const heading = document.createElement('div');
    heading.className = 'tab-icon-picker-heading';
    heading.textContent = 'Icon';
    pop.appendChild(heading);

    const grid = document.createElement('div');
    grid.className = 'tab-icon-picker-grid';
    const current = resolveTabIcon(tab);
    TAB_ICON_IDS.forEach(function (id) {
      const choice = document.createElement('button');
      choice.type = 'button';
      choice.className = 'tab-icon-picker-choice' + (id === current ? ' is-selected' : '');
      choice.title = TAB_ICON_LABELS[id] || id;
      choice.setAttribute('aria-label', TAB_ICON_LABELS[id] || id);
      choice.dataset.icon = id;
      choice.innerHTML = tabIconMarkup(id);
      choice.addEventListener('click', function (e) {
        e.preventDefault();
        e.stopPropagation();
        setTabIcon(tab, id);
      });
      grid.appendChild(choice);
    });
    pop.appendChild(grid);

    if (opts.showActions) {
      const actions = document.createElement('div');
      actions.className = 'tab-icon-picker-actions';
      const renameBtn = document.createElement('button');
      renameBtn.type = 'button';
      renameBtn.className = 'tab-icon-picker-action';
      renameBtn.textContent = 'Rename';
      renameBtn.addEventListener('click', function (e) {
        e.preventDefault();
        e.stopPropagation();
        closeTabIconPicker();
        startInlineRename(anchorBtn, tab);
      });
      const deleteBtn = document.createElement('button');
      deleteBtn.type = 'button';
      deleteBtn.className = 'tab-icon-picker-action is-danger';
      deleteBtn.textContent = 'Delete';
      deleteBtn.addEventListener('click', function (e) {
        e.preventDefault();
        e.stopPropagation();
        closeTabIconPicker();
        deleteTab(tab.id);
      });
      actions.appendChild(renameBtn);
      actions.appendChild(deleteBtn);
      pop.appendChild(actions);
    }

    document.body.appendChild(pop);
    tabIconPickerEl = pop;

    const pad = 6;
    let left;
    let top;
    if (Number.isFinite(opts.x) && Number.isFinite(opts.y)) {
      left = opts.x;
      top = opts.y;
    } else {
      const rect = anchorBtn.getBoundingClientRect();
      left = rect.left;
      top = rect.bottom + 4;
    }
    pop.style.left = '0px';
    pop.style.top = '0px';
    const size = pop.getBoundingClientRect();
    left = Math.max(pad, Math.min(left, window.innerWidth - size.width - pad));
    top = Math.max(pad, Math.min(top, window.innerHeight - size.height - pad));
    pop.style.left = left + 'px';
    pop.style.top = top + 'px';
  }

  function addTab() {
    pushHistory();
    const tab = makeTab(uid(), nextPartTitle());
    state.tabs.push(tab);
    toolsTabActive = false;
    state.activeTabId = tab.id;
    lastPartTabId = tab.id;
    renderTabs();
    renderGrid();
    renderMasterLibrary();
    scheduleSave();
    setStatus('Tab added');
  }

  function renameActiveTab() {
    const tab = toolsTargetTab();
    if (!tab || isMasterTab(tab)) return;
    const btn = el.tabBar.querySelector('.tab[data-id="' + tab.id + '"]');
    if (btn && !btn.classList.contains('tools-tab')) {
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
    if (trimmed.toLowerCase() === 'tools') {
      setStatus('Tools is reserved for the shared tools tab', 'err');
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
    clearMasterLibPrefsForPart(id);
    if (state.activeTabId === id) {
      const next = state.tabs[Math.min(idx, state.tabs.length - 1)];
      state.activeTabId = next.id;
    }
    if (lastPartTabId === id) {
      const active = state.tabs.find(function (t) { return t.id === state.activeTabId; });
      lastPartTabId = active && !isMasterTab(active)
        ? active.id
        : (partTabs()[0] ? partTabs()[0].id : null);
      if (toolsTabActive && !lastPartTabId) toolsTabActive = false;
      else if (toolsTabActive && lastPartTabId) state.activeTabId = lastPartTabId;
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
    ensureNestedCells(tab);
    ensureCellLocks(tab);
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
      remapNestedAfterColumnAdd(tab, oldCols, tab.cols);
      remapCellLocksAfterColumnAdd(tab, oldCols, tab.cols);
      remapCellPagesAfterColumnAdd(tab, oldCols, tab.cols);
      remapCellShadesAfterColumnAdd(tab, oldCols, tab.cols);
      remapConfirmedAfterColumnAdd(tab.id, oldCols, tab.cols);
    }
    while (tab.rows < rows) {
      ensureNestedCells(tab);
      ensureCellLocks(tab);
      ensureCellPages(tab);
      ensureCellShades(tab);
      for (let c = 0; c < tab.cols; c++) {
        tab.cells.push('');
        tab.nestedCells.push([]);
        tab.cellLocks.push(null);
        tab.cellPages.push(makeEmptyCellPages(''));
        tab.cellShades.push(null);
      }
      tab.rows += 1;
      ensureRowHeightsLength(tab);
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
        writeCellCurrentPage(tab, idx, matrix[r][c] == null ? '' : String(matrix[r][c]));
        // Sheet paste replaces cell text — drop Master lock on overwritten cells.
        if (!isMasterTab(tab)) clearCellMasterLock(tab, idx);
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

    if (!isMasterTab(tab) && isCellMasterLocked(tab, startIdx) &&
        !(isMasterCellEditSession() &&
          masterSegmentEdit.tabId === tab.id &&
          masterSegmentEdit.cellIndex === startIdx)) {
      e.preventDefault();
      setStatus('Locked Master cell — double-click to unlock before editing', 'err');
      return;
    }

    const matrix = parseClipboardMatrix(e.clipboardData);
    if (!matrix || !isMultiCellMatrix(matrix)) {
      // Single-cell / plain text: let the textarea handle a normal paste.
      return;
    }

    e.preventDefault();
    const startRow = Math.floor(startIdx / tab.cols);
    const startCol = startIdx % tab.cols;
    if (!isMasterTab(tab)) {
      const pr = matrix.length;
      const pc = matrix[0].length;
      for (let r = 0; r < pr; r++) {
        for (let c = 0; c < pc; c++) {
          const rr = startRow + r;
          const cc = startCol + c;
          if (rr < tab.rows && cc < tab.cols) {
            const i = rr * tab.cols + cc;
            if (isCellMasterLocked(tab, i)) {
              setStatus('Locked Master cell in paste range — double-click to unlock first', 'err');
              return;
            }
          }
        }
      }
    }
    pushHistory();
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
      'Pasted ' + result.rows + '×' + result.cols + ' cells at ' + cellAddress(startRow, startCol) +
      ' (' + tab.cols + '×' + tab.rows + ' grid)',
      'ok'
    );
  }

  function onCellCopy(e) {
    const ta = e.currentTarget;
    if (!ta || ta.tagName !== 'TEXTAREA') return;
    const tab = activeTab();
    const sticky = getStickyCellRange();
    if (tab && sticky && e.clipboardData) {
      const value = buildCellRangeTsv(
        tab, sticky.rMin, sticky.cMin, sticky.rMax, sticky.cMax
      );
      e.preventDefault();
      e.clipboardData.setData('text/plain', value);
      const rows = sticky.rMax - sticky.rMin + 1;
      const cols = sticky.cMax - sticky.cMin + 1;
      setStatus('Copied ' + rows + '×' + cols + ' cells', 'ok');
      return;
    }
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
    const idx = parseInt(ta.dataset.idx, 10);
    if (!Number.isNaN(idx) && idx >= 0 && idx < tab.cells.length &&
        !isMasterTab(tab) && isCellMasterLocked(tab, idx) &&
        !(isMasterCellEditSession() &&
          masterSegmentEdit.tabId === tab.id &&
          masterSegmentEdit.cellIndex === idx)) {
      e.preventDefault();
      setStatus('Locked Master cell — double-click to unlock before editing', 'err');
      return;
    }
    const value = ta.value == null ? '' : String(ta.value);
    if (!e.clipboardData) return;
    e.preventDefault();
    e.clipboardData.setData('text/plain', value);
    if (!Number.isNaN(idx) && idx >= 0 && idx < tab.cells.length) {
      pushHistory();
      writeCellCurrentPage(tab, idx, '');
      ta.value = '';
      clearCellMasterLock(tab, idx);
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
    const tab = toolsTargetTab();
    if (!tab) return;
    pushHistory();
    ensureNestedCells(tab);
    ensureCellLocks(tab);
    ensureCellPages(tab);
    ensureCellShades(tab);
    for (let c = 0; c < tab.cols; c++) {
      tab.cells.push('');
      tab.nestedCells.push([]);
      tab.cellLocks.push(null);
      tab.cellPages.push(makeEmptyCellPages(''));
      tab.cellShades.push(null);
    }
    tab.rows += 1;
    ensureRowHeightsLength(tab);
    renderTabs();
    renderGrid();
    renderMasterLibrary();
    scheduleSave();
    setStatus('Row added to ' + tab.title + ' (' + tab.cols + '×' + tab.rows + ')');
  }

  function rowIsEmpty(tab, rowIndex) {
    ensureNestedCells(tab);
    for (let col = 0; col < tab.cols; col++) {
      const idx = rowIndex * tab.cols + col;
      if ((tab.cells[idx] || '').trim()) return false;
      const nests = tab.nestedCells[idx] || [];
      for (let n = 0; n < nests.length; n++) {
        if (nestHasContent(nests[n])) return false;
      }
    }
    return true;
  }

  function sortRowsByColumn(colIndex, direction) {
    const tab = activeTab();
    if (!tab) return;
    if (colIndex < 0 || colIndex >= tab.cols) return;
    if (direction !== 'asc' && direction !== 'desc') return;

    pushHistory();
    ensureNestedCells(tab);

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

    ensureCellLocks(tab);
    ensureCellPages(tab);
    ensureCellShades(tab);
    const newCells = [];
    const newNested = [];
    const newLocks = [];
    const newPages = [];
    const newShades = [];
    for (let i = 0; i < order.length; i++) {
      const src = order[i];
      for (let c = 0; c < cols; c++) {
        const idx = src * cols + c;
        newCells.push(tab.cells[idx] || '');
        newNested.push(normalizeNestList(tab.nestedCells[idx]));
        newLocks.push(tab.cellLocks[idx] || null);
        newPages.push(cloneCellPagesEntry(tab.cellPages[idx]));
        newShades.push(tab.cellShades[idx] || null);
      }
    }

    const oldToNew = new Array(rows);
    for (let newR = 0; newR < order.length; newR++) {
      oldToNew[order[newR]] = newR;
    }

    tab.cells = newCells;
    tab.nestedCells = newNested;
    tab.cellLocks = newLocks;
    tab.cellPages = newPages;
    tab.cellShades = newShades;
    if (Array.isArray(tab.rowHeights) && tab.rowHeights.length === rows) {
      tab.rowHeights = order.map(function (src) { return tab.rowHeights[src]; });
    }

    state.confirmedLinks.forEach(function (link) {
      if (link.tabId !== tab.id) return;
      const row = Math.floor(link.cellIndex / cols);
      const col = link.cellIndex % cols;
      if (row < 0 || row >= rows) return;
      link.cellIndex = oldToNew[row] * cols + col;
    });

    // Master grid A–Z rewrites shared Master cells — remap other tabs' Master locks.
    if (isMasterTab(tab)) remapMasterCellIndices(oldToNew, cols);

    renderTabs();
    renderGrid();
    renderMasterLibrary();
    scheduleSave();
    setStatus(direction === 'desc' ? 'Sorted by column A Z–A' : 'Sorted by column A A–Z');
  }

  /**
   * Shift/move a cell row with ↑/↓.
   * ↓: move this row down one; push every row below further down (grow grid by 1);
   *    leave an empty row behind — never overwrite.
   * ↑: move this row up one via extract+insert so neighbors shift without clobbering.
   * Nested, Combined link indices, and rowHeights stay aligned.
   */
  function moveRow(rowIndex, direction) {
    const tab = activeTab();
    if (!tab || (direction !== -1 && direction !== 1)) return;
    const destination = rowIndex + direction;
    if (rowIndex < 0 || rowIndex >= tab.rows || destination < 0) return;

    pushHistory();
    ensureNestedCells(tab);

    const cols = tab.cols;

    if (direction < 0) {
      // ↑ Extract this row and insert it one slot higher; in-between rows shift down.
      const rows = tab.rows;
      const order = [];
      for (let r = 0; r < rows; r++) order.push(r);
      order.splice(rowIndex, 1);
      order.splice(destination, 0, rowIndex);

      ensureCellLocks(tab);
      ensureCellPages(tab);
      ensureCellShades(tab);
      const newCells = [];
      const newNested = [];
      const newLocks = [];
      const newPages = [];
      const newShades = [];
      for (let i = 0; i < order.length; i++) {
        const src = order[i];
        for (let c = 0; c < cols; c++) {
          const idx = src * cols + c;
          newCells.push(tab.cells[idx] || '');
          newNested.push(normalizeNestList(tab.nestedCells[idx]));
          newLocks.push(tab.cellLocks[idx] || null);
          newPages.push(cloneCellPagesEntry(tab.cellPages[idx]));
          newShades.push(tab.cellShades[idx] || null);
        }
      }

      const oldToNew = new Array(rows);
      for (let newR = 0; newR < order.length; newR++) {
        oldToNew[order[newR]] = newR;
      }

      tab.cells = newCells;
      tab.nestedCells = newNested;
      tab.cellLocks = newLocks;
      tab.cellPages = newPages;
      tab.cellShades = newShades;
      if (Array.isArray(tab.rowHeights) && tab.rowHeights.length === rows) {
        tab.rowHeights = order.map(function (src) { return tab.rowHeights[src]; });
      }
      remapConfirmedRowsByOrder(tab.id, cols, oldToNew);
      if (isMasterTab(tab)) remapMasterCellIndices(oldToNew, cols);
    } else {
      // ↓ Grow by one, push everything below further down, move this row into the gap.
      ensureCellLocks(tab);
      ensureCellPages(tab);
      ensureCellShades(tab);
      tab.cells.push.apply(tab.cells, emptyCells(cols, 1));
      tab.nestedCells.push.apply(tab.nestedCells, emptyNestedCells(cols, 1));
      tab.cellLocks.push.apply(tab.cellLocks, emptyCellLocks(cols, 1));
      for (let c = 0; c < cols; c++) {
        tab.cellPages.push(makeEmptyCellPages(''));
        tab.cellShades.push(null);
      }
      tab.rows += 1;
      ensureRowHeightsLength(tab);

      const last = tab.rows - 1;
      for (let row = last; row > rowIndex + 1; row--) {
        for (let col = 0; col < cols; col++) {
          tab.cells[row * cols + col] = tab.cells[(row - 1) * cols + col];
        }
        copyNestedRow(tab, row - 1, row);
        copyCellLockRow(tab, row - 1, row);
        copyCellPagesRow(tab, row - 1, row);
        copyCellShadeRow(tab, row - 1, row);
        if (Array.isArray(tab.rowHeights) && tab.rowHeights.length === tab.rows) {
          tab.rowHeights[row] = tab.rowHeights[row - 1];
        }
      }

      for (let col = 0; col < cols; col++) {
        tab.cells[(rowIndex + 1) * cols + col] = tab.cells[rowIndex * cols + col];
        tab.cells[rowIndex * cols + col] = '';
      }
      copyNestedRow(tab, rowIndex, rowIndex + 1);
      clearNestedRow(tab, rowIndex);
      copyCellLockRow(tab, rowIndex, rowIndex + 1);
      clearCellLockRow(tab, rowIndex);
      copyCellPagesRow(tab, rowIndex, rowIndex + 1);
      clearCellPagesRow(tab, rowIndex);
      copyCellShadeRow(tab, rowIndex, rowIndex + 1);
      clearCellShadeRow(tab, rowIndex);
      if (Array.isArray(tab.rowHeights) && tab.rowHeights.length === tab.rows) {
        tab.rowHeights[rowIndex + 1] = tab.rowHeights[rowIndex];
        tab.rowHeights[rowIndex] = MIN_ROW_HEIGHT;
      }

      // Links on this row and every row below move down one with the push.
      shiftConfirmedRowsFrom(tab.id, cols, rowIndex);
    }

    renderTabs();
    renderGrid();
    renderMasterLibrary();
    scheduleSave();
    setStatus(direction < 0 ? 'Row moved up' : 'Row moved down');
  }

  function addColumn() {
    const tab = toolsTargetTab();
    if (!tab) return;
    pushHistory();
    ensureNestedCells(tab);
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
    remapNestedAfterColumnAdd(tab, oldCols, tab.cols);
    remapCellLocksAfterColumnAdd(tab, oldCols, tab.cols);
    remapCellPagesAfterColumnAdd(tab, oldCols, tab.cols);
    remapCellShadesAfterColumnAdd(tab, oldCols, tab.cols);
    remapConfirmedAfterColumnAdd(tab.id, oldCols, tab.cols);
    renderTabs();
    renderGrid();
    renderMasterLibrary();
    scheduleSave();
    setStatus('Column added to ' + tab.title + ' (' + tab.cols + '×' + tab.rows + ')');
  }

  function rowConfirmedPiecesMissing(tab, rowIndex) {
    const scope = currentPromptScope();
    const pieces = [];
    const colSep = separatorValue(state.separators.column);
    for (let c = 0; c < tab.cols; c++) {
      const idx = rowIndex * tab.cols + c;
      const cellPieces = cellConfirmedPieces(tab, idx, {
        scope: scope,
        skipConfirmed: true
      });
      if (!cellPieces.length) continue;
      if (pieces.length) pieces.push({ type: 'plain', text: colSep });
      for (let i = 0; i < cellPieces.length; i++) pieces.push(cellPieces[i]);
    }
    return pieces;
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
    lastCombinedCaret = { start: 0, end: 0, scope: scope };
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

  // --- Part Find / Find All / Replace (active tab cells + nest pages) ---

  function setPartSearchStatus(text) {
    if (el.partSearchStatus) el.partSearchStatus.textContent = text || '';
  }

  function clearPartSearchDomHighlights() {
    if (!el.cellGrid) return;
    el.cellGrid.querySelectorAll('.search-hit, .search-hit-current, .search-hit-wrap').forEach(function (node) {
      node.classList.remove('search-hit', 'search-hit-current', 'search-hit-wrap');
    });
  }

  function clearPartSearch(opts) {
    partSearchHits = [];
    partSearchHitIndex = -1;
    partSearchFindAll = false;
    if (!(opts && opts.keepQuery)) partSearchQuery = '';
    clearPartSearchDomHighlights();
    setPartSearchStatus('');
    if (opts && opts.keepInputs) return;
    if (el.partFindInput) el.partFindInput.value = '';
    if (el.partReplaceInput) el.partReplaceInput.value = '';
  }

  function currentPartSearchTab() {
    return activeTab();
  }

  function collectPartSearchHits(tab, query) {
    const hits = [];
    if (!tab || !query) return hits;
    const q = String(query);
    const qLower = q.toLowerCase();
    if (!qLower) return hits;
    ensureNestedCells(tab);
    ensureCellPages(tab);
    const n = tab.cols * tab.rows;
    for (let i = 0; i < n; i++) {
      const pagesEntry = tab.cellPages[i] || makeEmptyCellPages(tab.cells[i] || '');
      const pages = pagesEntry.pages || [tab.cells[i] || ''];
      for (let cpi = 0; cpi < pages.length; cpi++) {
        const cellText = pages[cpi] == null ? '' : String(pages[cpi]);
        const cellLower = cellText.toLowerCase();
        let from = 0;
        while (from <= cellLower.length) {
          const at = cellLower.indexOf(qLower, from);
          if (at < 0) break;
          hits.push({
            cellIndex: i,
            nestIndex: null,
            pageIndex: cpi,
            start: at,
            end: at + q.length
          });
          from = at + Math.max(1, qLower.length);
        }
      }
      const nests = tab.nestedCells[i] || [];
      for (let ni = 0; ni < nests.length; ni++) {
        const nest = nests[ni];
        if (!nest || !Array.isArray(nest.pages)) continue;
        for (let pi = 0; pi < nest.pages.length; pi++) {
          const pageText = nest.pages[pi] == null ? '' : String(nest.pages[pi]);
          const pageLower = pageText.toLowerCase();
          from = 0;
          while (from <= pageLower.length) {
            const at = pageLower.indexOf(qLower, from);
            if (at < 0) break;
            hits.push({
              cellIndex: i,
              nestIndex: ni,
              pageIndex: pi,
              start: at,
              end: at + q.length
            });
            from = at + Math.max(1, qLower.length);
          }
        }
      }
    }
    return hits;
  }

  function rebuildPartSearchHits() {
    const tab = currentPartSearchTab();
    const query = el.partFindInput ? el.partFindInput.value : '';
    partSearchQuery = query;
    partSearchHits = collectPartSearchHits(tab, query);
    if (partSearchHitIndex >= partSearchHits.length) partSearchHitIndex = partSearchHits.length - 1;
    return partSearchHits;
  }

  function getHitText(tab, hit) {
    if (!tab || !hit) return '';
    if (hit.nestIndex == null) {
      ensureCellPages(tab);
      const entry = tab.cellPages[hit.cellIndex];
      if (entry && Array.isArray(entry.pages) && Number.isInteger(hit.pageIndex) &&
          hit.pageIndex >= 0 && hit.pageIndex < entry.pages.length) {
        const page = entry.pages[hit.pageIndex];
        return page == null ? '' : String(page);
      }
      return tab.cells[hit.cellIndex] == null ? '' : String(tab.cells[hit.cellIndex]);
    }
    ensureNestedCells(tab);
    const nest = (tab.nestedCells[hit.cellIndex] || [])[hit.nestIndex];
    if (!nest || !Array.isArray(nest.pages)) return '';
    const page = nest.pages[hit.pageIndex];
    return page == null ? '' : String(page);
  }

  function setHitText(tab, hit, text) {
    if (!tab || !hit) return;
    if (hit.nestIndex == null) {
      if (!isMasterTab(tab) && isCellMasterLocked(tab, hit.cellIndex)) return;
      ensureCellPages(tab);
      const entry = tab.cellPages[hit.cellIndex];
      if (entry && Array.isArray(entry.pages) && Number.isInteger(hit.pageIndex) &&
          hit.pageIndex >= 0 && hit.pageIndex < entry.pages.length) {
        entry.pages[hit.pageIndex] = text;
        if ((entry.page || 0) === hit.pageIndex) {
          tab.cells[hit.cellIndex] = text;
        }
        if (isMasterTab(tab)) syncLockedPartPagesFromMaster(hit.cellIndex);
        return;
      }
      writeCellCurrentPage(tab, hit.cellIndex, text);
      if (isMasterTab(tab)) syncLockedPartPagesFromMaster(hit.cellIndex);
      return;
    }
    ensureNestedCells(tab);
    const nest = (tab.nestedCells[hit.cellIndex] || [])[hit.nestIndex];
    if (!nest || !Array.isArray(nest.pages)) return;
    if (hit.pageIndex < 0 || hit.pageIndex >= nest.pages.length) return;
    nest.pages[hit.pageIndex] = text;
  }

  function queryHitTextarea(hit) {
    if (!el.cellGrid || !hit) return null;
    if (hit.nestIndex == null) {
      return el.cellGrid.querySelector('textarea.cell[data-idx="' + hit.cellIndex + '"]');
    }
    return el.cellGrid.querySelector(
      'textarea.cell-nest-input[data-idx="' + hit.cellIndex + '"][data-nest="' + hit.nestIndex + '"]'
    );
  }

  function applyPartSearchHighlights() {
    clearPartSearchDomHighlights();
    if (!partSearchHits.length) return;
    const markAll = partSearchFindAll;
    const current = partSearchHitIndex >= 0 ? partSearchHits[partSearchHitIndex] : null;
    const seen = Object.create(null);
    function markHit(hit, isCurrent) {
      if (!hit) return;
      const key = hit.cellIndex + ':' + (hit.nestIndex == null
        ? ('c:' + (hit.pageIndex == null ? 'x' : hit.pageIndex))
        : (hit.nestIndex + ':' + hit.pageIndex));
      const ta = queryHitTextarea(hit);
      // Only highlight textarea when its visible page matches.
      const tab = currentPartSearchTab();
      if (tab && hit.nestIndex != null) {
        ensureNestedCells(tab);
        const nest = (tab.nestedCells[hit.cellIndex] || [])[hit.nestIndex];
        if (!nest || nest.page !== hit.pageIndex) {
          if (!isCurrent && !markAll) return;
        }
      } else if (tab && hit.nestIndex == null && Number.isInteger(hit.pageIndex)) {
        ensureCellPages(tab);
        const entry = tab.cellPages[hit.cellIndex];
        if (!entry || entry.page !== hit.pageIndex) {
          if (!isCurrent && !markAll) return;
        }
      }
      if (ta) {
        if (markAll) ta.classList.add('search-hit');
        if (isCurrent) ta.classList.add('search-hit-current');
        const wrap = hit.nestIndex == null
          ? (ta.closest ? ta.closest('.cell-wrap') : null)
          : (ta.closest ? ta.closest('.cell-nest') : null);
        if (wrap && (markAll || isCurrent)) wrap.classList.add('search-hit-wrap');
      }
      seen[key] = true;
    }
    if (markAll) {
      for (let i = 0; i < partSearchHits.length; i++) markHit(partSearchHits[i], false);
    }
    if (current) markHit(current, true);
  }

  function revealPartSearchHit(hit, opts) {
    const tab = currentPartSearchTab();
    if (!tab || !hit) return;
    focusedCell = { tabId: tab.id, index: hit.cellIndex };
    if (hit.nestIndex != null) {
      ensureNestedCells(tab);
      const nest = (tab.nestedCells[hit.cellIndex] || [])[hit.nestIndex];
      if (nest && nest.page !== hit.pageIndex) {
        setNestPage(hit.cellIndex, hit.nestIndex, hit.pageIndex, { forceRender: true });
      }
    } else if (Number.isInteger(hit.pageIndex) && hit.pageIndex >= 0) {
      ensureCellPages(tab);
      const entry = tab.cellPages[hit.cellIndex];
      if (entry && entry.page !== hit.pageIndex) {
        setCellPage(hit.cellIndex, hit.pageIndex, { forceRender: true });
      }
    }
    applyPartSearchHighlights();
    const ta = queryHitTextarea(hit);
    if (ta) {
      ta.focus();
      try {
        const start = Math.max(0, hit.start);
        const end = Math.max(start, Math.min(hit.end, (ta.value || '').length));
        ta.setSelectionRange(start, end);
      } catch (err) { /* ignore */ }
      if (typeof ta.scrollIntoView === 'function') {
        ta.scrollIntoView({ block: 'nearest', inline: 'nearest' });
      }
    }
    if (!(opts && opts.silent)) {
      const n = partSearchHits.length;
      const i = partSearchHitIndex + 1;
      const addr = cellAddressFromIndex(tab, hit.cellIndex);
      const where = hit.nestIndex == null
        ? (addr + (Number.isInteger(hit.pageIndex) ? (' p' + (hit.pageIndex + 1)) : ''))
        : (addr + ' nest ' + (hit.nestIndex + 1) + ' p' + (hit.pageIndex + 1));
      setPartSearchStatus(n ? (i + ' / ' + n) : '0');
      setStatus('Find: ' + where + (n ? (' (' + i + '/' + n + ')') : ''));
    }
  }

  function partFindNext(opts) {
    const tab = currentPartSearchTab();
    if (!tab) {
      setStatus('No part tab to search', 'err');
      return;
    }
    if (toolsTabActive) {
      setStatus('Switch to a part tab to search cells', 'err');
      return;
    }
    const query = el.partFindInput ? el.partFindInput.value : '';
    if (!query) {
      setPartSearchStatus('');
      setStatus('Enter text to find', 'err');
      if (el.partFindInput) el.partFindInput.focus();
      return;
    }
    const rebuilt = query !== partSearchQuery || !partSearchHits.length;
    if (rebuilt) {
      partSearchFindAll = !!(opts && opts.findAll);
      rebuildPartSearchHits();
      partSearchHitIndex = -1;
    } else if (opts && opts.findAll) {
      partSearchFindAll = true;
    }
    if (!partSearchHits.length) {
      partSearchHitIndex = -1;
      applyPartSearchHighlights();
      setPartSearchStatus('0');
      setStatus('No matches in ' + tab.title, 'err');
      return;
    }
    if (opts && opts.findAll && rebuilt) {
      partSearchHitIndex = 0;
      revealPartSearchHit(partSearchHits[0]);
      setPartSearchStatus(partSearchHits.length + ' matches');
      setStatus('Find All: ' + partSearchHits.length + ' in ' + tab.title, 'ok');
      applyPartSearchHighlights();
      return;
    }
    partSearchHitIndex = (partSearchHitIndex + 1) % partSearchHits.length;
    revealPartSearchHit(partSearchHits[partSearchHitIndex]);
  }

  function partFindAll() {
    partFindNext({ findAll: true });
  }

  function replaceInHitText(text, hit, replacement) {
    const start = hit.start;
    const end = hit.end;
    if (start < 0 || end < start || end > text.length) return null;
    return text.slice(0, start) + replacement + text.slice(end);
  }

  function replaceAllInText(text, query, replacement) {
    if (!query) return { text: text, count: 0 };
    const qLower = query.toLowerCase();
    const src = String(text);
    let out = '';
    let i = 0;
    let count = 0;
    const lower = src.toLowerCase();
    while (i < src.length) {
      const at = lower.indexOf(qLower, i);
      if (at < 0) {
        out += src.slice(i);
        break;
      }
      out += src.slice(i, at) + replacement;
      i = at + query.length;
      count++;
    }
    return { text: out, count: count };
  }

  function partReplace(opts) {
    const tab = currentPartSearchTab();
    if (!tab) {
      setStatus('No part tab to search', 'err');
      return;
    }
    if (toolsTabActive) {
      setStatus('Switch to a part tab to replace in cells', 'err');
      return;
    }
    const query = el.partFindInput ? el.partFindInput.value : '';
    const replacement = el.partReplaceInput ? el.partReplaceInput.value : '';
    if (!query) {
      setStatus('Enter text to find before replacing', 'err');
      if (el.partFindInput) el.partFindInput.focus();
      return;
    }
    const replaceAll = !!(opts && opts.replaceAll);

    if (replaceAll) {
      rebuildPartSearchHits();
      if (!partSearchHits.length) {
        setPartSearchStatus('0');
        setStatus('No matches to replace in ' + tab.title, 'err');
        return;
      }
      pushHistory();
      // Unique fields (cell or nest page), replace all occurrences in each.
      const keysDone = Object.create(null);
      let total = 0;
      const touchedCells = Object.create(null);
      for (let h = 0; h < partSearchHits.length; h++) {
        const hit = partSearchHits[h];
        const key = hit.cellIndex + ':' + (hit.nestIndex == null
          ? ('c:' + (hit.pageIndex == null ? 'x' : hit.pageIndex))
          : (hit.nestIndex + ':' + hit.pageIndex));
        if (keysDone[key]) continue;
        keysDone[key] = true;
        const before = getHitText(tab, hit);
        const result = replaceAllInText(before, query, replacement);
        if (!result.count) continue;
        setHitText(tab, hit, result.text);
        total += result.count;
        touchedCells[hit.cellIndex] = true;
        liveSyncConfirmedLinksForCell(tab.id, hit.cellIndex, hit.nestIndex, { silent: true });
      }
      Object.keys(touchedCells).forEach(function (idxStr) {
        const idx = parseInt(idxStr, 10);
        revalidateLinksForCell(tab.id, idx, { silent: true });
      });
      partSearchFindAll = false;
      rebuildPartSearchHits();
      partSearchHitIndex = partSearchHits.length ? 0 : -1;
      renderGrid();
      if (isMasterTab(tab)) renderMasterLibrary();
      scheduleSave();
      setPartSearchStatus(total ? ('replaced ' + total) : '0');
      setStatus(total ? ('Replaced ' + total + ' in ' + tab.title) : ('No matches in ' + tab.title), total ? 'ok' : 'err');
      if (partSearchHitIndex >= 0) revealPartSearchHit(partSearchHits[partSearchHitIndex], { silent: true });
      return;
    }

    // Single replace: ensure a current hit, replace that occurrence, then find next.
    if (query !== partSearchQuery || !partSearchHits.length || partSearchHitIndex < 0) {
      partSearchFindAll = false;
      rebuildPartSearchHits();
      if (!partSearchHits.length) {
        setPartSearchStatus('0');
        setStatus('No matches to replace in ' + tab.title, 'err');
        return;
      }
      partSearchHitIndex = 0;
      revealPartSearchHit(partSearchHits[0], { silent: true });
    }
    const hit = partSearchHits[partSearchHitIndex];
    if (!hit) return;
    const before = getHitText(tab, hit);
    const after = replaceInHitText(before, hit, replacement);
    if (after == null) {
      setStatus('Replace failed — match out of range', 'err');
      return;
    }
    pushHistory();
    setHitText(tab, hit, after);
    liveSyncConfirmedLinksForCell(tab.id, hit.cellIndex, hit.nestIndex);
    // Refresh DOM value if visible
    if (hit.nestIndex != null) {
      ensureNestedCells(tab);
      const nest = (tab.nestedCells[hit.cellIndex] || [])[hit.nestIndex];
      if (nest && nest.page !== hit.pageIndex) {
        setNestPage(hit.cellIndex, hit.nestIndex, hit.pageIndex, { forceRender: true });
      } else {
        const ta = queryHitTextarea(hit);
        if (ta) ta.value = after;
      }
    } else {
      ensureCellPages(tab);
      const entry = tab.cellPages[hit.cellIndex];
      if (entry && Number.isInteger(hit.pageIndex) && entry.page !== hit.pageIndex) {
        setCellPage(hit.cellIndex, hit.pageIndex, { forceRender: true });
      } else {
        const ta = queryHitTextarea(hit);
        if (ta) ta.value = after;
      }
    }
    applyAppendCheckedState();
    if (isMasterTab(tab)) renderMasterLibrary();
    scheduleSave();

    // Rebuild and advance to next match after this point.
    const resumeCell = hit.cellIndex;
    const resumeNest = hit.nestIndex;
    const resumePage = hit.pageIndex;
    const resumeStart = hit.start + replacement.length;
    rebuildPartSearchHits();
    let nextIdx = -1;
    for (let i = 0; i < partSearchHits.length; i++) {
      const h = partSearchHits[i];
      if (h.cellIndex < resumeCell) continue;
      if (h.cellIndex > resumeCell) { nextIdx = i; break; }
      if (resumeNest == null) {
        if (h.nestIndex != null) { nextIdx = i; break; }
        if (Number.isInteger(resumePage) && Number.isInteger(h.pageIndex)) {
          if (h.pageIndex < resumePage) continue;
          if (h.pageIndex > resumePage) { nextIdx = i; break; }
        }
        if (h.start >= resumeStart) { nextIdx = i; break; }
        continue;
      }
      if (h.nestIndex == null) continue;
      if (h.nestIndex < resumeNest) continue;
      if (h.nestIndex > resumeNest) { nextIdx = i; break; }
      if (h.pageIndex < resumePage) continue;
      if (h.pageIndex > resumePage) { nextIdx = i; break; }
      if (h.start >= resumeStart) { nextIdx = i; break; }
    }
    if (nextIdx < 0 && partSearchHits.length) nextIdx = 0; // wrap
    partSearchHitIndex = nextIdx;
    if (partSearchHitIndex >= 0) {
      revealPartSearchHit(partSearchHits[partSearchHitIndex]);
      setStatus('Replaced in ' + cellAddressFromIndex(tab, hit.cellIndex), 'ok');
    } else {
      applyPartSearchHighlights();
      setPartSearchStatus('done');
      setStatus('Replaced — no more matches', 'ok');
    }
  }

  function focusPartFindInput() {
    if (toolsTabActive) return false;
    if (!el.partFindInput) return false;
    el.partFindInput.focus();
    el.partFindInput.select();
    return true;
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
  if (el.btnFitRows) el.btnFitRows.addEventListener('click', autoFitAllRowHeights);
  document.addEventListener('pointerdown', function (e) {
    if (col1FilterMenuOpen) {
      const wrap = el.cellGrid && el.cellGrid.querySelector('.column-col1-filter');
      if (!(wrap && wrap.contains(e.target))) closeCol1FilterMenu();
    }
    if (masterLibFilterMenuOpen) {
      const wrap = el.masterLibraryItems && el.masterLibraryItems.querySelector('.master-library-filter');
      if (!(wrap && wrap.contains(e.target))) closeMasterLibFilterMenu();
    }
    if (tabIconPickerEl && !tabIconPickerEl.contains(e.target)) {
      closeTabIconPicker();
    }
  });
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') {
      closeTabIconPicker();
      closeCol1FilterMenu();
      closeMasterLibFilterMenu();
      if (stickyCellRange) clearStickyCellRange();
      clearKeyboardCellRange();
      if (partSearchHits.length || partSearchFindAll) {
        clearPartSearch({ keepInputs: true, keepQuery: true });
        partSearchQuery = el.partFindInput ? el.partFindInput.value : '';
      }
    }
    // Fallback: sticky multi-cell Delete when focus left the cell textarea
    // (e.g. after chrome click) but the rectangle is still highlighted.
    if ((e.key === 'Delete' || e.key === 'Backspace') &&
        !e.ctrlKey && !e.metaKey && !e.altKey && getStickyCellRange()) {
      const t = e.target;
      if (t && t.tagName === 'TEXTAREA' && t.classList && t.classList.contains('cell')) {
        // onCellKeydown already handles this path.
      } else if (t && (
          t.tagName === 'INPUT' ||
          t.tagName === 'TEXTAREA' ||
          t.isContentEditable ||
          (t.closest && (t.closest('#combined-prompt') || t.closest('.combined-section') ||
            t.closest('.master-segment-dialog') || t.closest('.part-search-bar'))))) {
        // Leave typing surfaces alone.
      } else {
        e.preventDefault();
        clearStickyMultiCellContents();
      }
    }
    if ((e.key === 'f' || e.key === 'F') && (e.ctrlKey || e.metaKey) && !e.altKey) {
      if (focusPartFindInput()) {
        e.preventDefault();
      }
    }
  });
  window.addEventListener('resize', function () {
    closeTabIconPicker();
    closeCol1FilterMenu();
    closeMasterLibFilterMenu();
  });
  if (el.cellGrid && el.cellGrid.parentElement) {
    el.cellGrid.parentElement.addEventListener('scroll', closeCol1FilterMenu, { passive: true });
  }
  if (el.masterLibrary) {
    // Capture scroll from the results pane (or legacy items scroll). Do not close
    // when the user scrolls inside the fixed Values dropdown itself — that was the
    // "scroll down Values glitch" (menu closed / jumped mid-scroll).
    el.masterLibrary.addEventListener('scroll', function (e) {
      if (suppressMasterLibMenuScrollClose) return;
      const t = e.target;
      if (!t || !t.classList) return;
      const menu = el.masterLibraryItems && el.masterLibraryItems.querySelector('.master-library-filter-menu');
      if (menu && (t === menu || menu.contains(t))) return;
      if (t.classList.contains('master-library-results') ||
          t.classList.contains('master-library-items') ||
          t === el.masterLibrary) {
        closeMasterLibFilterMenu();
      }
    }, { passive: true, capture: true });
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
    resizeCombinedSectionByKeyboard(event.key === 'ArrowUp' ? -24 : 24);
  });

  function onSeparatorInput(key, input) {
    // Live value so the next append uses what is on screen. Combined itself
    // is rewritten only on commit (change/blur), not on each keystroke.
    input.addEventListener('input', function () {
      pushHistory({ coalesce: true });
      state.separators[key] = input.value;
      scheduleSave();
    });
    input.addEventListener('change', function () {
      state.separators[key] = input.value;
      if (rewriteCombinedForSeparators()) {
        setStatus('Updated Combined separators', 'ok');
      }
    });
  }

  if (el.partSeparator) onSeparatorInput('part', el.partSeparator);
  if (el.columnSeparator) onSeparatorInput('column', el.columnSeparator);
  if (el.rowSeparator) onSeparatorInput('row', el.rowSeparator);

  if (el.btnFind) el.btnFind.addEventListener('click', function () { partFindNext(); });
  if (el.btnFindAll) el.btnFindAll.addEventListener('click', partFindAll);
  if (el.btnReplace) {
    el.btnReplace.addEventListener('click', function (e) {
      partReplace({ replaceAll: !!(e && e.shiftKey) });
    });
  }
  if (el.partFindInput) {
    el.partFindInput.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') {
        e.preventDefault();
        if (e.shiftKey) partFindAll();
        else partFindNext();
      }
    });
    el.partFindInput.addEventListener('input', function () {
      partSearchQuery = '';
      partSearchHits = [];
      partSearchHitIndex = -1;
      partSearchFindAll = false;
      clearPartSearchDomHighlights();
      setPartSearchStatus('');
    });
  }
  if (el.partReplaceInput) {
    el.partReplaceInput.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') {
        e.preventDefault();
        partReplace({ replaceAll: !!e.shiftKey });
      }
    });
  }

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

  el.combined.addEventListener('pointerdown', function (e) {
    // Focus often lacks modifier flags; remember Ctrl/Cmd from the press that activates Combined.
    combinedActivateSkipCopy = !!(e.ctrlKey || e.metaKey);
  });
  el.combined.addEventListener('dblclick', function (e) {
    const seg = e.target && e.target.closest
      ? e.target.closest('.confirmed-segment.master-segment-locked')
      : null;
    if (!seg || !el.combined.contains(seg)) return;
    e.preventDefault();
    e.stopPropagation();
    beginMasterSegmentEdit(seg);
  });
  el.combined.addEventListener('beforeinput', function (e) {
    if (masterSegmentDialogOpen) {
      e.preventDefault();
      return;
    }
    // While editing an unlocked Master segment, allow input only outside other locked spans.
    if (typeof e.getTargetRanges === 'function') {
      const ranges = e.getTargetRanges();
      for (let i = 0; i < ranges.length; i++) {
        const hit = rangeTouchesLockedMasterSegment(ranges[i]);
        if (hit) {
          e.preventDefault();
          setStatus('Locked Master segment — double-click to unlock before editing', 'err');
          return;
        }
      }
    }
    if (selectionTouchesLockedMasterSegment()) {
      e.preventDefault();
      setStatus('Locked Master segment — double-click to unlock before editing', 'err');
    }
  });
  el.combined.addEventListener('focusout', function (e) {
    if (!isMasterCombinedEditSession() || masterSegmentDialogOpen) return;
    const next = e.relatedTarget;
    if (next && el.masterSegmentDialog && el.masterSegmentDialog.contains(next)) return;
    const span = findCombinedSegmentSpan(masterSegmentEdit.linkId);
    if (span && next && (next === span || span.contains(next))) return;
    // Defer so click-to-dialog / internal focus moves settle.
    window.setTimeout(function () {
      if (!masterSegmentEdit || masterSegmentDialogOpen) return;
      const active = document.activeElement;
      if (el.masterSegmentDialog && el.masterSegmentDialog.contains(active)) return;
      const live = findCombinedSegmentSpan(masterSegmentEdit.linkId);
      if (live && (active === live || live.contains(active))) return;
      requestMasterSegmentFinish();
    }, 0);
  });
  el.combined.addEventListener('focus', function (e) {
    copyCombinedOnActivate(e);
  });
  el.combined.addEventListener('click', function (e) {
    copyCombinedOnActivate(e);
    combinedActivateSkipCopy = false;
    rememberCombinedCaretFromDom();
  });
  el.combined.addEventListener('pointerup', function () {
    window.setTimeout(function () {
      combinedActivateSkipCopy = false;
      if (document.activeElement === el.combined) rememberCombinedCaretFromDom();
    }, 0);
  });
  el.combined.addEventListener('pointercancel', function () {
    combinedActivateSkipCopy = false;
  });
  el.combined.addEventListener('keyup', function () {
    rememberCombinedCaretFromDom();
  });
  el.combined.addEventListener('blur', function () {
    // selectionchange often clears before blur; keep last remembered offsets.
    rememberCombinedCaretFromDom();
  });
  document.addEventListener('selectionchange', function () {
    if (document.activeElement !== el.combined) return;
    rememberCombinedCaretFromDom();
  });

  el.combined.addEventListener('input', function () {
    dismissUndoClear();
    pushHistory({ coalesce: true });
    syncConfirmedFromCombinedDom();
    rememberCombinedCaretFromDom();
    syncCombinedOverflowY();
    scheduleSave();
  });

  el.combined.addEventListener('keydown', function (event) {
    if (isMasterCombinedEditSession() && !masterSegmentDialogOpen && !event.isComposing) {
      if (event.key === 'Escape') {
        event.preventDefault();
        cancelMasterSegmentEdit();
        return;
      }
      if (event.key === 'Enter' && !event.shiftKey) {
        event.preventDefault();
        requestMasterSegmentFinish();
        return;
      }
    }
    if (!event.isComposing && !masterSegmentDialogOpen) {
      const editingCombined = isMasterCombinedEditSession();
      if (event.key === 'Backspace' || event.key === 'Delete') {
        const touched = selectionTouchesLockedMasterSegment() ||
          lockedSegmentAdjacentToCaret(event.key);
        if (touched) {
          event.preventDefault();
          setStatus('Locked Master segment — double-click to unlock before editing', 'err');
          return;
        }
      } else if (!editingCombined) {
        // Block printable / modifying keys when caret sits in a locked span.
        const isMod = event.ctrlKey || event.metaKey || event.altKey;
        const isNav = event.key === 'Escape' || event.key === 'Tab' ||
          event.key === 'ArrowUp' || event.key === 'ArrowDown' ||
          event.key === 'ArrowLeft' || event.key === 'ArrowRight' ||
          event.key === 'Home' || event.key === 'End' ||
          event.key === 'PageUp' || event.key === 'PageDown';
        if (!isMod && !isNav && event.key.length === 1 &&
            selectionTouchesLockedMasterSegment()) {
          event.preventDefault();
          setStatus('Locked Master segment — double-click to unlock before editing', 'err');
          return;
        }
      }
    }
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
    if (masterSegmentDialogOpen || selectionTouchesLockedMasterSegment()) {
      event.preventDefault();
      setStatus('Locked Master segment — double-click to unlock before editing', 'err');
      return;
    }
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

  if (el.btnMasterOverwrite) {
    el.btnMasterOverwrite.addEventListener('click', function () {
      overwriteMasterFromSegmentEdit();
    });
  }
  if (el.btnMasterKeepLocal) {
    el.btnMasterKeepLocal.addEventListener('click', function () {
      keepMasterSegmentLocalOnly();
    });
  }
  if (el.btnMasterCancelEdit) {
    el.btnMasterCancelEdit.addEventListener('click', function () {
      cancelMasterSegmentEdit();
    });
  }
  if (el.masterSegmentDialog) {
    el.masterSegmentDialog.addEventListener('keydown', function (event) {
      if (event.key === 'Escape') {
        event.preventDefault();
        cancelMasterSegmentEdit();
      }
    });
  }

  el.globalCombined.addEventListener('change', function () {
    pushHistory();
    const next = el.globalCombined.checked;
    if (next && !state.globalCombined) mergePartPrompts();
    state.globalCombined = next;
    renderCombinedPrompt();
    renderMasterLibrary();
    scheduleSave();
  });

  if (el.matchSourceOrder) {
    el.matchSourceOrder.addEventListener('change', function () {
      pushHistory();
      state.matchSourceOrder = !!el.matchSourceOrder.checked;
      if (state.matchSourceOrder) {
        reorderCombinedToSourceOrder(currentPromptScope());
      }
      renderCombinedPrompt();
      applyConfirmedCellHighlights();
      renderMasterLibrary();
      scheduleSave();
      setStatus(
        state.matchSourceOrder
          ? 'Match source order on — Combined segments follow grid order'
          : 'Match source order off — Combined uses caret/append order',
        'ok'
      );
    });
  }

  async function init() {
    try {
      const savedHeight = Number(localStorage.getItem('click2copy-combined-height'));
      if (Number.isFinite(savedHeight) && savedHeight > 0) resizeCombinedSection(savedHeight);
    } catch (err) {
      console.warn('Could not load combined prompt height:', err);
    }

    if (typeof ResizeObserver === 'function' && el.combined) {
      const combinedOverflowObserver = new ResizeObserver(function () {
        syncCombinedOverflowY();
      });
      combinedOverflowObserver.observe(el.combined);
    }
    syncCombinedOverflowY();

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


  function showUpdateBanner(payload) {
    if (!el.updateBanner) return;
    const disk = payload && payload.diskVersion;
    const running = payload && payload.runningVersion;
    let text = 'Update ready — Restart';
    if (disk && running && disk !== running) {
      text = 'Update ready — v' + disk + ' (running v' + running + ') — Restart';
    } else if (disk) {
      text = 'Update ready — v' + disk + ' — Restart';
    }
    if (el.updateBannerText) el.updateBannerText.textContent = text;
    el.updateBanner.hidden = false;
    pushEvent(text, 'ok');
  }

  if (window.click2copy && typeof window.click2copy.onUpdateReady === 'function') {
    window.click2copy.onUpdateReady(function (payload) {
      showUpdateBanner(payload);
    });
  }
  if (el.btnUpdateRestart) {
    el.btnUpdateRestart.addEventListener('click', function () {
      if (!window.click2copy || typeof window.click2copy.relaunchApp !== 'function') {
        setStatus('Restart unavailable', 'err');
        return;
      }
      setStatus('Restarting…', 'ok');
      window.click2copy.relaunchApp().catch(function (err) {
        console.error(err);
        setStatus('Restart failed', 'err');
      });
    });
  }
  if (el.btnUpdateDismiss) {
    el.btnUpdateDismiss.addEventListener('click', function () {
      if (el.updateBanner) el.updateBanner.hidden = true;
    });
  }

  init();

})();
