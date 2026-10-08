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

  /** v0.140 single-source-of-truth Combined (checks → generated text + Notes). */
  const COMBINED_MODEL = 2;

  const state = {
    documents: [],
    activeDocumentId: null,
    tabs: [],
    activeTabId: null,
    combinedPrompt: '',
    // v0.140: Combined text is DERIVED from check keys (confirmedLinks) — the
    // only stored user text is Notes (global + per-tab when Global is off).
    combinedModel: COMBINED_MODEL,
    combinedNotes: '',
    partNotes: {},
    globalCombined: true,
    matchSourceOrder: true,
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
  /** v0.143: part Col A cell currently in type-edit mode ({ tabId, index }) — else click filters. */
  let col1EditCell = null;
  /** UI-only Master insert filter menu open (ephemeral; not per-tab) */
  let masterLibFilterMenuOpen = false;
  /**
   * v0.142: Values dropdowns live on <body> (portal). Inside the Master insert
   * toolbar (z-index 2 stacking context) or a sticky column header (z-index 4)
   * the grid's later sticky headers painted OVER the open list ("bar covering
   * master value"). On <body> with position:fixed + z-index they sit above
   * every grid header, row number and cell control.
   */
  const popoverMenus = { masterLib: null, grid: null };
  function portalPopover(kind, menu) {
    const prev = popoverMenus[kind];
    if (prev && prev !== menu && prev.parentNode) prev.parentNode.removeChild(prev);
    popoverMenus[kind] = menu;
    if (menu && menu.parentNode !== document.body) document.body.appendChild(menu);
  }
  function dropPopover(kind) {
    const prev = popoverMenus[kind];
    if (prev && prev.parentNode) prev.parentNode.removeChild(prev);
    popoverMenus[kind] = null;
  }
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
    btnAddRow: document.getElementById('btn-add-row'),
    btnAddCol: document.getElementById('btn-add-col'),
    btnFilterAll: document.getElementById('btn-filter-all'),
    btnFilterNonempty: document.getElementById('btn-filter-nonempty'),
    btnFilterIncluded: document.getElementById('btn-filter-included'),
    btnFitRows: document.getElementById('btn-fit-rows'),
    btnCopy: document.getElementById('btn-copy'),
    btnClear: document.getElementById('btn-clear'),
    btnUncheckAll: document.getElementById('btn-uncheck-all'),
    combinedNotes: document.getElementById('combined-notes'),
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
    partSearchStatus: document.getElementById('part-search-status'),
    updateBanner: document.getElementById('update-banner'),
    updateBannerText: document.getElementById('update-banner-text'),
    btnUpdateRestart: document.getElementById('btn-update-restart'),
    backupRemember: document.getElementById('backup-remember'),
    btnBackupFolder: document.getElementById('btn-backup-folder'),
    backupFolderPath: document.getElementById('backup-folder-path'),
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
    syncStoredPartPagesFromMaster(masterIdx);
  }

  /** Same push for part pages not currently shown (stored snapshots). */
  function syncStoredPartPagesFromMaster(masterIdx) {
    forEachPartLockSite(function (site, i, lock) {
      if (site.live || lock.masterCellIndex !== masterIdx) return;
      writeLockSiteFromMaster(site, i, masterIdx);
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
    syncStoredPartPagesFromMaster(masterIdx);
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
      combinedModel: COMBINED_MODEL,
      combinedNotes: '',
      partNotes: {},
      globalCombined: true,
      matchSourceOrder: true,
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
    return documentDataFromState();
  }

  /** Document payload from live state — no page flush (see snapshot / cloneCurrentDocument). */
  function documentDataFromState() {
    return {
      tabs: state.tabs,
      activeTabId: state.activeTabId,
      // combinedPrompt / partPrompts are a derived cache (what the checks
      // generate) — kept for older builds; never read back as user text.
      combinedPrompt: state.combinedPrompt,
      combinedModel: COMBINED_MODEL,
      combinedNotes: state.combinedNotes || '',
      partNotes: Object.assign({}, state.partNotes),
      globalCombined: state.globalCombined,
      matchSourceOrder: state.matchSourceOrder,
      partPrompts: Object.assign({}, state.partPrompts),
      separators: state.separators,
      // Check keys only: which cell / nest is checked, in which scope, and when.
      confirmedLinks: state.confirmedLinks.map(function (link) {
        const out = {
          id: link.id,
          tabId: link.tabId,
          cellIndex: link.cellIndex,
          scope: link.scope
        };
        if (Number.isInteger(link.nestIndex) && link.nestIndex >= 0) {
          out.nestIndex = link.nestIndex;
        }
        if (Number.isFinite(link.seq)) out.seq = link.seq;
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
      combinedModel: data.combinedModel,
      combinedNotes: typeof data.combinedNotes === 'string' ? data.combinedNotes : '',
      partNotes: data.partNotes && typeof data.partNotes === 'object' ? data.partNotes : {},
      globalCombined: typeof data.globalCombined === 'boolean' ? data.globalCombined : true,
      matchSourceOrder: typeof data.matchSourceOrder === 'boolean' ? data.matchSourceOrder : false,
      partPrompts: data.partPrompts && typeof data.partPrompts === 'object' ? data.partPrompts : {},
      separators: data.separators || DEFAULT_SEPARATORS,
      confirmedLinks: Array.isArray(data.confirmedLinks) ? data.confirmedLinks : []
    }));
  }

  /**
   * Undo/redo capture — must not touch live state. snapshot() flushes by
   * rebuilding tab.pages[page] / cellPages / nestedCells, so a caller holding
   * one of those objects across pushHistory() would then mutate an orphan
   * (part-page rename snapped back; cell/nest page add/remove no-ops).
   * Flush into the deep copy instead.
   */
  function cloneCurrentDocument() {
    const live = activeTab();
    const data = cloneDocumentData(documentDataFromState());
    data.tabs.forEach(function (tab, i) {
      if (i === 0 || tab.id === 'master') {
        ensureTabPages(tab);
        tab.pages = [captureTabPageGridOnly(tab)];
        tab.page = 0;
        return;
      }
      ensureTabPages(tab);
      if (live && tab.id === live.id) flushLiveCellInputs(tab);
      tab.pages[tab.page] = captureTabPage(tab);
    });
    return data;
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

  /**
   * Safety net for the single source of truth: every mutation ends in
   * scheduleSave, so re-derive the visible Combined here and repaint only
   * when it differs from what is on screen (paths that forgot to repaint).
   */
  let lastCombinedViewKey = '';
  function combinedViewKey(scope) {
    return scope + '\u0000' + getNotes(scope) + '\u0000' + getPromptText(scope) + '\u0000' +
      linksForScope(scope).map(function (link) { return link.id; }).join(',');
  }

  function refreshCombinedViewIfStale() {
    if (!el.combined || !state.tabs.length) return;
    const scope = currentPromptScope();
    regenerateCombinedScope(scope);
    if (combinedViewKey(scope) !== lastCombinedViewKey) renderCombinedPrompt();
  }

  function scheduleSave() {
    refreshCombinedViewIfStale();
    // Master Column B duplicate marks follow every edit (debounced, no re-render).
    const liveTab = activeTab();
    if (liveTab && isMasterTab(liveTab)) scheduleMasterDupeMarks();
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
      { label: 'Notes length', value: String(getNotes(currentPromptScope()).length) },
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
    return hasCell || !!(state.partNotes[tab.id] || '').trim();
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
      const cellIndex = toNonNegInt(item.cellIndex);
      if (cellIndex === null) continue;
      const scope = typeof item.scope === 'string' && item.scope ? item.scope : 'global';
      const nestIndex = toNonNegInt(item.nestIndex);
      // v0.140: a link is a CHECK KEY. text/start/end are a derived cache
      // (rebuilt by regenerateCombinedScope); legacy values are only read by
      // the one-time migration. Master lock fields live on cells only.
      const start = toNonNegInt(item.start);
      const end = toNonNegInt(item.end);
      const hasRange = typeof item.text === 'string' && start !== null && end !== null && end >= start;
      const link = {
        id: item.id,
        tabId: item.tabId,
        cellIndex: cellIndex,
        text: hasRange ? item.text : '',
        start: hasRange ? start : 0,
        end: hasRange ? end : 0,
        scope: scope
      };
      if (hasRange) link.legacyRange = true;
      if (nestIndex !== null) link.nestIndex = nestIndex;
      if (typeof item.seq === 'number' && Number.isFinite(item.seq)) link.seq = item.seq;
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

  /* ── v0.140 single source of truth ─────────────────────────────────────
   * Checks (state.confirmedLinks = check keys) are the ONLY truth. The
   * Combined text for a scope is generated from them every time:
   *   Match source order ON  → grid order (tab bar → row-major → parent, nests)
   *   Match source order OFF → the order cells were checked (link.seq)
   * joined with the Tools part / column / row separators. Col1 (filter-only)
   * contributes nothing. The user's own words live in Notes (per scope),
   * copied ahead of the generated text. Nothing can drift because there is no
   * stored Combined text to drift from.
   *
   * Part pages: only the CURRENT part page's checks are live (switching pages
   * swaps them via tab.pages[].confirmed — see restoreTabCombinedFromPage).
   * NOTE(all-pages): to include every part page later, also walk
   * tab.pages[p].confirmed for p !== tab.page here and read text from
   * tab.pages[p].cells / nestedCells instead of the live arrays.
   */
  let linkSeqCounter = 0;

  function nextLinkSeq() {
    linkSeqCounter += 1;
    return linkSeqCounter;
  }

  function syncLinkSeqCounter() {
    let max = 0;
    state.confirmedLinks.forEach(function (link) {
      if (Number.isFinite(link.seq) && link.seq > max) max = link.seq;
    });
    state.tabs.forEach(function (tab) {
      (Array.isArray(tab.pages) ? tab.pages : []).forEach(function (page) {
        (page && Array.isArray(page.confirmed) ? page.confirmed : []).forEach(function (key) {
          if (key && Number.isFinite(key.seq) && key.seq > max) max = key.seq;
        });
      });
    });
    if (max > linkSeqCounter) linkSeqCounter = max;
  }

  function compareLinksByCheckOrder(a, b) {
    const sa = Number.isFinite(a.seq) ? a.seq : Infinity;
    const sb = Number.isFinite(b.seq) ? b.seq : Infinity;
    if (sa !== sb) return sa < sb ? -1 : 1;
    return compareLinksBySourceOrder(a, b);
  }

  function getNotes(scope) {
    const target = scope || currentPromptScope();
    if (target === 'global') return state.combinedNotes || '';
    return state.partNotes[target] || '';
  }

  function setNotes(scope, text) {
    const target = scope || currentPromptScope();
    const value = typeof text === 'string' ? text : '';
    if (target === 'global') state.combinedNotes = value;
    else if (value) state.partNotes[target] = value;
    else delete state.partNotes[target];
  }

  function notesSeparator() {
    return separatorValue(state.separators.part) || '\n\n';
  }

  /** Text Copy / click-to-copy hand out: Notes first, then generated checks. */
  function composeCombinedOutput(scope) {
    const target = scope || currentPromptScope();
    const notes = getNotes(target).replace(/\s+$/, '');
    const generated = getPromptText(target);
    if (notes && generated) return notes + notesSeparator() + generated;
    return notes || generated;
  }

  /**
   * Rebuild one scope's generated text + link ranges from the check keys.
   * Drops only keys whose cell / nest no longer exists. Empty sources stay
   * checked but contribute nothing (e.g. a blank cell page).
   */
  function regenerateCombinedScope(scope) {
    const target = scope || currentPromptScope();
    const prevText = getPromptText(target);
    const seenKeys = Object.create(null);
    state.confirmedLinks = state.confirmedLinks.filter(function (link) {
      if (link.scope !== target) return true;
      if (sourceLinkText(link) === null) return false;
      // One check per cell / nest per scope (legacy data could hold repeats).
      const key = link.tabId + '|' + link.cellIndex + '|' + (linkNestIndex(link) === null ? '' : linkNestIndex(link));
      if (seenKeys[key]) return false;
      seenKeys[key] = true;
      return true;
    });
    const links = state.confirmedLinks.filter(function (link) { return link.scope === target; });
    links.forEach(function (link) {
      if (!Number.isFinite(link.seq)) link.seq = nextLinkSeq();
      link.text = sourceLinkText(link) || '';
      link.start = 0;
      link.end = 0;
      delete link.legacyRange;
    });
    const ordered = links.slice().sort(state.matchSourceOrder ? compareLinksBySourceOrder : compareLinksByCheckOrder);
    let text = '';
    let prev = null;
    for (let i = 0; i < ordered.length; i++) {
      const link = ordered[i];
      if (isCol1FilterOnlyLink(link) || isHiddenCol1NestLink(link) || !link.text) continue;
      if (prev) text += separatorBetweenSourceLinks(prev, link);
      link.start = text.length;
      text += link.text;
      link.end = text.length;
      prev = link;
    }
    if (target === 'global') state.combinedPrompt = text;
    else if (text) state.partPrompts[target] = text;
    else delete state.partPrompts[target];
    return text !== prevText;
  }

  function combinedScopesInUse() {
    const seen = Object.create(null);
    const scopes = ['global'];
    seen.global = true;
    function add(scope) {
      if (!scope || seen[scope]) return;
      seen[scope] = true;
      scopes.push(scope);
    }
    state.confirmedLinks.forEach(function (link) { add(link.scope); });
    Object.keys(state.partPrompts || {}).forEach(add);
    return scopes;
  }

  function regenerateAllCombined() {
    let changed = false;
    combinedScopesInUse().forEach(function (scope) {
      if (regenerateCombinedScope(scope)) changed = true;
    });
    return changed;
  }

  /** Legacy names — every caller now just regenerates. */
  function repairConfirmedLinksForScope(scope) {
    regenerateCombinedScope(scope);
  }

  function repairAllConfirmedLinks() {
    regenerateAllCombined();
  }

  /**
   * One-time v0.140 migration of a pre-0.140 document (stored Combined text +
   * offset links). Keeps every check the old build showed as checked (link
   * text still matches its cell and is found in the text), records their
   * visible order as check order, and moves every other character of the old
   * Combined text (typed text, orphaned unchecked segments) into Notes.
   */
  function migrateLegacyCombined() {
    const scopes = combinedScopesInUse();
    let movedChars = 0;
    let keptChecks = 0;
    let droppedChecks = 0;
    const colSep = separatorValue(state.separators.column);
    const rowSep = separatorValue(state.separators.row);
    const partSep = separatorValue(state.separators.part);
    function cleanGap(gap) {
      let g = gap;
      let changed = true;
      while (changed && g) {
        changed = false;
        [partSep, rowSep, colSep].forEach(function (sep) {
          if (!sep) return;
          if (g.startsWith(sep)) { g = g.slice(sep.length); changed = true; }
          if (g.endsWith(sep)) { g = g.slice(0, g.length - sep.length); changed = true; }
        });
      }
      return g.trim();
    }
    scopes.forEach(function (scope) {
      const text = getPromptText(scope);
      const valid = [];
      state.confirmedLinks = state.confirmedLinks.filter(function (link) {
        if (link.scope !== scope) return true;
        const expected = sourceLinkText(link);
        if (expected === null) { droppedChecks += 1; return false; }
        if (isCol1FilterOnlyLink(link)) {
          if (!expected) { droppedChecks += 1; return false; }
          // Legacy Col1 text spans (pre-filter-only) are cut from the text too.
          if (link.legacyRange && link.end > link.start &&
              text.slice(link.start, link.end) === link.text) {
            valid.push(link);
          } else {
            link.start = 0;
            link.end = 0;
            valid.push(link);
          }
          keptChecks += 1;
          return true;
        }
        if (!link.legacyRange || link.text !== expected || !expected ||
            !repairConfirmedLinkOffset(link, text)) {
          droppedChecks += 1;
          return false;
        }
        valid.push(link);
        keptChecks += 1;
        return true;
      });
      // Check order = the order the old build displayed them.
      valid.sort(function (a, b) { return a.start - b.start || compareLinksBySourceOrder(a, b); });
      valid.forEach(function (link) { link.seq = nextLinkSeq(); });
      // Everything not covered by a kept segment is stray user text → Notes.
      const ranges = valid
        .filter(function (link) { return link.end > link.start; })
        .map(function (link) { return [link.start, link.end]; })
        .sort(function (a, b) { return a[0] - b[0]; });
      const gaps = [];
      let pos = 0;
      ranges.forEach(function (r) {
        if (r[0] > pos) gaps.push(text.slice(pos, r[0]));
        pos = Math.max(pos, r[1]);
      });
      if (pos < text.length) gaps.push(text.slice(pos));
      const stray = gaps.map(cleanGap).filter(Boolean).join('\n');
      if (stray) {
        const existing = getNotes(scope);
        setNotes(scope, existing ? existing.replace(/\s+$/, '') + '\n' + stray : stray);
        movedChars += stray.length;
      }
    });
    state.confirmedLinks.forEach(function (link) { delete link.legacyRange; });
    return { keptChecks: keptChecks, droppedChecks: droppedChecks, movedChars: movedChars };
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

  /**
   * v0.140: a check stays valid while its cell / nest exists — the text is
   * always regenerated from the source, so there is nothing to mismatch.
   */
  function linkMatchesSource(link) {
    return sourceLinkText(link) !== null;
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
      if (isCol1FilterOnlyLink(link) || isHiddenCol1NestLink(link)) return false;
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

  /**
   * Visit every Master-origin lock on part tabs: the live grid (current part
   * page) AND every stored part-page snapshot. Master sync/remap used to walk
   * only the live arrays, so links on part pages not currently shown went stale.
   * cb(site, i, lock) — site = { tab, live, cells, cellLocks, cellPages, nestedCells }.
   */
  function forEachPartLockSite(cb) {
    state.tabs.forEach(function (tab) {
      if (!tab || isMasterTab(tab)) return;
      ensureCellLocks(tab);
      ensureCellPages(tab);
      ensureNestedCells(tab);
      const live = { tab: tab, live: true, cells: tab.cells, cellLocks: tab.cellLocks,
        cellPages: tab.cellPages, nestedCells: tab.nestedCells };
      for (let i = 0; i < tab.cellLocks.length; i++) {
        const lock = tab.cellLocks[i];
        if (lock && lock.masterOrigin === true) cb(live, i, lock);
      }
      if (!Array.isArray(tab.pages)) return;
      for (let k = 0; k < tab.pages.length; k++) {
        if (k === (tab.page || 0)) continue; // live arrays are authoritative for the shown page
        const snap = tab.pages[k];
        if (!snap || !Array.isArray(snap.cellLocks) || !Array.isArray(snap.cells)) continue;
        const n = snap.cells.length;
        if (!Array.isArray(snap.cellPages)) snap.cellPages = [];
        if (!Array.isArray(snap.nestedCells)) snap.nestedCells = [];
        const site = { tab: tab, live: false, cells: snap.cells, cellLocks: snap.cellLocks,
          cellPages: snap.cellPages, nestedCells: snap.nestedCells };
        for (let i = 0; i < Math.min(n, snap.cellLocks.length); i++) {
          const lock = snap.cellLocks[i];
          if (lock && lock.masterOrigin === true) cb(site, i, lock);
        }
      }
    });
  }

  /** Remap every part lock + Combined masterCellIndex through mapIndex(oldIdx) → newIdx. */
  function remapMasterIndexWith(mapIndex) {
    forEachPartLockSite(function (site, i, lock) {
      if (Number.isInteger(lock.masterCellIndex) && lock.masterCellIndex >= 0) {
        lock.masterCellIndex = mapIndex(lock.masterCellIndex);
      }
    });
    state.confirmedLinks.forEach(function (link) {
      if (Number.isInteger(link.masterCellIndex) && link.masterCellIndex >= 0) {
        link.masterCellIndex = mapIndex(link.masterCellIndex);
      }
    });
  }

  /** When Master rows reorder, keep part-tab locks + Combined masterCellIndex in sync. */
  function remapMasterCellIndices(oldToNew, cols) {
    if (!oldToNew || !cols) return;
    remapMasterIndexWith(function (idx) {
      const row = Math.floor(idx / cols);
      const col = idx % cols;
      if (row < 0 || row >= oldToNew.length) return idx;
      const next = oldToNew[row];
      if (!Number.isInteger(next) || next < 0) return idx;
      return next * cols + col;
    });
  }

  /** Master gained a column: row*oldCols+col → row*newCols+col. */
  function remapMasterIndicesAfterColumnAdd(oldCols, newCols) {
    if (!oldCols || !newCols) return;
    remapMasterIndexWith(function (idx) {
      return Math.floor(idx / oldCols) * newCols + (idx % oldCols);
    });
  }

  /** Master rows >= fromRow pushed down one. */
  function remapMasterIndicesAfterPushDown(cols, fromRow) {
    if (!cols) return;
    remapMasterIndexWith(function (idx) {
      const row = Math.floor(idx / cols);
      return row >= fromRow ? idx + cols : idx;
    });
  }

  function siteCellText(site, i) {
    const entry = site.cellPages[i];
    if (entry && Array.isArray(entry.pages) && entry.pages.length) {
      const pg = entry.page || 0;
      return entry.pages[pg] == null ? '' : String(entry.pages[pg]);
    }
    return site.cells[i] == null ? '' : String(site.cells[i]);
  }

  function masterSourceHasContent(master, idx) {
    if (!master || !Number.isInteger(idx) || idx < 0 || idx >= master.cells.length) return false;
    const entry = getCellPages(master, idx);
    for (let p = 0; p < entry.pages.length; p++) {
      if (String(entry.pages[p] || '').trim()) return true;
    }
    if ((master.cells[idx] || '').trim()) return true;
    const nests = getCellNests(master, idx);
    for (let n = 0; n < nests.length; n++) if (nestHasContent(nests[n])) return true;
    return false;
  }

  /** Copy Master[masterIdx] pages + nests into a lock site (live or stored page). */
  function writeLockSiteFromMaster(site, i, masterIdx) {
    const master = state.tabs.find(isMasterTab);
    if (!master) return false;
    const source = cloneCellPagesEntry(getCellPages(master, masterIdx));
    const live = master.cells[masterIdx] == null ? '' : String(master.cells[masterIdx]);
    if ((source.pages[source.page] || '') !== live) source.pages[source.page] = live;
    const nests = cloneNestList(getCellNests(master, masterIdx));
    const pagesSame = cellPagesContentEqual(site.cellPages[i], source) &&
      site.cellPages[i] && site.cellPages[i].page === source.page &&
      siteCellText(site, i) === (source.pages[source.page] || '');
    const nestsSame = nestsContentEqual(site.nestedCells[i], nests);
    if (pagesSame && nestsSame) return false;
    site.cellPages[i] = source;
    site.cells[i] = source.pages[source.page] == null ? '' : String(source.pages[source.page]);
    site.nestedCells[i] = nests;
    if (site.live) {
      if (!nestsSame) syncNestLinksAfterNestReplace(site.tab.id, i);
      liveSyncConfirmedLinksForCell(site.tab.id, i, null, { silent: true });
    }
    return true;
  }

  function findMasterIndexForText(text, preferCol, masterCols) {
    const master = state.tabs.find(isMasterTab);
    if (!master) return -1;
    const want = trimEndText(text);
    if (!want.trim()) return -1;
    let first = -1;
    for (let j = 0; j < master.cells.length; j++) {
      if (trimEndText(master.cells[j] || '') !== want) continue;
      if (Number.isInteger(preferCol) && masterCols && j % masterCols === preferCol) return j;
      if (first < 0) first = j;
    }
    return first;
  }

  /**
   * No locked cell may be left stale or orphaned. For every Master-origin lock
   * (all part tabs, all part pages):
   *  - source Master cell exists (in range, has content): follow it — if the cell
   *    text no longer matches, re-point to the Master cell that holds exactly this
   *    text (index drift) else re-sync pages + nests from the source (Master edited);
   *  - source missing/empty/unknown: re-link by exact text if a Master cell holds
   *    it, otherwise unlock and keep the text as a normal cell.
   * Skips the cell in an active unlock-edit session. Returns counts.
   */
  function reconcileMasterLocks(opts) {
    opts = opts || {};
    const out = { checked: 0, resynced: 0, repointed: 0, unlocked: 0 };
    const master = state.tabs.find(isMasterTab);
    if (!master) return out;
    ensureCellPages(master);
    ensureNestedCells(master);
    const unlockedLive = [];
    forEachPartLockSite(function (site, i, lock) {
      if (opts.onlyTabId && site.tab.id !== opts.onlyTabId) return;
      if (opts.onlyLive && !site.live) return;
      if (Number.isInteger(opts.onlyIndex) && (!site.live || i !== opts.onlyIndex)) return;
      if (site.live && isMasterCellEditFor(site.tab, i)) return;
      out.checked++;
      const text = siteCellText(site, i);
      const cols = site.tab.cols || 1;
      const idx = lock.masterCellIndex;
      const sourceOk = masterSourceHasContent(master, idx);
      if (sourceOk && trimEndText(master.cells[idx] || '') === trimEndText(text)) {
        if (writeLockSiteFromMaster(site, i, idx)) out.resynced++;
        return;
      }
      const match = findMasterIndexForText(text, i % cols, master.cols);
      if (match >= 0 && match !== idx) {
        lock.masterCellIndex = match;
        writeLockSiteFromMaster(site, i, match);
        out.repointed++;
        return;
      }
      if (sourceOk) {
        writeLockSiteFromMaster(site, i, idx);
        out.resynced++;
        return;
      }
      site.cellLocks[i] = null;
      out.unlocked++;
      if (site.live) unlockedLive.push({ tabId: site.tab.id, cellIndex: i });
    });
    // Orphans become normal cells: drop Master lock/origin on their Combined links.
    unlockedLive.forEach(function (u) {
      state.confirmedLinks.forEach(function (link) {
        if (link.tabId !== u.tabId || link.cellIndex !== u.cellIndex) return;
        if (!isMasterOriginLink(link)) return;
        link.masterOrigin = false;
        link.locked = false;
        if ('masterCellIndex' in link) delete link.masterCellIndex;
      });
    });
    return out;
  }

  function reportLockRepair(r) {
    if (!r || !(r.resynced || r.repointed || r.unlocked || r.relocked)) return;
    const msg = 'Repaired stale Master links: ' + reconcileSummary(r);
    pushEvent(msg, 'ok');
    scheduleSave();
  }

  function reconcileSummary(r) {
    const parts = [];
    if (r.resynced) parts.push(r.resynced + ' re-synced');
    if (r.repointed) parts.push(r.repointed + ' re-linked');
    if (r.unlocked) parts.push(r.unlocked + ' unlocked (Master source gone, text kept)');
    if (r.relocked) parts.push(r.relocked + ' re-locked (exact Master text)');
    return parts.join(', ');
  }

  /** Extra nests / other cell pages on a lock site (live or stored page snapshot). */
  function siteHasOwnExtraContent(site, index) {
    if (!site || index < 0) return false;
    const nests = (site.nestedCells && site.nestedCells[index]) || [];
    for (let n = 0; n < nests.length; n++) {
      const pages = (nests[n] && nests[n].pages) || [];
      for (let p = 0; p < pages.length; p++) {
        if (String(pages[p] || '').trim()) return true;
      }
    }
    const entry = site.cellPages && site.cellPages[index];
    if (entry && entry.pages && entry.pages.length > 1) {
      for (let p = 0; p < entry.pages.length; p++) {
        if (p === (entry.page || 0)) continue;
        if (String(entry.pages[p] || '').trim()) return true;
      }
    }
    return false;
  }

  /**
   * Exact-match Master index for a site cell (same scoring as findMasterMatchForCell,
   * but Col A comes from site.cells so stored part pages work).
   */
  function findMasterMatchForSite(site, index, text) {
    const master = state.tabs.find(isMasterTab);
    const tab = site && site.tab;
    if (!master || !tab || isMasterTab(tab)) return -1;
    const want = trimEndText(text);
    if (!want.trim()) return -1;
    const cols = master.cols || 1;
    const tabCols = tab.cols || 1;
    const targetCol = index % tabCols;
    const targetRow = Math.floor(index / tabCols);
    const cells = site.cells || tab.cells;
    const rowColA = targetCol === 0 ? '' : ((cells[targetRow * tabCols] || '').trim());
    const filter = getMasterLibPrefs(tab.id).valueFilter;
    let best = -1;
    let bestScore = -1;
    for (let i = 0; i < master.cells.length; i++) {
      if (trimEndText(master.cells[i]) !== want) continue;
      const mRow = Math.floor(i / cols);
      const mColA = (master.cells[mRow * cols] || '').trim();
      let score = 0;
      if (rowColA && mColA === rowColA) score += 4;
      if (filter === null || filter.has(mColA)) score += 2;
      if ((i % cols) === targetCol) score += 1;
      if (score > bestScore) {
        best = i;
        bestScore = score;
      }
    }
    return best;
  }

  /**
   * Visit every cell on every part page (live + stored snapshots).
   * cb(site, i) — site shape matches forEachPartLockSite.
   */
  function forEachPartCellSite(cb) {
    state.tabs.forEach(function (tab) {
      if (!tab || isMasterTab(tab)) return;
      ensureCellLocks(tab);
      ensureCellPages(tab);
      ensureNestedCells(tab);
      const live = {
        tab: tab, live: true, cells: tab.cells, cellLocks: tab.cellLocks,
        cellPages: tab.cellPages, nestedCells: tab.nestedCells
      };
      for (let i = 0; i < tab.cells.length; i++) cb(live, i);
      if (!Array.isArray(tab.pages)) return;
      for (let k = 0; k < tab.pages.length; k++) {
        if (k === (tab.page || 0)) continue;
        const snap = tab.pages[k];
        if (!snap || !Array.isArray(snap.cells)) continue;
        const n = snap.cells.length;
        if (!Array.isArray(snap.cellLocks)) snap.cellLocks = emptyCellLocks(snap.cols || tab.cols, snap.rows || 1);
        while (snap.cellLocks.length < n) snap.cellLocks.push(null);
        if (!Array.isArray(snap.cellPages)) snap.cellPages = [];
        if (!Array.isArray(snap.nestedCells)) snap.nestedCells = [];
        const site = {
          tab: tab, live: false, cells: snap.cells, cellLocks: snap.cellLocks,
          cellPages: snap.cellPages, nestedCells: snap.nestedCells
        };
        for (let i = 0; i < n; i++) cb(site, i);
      }
    });
  }

  /**
   * Startup / load: re-lock unlocked part cells whose text exactly matches a
   * Master cell (same rules as exact-match auto-link). Restores locks that
   * Duplicate-this-page used to clear. Never changes cell text / nests / pages.
   */
  function repairMissingMasterLocksFromExactMatch() {
    const out = { relocked: 0 };
    const master = state.tabs.find(isMasterTab);
    if (!master) return out;
    forEachPartCellSite(function (site, i) {
      const lock = site.cellLocks[i];
      if (lock && lock.masterOrigin === true) return;
      if (site.live && isMasterCellEditFor(site.tab, i)) return;
      const text = siteCellText(site, i);
      if (!String(text).trim()) return;
      if (siteHasOwnExtraContent(site, i)) return;
      const masterIdx = findMasterMatchForSite(site, i, text);
      if (masterIdx < 0) return;
      // Set lock only — do not rewrite text/nests/pages from Master.
      const next = { masterOrigin: true, locked: true, masterCellIndex: masterIdx };
      site.cellLocks[i] = next;
      out.relocked++;
    });
    return out;
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

  /** Last focused nest textarea (tab/cell/nest) — "current" nest for Copy current. */
  let lastFocusedNest = null;
  document.addEventListener('focusin', function (ev) {
    const t = ev.target;
    if (!t || !t.classList || !t.classList.contains('cell-nest-input')) return;
    const tab = activeTab();
    const idx = parseInt(t.dataset.idx, 10);
    const ni = parseInt(t.dataset.nest, 10);
    if (!tab || Number.isNaN(idx) || Number.isNaN(ni)) return;
    lastFocusedNest = { tabId: tab.id, cellIndex: idx, nestIndex: ni };
  });

  /**
   * Nest + → "Copy current": duplicate the current nest (last focused nest of
   * this cell, else its last nest) with all pages, insert right after it and
   * make it active. No nests yet → new nest seeded with the cell's own text.
   * Combined nest checks are copied (like Duplicate page). Locks live on the
   * cell (nests mirror Master), so the copy follows the cell's lock; on the
   * Master tab the new nest syncs to every locked part cell. One undo step.
   */
  function copyNestFromCurrent(cellIndex) {
    const tab = activeTab();
    if (!tab || cellIndex < 0 || cellIndex >= tab.cells.length) return;
    if (isCellMasterEditBlocked(tab, cellIndex)) {
      setStatus('Locked Master cell — double-click to unlock before editing', 'err');
      return;
    }
    ensureNestedCells(tab);
    const nests = tab.nestedCells[cellIndex] || [];
    let srcIdx = -1;
    if (nests.length) {
      srcIdx = nests.length - 1;
      if (lastFocusedNest && lastFocusedNest.tabId === tab.id &&
          lastFocusedNest.cellIndex === cellIndex &&
          lastFocusedNest.nestIndex >= 0 && lastFocusedNest.nestIndex < nests.length) {
        srcIdx = lastFocusedNest.nestIndex;
      }
    }
    pushHistory();
    ensureNestedCells(tab);
    const list = tab.nestedCells[cellIndex].slice();
    let insertAt;
    let copy;
    if (srcIdx >= 0) {
      copy = cloneNestList([list[srcIdx]])[0];
      insertAt = srcIdx + 1;
    } else {
      copy = { pages: [tab.cells[cellIndex] == null ? '' : String(tab.cells[cellIndex])], page: 0 };
      insertAt = 0;
    }
    list.splice(insertAt, 0, copy);
    tab.nestedCells[cellIndex] = list;

    // Shift nest links after the insert point; copy the source nest's checks.
    const scopes = Object.create(null);
    const clones = [];
    state.confirmedLinks.forEach(function (link) {
      if (link.tabId !== tab.id || link.cellIndex !== cellIndex) return;
      const ni = linkNestIndex(link);
      if (ni === null) return;
      if (ni >= insertAt) link.nestIndex = ni + 1;
      if (srcIdx >= 0 && ni === srcIdx) {
        clones.push({
          id: linkUid(),
          tabId: tab.id,
          cellIndex: cellIndex,
          nestIndex: insertAt,
          text: '',
          start: 0,
          end: 0,
          scope: link.scope,
          seq: nextLinkSeq()
        });
        scopes[link.scope] = true;
      }
    });
    clones.forEach(function (l) { state.confirmedLinks.push(l); });
    Object.keys(scopes).forEach(function (s) { regenerateCombinedScope(s); });
    liveSyncConfirmedLinksForCell(tab.id, cellIndex, undefined, { silent: true });
    if (isMasterTab(tab)) syncLockedPartNestsFromMaster(cellIndex);

    renderGrid();
    refitRowHeightAt(Math.floor(cellIndex / tab.cols));
    renderCombinedPrompt();
    applyConfirmedCellHighlights();
    applyAppendCheckedState();
    if (isMasterTab(tab)) renderMasterLibrary();
    scheduleSave();
    const nestTa = el.cellGrid.querySelector(
      'textarea.cell-nest-input[data-idx="' + cellIndex + '"][data-nest="' + insertAt + '"]'
    );
    if (nestTa) nestTa.focus();
    setStatus(srcIdx >= 0
      ? ('Added nest ' + (insertAt + 1) + ' — copy of nest ' + (srcIdx + 1) + ' (' + cellAddressFromIndex(tab, cellIndex) + ')')
      : ('Added nest — copy of ' + cellAddressFromIndex(tab, cellIndex) + ' text'));
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

  function removeNestedCell(cellIndex, nestIndex, anchor, confirmed) {
    const tab = activeTab();
    if (!tab || cellIndex < 0 || cellIndex >= tab.cells.length) return;
    if (isCellMasterEditBlocked(tab, cellIndex)) {
      setStatus('Locked Master cell — double-click to unlock before editing', 'err');
      return;
    }
    ensureNestedCells(tab);
    const nests = tab.nestedCells[cellIndex];
    if (nestIndex < 0 || nestIndex >= nests.length) return;
    if (!confirmed) {
      const pages = (nests[nestIndex] && nests[nestIndex].pages) || [];
      const withText = pages.filter(function (pg) { return String(pg || '').trim(); }).length;
      if (withText > 0) {
        openActionConfirm({
          anchor: anchor,
          message: 'Delete nest ' + (nestIndex + 1) + ' (' + pluralCells(withText, 'page') + ' with text)?',
          actionLabel: 'Delete',
          cancelStatus: 'Delete nest cancelled',
          ariaLabel: 'Confirm delete nest',
          onConfirm: function () {
            if (activeTab() !== tab) return;
            removeNestedCell(cellIndex, nestIndex, anchor, true);
          }
        });
        return;
      }
    }
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
    liveSyncConfirmedLinksForCell(tab.id, idx, nestIdx, { dropEmpty: true });
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

  function removeNestPage(cellIndex, nestIndex, anchor, confirmed) {
    const tab = activeTab();
    if (!tab || cellIndex < 0 || cellIndex >= tab.cells.length) return;
    if (isCellMasterEditBlocked(tab, cellIndex)) {
      setStatus('Locked Master cell — double-click to unlock before editing', 'err');
      return;
    }
    ensureNestedCells(tab);
    const nest = tab.nestedCells[cellIndex][nestIndex];
    if (!nest || nest.pages.length <= 1) {
      removeNestedCell(cellIndex, nestIndex, anchor);
      return;
    }
    if (!confirmed && String(nest.pages[nest.page || 0] || '').trim()) {
      openActionConfirm({
        anchor: anchor,
        message: 'Delete nest page ' + ((nest.page || 0) + 1) + ' (has text)?',
        actionLabel: 'Delete',
        cancelStatus: 'Delete page cancelled',
        ariaLabel: 'Confirm delete page',
        onConfirm: function () {
          if (activeTab() !== tab) return;
          removeNestPage(cellIndex, nestIndex, anchor, true);
        }
      });
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
    renderCombinedPrompt();
    refitRowHeightAt(Math.floor(cellIndex / tab.cols));
    scheduleSave();
    const ta = el.cellGrid.querySelector('textarea.cell[data-idx="' + cellIndex + '"]');
    if (ta) ta.focus();
    setStatus(copyMode
      ? ('Added cell page — copy of current (' + entry.pages.length + ')')
      : ('Added cell page — blank (' + entry.pages.length + ')'));
  }

  function removeCellPage(cellIndex, anchor, confirmed) {
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
    if (!confirmed && String(entry.pages[entry.page || 0] || '').trim()) {
      openActionConfirm({
        anchor: anchor,
        message: 'Delete page ' + ((entry.page || 0) + 1) + ' of ' +
          cellAddress(Math.floor(cellIndex / tab.cols), cellIndex % tab.cols) + ' (has text)?',
        actionLabel: 'Delete',
        cancelStatus: 'Delete page cancelled',
        ariaLabel: 'Confirm delete page',
        onConfirm: function () {
          if (activeTab() !== tab) return;
          removeCellPage(cellIndex, anchor, true);
        }
      });
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
      if (Number.isFinite(link.seq)) item.seq = link.seq;
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
      if (typeof item.seq === 'number' && Number.isFinite(item.seq)) entry.seq = item.seq;
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

  /**
   * Deep copy a page snapshot (cells, nests, cell pages, shades, Combined keys,
   * Master cellLocks). opts.unlock (legacy) clears locks — Duplicate no longer
   * uses it so copies stay real Master links.
   */
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
    // Default scope = the CURRENT part page only: drop this tab's live checks
    // and re-add the page's stored keys (keys, not text — Combined regenerates).
    // NOTE(all-pages): see regenerateCombinedScope for how to include every page.
    removeTabCombinedContribution(tab.id);
    const scope = currentPromptScope();
    const keys = normalizeTabConfirmedKeys(snap.confirmed);
    keys.sort(function (a, b) {
      if (a.cellIndex !== b.cellIndex) return a.cellIndex - b.cellIndex;
      const an = Number.isInteger(a.nestIndex) ? a.nestIndex : -1;
      const bn = Number.isInteger(b.nestIndex) ? b.nestIndex : -1;
      return an - bn;
    });
    for (let i = 0; i < keys.length; i++) {
      const item = keys[i];
      if (item.cellIndex >= tab.cells.length) continue;
      const nest = Number.isInteger(item.nestIndex) ? item.nestIndex : null;
      const exists = nest === null
        ? isCellConfirmedInScope(tab.id, item.cellIndex, scope)
        : isNestConfirmedInScope(tab.id, item.cellIndex, nest, scope);
      if (exists) continue;
      const link = {
        id: linkUid(),
        tabId: tab.id,
        cellIndex: item.cellIndex,
        text: '',
        start: 0,
        end: 0,
        scope: scope,
        seq: Number.isFinite(item.seq) ? item.seq : nextLinkSeq()
      };
      if (nest !== null) link.nestIndex = nest;
      state.confirmedLinks.push(link);
    }
    syncLinkSeqCounter();
    regenerateCombinedScope(scope);
  }

  function flushLiveCellInputs(tab) {
    if (!tab || !el.cellGrid) return;
    const tas = el.cellGrid.querySelectorAll('textarea.cell[data-idx]');
    for (let i = 0; i < tas.length; i++) {
      const ta = tas[i];
      const idx = parseInt(ta.dataset.idx, 10);
      if (Number.isNaN(idx) || idx < 0 || idx >= tab.cells.length) continue;
      // Only the focused textarea can hold a live user edit (onCellInput syncs
      // every keystroke). Others may lag a programmatic model write that has
      // not re-rendered yet (Master insert, Replace, clear) — model wins there.
      if (document.activeElement === ta && (ta.value || '') !== (tab.cells[idx] || '')) {
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
    reportLockRepair(reconcileMasterLocks({ onlyTabId: tab.id, onlyLive: true }));
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

  /**
   * Add a part page right after page `afterIndex` (default: the current page)
   * and switch to it. mode 'copy' = full copy including Master locks, else blank with
   * that page's grid size. Undoable (Ctrl+Z). Right-click a page chip to use.
   */
  function addTabPage(mode, afterIndex) {
    const tab = activeTab();
    if (!tab || isMasterTab(tab) || toolsTabActive) return;
    const copyMode = mode === 'copy';
    ensureTabPages(tab);
    pushHistory();
    flushLiveCellInputs(tab);
    tab.pages[tab.page] = captureTabPage(tab);
    const srcIndex = Number.isInteger(afterIndex) && afterIndex >= 0 && afterIndex < tab.pages.length
      ? afterIndex : (tab.page || 0);
    const src = tab.pages[srcIndex];
    let nextPage;
    if (copyMode) {
      // Keep Master cellLocks so duplicated cells stay true Master links
      // (sync when Master changes; same lock / Clear / overwrite rules).
      nextPage = cloneTabPage(src);
    } else {
      nextPage = makeBlankTabPage({ cols: src.cols, rows: src.rows, columnWidths: src.columnWidths });
    }
    if (!nextPage) return;
    // Fresh page (no name → chip shows its live position number).
    nextPage.name = '';
    const insertAt = srcIndex + 1;
    tab.pages.splice(insertAt, 0, nextPage);
    tab.page = insertAt;
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
      ? ('Added part page ' + (insertAt + 1) + ' — copy of page ' + (srcIndex + 1) + ' (Master locks kept)')
      : ('Added part page ' + (insertAt + 1) + ' — blank, after page ' + (srcIndex + 1)));
  }

  /**
   * Remove part page `index` (default: the current page). Removing the current
   * page selects the page that slides into its slot (or the new last page);
   * removing another page keeps the current page selected with its live edits.
   */
  function removeTabPage(index, anchor, confirmed) {
    const tab = activeTab();
    if (!tab || isMasterTab(tab) || toolsTabActive) return;
    ensureTabPages(tab);
    if (tab.pages.length <= 1) {
      setStatus('Part already has only one page', 'err');
      return;
    }
    const cur = tab.page || 0;
    const removeAt = Number.isInteger(index) ? index : cur;
    if (removeAt < 0 || removeAt >= tab.pages.length) return;
    if (!confirmed) {
      flushLiveCellInputs(tab);
      let filled = 0;
      if (removeAt === cur) {
        for (let i = 0; i < tab.cells.length; i++) if (cellHasExportableContent(tab, i)) filled++;
      } else {
        const snap = tab.pages[removeAt] || {};
        const cells = Array.isArray(snap.cells) ? snap.cells : [];
        const nested = Array.isArray(snap.nestedCells) ? snap.nestedCells : [];
        for (let i = 0; i < cells.length; i++) {
          let has = !!String(cells[i] || '').trim();
          if (!has) {
            const ns = nested[i] || [];
            for (let n = 0; n < ns.length && !has; n++) has = nestHasContent(ns[n]);
          }
          if (has) filled++;
        }
      }
      if (filled > 0) {
        openActionConfirm({
          anchor: anchor,
          message: 'Delete page ' + (removeAt + 1) + ' (' + pluralCells(filled) + ')?',
          actionLabel: 'Delete',
          cancelStatus: 'Delete page cancelled',
          ariaLabel: 'Confirm delete page',
          onConfirm: function () {
            if (activeTab() !== tab) return;
            removeTabPage(removeAt, anchor, true);
          }
        });
        return;
      }
    }
    pushHistory();
    flushLiveCellInputs(tab);
    if (removeAt !== cur) {
      // Save the live grid into its page before shifting indices.
      tab.pages[cur] = captureTabPage(tab);
    }
    tab.pages.splice(removeAt, 1);
    if (removeAt < cur) tab.page = cur - 1;
    else if (removeAt === cur && tab.page >= tab.pages.length) tab.page = tab.pages.length - 1;
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
    setStatus('Removed part page ' + (removeAt + 1) + ' (' + tab.pages.length + ' left)');
  }

  /* Right-click a page chip → small menu (like the tab right-click menu). */
  let partPageMenuEl = null;
  function closePartPageMenu() {
    if (partPageMenuEl && partPageMenuEl.parentNode) partPageMenuEl.parentNode.removeChild(partPageMenuEl);
    partPageMenuEl = null;
    document.removeEventListener('pointerdown', onPartPageMenuOutside, true);
    document.removeEventListener('keydown', onPartPageMenuKey, true);
    window.removeEventListener('blur', closePartPageMenu);
  }
  function onPartPageMenuOutside(e) {
    if (partPageMenuEl && !partPageMenuEl.contains(e.target)) closePartPageMenu();
  }
  function onPartPageMenuKey(e) {
    if (e.key === 'Escape') {
      e.preventDefault();
      e.stopPropagation();
      closePartPageMenu();
    }
  }
  function openPartPageMenu(tab, pageIndex, x, y) {
    ensureTabPages(tab);
    const count = tab.pages.length;
    openPageCtxMenu({
      heading: tabPageDisplayName(tab.pages[pageIndex], pageIndex) + ' (' + (pageIndex + 1) + '/' + count + ')',
      leading: [{
        label: 'Rename',
        title: 'Rename this part page (Enter saves, Esc cancels)',
        onClick: function () { startPartPageRename(tab, pageIndex); }
      }, {
        label: 'Add page after this',
        title: 'Insert a blank page right after this one and go to it (Ctrl+Z to undo)',
        onClick: function () { addTabPage('blank', pageIndex); }
      }, {
        label: 'Duplicate this page',
        title: 'Insert a copy of this page (keeps Master locks) right after it and go to it (Ctrl+Z to undo)',
        onClick: function () { addTabPage('copy', pageIndex); }
      }],
      disabled: count <= 1,
      disabledTitle: 'Only one page — the last page cannot be deleted',
      enabledTitle: 'Delete this part page (Ctrl+Z to undo)',
      onDelete: function () { removeTabPage(pageIndex, { x: x, y: y }); },
      x: x,
      y: y
    });
  }

  /** Shared small page menu (part-page chips, cell/nest page labels): heading + Delete page (+ extra). */
  function openPageCtxMenu(opts) {
    const items = (opts.leading || []).slice();
    items.push({
      label: 'Delete page',
      danger: opts.pageDanger !== false,
      disabled: !!opts.disabled,
      title: opts.disabled ? opts.disabledTitle : opts.enabledTitle,
      onClick: opts.onDelete
    });
    if (opts.extra) {
      items.push({
        label: opts.extra.label,
        danger: true,
        disabled: !!opts.extra.disabled,
        title: opts.extra.title || '',
        onClick: opts.extra.onClick
      });
    }
    openCtxMenu({ heading: opts.heading, items: items, x: opts.x, y: opts.y });
  }

  /**
   * Generic small context menu (tab right-click menu look).
   * items: [{ label, onClick, disabled, danger, current, swatch, title }]
   */
  function openCtxMenu(opts) {
    closePartPageMenu();
    closeTabIconPicker();
    const pop = document.createElement('div');
    pop.className = 'tab-icon-picker part-page-menu' + (opts.className ? ' ' + opts.className : '');
    pop.setAttribute('role', 'menu');
    const heading = document.createElement('div');
    heading.className = 'tab-icon-picker-heading';
    heading.textContent = opts.heading;
    pop.appendChild(heading);
    let firstEnabled = null;
    (opts.items || []).forEach(function (item) {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'tab-icon-picker-action' + (item.danger ? ' is-danger' : '') +
        (item.current ? ' is-current' : '');
      btn.setAttribute('role', 'menuitem');
      if (item.swatch) {
        const sw = document.createElement('span');
        sw.className = 'ctx-swatch ctx-swatch-' + item.swatch;
        sw.setAttribute('aria-hidden', 'true');
        btn.appendChild(sw);
      }
      const label = document.createElement('span');
      label.className = 'ctx-label';
      label.textContent = item.label;
      btn.appendChild(label);
      if (item.current) {
        const mark = document.createElement('span');
        mark.className = 'ctx-check';
        mark.textContent = '✓';
        mark.setAttribute('aria-label', 'current');
        btn.appendChild(mark);
      }
      btn.disabled = !!item.disabled;
      if (item.title) btn.title = item.title;
      btn.addEventListener('click', function (ev) {
        ev.preventDefault();
        ev.stopPropagation();
        closePartPageMenu();
        if (!item.disabled && typeof item.onClick === 'function') item.onClick();
      });
      pop.appendChild(btn);
      if (!firstEnabled && !btn.disabled) firstEnabled = btn;
    });
    document.body.appendChild(pop);
    partPageMenuEl = pop;
    const pad = 6;
    const r = pop.getBoundingClientRect();
    const left = Math.max(pad, Math.min(opts.x, window.innerWidth - r.width - pad));
    const top = Math.max(pad, Math.min(opts.y, window.innerHeight - r.height - pad));
    pop.style.left = left + 'px';
    pop.style.top = top + 'px';
    document.addEventListener('pointerdown', onPartPageMenuOutside, true);
    document.addEventListener('keydown', onPartPageMenuKey, true);
    window.addEventListener('blur', closePartPageMenu);
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

  /** Rename part page `pageIndex` of `tab` (empty = back to its number). Page + scroll stay put. */
  function renameTabPage(tab, pageIndex, name) {
    if (!tab || isMasterTab(tab)) return;
    ensureTabPages(tab);
    if (!tab.pages[pageIndex]) return;
    const next = typeof name === 'string' ? name.trim() : '';
    if ((tab.pages[pageIndex].name || '') === next) return;
    pushHistory();
    // Re-read after pushHistory (never hold a page object across it).
    const page = tab.pages[pageIndex];
    if (!page) return;
    page.name = next;
    if (tab === activeTab() && !toolsTabActive) {
      const oldStrip = el.partTabPageChrome && el.partTabPageChrome.querySelector('.part-page-strip');
      const keepScroll = oldStrip ? oldStrip.scrollLeft : 0;
      renderPartTabPageChrome();
      const strip = el.partTabPageChrome && el.partTabPageChrome.querySelector('.part-page-strip');
      if (strip) {
        strip.scrollLeft = keepScroll;
        requestAnimationFrame(function () { strip.scrollLeft = keepScroll; });
      }
    }
    scheduleSave();
    setStatus(next ? 'Renamed part page ' + (pageIndex + 1) + ' → ' + next
      : 'Part page ' + (pageIndex + 1) + ' name cleared', 'ok');
  }

  /**
   * Right-click → Rename on a page chip: a small input floats over the chip
   * (fixed position, nothing in the row moves). Enter/blur save, Esc cancels.
   */
  let partPageRenameEl = null;
  function startPartPageRename(tab, pageIndex) {
    if (partPageRenameEl) partPageRenameEl.blur();
    const host = el.partTabPageChrome;
    const chip = host && host.querySelector('.part-page-chip[data-page="' + pageIndex + '"]');
    if (!chip || !tab || !tab.pages || !tab.pages[pageIndex]) return;
    const snap = tab.pages[pageIndex];
    const rawName = typeof snap.name === 'string' ? snap.name.trim() : '';
    const r = chip.getBoundingClientRect();
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'part-page-rename-input';
    input.value = rawName && !/^page \d+$/i.test(rawName) ? rawName : '';
    input.placeholder = defaultTabPageName(pageIndex);
    input.spellcheck = false;
    input.setAttribute('aria-label', 'Rename part page ' + (pageIndex + 1));
    const w = Math.max(120, Math.ceil(r.width));
    const h = 22;
    input.style.width = w + 'px';
    input.style.left = Math.max(4, Math.min(r.left, window.innerWidth - w - 4)) + 'px';
    input.style.top = Math.round(r.top + r.height / 2 - h / 2) + 'px';
    document.body.appendChild(input);
    partPageRenameEl = input;
    input.focus();
    input.select();
    let done = false;
    function finish(commit) {
      if (done) return;
      done = true;
      const value = input.value;
      if (input.parentNode) input.parentNode.removeChild(input);
      if (partPageRenameEl === input) partPageRenameEl = null;
      if (commit) renameTabPage(tab, pageIndex, value);
    }
    input.addEventListener('keydown', function (ev) {
      if (ev.key === 'Enter') {
        ev.preventDefault();
        ev.stopPropagation();
        finish(true);
      } else if (ev.key === 'Escape') {
        ev.preventDefault();
        ev.stopPropagation();
        finish(false);
      }
    });
    input.addEventListener('blur', function () { finish(true); });
    input.addEventListener('mousedown', function (ev) { ev.stopPropagation(); });
    input.addEventListener('click', function (ev) { ev.stopPropagation(); });
  }

  /**
   * Inline page chips on the page-nav row: fixed-width strip (always shown,
   * even with one page) — click = jump, drag = reorder.
   */
  function buildPartPageStrip(tab, page) {
    const strip = document.createElement('div');
    strip.className = 'part-page-strip';
    strip.setAttribute('role', 'tablist');
    strip.setAttribute('aria-label', 'Part pages (drag to reorder)');
    const count = tab.pages.length;
    function clearDrop() {
      const marked = strip.querySelectorAll('.drop-before, .drop-after');
      for (let i = 0; i < marked.length; i++) marked[i].classList.remove('drop-before', 'drop-after');
    }
    for (let i = 0; i < count; i++) {
      const snap = tab.pages[i];
      // Auto names ("Page N" from older + adds) show the live position; custom names show as typed.
      const rawName = snap && typeof snap.name === 'string' ? snap.name.trim() : '';
      const named = rawName && !/^page \d+$/i.test(rawName);
      const chip = document.createElement('button');
      chip.type = 'button';
      chip.className = 'part-page-chip' + (i === page ? ' is-active' : '');
      chip.setAttribute('role', 'tab');
      chip.setAttribute('aria-selected', i === page ? 'true' : 'false');
      chip.textContent = named ? rawName : String(i + 1);
      chip.title = (named ? rawName + ' — ' : '') + 'page ' + (i + 1) + ' of ' + count +
        (count > 1 ? ' (drag to reorder; right-click to rename, add a page after, or delete)' : ' (right-click to rename or add a page after)');
      chip.dataset.page = String(i);
      chip.draggable = count > 1;
      chip.addEventListener('click', function (ev) {
        ev.preventDefault();
        if (i !== (tab.page || 0)) setTabPage(i);
      });
      chip.addEventListener('contextmenu', function (ev) {
        ev.preventDefault();
        ev.stopPropagation();
        openPartPageMenu(tab, i, ev.clientX, ev.clientY);
      });
      chip.addEventListener('dragstart', function (ev) {
        ev.stopPropagation();
        ev.dataTransfer.setData('application/x-c2c-part-page', String(i));
        ev.dataTransfer.effectAllowed = 'move';
        chip.classList.add('dragging');
      });
      chip.addEventListener('dragend', function () {
        chip.classList.remove('dragging');
        clearDrop();
      });
      chip.addEventListener('dragover', function (ev) {
        if (!ev.dataTransfer.types || ev.dataTransfer.types.indexOf('application/x-c2c-part-page') === -1) return;
        ev.preventDefault();
        ev.dataTransfer.dropEffect = 'move';
        clearDrop();
        const b = chip.getBoundingClientRect();
        chip.classList.add(ev.clientX < b.left + b.width / 2 ? 'drop-before' : 'drop-after');
      });
      chip.addEventListener('dragleave', function () {
        chip.classList.remove('drop-before', 'drop-after');
      });
      chip.addEventListener('drop', function (ev) {
        const raw = ev.dataTransfer.getData('application/x-c2c-part-page');
        if (raw === '') return;
        ev.preventDefault();
        ev.stopPropagation();
        const b = chip.getBoundingClientRect();
        clearDrop();
        reorderTabPage(parseInt(raw, 10), i, ev.clientX >= b.left + b.width / 2);
      });
      strip.appendChild(chip);
    }
    // Wheel scrolls the chip strip sideways when it overflows.
    strip.addEventListener('wheel', function (ev) {
      if (strip.scrollWidth <= strip.clientWidth) return;
      if (Math.abs(ev.deltaY) > Math.abs(ev.deltaX)) {
        strip.scrollLeft += ev.deltaY;
        ev.preventDefault();
      }
    }, { passive: false });
    requestAnimationFrame(function () {
      const active = strip.querySelector('.part-page-chip.is-active');
      if (active && strip.scrollWidth > strip.clientWidth) {
        const left = active.offsetLeft - strip.offsetLeft;
        if (left < strip.scrollLeft) strip.scrollLeft = left;
        else if (left + active.offsetWidth > strip.scrollLeft + strip.clientWidth) {
          strip.scrollLeft = left + active.offsetWidth - strip.clientWidth;
        }
      }
    });
    return strip;
  }

  /** Move part page `from` before/after `target`; page data moves, current page stays selected. */
  function reorderTabPage(from, target, afterTarget) {
    const tab = activeTab();
    if (!tab || isMasterTab(tab) || toolsTabActive) return;
    ensureTabPages(tab);
    const n = tab.pages.length;
    if (!Number.isInteger(from) || from < 0 || from >= n || target < 0 || target >= n) return;
    let insertAt = target + (afterTarget ? 1 : 0);
    if (from < insertAt) insertAt--;
    if (insertAt === from) return;
    pushHistory();
    // Flush the live grid into its page so the move carries current edits.
    flushLiveCellInputs(tab);
    tab.pages[tab.page] = captureTabPage(tab);
    const current = tab.pages[tab.page];
    const moved = tab.pages.splice(from, 1)[0];
    tab.pages.splice(insertAt, 0, moved);
    tab.page = tab.pages.indexOf(current);
    renderPartTabPageChrome();
    scheduleSave();
    setStatus('Moved part page ' + (from + 1) + ' → ' + (insertAt + 1), 'ok');
  }

  function renderPartTabPageChrome() {
    const host = el.partTabPageChrome;
    if (!host) return;
    const tab = activeTab();
    const show = !!(tab && !isMasterTab(tab) && !toolsTabActive);
    host.hidden = !show;
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

    chrome.appendChild(prevBtn);
    chrome.appendChild(label);
    chrome.appendChild(nextBtn);
    // No + here: right-click a page chip → Add page after this / Duplicate this page.

    host.appendChild(chrome);
    host.appendChild(buildPartPageStrip(tab, page));
  }



  function closeCellShadePicker() {
    if (partPageMenuEl && partPageMenuEl.classList.contains('cell-shade-menu')) closePartPageMenu();
  }

  /** Target cells for a shade action: the multi-cell range if the click is inside it, else the cell. */
  function shadeTargetIndices(tab, cellIndex) {
    const b = getStickyCellRange();
    const r = Math.floor(cellIndex / tab.cols);
    const c = cellIndex % tab.cols;
    if (b && r >= b.rMin && r <= b.rMax && c >= b.cMin && c <= b.cMax) {
      const out = [];
      for (let rr = b.rMin; rr <= b.rMax; rr++) {
        for (let cc = b.cMin; cc <= b.cMax; cc++) {
          const i = rr * tab.cols + cc;
          if (i >= 0 && i < tab.cells.length) out.push(i);
        }
      }
      return out;
    }
    return [cellIndex];
  }

  function applyShadeToCells(tab, indices, shade) {
    pushHistory();
    indices.forEach(function (i) { setCellShade(tab, i, shade); });
    indices.forEach(function (i) {
      const ta = el.cellGrid && el.cellGrid.querySelector('textarea.cell[data-idx="' + i + '"]');
      const wrapEl = ta && ta.closest ? ta.closest('.cell-wrap') : null;
      if (!wrapEl) return;
      wrapEl.classList.remove('cell-shade-green', 'cell-shade-yellow', 'cell-shade-red');
      if (shade) wrapEl.classList.add('cell-shade-' + shade);
    });
    scheduleSave();
    const what = indices.length === 1 ? 'Cell' : indices.length + ' cells';
    setStatus(shade ? (what + ' shaded ' + shade) : (what + ' shading cleared'), 'ok');
  }

  /**
   * Right-click Clear: empty current-page text + unlink Master lock on part tabs.
   * Never writes Master from a part-tab Clear. On the Master tab, empties like a
   * normal edit (existing sync pushes blank to locked linked part cells).
   * Nest textareas are not covered by this menu (nests keep default browser menu).
   */
  function cellHasClearableNestContent(tab, idx) {
    ensureNestedCells(tab);
    const nests = tab.nestedCells[idx] || [];
    for (let n = 0; n < nests.length; n++) {
      if (nestHasContent(nests[n])) return true;
    }
    return false;
  }

  /**
   * Empty cells (current page text) + unlink Master on part tabs. opts.clearNests
   * also drops nest lists (row Clear). Keeps shades + rowHeights. One undo step.
   */
  function clearCellsFromMenu(tab, indices, opts) {
    opts = opts || {};
    const clearNests = !!opts.clearNests;
    if (!tab || !indices || !indices.length) return;
    const uniq = [];
    const seen = {};
    indices.forEach(function (i) {
      if (!Number.isInteger(i) || i < 0 || i >= tab.cells.length || seen[i]) return;
      seen[i] = true;
      uniq.push(i);
    });
    if (!uniq.length) return;

    let anyWork = false;
    for (let i = 0; i < uniq.length; i++) {
      const idx = uniq[i];
      if ((tab.cells[idx] || '') !== '') { anyWork = true; break; }
      if (!isMasterTab(tab) && getCellLock(tab, idx)) { anyWork = true; break; }
      if (clearNests && cellHasClearableNestContent(tab, idx)) { anyWork = true; break; }
    }
    if (!anyWork) {
      setStatus('Nothing to clear', 'ok');
      return;
    }

    if (opts.confirm) {
      // v0.157: ask only when a cell would actually lose text (cell or nests).
      let filled = 0;
      for (let i = 0; i < uniq.length; i++) {
        const idx = uniq[i];
        if (String(tab.cells[idx] || '').trim() ||
            (clearNests && cellHasClearableNestContent(tab, idx))) filled++;
      }
      if (filled > 0) {
        const dims = tab.cols + 'x' + tab.rows;
        const again = Object.assign({}, opts, { confirm: false });
        openActionConfirm({
          anchor: opts.anchor,
          message: 'Clear ' + pluralCells(filled) + (opts.where || '') + '?' + linkedPartSuffix(tab, uniq),
          actionLabel: 'Clear',
          cancelStatus: 'Clear cancelled',
          ariaLabel: 'Confirm clear',
          onConfirm: function () {
            if (activeTab() !== tab || (tab.cols + 'x' + tab.rows) !== dims) {
              setStatus('Clear cancelled — grid changed', 'err');
              return;
            }
            clearCellsFromMenu(tab, uniq, again);
          }
        });
        return;
      }
    }

    // Drop an in-progress Master unlock session on any of these cells (no dialog).
    if (masterSegmentEdit && masterSegmentEdit.kind === 'cell' &&
        masterSegmentEdit.tabId === tab.id &&
        uniq.indexOf(masterSegmentEdit.cellIndex) >= 0) {
      masterSegmentEdit = null;
      hideMasterSegmentDialog();
    }

    pushHistory();
    ensureCellPages(tab);
    if (clearNests) ensureNestedCells(tab);
    for (let i = 0; i < uniq.length; i++) {
      const idx = uniq[i];
      writeCellCurrentPage(tab, idx, '');
      if (clearNests) {
        tab.nestedCells[idx] = [];
        syncNestLinksAfterNestReplace(tab.id, idx);
      }
      if (!isMasterTab(tab)) {
        clearCellMasterLock(tab, idx);
        // Drop Master origin on Combined links for this cell (parent + nests).
        state.confirmedLinks.forEach(function (other) {
          if (other.tabId !== tab.id || other.cellIndex !== idx) return;
          if (!isMasterOriginLink(other)) return;
          other.masterOrigin = false;
          other.locked = false;
          if ('masterCellIndex' in other) delete other.masterCellIndex;
        });
      }
      liveSyncConfirmedLinksForCell(tab.id, idx, undefined, { dropEmpty: true, silent: true });
      if (isMasterTab(tab)) syncLockedPartPagesFromMaster(idx);
    }
    if (clearNests) {
      // Nest DOM must rebuild; keep sticky highlight / scroll via normal render.
      renderGrid();
    } else if (el.cellGrid) {
      for (let i = 0; i < uniq.length; i++) {
        const idx = uniq[i];
        const ta = el.cellGrid.querySelector('textarea.cell[data-idx="' + idx + '"]');
        if (!ta) continue;
        ta.value = '';
        ta.readOnly = false;
        const confirmed = isCellConfirmed(tab.id, idx);
        ta.classList.toggle('cell-confirmed', confirmed);
        const wrap = ta.closest ? ta.closest('.cell-wrap') : null;
        if (wrap) {
          wrap.classList.toggle('cell-confirmed', confirmed);
          wrap.classList.remove('master-cell-locked', 'master-cell-editing');
        }
      }
    }
    renderCombinedPrompt();
    applyConfirmedCellHighlights();
    applyAppendCheckedState();
    if (getStickyCellRange()) restoreStickyCellRangeHighlight();
    if (isMasterTab(tab)) renderMasterLibrary();
    scheduleSave();
    if (opts.statusText) {
      setStatus(opts.statusText, 'ok');
    } else {
      const what = uniq.length === 1 ? 'Cell cleared' : (uniq.length + ' cells cleared');
      setStatus(isMasterTab(tab) ? what : (what + ' (unlinked from Master)'), 'ok');
    }
  }

  /** Right-click row number → Clear this row (current part page only). */
  function openRowClearMenu(rowIndex, x, y) {
    const tab = activeTab();
    if (!tab || rowIndex < 0 || rowIndex >= tab.rows) return;
    closeAddPageChoiceMenu();
    closeCellShadePicker();
    const indices = [];
    for (let c = 0; c < tab.cols; c++) {
      indices.push(rowIndex * tab.cols + c);
    }
    const anyClearable = indices.some(function (idx) {
      if ((tab.cells[idx] || '') !== '') return true;
      if (!isMasterTab(tab) && getCellLock(tab, idx)) return true;
      return cellHasClearableNestContent(tab, idx);
    });
    const heading = 'Row ' + (rowIndex + 1);
    openCtxMenu({
      heading: heading,
      className: 'row-clear-menu',
      x: x,
      y: y,
      items: [{
        label: 'Clear this row',
        title: isMasterTab(tab)
          ? 'Empty every cell in this Master row (linked part cells may sync)'
          : 'Empty every cell in this row and unlink from Master (Master library unchanged)',
        disabled: !anyClearable,
        onClick: function () {
          clearCellsFromMenu(tab, indices, {
            clearNests: true,
            confirm: true,
            anchor: rowNumberAt(rowIndex) || { x: x, y: y },
            where: ' in row ' + (rowIndex + 1),
            statusText: isMasterTab(tab)
              ? ('Row ' + (rowIndex + 1) + ' cleared')
              : ('Row ' + (rowIndex + 1) + ' cleared (unlinked from Master)')
          });
        }
      }, {
        label: 'Delete this row',
        title: isMasterTab(tab)
          ? 'Remove this Master row; rows below move up (part cells linked to it keep their text, unlocked)'
          : 'Remove this row; rows below move up one (empty row added at the bottom)',
        onClick: function () { deleteRowFromMenu(tab, rowIndex, rowNumberAt(rowIndex) || { x: x, y: y }); }
      }]
    });
  }

  /**
   * Row-number menu → Delete this row. Removes `rowIndex` on the current page;
   * rows below shift up one (reorderRowInTab remap: cells, nests, locks, cell
   * pages, shades, rowHeights, Combined links, Master→part indices). Row count
   * stays the same — an empty row is left at the bottom. Deleted row's Combined
   * checks are dropped. Part tab never touches Master. Master tab: part cells
   * linked to the deleted Master row keep their text and are unlocked (same as
   * stale-lock repair); links to rows below are remapped up. One undo step.
   */
  function deleteRowFromMenu(tab, rowIndex, anchor, confirmed) {
    if (!tab || tab !== activeTab() || rowIndex < 0 || rowIndex >= tab.rows) return;
    const cols = tab.cols;
    const rowStart = rowIndex * cols;
    const rowEnd = rowStart + cols; // exclusive
    if (!confirmed) {
      // v0.157: confirm only when a cell in the row has text (any cell page or nest).
      ensureNestedCells(tab);
      ensureCellPages(tab);
      let filled = 0;
      for (let i = rowStart; i < rowEnd; i++) {
        if (String(tab.cells[i] || '').trim() || siteHasOwnExtraContent(tab, i)) filled++;
      }
      if (filled > 0) {
        const dims = tab.cols + 'x' + tab.rows;
        const idxs = [];
        for (let i = rowStart; i < rowEnd; i++) idxs.push(i);
        openActionConfirm({
          anchor: anchor,
          message: 'Delete row ' + (rowIndex + 1) + ' (' + pluralCells(filled) + ')?' +
            linkedPartSuffix(tab, idxs, 'unlock, text kept'),
          actionLabel: 'Delete',
          cancelStatus: 'Delete row cancelled',
          ariaLabel: 'Confirm delete row',
          onConfirm: function () {
            if (activeTab() !== tab || (tab.cols + 'x' + tab.rows) !== dims) {
              setStatus('Delete row cancelled — grid changed', 'err');
              return;
            }
            deleteRowFromMenu(tab, rowIndex, anchor, true);
          }
        });
        return;
      }
    }
    // Drop an in-progress Master unlock session on this row (no dialog).
    if (masterSegmentEdit && masterSegmentEdit.kind === 'cell' &&
        masterSegmentEdit.tabId === tab.id &&
        masterSegmentEdit.cellIndex >= rowStart && masterSegmentEdit.cellIndex < rowEnd) {
      masterSegmentEdit = null;
      hideMasterSegmentDialog();
    }
    pushHistory();
    ensureNestedCells(tab);
    ensureCellLocks(tab);
    ensureCellPages(tab);
    ensureCellShades(tab);

    // 1) Drop Combined checks on the deleted row (all scopes).
    const dead = state.confirmedLinks.filter(function (link) {
      return link.tabId === tab.id && link.cellIndex >= rowStart && link.cellIndex < rowEnd;
    });
    if (dead.length) removeLinksFromCombined(dead);

    // 2) Master tab: unlink part cells that point at the deleted Master row.
    let unlocked = 0;
    if (isMasterTab(tab)) {
      const inDeleted = function (idx) {
        return Number.isInteger(idx) && idx >= rowStart && idx < rowEnd;
      };
      const liveUnlocked = [];
      forEachPartLockSite(function (site, i, lock) {
        if (!inDeleted(lock.masterCellIndex)) return;
        site.cellLocks[i] = null;
        unlocked++;
        if (site.live) liveUnlocked.push({ tabId: site.tab.id, cellIndex: i });
      });
      state.confirmedLinks.forEach(function (link) {
        const hitLive = liveUnlocked.some(function (u) {
          return u.tabId === link.tabId && u.cellIndex === link.cellIndex;
        });
        if (!hitLive && !inDeleted(link.masterCellIndex)) return;
        if (!isMasterOriginLink(link) && !inDeleted(link.masterCellIndex)) return;
        link.masterOrigin = false;
        link.locked = false;
        if ('masterCellIndex' in link) delete link.masterCellIndex;
      });
    }

    // 3) Empty the row, then move it to the bottom (rows below shift up).
    for (let c = 0; c < cols; c++) tab.cells[rowStart + c] = '';
    clearNestedRow(tab, rowIndex);
    clearCellLockRow(tab, rowIndex);
    clearCellPagesRow(tab, rowIndex);
    clearCellShadeRow(tab, rowIndex);
    const last = tab.rows - 1;
    if (rowIndex < last) reorderRowInTab(tab, rowIndex, last);
    if (Array.isArray(tab.rowHeights) && tab.rowHeights.length === tab.rows) {
      tab.rowHeights[last] = MIN_ROW_HEIGHT;
    }

    // Regenerate every Combined scope (indices moved; Match source order may reorder).
    const scopes = Object.create(null);
    state.confirmedLinks.forEach(function (l) { scopes[l.scope] = true; });
    Object.keys(scopes).forEach(function (s) { regenerateCombinedScope(s); });

    if (getStickyCellRange()) clearStickyCellRange();
    clearKeyboardCellRange();
    renderTabs();
    renderGrid();
    renderCombinedPrompt();
    applyConfirmedCellHighlights();
    applyAppendCheckedState();
    renderMasterLibrary();
    scheduleSave();
    setStatus('Row ' + (rowIndex + 1) + ' deleted — rows below moved up' +
      (unlocked ? (' (' + unlocked + ' linked part cell' + (unlocked === 1 ? '' : 's') + ' unlocked, text kept)') : ''), 'ok');
  }

  function openCellShadeMenu(cellIndex, x, y) {
    const tab = activeTab();
    if (!tab || cellIndex < 0 || cellIndex >= tab.cells.length) return;
    closeAddPageChoiceMenu();
    const indices = shadeTargetIndices(tab, cellIndex);
    const shades = indices.map(function (i) { return getCellShade(tab, i) || null; });
    const allSame = shades.every(function (v) { return v === shades[0]; });
    const current = allSame ? shades[0] : undefined;
    const anyShaded = shades.some(function (v) { return !!v; });
    const anyClearable = indices.some(function (i) {
      if ((tab.cells[i] || '') !== '') return true;
      return !isMasterTab(tab) && !!getCellLock(tab, i);
    });
    const r = Math.floor(cellIndex / tab.cols);
    const c = cellIndex % tab.cols;
    const heading = indices.length > 1
      ? (indices.length + ' cells')
      : ('Cell ' + cellAddress(r, c));
    // Snapshot text selection now — opening the menu can blur the textarea.
    const ta = el.cellGrid && el.cellGrid.querySelector('textarea.cell[data-idx="' + cellIndex + '"]');
    let textSel = null;
    if (ta && typeof ta.selectionStart === 'number' && typeof ta.selectionEnd === 'number' &&
        ta.selectionStart !== ta.selectionEnd) {
      textSel = {
        start: ta.selectionStart,
        end: ta.selectionEnd,
        value: ta.value == null ? '' : String(ta.value)
      };
    }
    const sticky = getStickyCellRange();
    const inStickyBlock = !!(sticky && indices.length > 1);

    function clipboardPayloadForCopy() {
      // Match Ctrl/Cmd+C: block TSV if right-click is inside sticky range; else
      // selected substring; else whole cell.
      if (inStickyBlock) {
        const tsv = buildCellRangeTsv(tab, sticky.rMin, sticky.cMin, sticky.rMax, sticky.cMax);
        const rows = sticky.rMax - sticky.rMin + 1;
        const cols = sticky.cMax - sticky.cMin + 1;
        return { text: tsv, status: 'Copied ' + rows + '\u00d7' + cols + ' cells' };
      }
      if (textSel) {
        return {
          text: textSel.value.slice(textSel.start, textSel.end),
          status: 'Copied selection'
        };
      }
      const value = (tab.cells[cellIndex] == null ? '' : String(tab.cells[cellIndex]));
      return { text: value, status: 'Copied cell' };
    }

    const items = [];
    items.push({
      label: 'Copy',
      title: inStickyBlock
        ? 'Copy selected cells as TSV (same as Ctrl+C)'
        : (textSel ? 'Copy selected text (same as Ctrl+C)' : 'Copy cell text (same as Ctrl+C)'),
      onClick: function () {
        const payload = clipboardPayloadForCopy();
        copyTextWithStatus(payload.text, payload.status);
      }
    });
    items.push({
      label: 'Cut',
      title: isMasterTab(tab)
        ? 'Copy, then empty (Master sync may update linked part cells)'
        : 'Copy, then empty and unlink from Master (Master library unchanged)',
      disabled: !anyClearable,
      onClick: function () {
        const payload = clipboardPayloadForCopy();
        // Clipboard write is async; Clear (v0.147) is the one undo step.
        writeTextToClipboard(payload.text).catch(function () {
          copyTextWithStatus(payload.text);
        });
        clearCellsFromMenu(tab, indices, {
          statusText: indices.length > 1
            ? ('Cut ' + indices.length + ' cells')
            : (isMasterTab(tab) ? 'Cut cell' : 'Cut cell (unlinked from Master)')
        });
      }
    });
    items.push.apply(items, [
      { id: 'green', label: 'Shade green' },
      { id: 'yellow', label: 'Shade yellow' },
      { id: 'red', label: 'Shade red' }
    ].map(function (opt) {
      return {
        label: opt.label,
        swatch: opt.id,
        current: current === opt.id,
        title: current === opt.id ? 'Current shade' : '',
        onClick: function () { applyShadeToCells(tab, indices, opt.id); }
      };
    }));
    items.push({
      label: 'Clear shading',
      swatch: 'none',
      disabled: !anyShaded,
      current: current === null,
      title: anyShaded ? '' : 'No shading to clear',
      onClick: function () { applyShadeToCells(tab, indices, null); }
    });
    items.push({
      label: 'Clear',
      title: isMasterTab(tab)
        ? 'Empty this cell (Master sync may update linked part cells)'
        : 'Empty cell and unlink from Master (Master library unchanged)',
      disabled: !anyClearable,
      onClick: function () {
        clearCellsFromMenu(tab, indices, { confirm: true, anchor: cellWrapAt(cellIndex) || { x: x, y: y } });
      }
    });
    openCtxMenu({ heading: heading, items: items, x: x, y: y, className: 'cell-shade-menu' });
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
         t.closest('.cell-nest-add'))) {
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
    const kind = opts.kind === 'part' ? 'part' : (opts.kind === 'nest' ? 'nest' : 'cell');
    const cellIndex = Number.isInteger(opts.cellIndex) ? opts.cellIndex : -1;
    // Same + clicked again while its menu is open → just close (toggle).
    const already = document.querySelector('.add-page-choice-menu');
    if (already && already.__anchor === anchorBtn) {
      closeAddPageChoiceMenu();
      return;
    }
    if (kind === 'cell' || kind === 'nest') {
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
    menu.setAttribute('aria-label', kind === 'part' ? 'Add part page' : (kind === 'nest' ? 'Add nest' : 'Add cell page'));
    menu.__anchor = anchorBtn;

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
        else if (kind === 'nest') {
          if (mode === 'copy') copyNestFromCurrent(cellIndex);
          else addNestedCell(cellIndex);
        } else addCellPage(cellIndex, mode);
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
      // Portal + fixed (part page +, nest +): no layout change in the grid.
      document.body.appendChild(menu);
      const rect = anchorBtn.getBoundingClientRect();
      const pad = 6;
      menu.style.position = 'fixed';
      menu.style.left = '0px';
      menu.style.top = '0px';
      menu.style.right = 'auto';
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
  /**
   * v0.143: Col A (part tabs) shows no nests. Old nest data is kept untouched
   * but hidden, so a nest check there must never feed Combined.
   */
  function isHiddenCol1NestLink(link) {
    if (!link || linkNestIndex(link) === null) return false;
    const tab = state.tabs.find(function (t) { return t.id === link.tabId; });
    return isPartTabCol1Cell(tab, link.cellIndex);
  }

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
      scope: target,
      seq: nextLinkSeq()
    });
    return true;
  }

  /**
   * Rebuild Combined scopes so legacy Col1 confirmed segments leave the prompt
   * text while keeping zero-length filter links for checkbox state.
   */
  function stripCol1TextFromAllCombinedScopes() {
    // v0.140: Col1 never contributes to the generated text; nothing to strip.
    return regenerateAllCombined();
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
    // v0.140: the generated Combined view is read-only (type in Notes), so
    // there is nothing to read back — just repaint from the checks.
    renderCombinedPrompt();
  }

  /**
   * Rewrite one confirmed segment in Combined to newText (in place), shifting
   * later links in the same scope. Returns false if the old span cannot be found.
   */
  function rewriteConfirmedLinkSegment(link, newText) {
    if (!link) return false;
    regenerateCombinedScope(link.scope);
    return true;
  }

  /**
   * Ash: editing a green/confirmed cell or nest in the table also edits the
   * matching Combined segment live (keeps the link). Empty source removes it.
   * nestIndexFilter: undefined = all links for the cell; null = parent only;
   * integer = that nest only.
   */
  function liveSyncConfirmedLinksForCell(tabId, cellIndex, nestIndexFilter, opts) {
    // v0.140: Combined text is regenerated from the cell, so a live edit just
    // repaints. opts.dropEmpty (user typed the cell empty) unchecks it, the
    // familiar behaviour; page flips / Master syncs keep the check.
    const silent = !!(opts && opts.silent);
    const dropEmpty = !!(opts && opts.dropEmpty);
    const candidates = state.confirmedLinks.filter(function (link) {
      if (link.tabId !== tabId || link.cellIndex !== cellIndex) return false;
      if (nestIndexFilter === undefined) return true;
      return linkNestIndex(link) === nestIndexFilter;
    });
    if (!candidates.length) return false;
    const removeIds = {};
    candidates.forEach(function (link) {
      const expected = sourceLinkText(link);
      if (expected === null || (dropEmpty && expected === '')) removeIds[link.id] = true;
    });
    const scopes = {};
    candidates.forEach(function (link) { scopes[link.scope] = true; });
    if (Object.keys(removeIds).length) {
      state.confirmedLinks = state.confirmedLinks.filter(function (link) { return !removeIds[link.id]; });
    }
    let changed = Object.keys(removeIds).length > 0;
    Object.keys(scopes).forEach(function (scope) {
      if (regenerateCombinedScope(scope)) changed = true;
    });
    if (!silent) {
      renderCombinedPrompt();
      applyConfirmedCellHighlights();
    }
    return changed;
  }

  function revalidateLinksForCell(tabId, cellIndex, opts) {
    // v0.140: checks follow the cell; only a vanished cell drops its check.
    const before = state.confirmedLinks.length;
    regenerateAllCombined();
    const removed = state.confirmedLinks.length !== before;
    if (!(opts && opts.silent)) {
      renderCombinedPrompt();
      applyConfirmedCellHighlights();
    }
    return removed;
  }

  function revalidateAllConfirmedLinks() {
    const before = state.confirmedLinks.length;
    regenerateAllCombined();
    renderCombinedPrompt();
    applyConfirmedCellHighlights();
    return state.confirmedLinks.length !== before;
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
  function rejoinConfirmedLinks(scope) {
    return regenerateCombinedScope(scope || currentPromptScope());
  }

  /**
   * When Match source order is ON, rebuild Combined confirmed segments in
   * grid/tab source order (row-major). Preserves plain text before the first
   * and after the last confirmed segment; replaces the middle.
   */
  function reorderCombinedToSourceOrder(scope) {
    return regenerateCombinedScope(scope || currentPromptScope());
  }

  /**
   * After a Tools separator is committed, rejoin every scope that has confirmed
   * links using the new part/column/row separators. Match source order also
   * re-sorts; otherwise document order is kept so caret/append order is not wiped.
   */
  function rewriteCombinedForSeparators() {
    const changed = regenerateAllCombined();
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
    const removeIds = {};
    const scopes = {};
    for (let i = 0; i < linksToRemove.length; i++) {
      if (!linksToRemove[i]) continue;
      removeIds[linksToRemove[i].id] = true;
      scopes[linksToRemove[i].scope] = true;
    }
    const before = state.confirmedLinks.length;
    state.confirmedLinks = state.confirmedLinks.filter(function (link) {
      return !removeIds[link.id];
    });
    if (state.confirmedLinks.length === before) return false;
    Object.keys(scopes).forEach(function (scope) { regenerateCombinedScope(scope); });
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
    const wantOn = !isCellConfirmedInScope(tab.id, cellIndex, scope);
    if (wantOn && !cellParentText(tab, cellIndex)) {
      setStatus('Nothing to append', 'err');
      return;
    }
    // ONE undo snapshot taken BEFORE anything changes — Col1 exclusive
    // cascade used to uncheck rival categories before the snapshot (v0.140 fix).
    const depth = undoStack.length;
    pushHistory();
    const tookSnapshot = undoStack.length > depth;
    const wasSuspended = historySuspended;
    historySuspended = true;
    let changed = false;
    try {
      // Route through ensureCellConfirmedState so Col1 Combined cascades on both
      // check and uncheck (nests stay independent inside that helper).
      changed = ensureCellConfirmedState(cellIndex, wantOn, { col1Cascade: true });
    } finally {
      historySuspended = wasSuspended;
    }
    if (!changed) {
      if (tookSnapshot) undoStack.pop();
      applyConfirmedCellHighlights();
      return;
    }
    refreshAfterConfirmedChange();
    if (!wantOn) {
      const row = Math.floor(cellIndex / tab.cols) + 1;
      const col = (cellIndex % tab.cols) + 1;
      setStatus('Removed ' + cellAddress(row - 1, col - 1) + ' from combined', 'ok');
    }
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
    // v0.140: "append" = add check keys. Plain separator pieces are ignored —
    // the generated text inserts separators itself. Order comes from Match
    // source order (grid) or check order (link.seq), never from a caret.
    const quiet = !!(opts && opts.quiet);
    if (!pieces || !pieces.length) {
      if (!quiet) setStatus('Nothing to append', 'err');
      return;
    }
    const scope = currentPromptScope();
    if (scope !== 'global' && !activeTab()) return;
    const toAdd = [];
    for (let i = 0; i < pieces.length; i++) {
      const piece = pieces[i];
      if (!piece || piece.type !== 'confirmed') continue;
      const nestIdx = Number.isInteger(piece.nestIndex) && piece.nestIndex >= 0 ? piece.nestIndex : null;
      const exists = nestIdx === null
        ? isCellConfirmedInScope(piece.tabId, piece.cellIndex, scope)
        : isNestConfirmedInScope(piece.tabId, piece.cellIndex, nestIdx, scope);
      if (exists) continue;
      const pieceTab = state.tabs.find(function (t) { return t.id === piece.tabId; });
      if (!pieceTab) continue;
      const source = nestIdx === null
        ? sourceCellText(piece.tabId, piece.cellIndex)
        : sourceNestText(piece.tabId, piece.cellIndex, nestIdx);
      if (!source) continue;
      toAdd.push({ piece: piece, nestIdx: nestIdx, tab: pieceTab });
    }
    if (!toAdd.length) {
      if (!quiet) setStatus('Nothing to append', 'err');
      return;
    }
    pushHistory();
    toAdd.forEach(function (item) {
      if (item.nestIdx === null && isPartTabCol1Cell(item.tab, item.piece.cellIndex)) {
        ensureCol1FilterLink(item.tab, item.piece.cellIndex, scope);
        return;
      }
      const link = {
        id: linkUid(),
        tabId: item.piece.tabId,
        cellIndex: item.piece.cellIndex,
        text: '',
        start: 0,
        end: 0,
        scope: scope,
        seq: nextLinkSeq()
      };
      if (item.nestIdx !== null) link.nestIndex = item.nestIdx;
      state.confirmedLinks.push(link);
    });
    regenerateCombinedScope(scope);
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
    // Never open an edit session (and its Overwrite-Master prompt) on a stale or
    // orphaned lock: re-sync it from Master first, or unlock it if the source is gone.
    const fix = reconcileMasterLocks({ onlyTabId: tab.id, onlyIndex: cellIndex });
    if (fix.resynced || fix.repointed || fix.unlocked) {
      renderGrid();
      renderCombinedPrompt();
      applyConfirmedCellHighlights();
      applyAppendCheckedState();
      scheduleSave();
      if (fix.unlocked) {
        setStatus('Master source for this cell is gone — unlocked; your text is kept and editable', 'ok');
        const fresh = el.cellGrid && el.cellGrid.querySelector('textarea.cell[data-idx="' + cellIndex + '"]');
        if (fresh) fresh.focus();
        return;
      }
      setStatus('This linked cell was out of date — re-synced from Master; double-click again to edit', 'ok');
      return;
    }
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

  /** Live text of the in-progress Master edit (cell or Combined segment). */
  function getMasterEditLiveNewText() {
    if (!masterSegmentEdit) return '';
    if (masterSegmentEdit.kind === 'cell') {
      const tab = state.tabs.find(function (t) { return t.id === masterSegmentEdit.tabId; });
      if (!tab) return '';
      const idx = masterSegmentEdit.cellIndex;
      if (idx < 0 || idx >= tab.cells.length) return masterSegmentEdit.originalText || '';
      return tab.cells[idx] == null ? '' : String(tab.cells[idx]);
    }
    const link = linkById(masterSegmentEdit.linkId, masterSegmentEdit.scope);
    if (!link) return '';
    const span = findCombinedSegmentSpan(masterSegmentEdit.linkId);
    return span ? (span.textContent || '') : (link.text || '');
  }

  function showMasterSegmentDialog() {
    if (!el.masterSegmentDialog) return;
    updateMasterEditDialogCopy();
    const emptyEdit = !String(getMasterEditLiveNewText() || '').trim();
    if (el.btnMasterOverwrite) {
      el.btnMasterOverwrite.hidden = emptyEdit;
    }
    if (el.btnMasterKeepLocal) {
      el.btnMasterKeepLocal.classList.toggle('btn-primary', emptyEdit);
      el.btnMasterKeepLocal.classList.toggle('btn-secondary', !emptyEdit);
    }
    if (emptyEdit) {
      const title = document.getElementById('master-segment-dialog-title');
      const body = el.masterSegmentDialog.querySelector('.modal-body');
      if (isMasterCellEditSession()) {
        if (title) title.textContent = 'Master cell cleared';
        if (body) {
          body.textContent =
            'Keep this as a local empty cell (unlinked from Master), or cancel and discard the clear? Overwrite Master is not offered for empty text.';
        }
      } else {
        if (title) title.textContent = 'Master segment cleared';
        if (body) {
          body.textContent =
            'Keep this as local empty Combined text (unlinked from Master), or cancel and discard the clear? Overwrite Master is not offered for empty text.';
        }
      }
    }
    masterSegmentDialogOpen = true;
    el.masterSegmentDialog.hidden = false;
    // Empty edits must never default to Overwrite — focus Keep as local only.
    if (emptyEdit) {
      if (el.btnMasterKeepLocal) el.btnMasterKeepLocal.focus();
    } else if (el.btnMasterOverwrite) {
      el.btnMasterOverwrite.focus();
    }
  }

  function hideMasterSegmentDialog() {
    masterSegmentDialogOpen = false;
    if (el.masterSegmentDialog) el.masterSegmentDialog.hidden = true;
    if (el.btnMasterOverwrite) el.btnMasterOverwrite.hidden = false;
    if (el.btnMasterKeepLocal) {
      el.btnMasterKeepLocal.classList.add('btn-secondary');
      el.btnMasterKeepLocal.classList.remove('btn-primary');
    }
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

    // No Master source found: do NOT rewrite/lock other cells (that created orphans).
    if (masterIdx < 0) return -1;
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

  /** Master index applyMasterOverwriteText would pick for oldText (same rule). */
  function resolveMasterOverwriteIdx(oldText, preferredMasterIdx) {
    const master = state.tabs.find(isMasterTab);
    if (!master) return -1;
    let masterIdx = Number.isInteger(preferredMasterIdx) ? preferredMasterIdx : -1;
    if (masterIdx < 0 || masterIdx >= master.cells.length ||
        cellParentText(master, masterIdx) !== oldText) {
      const found = findMasterCellIndexByText(oldText);
      masterIdx = found === null ? -1 : found;
    }
    return masterIdx;
  }

  /**
   * Cells Overwrite Master would ALSO rewrite that are not linked to that Master
   * cell: any tab cell (incl. Master duplicates) whose text equals the old text.
   * Linked part cells (lock → masterIdx) and the edited cell itself don't count.
   */
  function countUnlinkedSameTextCells(oldText, masterIdx, skipTabId, skipIndex) {
    const master = state.tabs.find(isMasterTab);
    if (!master || masterIdx < 0) return 0;
    let n = 0;
    state.tabs.forEach(function (tab) {
      if (!tab || !Array.isArray(tab.cells)) return;
      for (let i = 0; i < tab.cells.length; i++) {
        if (tab.id === master.id && i === masterIdx) continue;
        if (tab.id === skipTabId && i === skipIndex) continue;
        if (cellParentText(tab, i) !== oldText) continue;
        if (!isMasterTab(tab)) {
          const lock = getCellLock(tab, i);
          if (lock && lock.masterOrigin === true && lock.masterCellIndex === masterIdx) continue;
        }
        n++;
      }
    });
    return n;
  }

  function overwriteMasterFromSegmentEdit(confirmedOthers) {
    if (!masterSegmentEdit) {
      hideMasterSegmentDialog();
      return;
    }
    const edit = masterSegmentEdit;
    if (!confirmedOthers) {
      // v0.157: Overwrite also rewrites unlinked cells holding the same old text —
      // ask first (Cancel keeps the dialog + edit open, nothing changes).
      let preferredIdx = -1;
      let skipTabId = null;
      let skipIndex = -1;
      if (edit.kind === 'cell') {
        const etab = state.tabs.find(function (t) { return t.id === edit.tabId; });
        const elock = etab ? getCellLock(etab, edit.cellIndex) : null;
        preferredIdx = elock && Number.isInteger(elock.masterCellIndex) ? elock.masterCellIndex : -1;
        skipTabId = edit.tabId;
        skipIndex = edit.cellIndex;
      } else {
        const elink = linkById(edit.linkId, edit.scope);
        preferredIdx = elink && Number.isInteger(elink.masterCellIndex) ? elink.masterCellIndex : -1;
      }
      const mIdx = resolveMasterOverwriteIdx(edit.originalText, preferredIdx);
      const others = countUnlinkedSameTextCells(edit.originalText, mIdx, skipTabId, skipIndex);
      if (others > 0) {
        openActionConfirm({
          anchor: el.btnMasterOverwrite,
          focusBack: el.btnMasterOverwrite,
          message: 'Also update ' + pluralCells(others, 'other cell') + ' with the same text?',
          actionLabel: 'Update',
          cancelStatus: 'Overwrite cancelled — nothing changed',
          ariaLabel: 'Confirm overwrite other cells',
          onConfirm: function () {
            if (masterSegmentEdit !== edit) return;
            overwriteMasterFromSegmentEdit(true);
          }
        });
        return;
      }
    }
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
      const master = state.tabs.find(isMasterTab);
      if (!master || !masterSourceHasContent(master, preferred) ||
          trimEndText(master.cells[preferred] || '') !== trimEndText(oldText)) {
        // Source missing or out of date — never "overwrite" a Master cell we can't
        // identify; keep the edit as a normal local cell instead.
        masterSegmentEdit = edit;
        keepMasterSegmentLocalOnly();
        setStatus('No matching Master cell — kept as local text (Master unchanged)', 'err');
        return;
      }
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
    const scope = currentPromptScope();
    // v0.140: always regenerate from the checks before painting.
    regenerateCombinedScope(scope);
    const text = getPromptText(scope);
    const notes = getNotes(scope).replace(/\s+$/, '');
    const links = linksForScope(scope).filter(function (link) {
      return !isCol1FilterOnlyLink(link) && !isHiddenCol1NestLink(link) && link.end > link.start;
    });

    el.combined.replaceChildren();
    // Notes are shown first (muted) so the box shows exactly what Copy gives.
    if (notes) {
      const notesSpan = document.createElement('span');
      notesSpan.className = 'combined-notes-preview';
      notesSpan.textContent = notes + (text ? notesSeparator() : '');
      notesSpan.title = 'Your Notes (edit them in the Notes box)';
      el.combined.appendChild(notesSpan);
    }
    let pos = 0;
    links.forEach(function (link) {
      if (link.start > pos) {
        el.combined.appendChild(document.createTextNode(text.slice(pos, link.start)));
      }
      const span = document.createElement('span');
      span.className = 'confirmed-segment';
      span.dataset.linkId = link.id;
      span.textContent = text.slice(link.start, link.end);
      span.title = 'From a checked cell — edit the cell to change it, uncheck it to remove it';
      el.combined.appendChild(span);
      pos = link.end;
    });
    if (pos < text.length) {
      el.combined.appendChild(document.createTextNode(text.slice(pos)));
    }

    el.combined.classList.toggle('is-empty', !text && !notes);
    lastCombinedViewKey = combinedViewKey(scope);
    if (el.combinedNotes && document.activeElement !== el.combinedNotes &&
        el.combinedNotes.value !== getNotes(scope)) {
      el.combinedNotes.value = getNotes(scope);
    }
    el.globalCombined.checked = state.globalCombined;
    if (el.matchSourceOrder) el.matchSourceOrder.checked = !!state.matchSourceOrder;
    applyConfirmedCellHighlights();
    syncCombinedOverflowY();
  }

  function mergePartPrompts() {
    // v0.140: Global ON pulls every per-tab check (and Notes) into Global.
    state.confirmedLinks.forEach(function (link) {
      if (link.scope === 'global') return;
      link.scope = 'global';
    });
    Object.keys(state.partNotes || {}).forEach(function (tabId) {
      const notes = (state.partNotes[tabId] || '').replace(/\s+$/, '');
      if (!notes) return;
      const base = (state.combinedNotes || '').replace(/\s+$/, '');
      state.combinedNotes = base ? base + '\n' + notes : notes;
    });
    state.partNotes = {};
    state.partPrompts = {};
    regenerateCombinedScope('global');
  }

  /** Pre-0.140 text-based merge (only used before migrating legacy data). */
  function legacyMergePartPrompts() {
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
    return composeCombinedOutput(currentPromptScope());
  }

  let lastLockRepair = null;
  let lastCombinedMigration = null;
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
    state.combinedNotes = typeof data.combinedNotes === 'string' ? data.combinedNotes : '';
    state.partNotes = data.partNotes && typeof data.partNotes === 'object'
      ? Object.keys(data.partNotes).reduce(function (notes, id) {
        if (typeof data.partNotes[id] === 'string' && data.partNotes[id]) notes[id] = data.partNotes[id];
        return notes;
      }, {})
      : {};
    syncLinkSeqCounter();
    lastCombinedMigration = null;
    if (data.combinedModel !== COMBINED_MODEL) {
      // One-time: stored Combined text → checks + Notes (nothing lost).
      if (state.globalCombined) legacyMergePartPrompts();
      lastCombinedMigration = migrateLegacyCombined();
    }
    state.combinedModel = COMBINED_MODEL;
    if (state.globalCombined) mergePartPrompts();
    // Generated text is rebuilt from the checks (Col1 contributes nothing).
    regenerateAllCombined();
    // Repair stale / orphaned Master locks (all part tabs + stored part pages).
    lastLockRepair = reconcileMasterLocks();
    // Re-lock cells that still match Master text exactly (e.g. old Duplicate
    // pages that cleared cellLocks). Never changes user text.
    const relock = repairMissingMasterLocksFromExactMatch();
    if (relock.relocked) {
      lastLockRepair = lastLockRepair || { checked: 0, resynced: 0, repointed: 0, unlocked: 0 };
      lastLockRepair.relocked = (lastLockRepair.relocked || 0) + relock.relocked;
    }
    if (el.partSeparator) el.partSeparator.value = state.separators.part;
    if (el.columnSeparator) el.columnSeparator.value = state.separators.column;
    if (el.rowSeparator) el.rowSeparator.value = state.separators.row;

    renderCombinedPrompt();
    // Undo/redo/doc switch: Notes box shows the restored text even if focused.
    if (el.combinedNotes) el.combinedNotes.value = getNotes(currentPromptScope());
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
      !data.combinedNotes &&
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
    const menu = popoverMenus.masterLib;
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
    const menu = popoverMenus.masterLib;
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
        const row = document.createElement('div');
        row.className = 'column-col1-filter-option';

        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.checked = isMasterLibValueSelected(value);
        if (value) cb.dataset.value = value;
        else cb.dataset.blank = '1';
        cb.title = 'Toggle filter for this Master Column A value';
        cb.setAttribute('aria-label', 'Filter Master Column A “' + (value || '(blank)') + '”');
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

        const nameBtn = document.createElement('button');
        nameBtn.type = 'button';
        nameBtn.className = 'column-col1-filter-option-text';
        nameBtn.textContent = value ? value : '(blank)';
        if (!value) nameBtn.classList.add('is-blank');
        nameBtn.title = 'Click to jump · checkbox to filter';
        nameBtn.addEventListener('click', function (e) {
          e.preventDefault();
          e.stopPropagation();
          jumpToMasterLibCol1Value(master, value);
        });

        row.appendChild(cb);
        row.appendChild(nameBtn);
        list.appendChild(row);
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
              insertMasterText(cellText, masterIdx, { anchor: button });
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
      dropPopover('masterLib');
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
    portalPopover('masterLib', filterMenu);
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
    sizeMasterLibraryPane();
  }

  /**
   * v0.142: bound the Master insert pane in px from JS. The CSS-only bound
   * relied on ::details-content (Chromium 131+); Ash's Electron 33 is
   * Chromium 130, which ignores that rule, so the <details> content box grew
   * to the full list height and was merely clipped by max-height:30% — the
   * results pane never got a bounded height and could not scroll (worse the
   * more Values were selected). An explicit max-height on the items flex box
   * works in every engine: toolbar stays, results flex/shrink and scroll.
   */
  let masterLibSizeRaf = 0;
  function sizeMasterLibraryPane() {
    masterLibSizeRaf = 0;
    const lib = el.masterLibrary;
    const items = el.masterLibraryItems;
    if (!lib || !items) return;
    if (!lib.open || lib.hidden || !el.partSection) {
      items.style.maxHeight = '';
      lib.style.maxHeight = '';
      return;
    }
    const summary = lib.querySelector('summary');
    const toolbar = items.querySelector('.master-library-toolbar');
    const summaryH = summary ? summary.offsetHeight : 0;
    const borders = lib.offsetHeight - lib.clientHeight;
    // Fixed chrome = header + search/sort bar + bottom padding. The results
    // always get >= ~3 rows (110px) when the part area can spare it (max 50%),
    // else the plain 30% cap. Depends only on window/section size, never on
    // how many Values are selected → no bounce.
    const chrome = summaryH + borders + (toolbar ? toolbar.offsetHeight : 0) + 10;
    const partH = el.partSection.clientHeight;
    const cap = Math.max(Math.floor(partH * 0.3), Math.min(chrome + 110, Math.floor(partH * 0.5)));
    const libValue = cap + 'px';
    if (lib.style.maxHeight !== libValue) lib.style.maxHeight = libValue;
    const value = Math.max(48, cap - summaryH - borders) + 'px';
    if (items.style.maxHeight !== value) items.style.maxHeight = value;
  }

  function scheduleMasterLibrarySize() {
    if (masterLibSizeRaf) return;
    masterLibSizeRaf = requestAnimationFrame(sizeMasterLibraryPane);
  }

  /**
   * Model-only core of a Master link (shared by Master insert, typed/pasted
   * exact matches). Caller pushes history and renders. tab must be active
   * (Combined auto-check targets the active tab).
   */
  function applyMasterLinkToCell(tab, index, masterIdx, insertText) {
    setCellMasterLock(tab, index, masterIdx >= 0 ? masterIdx : null);
    // Replace/sync nests + parent pages from the Master cell (Ash).
    copyNestsFromMasterToCell(tab, index, masterIdx);
    copyPagesFromMasterToCell(tab, index, masterIdx);
    // Master text is authoritative for the active page slot.
    // (copyPages alone can leave the cell blank when Master cellPages desyncs
    // from master.cells, especially after part-page switches / empty page 2.)
    writeCellCurrentPage(tab, index, insertText);
    syncNestLinksAfterNestReplace(tab.id, index);
    revalidateLinksForCell(tab.id, index, { silent: true });
    // Auto-check Combined for the inserted cell (no-op if already included).
    const wasSuspended = historySuspended;
    historySuspended = true;
    try {
      ensureCellConfirmedState(index, true, { quiet: true, col1Cascade: false });
    } finally {
      historySuspended = wasSuspended;
    }
  }

  function trimEndText(value) {
    return String(value == null ? '' : value).replace(/\s+$/, '');
  }

  /**
   * Master cell whose text exactly equals typed/pasted text (trailing
   * whitespace ignored; leading whitespace and inner newlines must match).
   * Several matches → prefer, in order: Master row whose Column A equals this
   * part row's Column A; row visible under this part's Master-insert Values
   * filter; same column as the target cell; then first in row-major order.
   * Returns -1 when nothing matches.
   */
  function findMasterMatchForCell(tab, index, text) {
    const master = state.tabs.find(isMasterTab);
    if (!master || !tab || isMasterTab(tab)) return -1;
    const want = trimEndText(text);
    if (!want.trim()) return -1;
    const cols = master.cols || 1;
    const targetCol = index % (tab.cols || 1);
    const targetRow = Math.floor(index / (tab.cols || 1));
    const rowColA = targetCol === 0 ? '' : (tab.cells[targetRow * tab.cols] || '').trim();
    const filter = getMasterLibPrefs(tab.id).valueFilter;
    let best = -1;
    let bestScore = -1;
    for (let i = 0; i < master.cells.length; i++) {
      if (trimEndText(master.cells[i]) !== want) continue;
      const mRow = Math.floor(i / cols);
      const mColA = (master.cells[mRow * cols] || '').trim();
      let score = 0;
      if (rowColA && mColA === rowColA) score += 4;
      if (filter === null || filter.has(mColA)) score += 2;
      if ((i % cols) === targetCol) score += 1;
      if (score > bestScore) {
        best = i;
        bestScore = score;
      }
    }
    return best;
  }

  /** True when linking would discard this cell's own nests / other pages. */
  function cellHasOwnExtraContent(tab, index) {
    ensureNestedCells(tab);
    const nests = tab.nestedCells[index] || [];
    for (let n = 0; n < nests.length; n++) {
      const pages = (nests[n] && nests[n].pages) || [];
      for (let p = 0; p < pages.length; p++) {
        if (String(pages[p] || '').trim()) return true;
      }
    }
    ensureCellPagesArray(tab);
    const entry = tab.cellPages[index];
    if (entry && entry.pages && entry.pages.length > 1) {
      for (let p = 0; p < entry.pages.length; p++) {
        if (p === (entry.page || 0)) continue;
        if (String(entry.pages[p] || '').trim()) return true;
      }
    }
    return false;
  }

  /**
   * Typed/pasted text that exactly matches a Master cell becomes a real
   * Master link (same as Master insert). Returns 'linked', 'blocked'
   * (would discard own nests/pages) or null. Model only — no history/render.
   */
  function linkCellIfMasterMatch(tab, index) {
    if (!tab || isMasterTab(tab) || tab !== activeTab()) return null;
    if (index < 0 || index >= tab.cells.length) return null;
    if (getCellLock(tab, index)) return null;
    if (isMasterCellEditFor(tab, index)) return null;
    const text = tab.cells[index] || '';
    if (!text.trim()) return null;
    const masterIdx = findMasterMatchForCell(tab, index, text);
    if (masterIdx < 0) return null;
    if (cellHasOwnExtraContent(tab, index)) return 'blocked';
    const master = state.tabs.find(isMasterTab);
    applyMasterLinkToCell(tab, index, masterIdx, master.cells[masterIdx] || '');
    return 'linked';
  }

  /* ── Col1 data validation (part tabs): list = Master Column A values ── */

  /** Unique trimmed non-empty Master Column A values (Master row order). */
  function masterCol1Values() {
    const master = state.tabs.find(isMasterTab);
    const out = [];
    if (!master || !master.cols) return out;
    const seen = Object.create(null);
    for (let r = 0; r < master.rows; r++) {
      const v = (master.cells[r * master.cols] || '').trim();
      if (!v || seen[v]) continue;
      seen[v] = true;
      out.push(v);
    }
    return out;
  }

  /**
   * Col1 value allowed? Empty is allowed (clearing). With no Master Column A
   * values at all, validation is off so nobody gets locked out.
   */
  function isCol1ValueAllowed(value, allowed) {
    const v = String(value == null ? '' : value).trim();
    if (!v) return true;
    const list = allowed || masterCol1Values();
    if (!list.length) return true;
    return list.indexOf(v) !== -1;
  }

  function countInvalidCol1Cells(tab) {
    if (!tab || isMasterTab(tab) || !tab.cols) return 0;
    const allowed = masterCol1Values();
    if (!allowed.length) return 0;
    let n = 0;
    for (let r = 0; r < tab.rows; r++) {
      if (!isCol1ValueAllowed(tab.cells[r * tab.cols], allowed)) n++;
    }
    return n;
  }

  /** Master Column A cell index for a Col1 value (first row), or -1. */
  function masterCol1IndexForValue(value) {
    const master = state.tabs.find(isMasterTab);
    const v = String(value == null ? '' : value).trim();
    if (!master || !master.cols || !v) return -1;
    for (let r = 0; r < master.rows; r++) {
      const idx = r * master.cols;
      if ((master.cells[idx] || '').trim() === v) return idx;
    }
    return -1;
  }

  /** Values to remember per Col1 cell at focus, so an invalid commit can revert. */
  const col1FocusValues = Object.create(null);

  function onCol1CellFocusRemember(e) {
    const tab = activeTab();
    const ta = e.currentTarget;
    if (!tab || !ta) return;
    const idx = parseInt(ta.dataset.idx, 10);
    if (!isPartTabCol1Cell(tab, idx)) return;
    col1FocusValues[tab.id + ':' + idx] = tab.cells[idx] || '';
  }

  function flashCol1Invalid(idx) {
    const ta = el.cellGrid && el.cellGrid.querySelector('textarea.cell[data-idx="' + idx + '"]');
    const wrap = ta && ta.closest ? ta.closest('.cell-wrap') : null;
    if (!wrap) return;
    wrap.classList.remove('col1-invalid-flash');
    void wrap.offsetWidth;
    wrap.classList.add('col1-invalid-flash');
    window.setTimeout(function () { wrap.classList.remove('col1-invalid-flash'); }, 900);
  }

  /** Typed/pasted Col1 value not in the Master list → revert (returns true). */
  function revertInvalidCol1Commit(tab, idx, ta) {
    if (!isPartTabCol1Cell(tab, idx)) return false;
    const value = tab.cells[idx] || '';
    if (isCol1ValueAllowed(value)) return false;
    const key = tab.id + ':' + idx;
    const prev = Object.prototype.hasOwnProperty.call(col1FocusValues, key) ? col1FocusValues[key] : '';
    writeCellCurrentPage(tab, idx, prev);
    if (ta) ta.value = prev;
    liveSyncConfirmedLinksForCell(tab.id, idx, null, { silent: true });
    applyAppendCheckedState();
    scheduleSave();
    flashCol1Invalid(idx);
    setStatus('“' + value.trim() + '” is not in Master Column A — reverted (use the ▾ list)', 'err');
    return true;
  }

  let masterValueMenuEl = null;

  function closeMasterValueMenu() {
    if (masterValueMenuEl && masterValueMenuEl.parentNode) {
      masterValueMenuEl.parentNode.removeChild(masterValueMenuEl);
    }
    masterValueMenuEl = null;
    col1MenuEl = null;
    document.removeEventListener('pointerdown', onMasterValueMenuOutside, true);
    document.removeEventListener('keydown', onMasterValueMenuKey, true);
  }
  let col1MenuEl = null;
  function closeCol1ValueMenu() { closeMasterValueMenu(); }

  function onMasterValueMenuOutside(ev) {
    const t = ev.target;
    if (masterValueMenuEl && masterValueMenuEl.contains(t)) return;
    if (t && t.closest && (t.closest('.col1-dropdown-btn') || t.closest('.col2-dropdown-btn'))) return;
    closeMasterValueMenu();
  }

  function masterValueMenuVisibleOptions() {
    if (!masterValueMenuEl) return [];
    return Array.prototype.slice.call(
      masterValueMenuEl.querySelectorAll('.col1-value-option:not([hidden])')
    );
  }

  function onMasterValueMenuKey(ev) {
    if (!masterValueMenuEl) return;
    const search = masterValueMenuEl.querySelector('.col1-value-search');
    const inMenu = masterValueMenuEl.contains(document.activeElement) ||
      (search && document.activeElement === search);
    // Keep focus on the search box so every keystroke filters live. Arrow keys
    // used to .focus() option buttons, after which typing never reached search
    // (looked like filter needed Enter / a re-click).
    if (search && document.activeElement !== search &&
        masterValueMenuEl.contains(document.activeElement)) {
      const printable = !ev.ctrlKey && !ev.metaKey && !ev.altKey &&
        (ev.key.length === 1 || ev.key === 'Backspace' || ev.key === 'Delete' ||
          ev.key === 'Process' || ev.isComposing);
      if (printable) {
        search.focus();
        // Let the character land in the input (input event → applyFilter).
        return;
      }
    }
    if (ev.key === 'Escape' || (ev.key === 'Tab' && !(search && document.activeElement === search))) {
      ev.preventDefault();
      ev.stopPropagation();
      const idx = masterValueMenuEl.dataset.idx;
      closeMasterValueMenu();
      const ta = el.cellGrid && el.cellGrid.querySelector('textarea.cell[data-idx="' + idx + '"]');
      if (ta) ta.focus();
      return;
    }
    const opts = masterValueMenuVisibleOptions();
    const inSearch = search && document.activeElement === search;
    if (ev.key === 'Enter' && (inSearch || inMenu)) {
      ev.preventDefault();
      ev.stopPropagation();
      const pick = opts.find(function (o) { return o.classList.contains('is-kbd'); }) || opts[0];
      if (pick) pick.click();
      return;
    }
    const navKeys = { ArrowDown: 1, ArrowUp: -1, PageDown: 8, PageUp: -8, Home: -1e6, End: 1e6 };
    if (!Object.prototype.hasOwnProperty.call(navKeys, ev.key)) return;
    if (!opts.length) return;
    ev.preventDefault();
    ev.stopPropagation();
    let at = opts.findIndex(function (o) { return o.classList.contains('is-kbd'); });
    if (at < 0) at = opts.findIndex(function (o) { return o.classList.contains('is-selected'); });
    if (at < 0) at = -1;
    const delta = navKeys[ev.key];
    let next;
    if (ev.key === 'Home') next = 0;
    else if (ev.key === 'End') next = opts.length - 1;
    else next = Math.max(0, Math.min(opts.length - 1, (at < 0 ? (delta > 0 ? -1 : 0) : at) + delta));
    opts.forEach(function (o) { o.classList.remove('is-kbd'); });
    opts[next].classList.add('is-kbd');
    opts[next].scrollIntoView({ block: 'nearest' });
    if (search) search.focus();
  }

  /**
   * Shared searchable Master-value picker (Col A + Col B).
   * opts: { kind, idx, values, current, ariaLabel, emptyText, onPick, anchorBtn }
   */
  function openMasterValueMenu(opts) {
    opts = opts || {};
    const kind = opts.kind || 'col1';
    const idx = opts.idx;
    const values = opts.values || [];
    const current = opts.current || '';
    if (masterValueMenuEl && masterValueMenuEl.dataset.idx === String(idx) &&
        masterValueMenuEl.dataset.kind === kind) {
      closeMasterValueMenu();
      return;
    }
    closeMasterValueMenu();
    closeAddPageChoiceMenu();
    closeCellShadePicker();
    const menu = document.createElement('div');
    menu.className = 'col1-value-menu';
    menu.dataset.idx = String(idx);
    menu.dataset.kind = kind;
    menu.setAttribute('role', 'listbox');
    menu.setAttribute('aria-label', opts.ariaLabel || 'Values');

    const search = document.createElement('input');
    search.type = 'search';
    search.className = 'col1-value-search';
    search.placeholder = 'Search…';
    search.setAttribute('aria-label', 'Search values');
    search.autocomplete = 'off';
    search.spellcheck = false;
    menu.appendChild(search);

    const list = document.createElement('div');
    list.className = 'col1-value-list';
    menu.appendChild(list);

    const empty = document.createElement('div');
    empty.className = 'col1-value-empty';
    empty.hidden = true;
    list.appendChild(empty);

    if (current && values.indexOf(current) === -1) {
      const note = document.createElement('div');
      note.className = 'col1-value-empty col1-value-note';
      note.textContent = '“' + current + '” is not in Master — pick a value';
      list.insertBefore(note, empty);
    }

    values.forEach(function (v) {
      const opt = document.createElement('button');
      opt.type = 'button';
      opt.className = 'col1-value-option' + (v === current ? ' is-selected' : '');
      opt.setAttribute('role', 'option');
      opt.setAttribute('aria-selected', v === current ? 'true' : 'false');
      opt.dataset.value = v;
      opt.textContent = v;
      opt.title = v === current ? (v + ' (current)') : v;
      opt.addEventListener('click', function (ev) {
        ev.preventDefault();
        ev.stopPropagation();
        closeMasterValueMenu();
        if (typeof opts.onPick === 'function') opts.onPick(v);
      });
      list.appendChild(opt);
    });

    function applyFilter() {
      const q = (search.value || '').trim().toLowerCase();
      let shown = 0;
      const optsEl = list.querySelectorAll('.col1-value-option');
      for (let i = 0; i < optsEl.length; i++) {
        const hit = !q || (optsEl[i].dataset.value || '').toLowerCase().indexOf(q) !== -1;
        optsEl[i].hidden = !hit;
        optsEl[i].classList.remove('is-kbd');
        if (hit) shown++;
      }
      // Keep empty row in-flow so menu height stays stable while filtering.
      empty.hidden = false;
      if (!values.length) {
        empty.textContent = opts.emptyText || 'No values in Master yet';
      } else if (shown === 0) {
        empty.textContent = q
          ? ('No matches for “' + search.value.trim() + '”')
          : (opts.emptyText || 'No values');
      } else {
        empty.textContent = '';
        empty.hidden = true;
      }
      const first = list.querySelector('.col1-value-option:not([hidden])');
      if (first) first.classList.add('is-kbd');
    }
    applyFilter();
    // Live filter on every keystroke / IME / paste (not Enter-gated).
    search.addEventListener('input', applyFilter);
    search.addEventListener('compositionend', applyFilter);
    search.addEventListener('keydown', function (ev) {
      // Stop keys from reaching locked/readonly cell handlers underneath.
      ev.stopPropagation();
      if (ev.key === 'Escape' || ev.key === 'Enter' ||
          ev.key === 'ArrowDown' || ev.key === 'ArrowUp' ||
          ev.key === 'PageDown' || ev.key === 'PageUp' ||
          ev.key === 'Home' || ev.key === 'End') {
        // Handled by capture-phase onMasterValueMenuKey.
        return;
      }
    });
    // option buttons: don't steal focus on pointerdown so search keeps receiving keys
    list.addEventListener('pointerdown', function (ev) {
      const opt = ev.target && ev.target.closest && ev.target.closest('.col1-value-option');
      if (opt) {
        // Allow click; keep search focused for further typing after a miss-click.
        // (Actual pick still happens on click.)
      }
    });

    document.body.appendChild(menu);
    masterValueMenuEl = menu;
    col1MenuEl = menu; // busy checks / outside click
    const rect = opts.anchorBtn.getBoundingClientRect();
    const size = menu.getBoundingClientRect();
    const pad = 6;
    const left = Math.max(pad, Math.min(rect.right - size.width, window.innerWidth - size.width - pad));
    let top = rect.bottom + 3;
    if (top + size.height > window.innerHeight - pad) top = Math.max(pad, rect.top - size.height - 3);
    menu.style.left = left + 'px';
    menu.style.top = top + 'px';
    // Pin height so filtering rows does not bounce the menu (scroll inside list).
    const pinned = Math.max(size.height, Math.min(320, window.innerHeight - pad * 2));
    menu.style.height = pinned + 'px';
    menu.style.maxHeight = pinned + 'px';
    search.focus();
    const sel = list.querySelector('.col1-value-option.is-selected:not([hidden])') ||
      list.querySelector('.col1-value-option:not([hidden])');
    if (sel) {
      sel.classList.add('is-kbd');
      sel.scrollIntoView({ block: 'nearest' });
    }
    document.addEventListener('pointerdown', onMasterValueMenuOutside, true);
    document.addEventListener('keydown', onMasterValueMenuKey, true);
  }

  /** Pick a Master Column A value for a part-tab Col1 cell (= Master insert). */
  function pickCol1Value(tab, idx, value) {
    if (!tab || activeTab() !== tab || !isPartTabCol1Cell(tab, idx)) return;
    const master = state.tabs.find(isMasterTab);
    const masterIdx = masterCol1IndexForValue(value);
    if (!master || masterIdx < 0) return;
    if ((tab.cells[idx] || '').trim() === value && getCellLock(tab, idx)) {
      setStatus('Column A already “' + value + '”');
      return;
    }
    focusedCell = { tabId: tab.id, index: idx };
    insertMasterText(master.cells[masterIdx] || value, masterIdx);
  }

  function openCol1ValueMenu(anchorBtn, idx) {
    const tab = activeTab();
    if (!tab || !isPartTabCol1Cell(tab, idx)) return;
    const values = masterCol1Values();
    openMasterValueMenu({
      kind: 'col1',
      idx: idx,
      values: values,
      current: (tab.cells[idx] || '').trim(),
      ariaLabel: 'Master Column A values',
      emptyText: 'No Column A values in Master yet',
      anchorBtn: anchorBtn,
      onPick: function (v) { pickCol1Value(tab, idx, v); }
    });
  }

  function openCol2ValueMenu(anchorBtn, idx) {
    const tab = activeTab();
    if (!tab || !isPartTabCol2Cell(tab, idx)) return;
    const row = Math.floor(idx / tab.cols);
    if (!isCol2ValidatedForRow(tab, row)) return;
    const colA = (tab.cells[row * tab.cols] || '').trim();
    const values = masterCol2ValuesForCol1(colA);
    openMasterValueMenu({
      kind: 'col2',
      idx: idx,
      values: values,
      current: (tab.cells[idx] || '').trim(),
      ariaLabel: 'Master Column B values for ' + colA,
      emptyText: 'No Column B values in Master for “' + colA + '”',
      anchorBtn: anchorBtn,
      onPick: function (v) { pickCol2Value(tab, idx, v); }
    });
  }


  /* ── Col2 data validation (part tabs): list = Master Column B for this row's Col A ── */

  /** Part-tab Column B (1): validated against Master when Col A is a Master value. */
  function isPartTabCol2Cell(tab, cellIndex) {
    if (!tab || isMasterTab(tab)) return false;
    const cols = tab.cols || 0;
    if (cols < 2) return false;
    if (!Number.isInteger(cellIndex) || cellIndex < 0 ||
        cellIndex >= (tab.cells ? tab.cells.length : 0)) {
      return false;
    }
    return (cellIndex % cols) === 1;
  }

  /** Unique trimmed non-empty Master Column B values for a given Master Column A (Master row order). */
  function masterCol2ValuesForCol1(col1Value) {
    const master = state.tabs.find(isMasterTab);
    const out = [];
    const a = String(col1Value == null ? '' : col1Value).trim();
    if (!master || (master.cols || 0) < 2 || !a) return out;
    const seen = Object.create(null);
    for (let r = 0; r < master.rows; r++) {
      const base = r * master.cols;
      if ((master.cells[base] || '').trim() !== a) continue;
      const v = (master.cells[base + 1] || '').trim();
      if (!v || seen[v]) continue;
      seen[v] = true;
      out.push(v);
    }
    return out;
  }

  /**
   * Col B is validated on this row when Col A is a Master Column A value that has
   * at least one Master Column B value. Empty Col A or no Master B → free text.
   */
  function isCol2ValidatedForRow(tab, row) {
    if (!tab || isMasterTab(tab) || (tab.cols || 0) < 2) return false;
    if (!Number.isInteger(row) || row < 0 || row >= tab.rows) return false;
    const colA = (tab.cells[row * tab.cols] || '').trim();
    if (!colA || !isCol1ValueAllowed(colA)) return false;
    return masterCol2ValuesForCol1(colA).length > 0;
  }

  function isCol2ValueAllowed(tab, idx, value, allowed) {
    const v = String(value == null ? '' : value).trim();
    if (!v) return true;
    if (!tab || !isPartTabCol2Cell(tab, idx)) return true;
    const row = Math.floor(idx / tab.cols);
    if (!isCol2ValidatedForRow(tab, row)) return true;
    const list = allowed || masterCol2ValuesForCol1(tab.cells[row * tab.cols]);
    return list.indexOf(v) !== -1;
  }

  function countInvalidCol2Cells(tab) {
    if (!tab || isMasterTab(tab) || (tab.cols || 0) < 2) return 0;
    let n = 0;
    for (let r = 0; r < tab.rows; r++) {
      if (!isCol2ValidatedForRow(tab, r)) continue;
      const idx = r * tab.cols + 1;
      if (!isCol2ValueAllowed(tab, idx, tab.cells[idx])) n++;
    }
    return n;
  }

  /** Master Column B cell index for (colA, colB) — first matching Master row, or -1. */
  function masterCol2IndexForValue(col1Value, col2Value) {
    const master = state.tabs.find(isMasterTab);
    const a = String(col1Value == null ? '' : col1Value).trim();
    const b = String(col2Value == null ? '' : col2Value).trim();
    if (!master || (master.cols || 0) < 2 || !a || !b) return -1;
    for (let r = 0; r < master.rows; r++) {
      const base = r * master.cols;
      if ((master.cells[base] || '').trim() !== a) continue;
      if ((master.cells[base + 1] || '').trim() === b) return base + 1;
    }
    return -1;
  }

  const col2FocusValues = Object.create(null);

  function onCol2CellFocusRemember(e) {
    const tab = activeTab();
    const ta = e.currentTarget;
    if (!tab || !ta) return;
    const idx = parseInt(ta.dataset.idx, 10);
    if (!isPartTabCol2Cell(tab, idx)) return;
    col2FocusValues[tab.id + ':' + idx] = tab.cells[idx] || '';
  }

  function flashCol2Invalid(idx) {
    flashCol1Invalid(idx); // same amber flash look
  }

  /** Typed/pasted Col B value not allowed for this row's Col A → revert. */
  function revertInvalidCol2Commit(tab, idx, ta) {
    if (!isPartTabCol2Cell(tab, idx)) return false;
    const value = tab.cells[idx] || '';
    if (isCol2ValueAllowed(tab, idx, value)) return false;
    const key = tab.id + ':' + idx;
    const prev = Object.prototype.hasOwnProperty.call(col2FocusValues, key) ? col2FocusValues[key] : '';
    writeCellCurrentPage(tab, idx, prev);
    if (ta) ta.value = prev;
    liveSyncConfirmedLinksForCell(tab.id, idx, null, { silent: true });
    applyAppendCheckedState();
    scheduleSave();
    flashCol2Invalid(idx);
    const colA = (tab.cells[Math.floor(idx / tab.cols) * tab.cols] || '').trim();
    setStatus('“' + value.trim() + '” is not a Master Column B for “' + colA +
      '” — reverted (use the ▾ list)', 'err');
    return true;
  }

  /** Pick a Master Column B value for a part-tab Col B cell (= Master insert + lock). */
  function pickCol2Value(tab, idx, value) {
    if (!tab || activeTab() !== tab || !isPartTabCol2Cell(tab, idx)) return;
    const row = Math.floor(idx / tab.cols);
    const colA = (tab.cells[row * tab.cols] || '').trim();
    const master = state.tabs.find(isMasterTab);
    const masterIdx = masterCol2IndexForValue(colA, value);
    if (!master || masterIdx < 0) return;
    if ((tab.cells[idx] || '').trim() === value && getCellLock(tab, idx)) {
      setStatus('Column B already “' + value + '”');
      return;
    }
    focusedCell = { tabId: tab.id, index: idx };
    insertMasterText(master.cells[masterIdx] || value, masterIdx);
  }


  /** Cell 'change' (blur after a typed/pasted edit, incl. Enter/Tab moves). */
  function onCellCommitMasterMatch(e) {
    const tab = activeTab();
    const ta = e.currentTarget;
    if (!tab || !ta || isMasterTab(tab)) return;
    const idx = parseInt(ta.dataset.idx, 10);
    if (Number.isNaN(idx)) return;
    if (revertInvalidCol1Commit(tab, idx, ta)) return;
    if (revertInvalidCol2Commit(tab, idx, ta)) return;
    if (getCellLock(tab, idx) || findMasterMatchForCell(tab, idx, tab.cells[idx]) < 0) return;
    if (cellHasOwnExtraContent(tab, idx)) {
      setStatus('Matches a Master cell — not linked (this cell has its own nests/pages)', 'err');
      return;
    }
    pushHistory();
    if (linkCellIfMasterMatch(tab, idx) !== 'linked') return;
    const addr = cellAddressFromIndex(tab, idx);
    // Render after focus settles so Enter/Tab/click targets keep focus.
    window.setTimeout(function () {
      if (activeTab() === tab) {
        const active = document.activeElement;
        const refocus = active && active.closest && el.cellGrid.contains(active) && active.dataset
          ? {
            idx: active.dataset.idx,
            nest: active.dataset.nest,
            start: active.selectionStart,
            end: active.selectionEnd
          }
          : null;
        renderGrid();
        flushTabPage(tab);
        if (refocus && refocus.idx != null) {
          const sel = refocus.nest != null
            ? 'textarea.cell-nest-input[data-idx="' + refocus.idx + '"][data-nest="' + refocus.nest + '"]'
            : 'textarea.cell[data-idx="' + refocus.idx + '"]';
          const next = el.cellGrid.querySelector(sel);
          if (next) {
            next.focus();
            try {
              if (refocus.start != null) next.setSelectionRange(refocus.start, refocus.end);
            } catch (err) { /* ignore */ }
          }
        }
      }
      renderCombinedPrompt();
      applyConfirmedCellHighlights();
      renderMasterLibrary();
      scheduleSave();
      setStatus('Matched Master text in ' + addr + ' — linked (Combined checked; locked until double-click unlock)', 'ok');
    }, 0);
  }

  function insertMasterText(text, masterCellIndex, opts) {
    opts = opts || {};
    const tab = activeTab();
    if (!tab || isMasterTab(tab) || tab.cells.length === 0) return;

    let index = -1;
    let addedRow = false;
    if (Number.isInteger(opts.index) && opts.index >= 0 && opts.index < tab.cells.length) {
      index = opts.index;
    } else if (focusedCell && focusedCell.tabId === tab.id &&
        Number.isInteger(focusedCell.index) &&
        focusedCell.index >= 0 && focusedCell.index < tab.cells.length) {
      index = focusedCell.index;
    } else {
      index = tab.cells.findIndex(function (cell) {
        return !cell || !cell.trim();
      });
    }

    const insertText = text == null ? '' : String(text);

    // v0.157: confirm only when the cell would lose the user's own text (or its
    // own nests/pages). Swapping one Master value for another never asks (Ash).
    if (index >= 0 && !opts.confirmed) {
      const cur = String(tab.cells[index] || '');
      const lock = getCellLock(tab, index);
      const linked = !!(lock && lock.masterOrigin === true);
      const ownText = !!cur.trim() && trimEndText(cur) !== trimEndText(insertText) &&
        !partCellHoldsMasterValue(tab, index);
      const ownExtras = !linked && siteHasOwnExtraContent(tab, index);
      if (ownText || ownExtras) {
        const addr = cellAddress(Math.floor(index / tab.cols), index % tab.cols);
        const dims = tab.cols + 'x' + tab.rows;
        const message = ownText
          ? ('Replace text in ' + addr + (ownExtras ? ' (its nests/pages too)' : '') + '?')
          : ('Replace nests/pages in ' + addr + '?');
        openActionConfirm({
          anchor: opts.anchor || cellWrapAt(index),
          focusBack: el.cellGrid && el.cellGrid.querySelector('textarea.cell[data-idx="' + index + '"]'),
          message: message,
          actionLabel: 'Replace',
          cancelStatus: 'Insert cancelled',
          ariaLabel: 'Confirm replace',
          onConfirm: function () {
            if (activeTab() !== tab || (tab.cols + 'x' + tab.rows) !== dims) {
              setStatus('Insert cancelled — grid changed', 'err');
              return;
            }
            insertMasterText(text, masterCellIndex,
              Object.assign({}, opts, { index: index, confirmed: true }));
          }
        });
        return;
      }
    }

    if (index < 0) {
      // No selected cell and no empty cell: never overwrite A1 — add a row at
      // the bottom (same as ⤓ when no empty row is left) and insert there.
      // appendRowToTab is the single undo step for the whole insert.
      const newRow = appendRowToTab(tab);
      index = newRow * tab.cols;
      addedRow = true;
    } else {
      pushHistory();
    }
    let masterIdx = Number.isInteger(masterCellIndex) && masterCellIndex >= 0
      ? masterCellIndex
      : findMasterCellIndexByText(insertText);
    if (masterIdx === null) masterIdx = -1;
    focusedCell = { tabId: tab.id, index: index };
    applyMasterLinkToCell(tab, index, masterIdx, insertText);
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
      addedRow
        ? 'Added Master text (+ nests) to ' + tab.title + ' in new row ' + (Math.floor(index / tab.cols) + 1) + ' (no empty cell; Combined checked; locked until double-click unlock)'
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
    // Grow-only while typing: keep a dragged / Fit height; Fit shrinks.
    const kept = Number(tab.rowHeights[row]);
    refitRowHeightAt(row);
    if (Number.isFinite(kept) && kept > tab.rowHeights[row]) {
      tab.rowHeights[row] = kept;
      applyHeightToGridRow(row, kept);
    }
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
      // v0.143: the Col A cell being typed in stays visible until the edit ends
      // (hiding a focused cell blurs it mid-typing).
      const pinned = !!(col1EditCell && col1EditCell.tabId === tab.id &&
        Math.floor(col1EditCell.index / tab.cols) === row);
      const hidden = !pinned && !rowMatchesFilter(tab, row);
      controls.classList.toggle('is-row-filtered', hidden);
      const wraps = el.cellGrid.querySelectorAll('.cell-wrap[data-row="' + row + '"]');
      for (let w = 0; w < wraps.length; w++) wraps[w].classList.toggle('is-row-filtered', hidden);
    }
    // v0.143: tint Col A cells whose value the Col A filter is showing.
    const filter = getGridFilterPrefs(tab.id).valueFilter;
    const col1Wraps = el.cellGrid.querySelectorAll('.cell-wrap[data-col="0"]');
    for (let i = 0; i < col1Wraps.length; i++) {
      const w = col1Wraps[i];
      const row = parseInt(w.dataset.row, 10);
      let on = false;
      if (filter !== null && !isMasterTab(tab) && !Number.isNaN(row)) {
        on = filter.has((tab.cells[row * tab.cols] || '').trim());
      }
      w.classList.toggle('col1-filter-match', on);
    }
  }

  /**
   * v0.143 click-to-filter on part-tab Col A. Same state as the Values
   * checkboxes (prefs.valueFilter): click a value → only that value; click it
   * again (filter is exactly that value) → all rows. Empty Col A cell → edit.
   * Keeps the clicked cell at the same screen spot (no jump).
   */
  function setCol1FilterKeepingAnchor(tab, nextFilter, anchorEl) {
    const prefs = getGridFilterPrefs(tab.id);
    const scroller = el.cellGrid ? el.cellGrid.parentElement : null;
    const before = anchorEl ? anchorEl.getBoundingClientRect().top : null;
    prefs.valueFilter = nextFilter;
    applyRowFilterVisibility();
    syncCol1FilterControls();
    if (col1FilterMenuOpen && popoverMenus.grid) buildCol1FilterMenu(tab, popoverMenus.grid);
    if (scroller && anchorEl && before !== null && anchorEl.isConnected) {
      const after = anchorEl.getBoundingClientRect().top;
      if (Math.abs(after - before) > 0.5) scroller.scrollTop += (after - before);
    }
    captureLiveGridScroll();
  }

  function toggleCol1ClickFilter(tab, idx, anchorEl) {
    if (!tab || !isPartTabCol1Cell(tab, idx)) return;
    const value = (tab.cells[idx] || '').trim();
    const prefs = getGridFilterPrefs(tab.id);
    const f = prefs.valueFilter;
    const isExactly = f !== null && f.size === 1 && f.has(value);
    if (isExactly) {
      setCol1FilterKeepingAnchor(tab, null, anchorEl);
      setStatus('Column A filter off — showing all rows', 'ok');
      return;
    }
    setCol1FilterKeepingAnchor(tab, new Set([value]), anchorEl);
    let n = 0;
    for (let r = 0; r < tab.rows; r++) if (rowMatchesFilter(tab, r)) n++;
    setStatus('Showing Column A “' + value + '” (' + n + ' row' + (n === 1 ? '' : 's') +
      ') — click it again or Esc to clear', 'ok');
  }

  function clearCol1ValueFilter(anchorEl) {
    const tab = activeTab();
    if (!tab || isMasterTab(tab)) return false;
    const prefs = getGridFilterPrefs(tab.id);
    if (prefs.valueFilter === null) return false;
    setCol1FilterKeepingAnchor(tab, null, anchorEl || null);
    setStatus('Column A filter off — showing all rows', 'ok');
    return true;
  }

  function beginCol1TypeEdit(ta, caretEnd) {
    const tab = activeTab();
    if (!tab || !ta) return false;
    const idx = parseInt(ta.dataset.idx, 10);
    if (!isPartTabCol1Cell(tab, idx) || isCellMasterLocked(tab, idx)) return false;
    col1EditCell = { tabId: tab.id, index: idx };
    ta.readOnly = false;
    ta.classList.add('col1-editing');
    if (caretEnd) {
      try { const len = ta.value.length; ta.setSelectionRange(len, len); } catch (err) { /* no-op */ }
    }
    return true;
  }

  function endCol1TypeEdit(ta) {
    col1EditCell = null;
    if (!ta) return;
    ta.classList.remove('col1-editing');
    const tab = activeTab();
    const idx = parseInt(ta.dataset.idx, 10);
    if (tab && isPartTabCol1Cell(tab, idx) && !isCellMasterLocked(tab, idx)) ta.readOnly = true;
    // Re-apply after the commit/revert of this blur has run.
    window.setTimeout(applyRowFilterVisibility, 0);
  }

  function onCol1CellClickFilter(e) {
    const ta = e.currentTarget;
    if (e.detail > 1 || suppressCellAutoCopy) return;
    if (ta.classList.contains('col1-editing') || ta.classList.contains('master-cell-editing')) return;
    if (cellRangeDrag && cellRangeDrag.mode) return;
    // Selecting text / dragging / Shift+click in Col A must not toggle the row filter.
    // Use the press's own gesture (not the leftover caret range — a plain click collapses
    // selection after mousedown but before click, so that check would race).
    const tab = activeTab();
    const idx = parseInt(ta.dataset.idx, 10);
    if (lastCellGesture && lastCellGesture.idx === idx && (lastCellGesture.moved || lastCellGesture.shift)) return;
    if (!tab || !isPartTabCol1Cell(tab, idx)) return;
    if (!(tab.cells[idx] || '').trim()) {
      // Nothing to filter by: an empty Col A cell just starts typing.
      if (beginCol1TypeEdit(ta, true)) setStatus('Type a Column A value (or pick from ▾)', 'ok');
      return;
    }
    toggleCol1ClickFilter(tab, idx, ta);
  }

  function onCol1CellKeyEdit(e) {
    if (e.defaultPrevented) return;
    const ta = e.currentTarget;
    if (ta.classList.contains('col1-editing')) {
      if (e.key === 'Escape') {
        e.preventDefault();
        endCol1TypeEdit(ta);
        setStatus('Column A edit done', 'ok');
      }
      return;
    }
    if (!ta.readOnly) return;
    const mod = e.ctrlKey || e.metaKey;
    const printable = !mod && !e.altKey && (e.key.length === 1 || e.key === 'Process' || e.isComposing);
    const pasteOrCut = mod && !e.altKey && (e.key === 'v' || e.key === 'V' || e.key === 'x' || e.key === 'X');
    const clearKey = !mod && !e.altKey && (e.key === 'Backspace' || e.key === 'Delete');
    if (e.key === 'F2') {
      e.preventDefault();
      beginCol1TypeEdit(ta, true);
      return;
    }
    // Typing / paste / delete edit in place as before. Chromium decides
    // editability before keydown handlers run, so insert the first key here
    // (execCommand keeps native undo + fires the normal input event).
    if (pasteOrCut) { beginCol1TypeEdit(ta, true); return; }
    if ((printable && e.key.length === 1) || clearKey) {
      if (!beginCol1TypeEdit(ta, true)) return;
      e.preventDefault();
      if (clearKey) document.execCommand(e.key === 'Delete' ? 'forwardDelete' : 'delete', false);
      else document.execCommand('insertText', false, e.key);
      return;
    }
    if (printable) beginCol1TypeEdit(ta, true);
  }

  function onCol1CellBlurEdit(e) {
    const ta = e.currentTarget;
    if (!ta.classList.contains('col1-editing')) return;
    endCol1TypeEdit(ta);
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
    // Keep the whole list reachable: cap to the space below the button, or
    // open upward when there is more room above (small windows, low panel).
    const below = window.innerHeight - rect.bottom - 4 - pad;
    const above = rect.top - 4 - pad;
    const useAbove = below < 180 && above > below;
    const room = Math.max(120, Math.min(320, useAbove ? above : below));
    menu.style.maxHeight = room + 'px';
    menu.style.top = useAbove
      ? Math.round(Math.max(pad, rect.top - 4 - room)) + 'px'
      : Math.round(rect.bottom + 4) + 'px';
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
    const menu = popoverMenus.grid;
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


  /** Brief flash on a cell-wrap after Values-name jump. */
  let col1JumpFlashTimer = 0;
  function flashCol1Jump(wrap) {
    if (!wrap) return;
    wrap.classList.remove('col1-jump-flash');
    void wrap.offsetWidth;
    wrap.classList.add('col1-jump-flash');
    if (col1JumpFlashTimer) clearTimeout(col1JumpFlashTimer);
    col1JumpFlashTimer = window.setTimeout(function () {
      wrap.classList.remove('col1-jump-flash');
      col1JumpFlashTimer = 0;
    }, 900);
  }

  /**
   * Values dropdown: click the name (not the checkbox) → ensure the value is
   * visible in the filter, scroll its first Col A cell near the top of the
   * grid, flash it, close the menu. Checkbox still only toggles the filter.
   */
  function jumpToCol1ValueInGrid(tab, value) {
    if (!tab || !el.cellGrid) return;
    const want = value == null ? '' : String(value);
    const prefs = getGridFilterPrefs(tab.id);
    if (prefs.valueFilter !== null && !prefs.valueFilter.has(want)) {
      const next = new Set(prefs.valueFilter);
      next.add(want);
      const all = uniqueCol1Values(tab);
      let allOn = next.size === all.length;
      if (allOn) {
        for (let i = 0; i < all.length; i++) {
          if (!next.has(all[i])) { allOn = false; break; }
        }
      }
      prefs.valueFilter = allOn ? null : next;
      applyRowFilterVisibility();
      syncCol1FilterControls();
    }
    closeCol1FilterMenu();

    let row = -1;
    for (let r = 0; r < tab.rows; r++) {
      if ((tab.cells[r * tab.cols] || '').trim() === want) {
        row = r;
        break;
      }
    }
    if (row < 0) {
      setStatus('No Column A cell for “' + (want || '(blank)') + '”', 'err');
      return;
    }
    const idx = row * tab.cols;
    const ta = el.cellGrid.querySelector('textarea.cell[data-idx="' + idx + '"]');
    const wrap = ta && ta.closest ? ta.closest('.cell-wrap') : null;
    const scroller = el.cellGrid.parentElement;
    if (wrap && scroller) {
      // Sticky header row sits above cells — leave a little room under it.
      const header = el.cellGrid.querySelector('.column-header');
      const headerH = header ? header.getBoundingClientRect().height : 28;
      const sRect = scroller.getBoundingClientRect();
      const wRect = wrap.getBoundingClientRect();
      scroller.scrollTop += (wRect.top - sRect.top) - headerH - 4;
    }
    flashCol1Jump(wrap);
    if (ta && typeof ta.focus === 'function') {
      try { ta.focus({ preventScroll: true }); } catch (err) { ta.focus(); }
    }
    setStatus('Jumped to Column A “' + (want || '(blank)') + '”', 'ok');
  }

  /**
   * Master-insert Values: jump inside the results pane to the first library
   * cell whose Master row Column A equals `value` (scroll near top + flash).
   */
  function jumpToMasterLibCol1Value(master, value) {
    if (!master) return;
    const want = value == null ? '' : String(value);
    const prefs = getMasterLibPrefs(masterLibPrefsPartId());
    if (prefs.valueFilter !== null && !prefs.valueFilter.has(want)) {
      prefs.valueFilter.add(want);
      const all = uniqueMasterLibCol1Values(master);
      let allOn = prefs.valueFilter.size === all.length;
      if (allOn) {
        for (let i = 0; i < all.length; i++) {
          if (!prefs.valueFilter.has(all[i])) { allOn = false; break; }
        }
      }
      if (allOn) prefs.valueFilter = null;
      const current = activeTab();
      refreshMasterLibraryResults(master, current);
      syncMasterLibFilterControls();
    }
    masterLibFilterMenuOpen = false;
    syncMasterLibFilterControls();

    let row = -1;
    for (let r = 0; r < master.rows; r++) {
      if ((master.cells[r * master.cols] || '').trim() === want) {
        row = r;
        break;
      }
    }
    if (row < 0) {
      setStatus('No Master Column A “' + (want || '(blank)') + '”', 'err');
      return;
    }
    const results = masterLibraryResultsEl();
    if (!results) return;
    const cell = results.querySelector(
      '.master-library-cell[data-row="' + row + '"][data-col="0"]'
    );
    if (!cell) {
      setStatus('Master “' + (want || '(blank)') + '” is hidden by the current filter', 'err');
      return;
    }
    const sRect = results.getBoundingClientRect();
    const cRect = cell.getBoundingClientRect();
    results.scrollTop += (cRect.top - sRect.top) - 4;
    cell.classList.remove('master-lib-jump-flash');
    void cell.offsetWidth;
    cell.classList.add('master-lib-jump-flash');
    window.setTimeout(function () { cell.classList.remove('master-lib-jump-flash'); }, 900);
    setStatus('Jumped to Master Column A “' + (want || '(blank)') + '”', 'ok');
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
        // div (not label): checkbox toggles filter; name click jumps to the cell.
        const row = document.createElement('div');
        row.className = 'column-col1-filter-option';

        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.checked = isCol1ValueSelected(value);
        if (value) cb.dataset.value = value;
        else cb.dataset.blank = '1';
        cb.title = 'Toggle filter for this Column A value';
        cb.setAttribute('aria-label', 'Filter Column A “' + (value || '(blank)') + '”');
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

        const nameBtn = document.createElement('button');
        nameBtn.type = 'button';
        nameBtn.className = 'column-col1-filter-option-text';
        nameBtn.textContent = value ? value : '(blank)';
        if (!value) nameBtn.classList.add('is-blank');
        nameBtn.title = 'Click to jump · checkbox to filter';
        nameBtn.addEventListener('click', function (e) {
          e.preventDefault();
          e.stopPropagation();
          jumpToCol1ValueInGrid(tab, value);
        });

        row.appendChild(cb);
        row.appendChild(nameBtn);
        list.appendChild(row);
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

  // Per tab/part-page grid scroll memory. One .cell-grid-wrap scroller is
  // shared by every tab, so without this each tab inherited the previous
  // tab's offset (clamped to its height → usually the top).
  const gridScrollPos = Object.create(null);
  let gridScrollKey = null;

  // Keys are "documentId|tabId:page" so documents with the same tab ids don't collide.
  // Persisted to userData grid-scroll.json (v0.138) so it survives restarts/updates.
  function gridScrollKeyFor(tab) {
    if (!tab) return null;
    const doc = activeDocument();
    return (doc ? doc.id : '') + '|' + tab.id + ':' + (isMasterTab(tab) ? 0 : (tab.page || 0));
  }

  let gridScrollSaveTimer = null;
  let gridScrollCaptureTimer = null;

  /** Record the live grid scroll for the current tab/page (skips while hidden). */
  function captureLiveGridScroll() {
    const wrap = el.cellGrid && el.cellGrid.parentElement;
    if (!wrap || !gridScrollKey || wrap.clientHeight <= 0) return;
    gridScrollPos[gridScrollKey] = { top: wrap.scrollTop, left: wrap.scrollLeft };
  }

  /** Nested { docId: { "tabId:page": [top, left] } } for open documents only. */
  function gridScrollPayload() {
    const open = Object.create(null);
    (state.documents || []).forEach(function (d) { open[d.id] = true; });
    const out = {};
    Object.keys(gridScrollPos).forEach(function (key) {
      const bar = key.indexOf('|');
      if (bar < 0) return;
      const docId = key.slice(0, bar);
      if (!open[docId]) return;
      const pos = gridScrollPos[key];
      if (!pos) return;
      if (!out[docId]) out[docId] = {};
      out[docId][key.slice(bar + 1)] = [Math.round(pos.top) || 0, Math.round(pos.left) || 0];
    });
    return out;
  }

  function seedGridScroll(raw) {
    if (!raw || typeof raw !== 'object') return;
    Object.keys(raw).forEach(function (docId) {
      const map = raw[docId];
      if (!map || typeof map !== 'object') return;
      Object.keys(map).forEach(function (k) {
        const v = map[k];
        const key = docId + '|' + k;
        if (gridScrollPos[key] || !Array.isArray(v)) return; // live memory wins
        gridScrollPos[key] = { top: Number(v[0]) || 0, left: Number(v[1]) || 0 };
      });
    });
  }

  function scheduleGridScrollSave() {
    if (!window.click2copy || typeof window.click2copy.saveGridScroll !== 'function') return;
    clearTimeout(gridScrollSaveTimer);
    gridScrollSaveTimer = setTimeout(function () {
      gridScrollSaveTimer = null;
      window.click2copy.saveGridScroll(gridScrollPayload()).catch(function () {});
    }, 800);
  }

  /** Immediate save (Restart). Returns a promise; never throws. */
  function flushGridScroll() {
    captureLiveGridScroll();
    clearTimeout(gridScrollSaveTimer);
    gridScrollSaveTimer = null;
    if (!window.click2copy || typeof window.click2copy.saveGridScroll !== 'function') return Promise.resolve(false);
    return Promise.resolve(window.click2copy.saveGridScroll(gridScrollPayload())).catch(function () { return false; });
  }

  function onGridScrolled() {
    clearTimeout(gridScrollCaptureTimer);
    gridScrollCaptureTimer = setTimeout(function () {
      captureLiveGridScroll();
      scheduleGridScrollSave();
    }, 150);
  }

  function restoreGridScroll(tab) {
    const wrap = el.cellGrid.parentElement;
    const key = gridScrollKeyFor(tab);
    gridScrollKey = key;
    if (!wrap || !key) return;
    const pos = gridScrollPos[key];
    const top = pos ? pos.top : 0;
    const left = pos ? pos.left : 0;
    wrap.scrollTop = top;
    wrap.scrollLeft = left;
    if (wrap.scrollTop === top && wrap.scrollLeft === left) return;
    // Late row fits can grow content after this frame — retry once.
    const clampedTop = wrap.scrollTop;
    const clampedLeft = wrap.scrollLeft;
    requestAnimationFrame(function () {
      if (gridScrollKey !== key) return;
      if (wrap.scrollTop !== clampedTop || wrap.scrollLeft !== clampedLeft) return;
      wrap.scrollTop = top;
      wrap.scrollLeft = left;
    });
  }

  /* ── Master Column B duplicate highlight (v0.141, display only) ─────────
   * Compares what each Master B cell shows now (its current cell page),
   * trimmed, inner whitespace collapsed, case-insensitive. Highlight only —
   * never blocks edits or changes data.
   */
  function masterDupeKey(value) {
    return String(value == null ? '' : value).trim().replace(/\s+/g, ' ').toLowerCase();
  }

  function computeMasterColBDupes(tab) {
    const out = { byIdx: Object.create(null), count: 0 };
    if (!tab || !isMasterTab(tab) || (tab.cols || 0) < 2) return out;
    const groups = Object.create(null);
    for (let r = 0; r < tab.rows; r++) {
      const idx = r * tab.cols + 1;
      if (idx >= tab.cells.length) break;
      const key = masterDupeKey(tab.cells[idx]);
      if (!key) continue;
      (groups[key] = groups[key] || []).push(idx);
    }
    Object.keys(groups).forEach(function (key) {
      const list = groups[key];
      if (list.length < 2) return;
      list.forEach(function (idx) {
        out.byIdx[idx] = list.filter(function (other) { return other !== idx; });
        out.count += 1;
      });
    });
    return out;
  }

  function applyMasterDupeMarks() {
    if (masterDupeTimer) {
      clearTimeout(masterDupeTimer);
      masterDupeTimer = null;
    }
    const tab = activeTab();
    if (!tab || !el.cellGrid) return;
    const isMaster = isMasterTab(tab);
    const dupes = isMaster ? computeMasterColBDupes(tab) : { byIdx: {}, count: 0 };
    const wraps = el.cellGrid.querySelectorAll('.cell-wrap.master-dupe, textarea.cell[data-col="1"]');
    for (let i = 0; i < wraps.length; i++) {
      const node = wraps[i];
      const ta = node.matches('textarea') ? node : node.querySelector('textarea.cell');
      const wrap = node.matches('textarea') ? node.closest('.cell-wrap') : node;
      if (!ta || !wrap) continue;
      const idx = parseInt(ta.dataset.idx, 10);
      const others = isMaster && !Number.isNaN(idx) ? dupes.byIdx[idx] : null;
      const on = !!(others && others.length);
      wrap.classList.toggle('master-dupe', on);
      if (!('baseTitle' in ta.dataset)) ta.dataset.baseTitle = ta.title || '';
      if (on) {
        ta.title = 'Duplicate of ' + others.map(function (o) {
          return cellAddressFromIndex(tab, o);
        }).join(', ') + ' (same text, ignoring case and extra spaces)';
      } else {
        ta.title = ta.dataset.baseTitle;
      }
    }
    const badge = el.cellGrid.querySelector('.master-dupe-count');
    if (badge) {
      badge.textContent = dupes.count ? '⚠ ' + dupes.count + ' dup' : '';
      badge.title = dupes.count
        ? dupes.count + ' Column B cells share their text with another B cell (highlighted; nothing is blocked)'
        : '';
      badge.style.visibility = dupes.count ? 'visible' : 'hidden';
    }
  }

  let masterDupeTimer = null;
  /** Debounced live refresh while typing (no re-render, so no layout bounce). */
  function scheduleMasterDupeMarks() {
    if (masterDupeTimer) clearTimeout(masterDupeTimer);
    masterDupeTimer = setTimeout(applyMasterDupeMarks, 160);
  }

  function renderGrid() {
    const tab = activeTab();
    const scrollWrap = el.cellGrid.parentElement;
    // Remember the outgoing view (skip while hidden, e.g. Tools view → reads 0).
    if (scrollWrap && gridScrollKey && scrollWrap.clientHeight > 0) {
      gridScrollPos[gridScrollKey] = { top: scrollWrap.scrollTop, left: scrollWrap.scrollLeft };
    }
    el.cellGrid.innerHTML = '';
    if (!tab) {
      gridScrollKey = null;
      return;
    }

    // The first grid row is a header row. Its handles resize the matching cell column.
    applyGridColumns(tab);
    // Col1 data validation list (part tabs) — Master Column A values.
    const col1Allowed = isMasterTab(tab) ? [] : masterCol1Values();
    const col1InvalidCount = isMasterTab(tab) ? 0 : countInvalidCol1Cells(tab);
    const col2InvalidCount = isMasterTab(tab) ? 0 : countInvalidCol2Cells(tab);

    const corner = document.createElement('div');
    corner.className = 'grid-corner';
    const jumpBtn = document.createElement('button');
    jumpBtn.type = 'button';
    jumpBtn.className = 'grid-jump-empty-btn';
    jumpBtn.textContent = '⤓';
    jumpBtn.title = 'Jump to next empty row';
    jumpBtn.setAttribute('aria-label', 'Jump to next empty row');
    jumpBtn.addEventListener('pointerdown', function (ev) {
      // Keep the current cell as the search start (focus moves to the target cell).
      ev.preventDefault();
    });
    jumpBtn.addEventListener('click', function (ev) {
      ev.preventDefault();
      jumpToNextEmptyRow();
    });
    corner.appendChild(jumpBtn);
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

      if (c === 1 && isMasterTab(tab)) {
        // Always present (visibility toggled) so the count never shifts the header.
        const dupBadge = document.createElement('span');
        dupBadge.className = 'master-dupe-count';
        dupBadge.style.visibility = 'hidden';
        header.appendChild(dupBadge);
      }
      if (c === 1 && !isMasterTab(tab) && col2InvalidCount > 0) {
        const bad = document.createElement('span');
        bad.className = 'col1-invalid-count';
        bad.textContent = '⚠ ' + col2InvalidCount;
        bad.title = col2InvalidCount + ' Column B value' + (col2InvalidCount === 1 ? '' : 's') +
          ' not allowed for that row\'s Column A (kept; marked). Pick from ▾ to fix.';
        header.appendChild(bad);
      }

      if (c === 0) {
        header.classList.add('column-header-sortable');
        if (col1InvalidCount > 0) {
          const bad = document.createElement('span');
          bad.className = 'col1-invalid-count';
          bad.textContent = '⚠ ' + col1InvalidCount;
          bad.title = col1InvalidCount + ' Column A value' + (col1InvalidCount === 1 ? '' : 's') +
            ' not in Master Column A (kept; marked). Pick from ▾ to fix.';
          header.appendChild(bad);
        }
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
        portalPopover('grid', filterMenu);
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
      rowNum.title = 'Row ' + (r + 1) + ' — right-click to clear';
      rowNum.setAttribute('aria-hidden', 'true');
      rowControls.appendChild(rowNum);

      // Right-click the row number / move controls (not the bottom resize handle).
      rowControls.addEventListener('contextmenu', function (ev) {
        const t = ev.target;
        if (t && t.closest && t.closest('.row-resize-handle')) return;
        ev.preventDefault();
        ev.stopPropagation();
        openRowClearMenu(r, ev.clientX, ev.clientY);
      });

      const moveControls = document.createElement('span');
      moveControls.className = 'row-move-controls';

      const moveUp = document.createElement('button');
      moveUp.type = 'button';
      moveUp.className = 'row-move';
      moveUp.textContent = '↑';
      moveUp.disabled = r === 0;
      moveUp.title = 'Move row ' + (r + 1) + ' up · hold + drag to move anywhere';
      moveUp.setAttribute('aria-label', 'Move row ' + (r + 1) + ' up');
      moveUp.addEventListener('pointerdown', function (ev) {
        if (moveUp.disabled) return;
        beginRowDragFromArrow(ev, r, moveUp);
      });
      moveUp.addEventListener('click', function () {
        moveRow(r, -1);
      });
      moveControls.appendChild(moveUp);

      const moveDown = document.createElement('button');
      moveDown.type = 'button';
      moveDown.className = 'row-move';
      moveDown.textContent = '↓';
      moveDown.title = 'Move row ' + (r + 1) + ' down · hold + drag to move anywhere';
      moveDown.setAttribute('aria-label', 'Move row ' + (r + 1) + ' down');
      moveDown.addEventListener('pointerdown', function (ev) {
        beginRowDragFromArrow(ev, r, moveDown);
      });
      moveDown.addEventListener('click', function () {
        moveRow(r, 1);
      });
      moveControls.appendChild(moveDown);

      rowControls.appendChild(moveControls);

      const rowResize = document.createElement('div');
      rowResize.className = 'row-resize-handle';
      rowResize.setAttribute('role', 'separator');
      rowResize.setAttribute('aria-orientation', 'horizontal');
      rowResize.setAttribute('aria-label', 'Resize row ' + (r + 1));
      rowResize.title = 'Drag to resize row ' + (r + 1);
      rowResize.addEventListener('pointerdown', function (e) {
        beginRowResize(e, r, rowResize);
      });
      rowControls.appendChild(rowResize);
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
        // Right-click the cell body → shade menu (page labels have their own menu;
        // nests have no shading and keep default behaviour).
        wrap.addEventListener('contextmenu', function (ev) {
          const t = ev.target;
          if (t && t.closest && (t.closest('.cell-nest') || t.closest('.cell-page-label'))) return;
          ev.preventDefault();
          ev.stopPropagation();
          openCellShadeMenu(idx, ev.clientX, ev.clientY);
        });
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
        nestAddBtn.title = 'Add nest under this cell (blank or copy current)';
        nestAddBtn.setAttribute('aria-label', 'Add nest under ' + cellAddress(r, c));
        nestAddBtn.innerHTML = '<span class="cell-nest-add-mark" aria-hidden="true">+</span>';
        if (isPartTabCol1Cell(tab, idx)) {
          // v0.143: Col A is filter-only — no nests. Keep the (invisible) slot so
          // the gutter/row layout is identical to other columns.
          nestAddBtn.classList.add('is-col1-off');
          nestAddBtn.disabled = true;
          nestAddBtn.tabIndex = -1;
          nestAddBtn.setAttribute('aria-hidden', 'true');
        }
        nestAddBtn.addEventListener('pointerdown', function (ev) {
          ev.preventDefault();
          ev.stopPropagation();
        });
        nestAddBtn.addEventListener('click', function (ev) {
          ev.preventDefault();
          ev.stopPropagation();
          const targetIdx = cellIndexFromNestAddEvent(ev, idx);
          if (targetIdx < 0) return;
          openAddPageChoiceMenu(nestAddBtn, { kind: 'nest', cellIndex: targetIdx });
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

        if (isPartTabCol1Cell(tab, idx)) {
          const ddBtn = document.createElement('button');
          ddBtn.type = 'button';
          ddBtn.className = 'col1-dropdown-btn';
          ddBtn.innerHTML = '<svg viewBox="0 0 12 12" width="12" height="12" aria-hidden="true" focusable="false">' +
            '<path d="M2.5 4.5 L6 8 L9.5 4.5" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"/></svg>';
          ddBtn.title = 'Pick a Master Column A value (Alt+↓)';
          ddBtn.setAttribute('aria-label', 'Pick Column A value for ' + cellAddress(r, c));
          ddBtn.addEventListener('pointerdown', function (ev) {
            ev.preventDefault();
            ev.stopPropagation();
          });
          ddBtn.addEventListener('click', function (ev) {
            ev.preventDefault();
            ev.stopPropagation();
            openCol1ValueMenu(ddBtn, idx);
          });
          corner.appendChild(ddBtn);
        }
        if (isPartTabCol2Cell(tab, idx)) {
          // Always reserve the ▾ slot (visibility toggled) so Col A changes never
          // jump Col B width / wrap (Ash).
          const dd2 = document.createElement('button');
          dd2.type = 'button';
          dd2.className = 'col1-dropdown-btn col2-dropdown-btn';
          dd2.innerHTML = '<svg viewBox="0 0 12 12" width="12" height="12" aria-hidden="true" focusable="false">' +
            '<path d="M2.5 4.5 L6 8 L9.5 4.5" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"/></svg>';
          const col2On = isCol2ValidatedForRow(tab, r);
          if (!col2On) {
            dd2.classList.add('is-off');
            dd2.disabled = true;
            dd2.tabIndex = -1;
            dd2.setAttribute('aria-hidden', 'true');
            dd2.title = '';
          } else {
            const colA = (tab.cells[r * tab.cols] || '').trim();
            dd2.title = 'Pick a Master Column B for “' + colA + '” (Alt+↓)';
            dd2.setAttribute('aria-label', 'Pick Column B value for ' + cellAddress(r, c));
            dd2.addEventListener('pointerdown', function (ev) {
              ev.preventDefault();
              ev.stopPropagation();
            });
            dd2.addEventListener('click', function (ev) {
              ev.preventDefault();
              ev.stopPropagation();
              openCol2ValueMenu(dd2, idx);
            });
          }
          corner.appendChild(dd2);
        }

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
        pageLabel.title = 'Cell page ' + (page + 1) + ' of ' + pageCount + ' — right-click to delete';
        pageLabel.addEventListener('contextmenu', function (ev) {
          ev.preventDefault();
          ev.stopPropagation();
          const blocked = isCellMasterEditBlocked(tab, idx);
          const onlyOne = pageCount <= 1;
          openPageCtxMenu({
            heading: 'Page ' + (page + 1) + ' (' + (page + 1) + '/' + pageCount + ')',
            disabled: onlyOne || blocked,
            disabledTitle: onlyOne
              ? 'Only one page — the last page cannot be deleted'
              : 'Locked Master cell — double-click to unlock before deleting a page',
            enabledTitle: 'Delete this cell page (Ctrl+Z to undo)',
            onDelete: function () { removeCellPage(idx, pageLabel); },
            x: ev.clientX,
            y: ev.clientY
          });
        });

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


        corner.appendChild(pageChrome);
        stack.appendChild(corner);

        const ta = document.createElement('textarea');
        ta.className = 'cell';
        if (isCellConfirmed(tab.id, idx)) ta.classList.add('cell-confirmed');
        if (isPartTabCol1Cell(tab, idx)) {
          ta.classList.add('cell-col1-validated');
          if (!isCol1ValueAllowed(tab.cells[idx], col1Allowed)) {
            wrap.classList.add('col1-invalid');
            ta.title = 'Not in Master Column A — keep or pick a value from ▾';
          }
        }
        if (isPartTabCol2Cell(tab, idx)) {
          // Always reserve ▾ padding on Col B (slot may be hidden when not validated).
          ta.classList.add('cell-col2-validated');
          if (isCol2ValidatedForRow(tab, r) && !isCol2ValueAllowed(tab, idx, tab.cells[idx])) {
            wrap.classList.add('col1-invalid');
            const colA = (tab.cells[r * tab.cols] || '').trim();
            ta.title = 'Not a Master Column B for “' + colA + '” — keep or pick a value from ▾';
          }
        }
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
        if (isPartTabCol1Cell(tab, idx)) {
          ta.classList.add('col1-click-filter');
          if (!cellEditing && !cellLocked) {
            const editingHere = col1EditCell && col1EditCell.tabId === tab.id && col1EditCell.index === idx;
            ta.readOnly = !editingHere;
            if (editingHere) ta.classList.add('col1-editing');
            if (!ta.title) ta.title = 'Click: filter rows to this value (again = clear) · type or F2 to edit · ▾ / Alt+↓ to pick';
          }
        }
        ta.addEventListener('input', onCellInput);
        ta.addEventListener('keydown', onCellKeydown);
        ta.addEventListener('pointerdown', onCellPointerDownSelect);
        ta.addEventListener('focus', onCellFocusSelect);
        ta.addEventListener('click', onCellClickSelect);
        ta.addEventListener('dblclick', onCellDblClickMasterUnlock);
        ta.addEventListener('blur', onCellBlurMasterFinish);
        ta.addEventListener('paste', onCellPaste);
        ta.addEventListener('change', onCellCommitMasterMatch);
        if (isPartTabCol1Cell(tab, idx)) {
          ta.addEventListener('focus', onCol1CellFocusRemember);
          ta.addEventListener('click', onCol1CellClickFilter);
          ta.addEventListener('keydown', onCol1CellKeyEdit);
          ta.addEventListener('blur', onCol1CellBlurEdit);
          ta.addEventListener('keydown', function (ev) {
            if (ev.altKey && ev.key === 'ArrowDown') {
              ev.preventDefault();
              const btn = ev.currentTarget.closest('.cell-wrap').querySelector('.col1-dropdown-btn:not(.col2-dropdown-btn)');
              if (btn) openCol1ValueMenu(btn, idx);
            }
          });
        }
        if (isPartTabCol2Cell(tab, idx)) {
          ta.addEventListener('focus', onCol2CellFocusRemember);
          ta.addEventListener('keydown', function (ev) {
            if (ev.altKey && ev.key === 'ArrowDown') {
              ev.preventDefault();
              const btn = ev.currentTarget.closest('.cell-wrap').querySelector('.col2-dropdown-btn:not(.is-off)');
              if (btn) openCol2ValueMenu(btn, idx);
            }
          });
        }
        ta.addEventListener('copy', onCellCopy);
        ta.addEventListener('cut', onCellCut);
        ta.addEventListener('beforeinput', onCellBeforeInputMasterLock);
        stack.appendChild(ta);

        // v0.143: Col A nests are not shown (data kept as-is, see isHiddenCol1NestLink).
        const nests = isPartTabCol1Cell(tab, idx) ? [] : getCellNests(tab, idx);
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
              pageLabel.title = 'Nest page ' + (page + 1) + ' of ' + pageCount + ' — right-click to delete';
              pageLabel.addEventListener('contextmenu', function (ev) {
                ev.preventDefault();
                ev.stopPropagation();
                const blocked = isCellMasterEditBlocked(tab, idx);
                const lockTip = 'Locked Master cell — double-click to unlock before deleting';
                openPageCtxMenu({
                  heading: 'Nest page ' + (page + 1) + ' (' + (page + 1) + '/' + pageCount + ')',
                  disabled: pageCount <= 1 || blocked,
                  pageDanger: false,
                  disabledTitle: blocked ? lockTip : 'Only one page — use Delete nest to remove it',
                  enabledTitle: 'Delete this nest page (Ctrl+Z to undo)',
                  onDelete: function () { removeNestPage(idx, nestIndex, pageLabel); },
                  extra: {
                    label: 'Delete nest',
                    disabled: blocked,
                    title: blocked ? lockTip : 'Delete this whole nest, all pages (Ctrl+Z to undo)',
                    onClick: function () { removeNestedCell(idx, nestIndex, pageLabel); }
                  },
                  x: ev.clientX,
                  y: ev.clientY
                });
              });

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


              nestChrome.appendChild(prevBtn);
              nestChrome.appendChild(pageLabel);
              nestChrome.appendChild(nextBtn);
              nestChrome.appendChild(addPageBtn);

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
    applyRowFilterVisibility(); // v0.143: Col A filter tint + edit-row pin
    applyAppendCheckedState();
    applyMasterDupeMarks();
    if (col1FilterMenuOpen) {
      const menu = popoverMenus.grid;
      if (menu) buildCol1FilterMenu(tab, menu);
    }
    syncCol1FilterControls();
    applyPersistedRowHeights(tab);
    restoreStickyCellRangeHighlight();
    applyPartSearchHighlights();
    restoreGridScroll(tab);
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

  /** Drag the row-number bottom edge → tab.rowHeights[row] (same store as Fit). */
  function beginRowResize(e, row, handle) {
    const tab = activeTab();
    if (!tab || !el.cellGrid || e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    const controls = el.cellGrid.querySelector('.row-controls[data-row="' + row + '"]');
    if (!controls) return;
    const startY = e.clientY;
    const startH = Math.max(MIN_ROW_HEIGHT, Math.round(controls.getBoundingClientRect().height));
    let started = false;
    document.body.classList.add('resizing-rows');
    handle.classList.add('is-active');

    function onMove(moveEvent) {
      const delta = moveEvent.clientY - startY;
      if (!started) {
        if (Math.abs(delta) < 2) return;
        started = true;
        pushHistory();
        // First manual size: pin every row at its current rendered height so
        // only the dragged row changes (no layout jump elsewhere).
        if (!normalizeRowHeights(tab.rowHeights, tab.rows)) {
          const heights = [];
          for (let r = 0; r < tab.rows; r++) {
            const rc = el.cellGrid.querySelector('.row-controls[data-row="' + r + '"]');
            const h = rc ? Math.round(rc.getBoundingClientRect().height) : MIN_ROW_HEIGHT;
            heights.push(Math.max(MIN_ROW_HEIGHT, h));
          }
          tab.rowHeights = heights;
        }
      }
      const next = Math.max(MIN_ROW_HEIGHT, Math.round(startH + delta));
      tab.rowHeights[row] = next;
      applyHeightToGridRow(row, next);
    }

    function finish() {
      document.removeEventListener('pointermove', onMove);
      document.removeEventListener('pointerup', finish);
      document.removeEventListener('pointercancel', finish);
      document.body.classList.remove('resizing-rows');
      handle.classList.remove('is-active');
      if (started) {
        scheduleSave();
        setStatus('Row ' + (row + 1) + ' height ' + tab.rowHeights[row] + 'px', 'ok');
      }
    }

    document.addEventListener('pointermove', onMove);
    document.addEventListener('pointerup', finish);
    document.addEventListener('pointercancel', finish);
    if (handle.setPointerCapture) {
      try { handle.setPointerCapture(e.pointerId); } catch (err) { /* no-op */ }
    }
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
  // drag that goes clearly into another cell selects a rectangle (never moves/swaps
  // cells). Shift+Arrow / Shift+click extend the sticky multi-cell highlight. A block
  // and a native text selection are never shown together.
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
  /** A text-select drag only becomes a cell block this far past the start cell's edge. */
  const CELL_DRIFT_PX = 20;
  /** Movement above this between press and release is a drag, not a click (Col A filter). */
  const CELL_CLICK_SLOP_PX = 4;
  /** Last press on a cell: { idx, shift, moved } — lets click handlers skip drags / Shift+click. */
  let lastCellGesture = null;
  /** Drag-start textarea of the last block drag: its leftover native selection is
   *  collapsed (not treated as a new text selection) until the next key/press. */
  let blockDragSelGuard = null;

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
    let filled = 0;
    for (let i = 0; i < indices.length; i++) {
      const v = tab.cells[indices[i]] || '';
      if (v !== '') anyContent = true;
      if (v.trim()) filled++;
    }
    if (!anyContent) return true;

    if (filled > 0) {
      const active = document.activeElement;
      const anchor = active && active.tagName === 'TEXTAREA' && el.cellGrid && el.cellGrid.contains(active)
        ? active : cellWrapAt(b.rMin * tab.cols + b.cMin);
      const dims = tab.cols + 'x' + tab.rows;
      openActionConfirm({
        anchor: anchor,
        message: 'Clear ' + pluralCells(filled) + '?' + linkedPartSuffix(tab, indices),
        actionLabel: 'Clear',
        cancelStatus: 'Clear cancelled',
        ariaLabel: 'Confirm clear',
        onConfirm: function () {
          if (activeTab() !== tab || (tab.cols + 'x' + tab.rows) !== dims) {
            setStatus('Clear cancelled — grid changed', 'err');
            return;
          }
          doClearStickyBlock(tab, b, indices);
        }
      });
      return true;
    }
    doClearStickyBlock(tab, b, indices);
    return true;
  }

  function doClearStickyBlock(tab, b, indices) {
    pushHistory();
    ensureCellPages(tab);
    for (let i = 0; i < indices.length; i++) {
      const idx = indices[i];
      writeCellCurrentPage(tab, idx, '');
      if (!isMasterTab(tab)) clearCellMasterLock(tab, idx);
      revalidateLinksForCell(tab.id, idx, { silent: true });
      // Same as right-click Clear: a Master clear updates linked part cells.
      if (isMasterTab(tab)) syncLockedPartPagesFromMaster(idx);
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

    // No lit block → start fresh from this cell (a stale anchor must not jump).
    if (!getStickyCellRange() || !keyboardRangeAnchor || keyboardRangeAnchor.tabId !== tab.id) {
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
    if (isMultiCellBounds(b)) {
      // Never show the block and a text selection together.
      collapseCellTextSelection();
      setStickyCellRange(b, tab.id);
    } else {
      clearStickyCellRange();
    }
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

  // Last block copied inside Click2Copy: its exact cell matrix, so a paste of
  // that same text keeps the block shape (even one column, or cells with line
  // breaks) while other plain text with line breaks stays one cell.
  let lastBlockCopy = null;

  function buildCellRangeTsv(tab, r0, c0, r1, c1) {
    const b = normalizeRangeBounds(r0, c0, r1, c1);
    const lines = [];
    const matrix = [];
    for (let r = b.rMin; r <= b.rMax; r++) {
      const cells = [];
      for (let c = b.cMin; c <= b.cMax; c++) {
        cells.push(tab.cells[r * tab.cols + c] || '');
      }
      matrix.push(cells.slice());
      lines.push(cells.join('\t'));
    }
    const text = lines.join('\n');
    lastBlockCopy = { text: text, matrix: matrix };
    return text;
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

  function collapseCellTextSelection(extraTa) {
    const act = document.activeElement;
    if (!act || act.tagName !== 'TEXTAREA') {
      // Document (non-textarea) selection only. removeAllRanges while a cell
      // textarea is focused leaves Chromium's cached textarea selection stale,
      // so the blue comes back on the next key/selection event.
      try {
        const sel = window.getSelection();
        if (sel && sel.removeAllRanges) sel.removeAllRanges();
      } catch (err) { /* no-op */ }
    }
    const targets = [act, extraTa];
    for (let i = 0; i < targets.length; i++) {
      const ta = targets[i];
      if (!ta || ta.tagName !== 'TEXTAREA' || !ta.classList.contains('cell')) continue;
      if (ta.selectionStart === ta.selectionEnd) continue;
      try {
        const pos = typeof ta.selectionStart === 'number' ? ta.selectionStart : 0;
        ta.setSelectionRange(pos, pos);
      } catch (err2) { /* no-op */ }
    }
  }

  function beginMultiCellRangeDrag() {
    if (!cellRangeDrag || cellRangeDrag.mode === 'multi') return;
    // Remember the in-cell text selection so returning to the start cell can
    // put it back (drifting past the edge must not lose the user's selection).
    const ta = cellRangeDrag.captureEl;
    if (ta && ta.tagName === 'TEXTAREA') {
      cellRangeDrag.savedSel = {
        start: ta.selectionStart,
        end: ta.selectionEnd,
        dir: ta.selectionDirection || 'none',
        scrollTop: ta.scrollTop
      };
    }
    cellRangeDrag.mode = 'multi';
    cellRangeDrag.multi = true;
    suppressCellAutoCopy = true;
    document.body.classList.add('selecting-cell-range');
    clearStickyCellRange();
    clearCellRelocateHighlight();
    collapseCellTextSelection();
  }

  /** Pointer came back into the start cell: drop the block, restore the text selection. */
  function returnToTextSelectDrag() {
    const drag = cellRangeDrag;
    if (!drag) return;
    drag.mode = null;
    drag.multi = false;
    drag.endRow = drag.startRow;
    drag.endCol = drag.startCol;
    suppressCellAutoCopy = false;
    document.body.classList.remove('selecting-cell-range');
    clearCellRangeHighlight();
    const ta = drag.captureEl;
    const s = drag.savedSel;
    if (ta && s && ta.isConnected && typeof ta.setSelectionRange === 'function') {
      try {
        ta.setSelectionRange(s.start, s.end, s.dir);
        ta.scrollTop = s.scrollTop;
      } catch (err) { /* no-op */ }
    }
  }

  function onCellRangePointerMove(e) {
    if (!cellRangeDrag || !cellRangeDrag.active) return;
    const dx = e.clientX - cellRangeDrag.startX;
    const dy = e.clientY - cellRangeDrag.startY;
    const distSq = dx * dx + dy * dy;
    if (distSq > CELL_CLICK_SLOP_PX * CELL_CLICK_SLOP_PX) cellRangeDrag.movedSlop = true;
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

    // Stay in native text-select while the pointer is in (or just past) the
    // start cell. The block only starts once the pointer is clearly inside
    // another cell (CELL_DRIFT_PX past the start cell's edge); coming back into
    // the start cell drops the block and restores the text selection.
    const startWrap = cellRangeDrag.startWrap;
    const sr = startWrap && startWrap.isConnected ? startWrap.getBoundingClientRect() : null;
    const x = e.clientX;
    const y = e.clientY;
    const insideStart = sr
      ? (x >= sr.left && x <= sr.right && y >= sr.top && y <= sr.bottom)
      : (row === cellRangeDrag.startRow && col === cellRangeDrag.startCol);
    const clearlyOut = sr
      ? (x < sr.left - CELL_DRIFT_PX || x > sr.right + CELL_DRIFT_PX ||
         y < sr.top - CELL_DRIFT_PX || y > sr.bottom + CELL_DRIFT_PX)
      : true;
    const leftStartCell =
      row !== cellRangeDrag.startRow || col !== cellRangeDrag.startCol;
    if (!cellRangeDrag.mode && moved && leftStartCell && clearlyOut) {
      beginMultiCellRangeDrag();
    }
    if (cellRangeDrag.mode === 'multi' && insideStart) {
      returnToTextSelectDrag();
      return;
    }
    if (cellRangeDrag.mode !== 'multi') return;
    // Chromium keeps extending the start textarea's native selection on every
    // mousemove while the button is held, so collapsing once at block start
    // let the blue selection come back under the teal block. Collapse each move.
    if (e.cancelable) e.preventDefault();
    const dragTa = cellRangeDrag.captureEl;
    collapseCellTextSelection(dragTa);
    // The compat mousemove that follows re-extends it again — collapse once more
    // before the next paint so the blue never shows under the teal block.
    window.requestAnimationFrame(function () {
      if (cellRangeDrag && cellRangeDrag.mode === 'multi') collapseCellTextSelection(dragTa);
    });

    if (row === cellRangeDrag.endRow && col === cellRangeDrag.endCol) {
      return;
    }
    cellRangeDrag.endRow = row;
    cellRangeDrag.endCol = col;
    if (e.cancelable) e.preventDefault();
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
    if (lastCellGesture) lastCellGesture.moved = !!drag.movedSlop;

    if (drag.mode === 'multi') {
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
      if (isMultiCellBounds(b)) {
        if (e.cancelable) e.preventDefault();
        suppressCellAutoCopy = true;
        const tab = activeTab();
        // Sticky multi-cell selection only — copy via Ctrl/Cmd+C (no auto-copy on release).
        setStickyCellRange(b, tab && tab.id);
        // Shift+Arrow / Shift+click keep growing this block from the start cell.
        if (tab) {
          keyboardRangeAnchor = { tabId: tab.id, row: drag.startRow, col: drag.startCol };
          keyboardRangeFocus = { row: drag.endRow, col: drag.endCol };
        }
        const startTa = drag.captureEl;
        blockDragSelGuard = startTa;
        collapseCellTextSelection(startTa);
        window.setTimeout(function () { collapseCellTextSelection(startTa); }, 0);
        window.setTimeout(function () {
          suppressCellAutoCopy = false;
          cellRangeDrag = null;
          collapseCellTextSelection(startTa);
          restoreStickyCellRangeHighlight();
        }, 120);
        return;
      }
      clearCellRangeHighlight();
      suppressCellAutoCopy = false;
      cellRangeDrag = null;
      return;
    }

    // Short click / in-cell text drag: leave the caret / text selection as is.
    clearCellRelocateHighlight();
    suppressCellAutoCopy = false;
    cellRangeDrag = null;
  }

  /**
   * Shift+click another cell: extend the block from the current cell (or the
   * existing block's anchor), Excel-style. Focus stays on the anchor cell.
   */
  function extendRangeToCellByShiftClick(ta, row, col) {
    const tab = activeTab();
    const active = document.activeElement;
    if (!tab || !el.cellGrid || !active || active === ta || active.tagName !== 'TEXTAREA' ||
        !active.classList.contains('cell') || !el.cellGrid.contains(active)) {
      return false;
    }
    const ar = parseInt(active.dataset.row, 10);
    const ac = parseInt(active.dataset.col, 10);
    if (Number.isNaN(ar) || Number.isNaN(ac)) return false;
    if (!getStickyCellRange() || !keyboardRangeAnchor || keyboardRangeAnchor.tabId !== tab.id) {
      keyboardRangeAnchor = { tabId: tab.id, row: ar, col: ac };
    }
    keyboardRangeFocus = { row: row, col: col };
    const b = normalizeRangeBounds(keyboardRangeAnchor.row, keyboardRangeAnchor.col, row, col);
    collapseCellTextSelection();
    if (isMultiCellBounds(b)) setStickyCellRange(b, tab.id);
    else clearStickyCellRange();
    return true;
  }

  function onCellPointerDownSelect(e) {
    if (e.button != null && e.button !== 0) return;
    const ta = e.currentTarget;
    const row = parseInt(ta.dataset.row, 10);
    const col = parseInt(ta.dataset.col, 10);
    if (Number.isNaN(row) || Number.isNaN(col)) return;
    if (cellRangeDrag && cellRangeDrag.active) {
      finishCellRangeListeners();
    }
    const idx = parseInt(ta.dataset.idx, 10);

    if (e.shiftKey && extendRangeToCellByShiftClick(ta, row, col)) {
      // Keep focus + caret on the anchor cell; no native text extend.
      e.preventDefault();
      cellRangeDrag = null;
      lastCellGesture = { idx: idx, shift: true, moved: false };
      return;
    }

    clearKeyboardCellRange();
    // A plain press always starts fresh: drop any lit block right away so a block
    // and a text selection never show together. Drag never moves/swaps cells (Ash).
    if (stickyCellRange) clearStickyCellRange();
    lastCellGesture = { idx: idx, shift: false, moved: false };

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
      startWrap: ta.closest ? ta.closest('.cell-wrap') : null,
      movedSlop: false,
      savedSel: null
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
    // Shift+click that extended the block keeps the anchor cell as the focused one.
    if (e.shiftKey && lastCellGesture && lastCellGesture.shift) return;
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
    liveSyncConfirmedLinksForCell(tab.id, idx, null, { dropEmpty: true });
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

  /**
   * Is the collapsed caret on the first / last *visual* line of the textarea?
   * Measures with a hidden mirror (same width/font/padding) so wrapped lines count.
   */
  function caretVisualLineEdges(ta) {
    const out = { first: true, last: true };
    if (!ta || ta.tagName !== 'TEXTAREA') return out;
    try {
      const cs = window.getComputedStyle(ta);
      const val = ta.value == null ? '' : String(ta.value);
      const pos = typeof ta.selectionStart === 'number' ? ta.selectionStart : 0;
      const padL = parseFloat(cs.paddingLeft) || 0;
      const padR = parseFloat(cs.paddingRight) || 0;
      const fontSize = parseFloat(cs.fontSize) || 12;
      const lh = parseFloat(cs.lineHeight) || fontSize * 1.4;
      const m = document.createElement('div');
      const st = m.style;
      st.position = 'absolute';
      st.visibility = 'hidden';
      st.left = '-10000px';
      st.top = '0';
      st.boxSizing = 'content-box';
      st.width = Math.max(1, ta.clientWidth - padL - padR) + 'px';
      st.padding = '0';
      st.border = '0';
      st.whiteSpace = 'pre-wrap';
      st.overflowWrap = 'break-word';
      st.wordBreak = cs.wordBreak;
      ['fontFamily', 'fontSize', 'fontWeight', 'fontStyle', 'letterSpacing', 'wordSpacing',
        'lineHeight', 'textTransform', 'textIndent', 'tabSize'].forEach(function (k) { st[k] = cs[k]; });
      m.appendChild(document.createTextNode(val.slice(0, pos)));
      const mark = document.createElement('span');
      m.appendChild(mark);
      m.appendChild(document.createTextNode(val.slice(pos) + (val.endsWith('\n') ? ' ' : '')));
      document.body.appendChild(m);
      const top = mark.offsetTop;
      const h = m.offsetHeight;
      m.remove();
      out.first = top < lh / 2;
      out.last = top + lh > h - lh / 2;
    } catch (err) { /* fall back to "at edge" */ }
    return out;
  }

  /** Shift+Arrow may start a cell block only from a real text edge (collapsed caret). */
  function caretAtShiftBlockEdge(ta, key) {
    if (!ta || ta.tagName !== 'TEXTAREA') return false;
    const start = ta.selectionStart;
    if (start !== ta.selectionEnd) return false;
    const len = (ta.value == null ? '' : String(ta.value)).length;
    if (key === 'ArrowLeft') return start === 0;
    if (key === 'ArrowRight') return start === len;
    if (key === 'ArrowUp') return start === 0 || caretVisualLineEdges(ta).first;
    if (key === 'ArrowDown') return start === len || caretVisualLineEdges(ta).last;
    return false;
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
      // Ctrl/Cmd+V reaches onCellPaste, which allows Master-value pastes only (v0.157).
      const modNav = (e.ctrlKey || e.metaKey) && (e.key === 'c' || e.key === 'C' ||
        e.key === 'a' || e.key === 'A' || e.key === 'v' || e.key === 'V');
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
        // Once a block is lit, Shift+Arrow always grows/shrinks the block (never text).
        // Otherwise a block starts only from a real text edge: Left/Right at the very
        // start/end, Up/Down on the first/last *visual* (wrapped) line.
        const sta = e.currentTarget;
        if (!getStickyCellRange() && !isWholeCellSelected(sta) && !caretAtShiftBlockEdge(sta, e.key)) return;
        e.preventDefault();
        extendKeyboardCellRange(sta, e.key);
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
    // Safety net: linked cells always show their current Master source.
    reportLockRepair(reconcileMasterLocks());
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
    // Keep the tab at its current width while renaming (input fills it, text
    // scrolls inside) so the tab row never re-wraps or bounces (Ash).
    const lockedWidth = btn.getBoundingClientRect().width; // exact (sub-pixel) — no 1px shift
    if (lockedWidth > 0) {
      btn.style.width = lockedWidth + 'px';
      btn.style.minWidth = lockedWidth + 'px';
      btn.style.maxWidth = lockedWidth + 'px';
    }
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
      btn.style.width = '';
      btn.style.minWidth = '';
      btn.style.maxWidth = '';
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
    // Our own block copy → paste back exactly that block.
    if (lastBlockCopy && plain != null &&
        String(plain).replace(/\r\n/g, '\n') === lastBlockCopy.text &&
        isMultiCellMatrix(lastBlockCopy.matrix)) {
      return lastBlockCopy.matrix.map(function (row) { return row.slice(); });
    }
    // Plain text with line breaks but no tabs is one cell's text (paragraphs),
    // not a column of cells — Ash. Tabs still mean a sheet-style block.
    const fromTsv = (plain != null && String(plain).indexOf('\t') >= 0) ? parseTsv(plain) : null;
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
      if (isMasterTab(tab)) remapMasterIndicesAfterColumnAdd(oldCols, tab.cols);
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
        // Pasting into Master must push to linked part cells like typing does.
        if (isMasterTab(tab)) syncLockedPartPagesFromMaster(idx);
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
    let startIdx = parseInt(e.currentTarget.dataset.idx, 10);
    if (Number.isNaN(startIdx) || startIdx < 0 || startIdx >= tab.cells.length) return;
    // A lit block owns the paste: it lands at the block's top-left, not at the
    // focused cell (which is the drag START — the bottom cell for a bottom-up drag,
    // so the upper cell used to stay blank).
    const pasteBlock = getStickyCellRange();
    let intoBlock = false;
    if (pasteBlock) {
      const fr = Math.floor(startIdx / tab.cols);
      const fc = startIdx % tab.cols;
      if (fr >= pasteBlock.rMin && fr <= pasteBlock.rMax && fc >= pasteBlock.cMin && fc <= pasteBlock.cMax) {
        startIdx = pasteBlock.rMin * tab.cols + pasteBlock.cMin;
        intoBlock = true;
      }
    }

    // v0.157: a locked Master cell may take a paste only when the pasted value
    // is itself a Master value (relinks to that Master cell, no warning — Ash).
    const startLocked = !isMasterTab(tab) && isCellMasterEditBlocked(tab, startIdx);
    const refuseLocked = function () {
      e.preventDefault();
      setStatus('Locked Master cell — double-click to unlock before editing', 'err');
    };

    let matrix = parseClipboardMatrix(e.clipboardData);
    if (!matrix && (intoBlock || startLocked) && e.clipboardData) {
      // Single value into a lit block → goes to the block's top-left cell.
      const plain = e.clipboardData.getData('text/plain');
      if (plain) matrix = [[plain.replace(/\r?\n$/, '')]];
    }
    if (startLocked && (!matrix || (!isMultiCellMatrix(matrix) && !isMasterValueText(matrix[0][0])))) {
      refuseLocked();
      return;
    }
    if (!matrix || (!isMultiCellMatrix(matrix) && !intoBlock && !startLocked)) {
      // Single-cell / plain text: let the textarea handle a normal paste.
      return;
    }

    e.preventDefault();
    const startRow = Math.floor(startIdx / tab.cols);
    const startCol = startIdx % tab.cols;
    // Col A / Col B data validation: values not allowed keep the cell's
    // previous value (never silently stored).
    let col1Rejected = 0;
    let col2Rejected = 0;
    const pastedRaw = matrix.map(function (row) { return row.slice(); });
    if (!isMasterTab(tab)) {
      const allowedA = masterCol1Values();
      for (let r = 0; r < matrix.length; r++) {
        for (let c = 0; c < matrix[r].length; c++) {
          const rr = startRow + r;
          const cc = startCol + c;
          const v = matrix[r][c];
          if (cc === 0) {
            if (isCol1ValueAllowed(v, allowedA)) continue;
            matrix[r][c] = rr < tab.rows ? (tab.cells[rr * tab.cols] || '') : '';
            col1Rejected++;
          } else if (cc === 1 && (tab.cols || 0) >= 2) {
            // Col A for this row: prefer the value being pasted into Col A, else current.
            let colA = '';
            if (startCol === 0) colA = String(matrix[r][0] == null ? '' : matrix[r][0]).trim();
            else if (rr < tab.rows) colA = (tab.cells[rr * tab.cols] || '').trim();
            if (!colA || !isCol1ValueAllowed(colA, allowedA)) continue;
            const allowedB = masterCol2ValuesForCol1(colA);
            if (!allowedB.length) continue;
            if (isCol2ValueAllowed(tab, rr * tab.cols + 1, v, allowedB)) continue;
            matrix[r][c] = (rr < tab.rows && tab.cols > 1) ? (tab.cells[rr * tab.cols + 1] || '') : '';
            col2Rejected++;
          }
        }
      }
    }
    const lockedTargets = [];
    if (!isMasterTab(tab)) {
      const pr = matrix.length;
      const pc = matrix[0].length;
      for (let r = 0; r < pr; r++) {
        for (let c = 0; c < pc; c++) {
          const rr = startRow + r;
          const cc = startCol + c;
          if (rr < tab.rows && cc < tab.cols) {
            const i = rr * tab.cols + cc;
            if (isCellMasterEditBlocked(tab, i)) {
              // Judge what the user pasted (before Col A/B validation reverts it).
              const v = pastedRaw[r][c] == null ? '' : String(pastedRaw[r][c]);
              if (!isMasterValueText(v)) {
                setStatus('Locked Master cell in paste range — double-click to unlock first', 'err');
                return;
              }
              lockedTargets.push({ idx: i, prevText: String(tab.cells[i] || ''),
                prevLock: Object.assign({}, getCellLock(tab, i)) });
            }
          }
        }
      }
    }
    const pasteAnchor = e.currentTarget;
    function doPaste() {
      pushHistory();
      const result = pasteMatrixAt(tab, startRow, startCol, matrix);
      if (!result) return;
      // Push pasted text into the visible textareas now: the focused one would
      // otherwise "win" over the model in flushLiveCellInputs (run while Master
      // links are applied below) and put the old text back.
      if (el.cellGrid) {
        for (let r = 0; r < result.rows; r++) {
          for (let c = 0; c < result.cols; c++) {
            const pi = (startRow + r) * tab.cols + (startCol + c);
            const pta = el.cellGrid.querySelector('textarea.cell[data-idx="' + pi + '"]');
            if (pta && pi < tab.cells.length) pta.value = tab.cells[pi] || '';
          }
        }
      }
      // Pasted cells that exactly match Master text become real Master links.
      let pasteLinked = 0;
      let pasteBlocked = 0;
      if (!isMasterTab(tab) && lockedTargets.length) {
        // Locked cells took a Master value: relink/lock to that Master cell
        // (nests + pages follow, like a ▾ pick). Same text → keep the old link.
        const masterTab = state.tabs.find(isMasterTab);
        lockedTargets.forEach(function (lt) {
          const now = String(tab.cells[lt.idx] || '');
          if (trimEndText(now) === trimEndText(lt.prevText)) {
            ensureCellLocks(tab);
            tab.cellLocks[lt.idx] = lt.prevLock;
            return;
          }
          const mIdx = findMasterMatchForCell(tab, lt.idx, now);
          if (mIdx < 0 || !masterTab) return;
          applyMasterLinkToCell(tab, lt.idx, mIdx, masterTab.cells[mIdx] || '');
          pasteLinked++;
        });
      }
      if (!isMasterTab(tab)) {
        for (let r = 0; r < result.rows; r++) {
          for (let c = 0; c < result.cols; c++) {
            const outcome = linkCellIfMasterMatch(tab, (startRow + r) * tab.cols + (startCol + c));
            if (outcome === 'linked') pasteLinked++;
            else if (outcome === 'blocked') pasteBlocked++;
          }
        }
        if (pasteLinked) {
          renderCombinedPrompt();
          applyConfirmedCellHighlights();
        }
      }

      focusedCell = { tabId: tab.id, index: startIdx };
      renderTabs();
      renderGrid();
      renderMasterLibrary();
      scheduleSave();
      const cell = el.cellGrid.querySelector('textarea.cell[data-idx="' + startIdx + '"]');
      if (cell) cell.focus();
      setStatus(
        'Pasted ' + result.rows + '×' + result.cols + ' cells at ' + cellAddress(startRow, startCol) +
        ' (' + tab.cols + '×' + tab.rows + ' grid)' +
        (pasteLinked ? ' — ' + pasteLinked + ' matched Master (linked + locked)' : '') +
        (pasteBlocked ? ' — ' + pasteBlocked + ' Master match(es) not linked (own nests/pages)' : '') +
        (col1Rejected ? ' — ' + col1Rejected + ' Column A value(s) not in Master kept previous' : '') +
        (col2Rejected ? ' — ' + col2Rejected + ' Column B value(s) not allowed for that Column A kept previous' : ''),
        (col1Rejected || col2Rejected) ? 'err' : 'ok'
      );
    }

    // Ash: confirm before a multi-cell paste overwrites filled cells (own text
    // only — Master value → Master value swaps and identical text don't ask).
    const filled = countFilledCellsInPasteRange(tab, startRow, startCol, matrix);
    if (filled > 0) {
      const dims = tab.cols + 'x' + tab.rows;
      openPasteOverwriteConfirm(pasteAnchor, filled, function () {
        if (activeTab() !== tab || (tab.cols + 'x' + tab.rows) !== dims) {
          setStatus('Paste cancelled — grid changed', 'err');
          return;
        }
        doPaste();
      });
      return;
    }
    doPaste();
  }

  function countFilledCellsInPasteRange(tab, startRow, startCol, matrix) {
    let filled = 0;
    for (let r = 0; r < matrix.length; r++) {
      for (let c = 0; c < matrix[r].length; c++) {
        const rr = startRow + r;
        const cc = startCol + c;
        if (rr >= tab.rows || cc >= tab.cols) continue;
        const idx = rr * tab.cols + cc;
        const cur = String(tab.cells[idx] || '');
        if (!cur.trim()) continue;
        const next = matrix[r][c] == null ? '' : String(matrix[r][c]);
        // Same text → nothing lost (also covers validation-rejected values).
        if (trimEndText(cur) === trimEndText(next)) continue;
        // Ash: swapping one Master value for another is fine (part tabs).
        if (partCellHoldsMasterValue(tab, idx) && isMasterValueText(next)) continue;
        filled++;
      }
    }
    return filled;
  }

  let pasteConfirmEl = null;
  let pasteConfirmCancelStatus = 'Paste cancelled';
  let pasteConfirmFocusBack = null;

  function closePasteOverwriteConfirm() {
    if (pasteConfirmEl && pasteConfirmEl.parentNode) pasteConfirmEl.parentNode.removeChild(pasteConfirmEl);
    pasteConfirmEl = null;
    document.removeEventListener('pointerdown', onPasteConfirmOutside, true);
    document.removeEventListener('keydown', onPasteConfirmKey, true);
  }

  function cancelActionConfirm(refocus) {
    const status = pasteConfirmCancelStatus;
    const back = pasteConfirmFocusBack;
    closePasteOverwriteConfirm();
    setStatus(status);
    if (refocus && back && back.isConnected && typeof back.focus === 'function') back.focus();
  }

  function onPasteConfirmOutside(ev) {
    if (pasteConfirmEl && pasteConfirmEl.contains(ev.target)) return;
    cancelActionConfirm(false);
  }

  function onPasteConfirmKey(ev) {
    if (!pasteConfirmEl) return;
    if (ev.key === 'Escape') {
      ev.preventDefault();
      ev.stopPropagation();
      if (ev.stopImmediatePropagation) ev.stopImmediatePropagation();
      cancelActionConfirm(true);
    }
  }

  /** Anchor → viewport rect: an element, a rect-like {left,top,bottom}, or a point {x,y}. */
  function confirmAnchorRect(anchor) {
    if (anchor && typeof anchor.getBoundingClientRect === 'function' && anchor.isConnected !== false) {
      const r = anchor.getBoundingClientRect();
      if (r.width || r.height) return r;
    }
    if (anchor && typeof anchor.x === 'number' && typeof anchor.y === 'number' &&
        typeof anchor.getBoundingClientRect !== 'function') {
      return { left: anchor.x, right: anchor.x, top: anchor.y, bottom: anchor.y };
    }
    if (anchor && typeof anchor.left === 'number' && typeof anchor.getBoundingClientRect !== 'function') {
      const top = typeof anchor.top === 'number' ? anchor.top : 20;
      return { left: anchor.left, right: anchor.left, top: top,
        bottom: typeof anchor.bottom === 'number' ? anchor.bottom : top };
    }
    return { left: 20, right: 20, top: 20, bottom: 20 };
  }

  /**
   * v0.157 shared overwrite/destroy confirm — the paste-confirm popover (same
   * look): "<message>" + action button + Cancel. Cancel is focused (Enter on it
   * cancels, as before); Esc / outside press cancel. position:fixed → no layout
   * shift. opts: anchor, message, actionLabel, onConfirm, cancelStatus, ariaLabel,
   * focusBack (refocused on Cancel/Esc; defaults to the focused element).
   */
  function openActionConfirm(opts) {
    closePasteOverwriteConfirm();
    const prevFocus = document.activeElement && document.activeElement !== document.body
      ? document.activeElement : null;
    pasteConfirmFocusBack = opts.focusBack || prevFocus;
    pasteConfirmCancelStatus = opts.cancelStatus || 'Cancelled — nothing changed';
    const pop = document.createElement('div');
    pop.className = 'paste-confirm';
    pop.setAttribute('role', 'alertdialog');
    pop.setAttribute('aria-label', opts.ariaLabel || 'Confirm');
    const msg = document.createElement('span');
    msg.className = 'paste-confirm-text';
    msg.textContent = opts.message;
    const ok = document.createElement('button');
    ok.type = 'button';
    ok.className = 'btn btn-danger paste-confirm-ok';
    ok.textContent = opts.actionLabel || 'OK';
    const cancel = document.createElement('button');
    cancel.type = 'button';
    cancel.className = 'btn btn-secondary paste-confirm-cancel';
    cancel.textContent = 'Cancel';
    ok.addEventListener('click', function (ev) {
      ev.preventDefault();
      const back = pasteConfirmFocusBack;
      closePasteOverwriteConfirm();
      opts.onConfirm();
      if (back && back.isConnected && typeof back.focus === 'function' &&
          (!document.activeElement || document.activeElement === document.body)) {
        back.focus();
      }
    });
    cancel.addEventListener('click', function (ev) {
      ev.preventDefault();
      cancelActionConfirm(true);
    });
    pop.appendChild(msg);
    pop.appendChild(ok);
    pop.appendChild(cancel);
    document.body.appendChild(pop);
    pasteConfirmEl = pop;
    const rect = confirmAnchorRect(opts.anchor);
    const size = pop.getBoundingClientRect();
    const pad = 6;
    const left = Math.max(pad, Math.min(rect.left, window.innerWidth - size.width - pad));
    let top = rect.bottom + 4;
    if (top + size.height > window.innerHeight - pad) top = Math.max(pad, rect.top - size.height - 4);
    pop.style.left = left + 'px';
    pop.style.top = top + 'px';
    document.addEventListener('pointerdown', onPasteConfirmOutside, true);
    document.addEventListener('keydown', onPasteConfirmKey, true);
    cancel.focus();
  }

  function openPasteOverwriteConfirm(anchorEl, count, onConfirm) {
    openActionConfirm({
      anchor: anchorEl,
      focusBack: anchorEl,
      message: 'Overwrite ' + count + ' filled cell' + (count === 1 ? '' : 's') + '?',
      actionLabel: 'Overwrite',
      cancelStatus: 'Paste cancelled',
      ariaLabel: 'Confirm paste overwrite',
      onConfirm: onConfirm
    });
  }

  function pluralCells(n, word) {
    return n + ' ' + (word || 'filled cell') + (n === 1 ? '' : 's');
  }

  /** Text exactly equal (trailing whitespace ignored) to some non-empty Master cell. */
  function isMasterValueText(text) {
    const want = trimEndText(text);
    if (!want.trim()) return false;
    const master = state.tabs.find(isMasterTab);
    if (!master) return false;
    for (let i = 0; i < master.cells.length; i++) {
      if (trimEndText(master.cells[i] || '') === want) return true;
    }
    return false;
  }

  /** Part cell currently holding Master text: a Master link, or exact Master value text. */
  function partCellHoldsMasterValue(tab, idx) {
    if (!tab || isMasterTab(tab)) return false;
    const lock = getCellLock(tab, idx);
    if (lock && lock.masterOrigin === true) return true;
    return isMasterValueText(tab.cells[idx] || '');
  }

  /** Part cells (live + stored part pages) linked to any of these Master indices. */
  function countLinkedPartCells(masterIndices) {
    const want = Object.create(null);
    (masterIndices || []).forEach(function (i) { want[i] = true; });
    let n = 0;
    forEachPartLockSite(function (site, i, lock) {
      if (Number.isInteger(lock.masterCellIndex) && want[lock.masterCellIndex]) n++;
    });
    return n;
  }

  /** Master tab: " (N linked part cells will update)" for clear/replace confirms. */
  function linkedPartSuffix(tab, indices, verb) {
    if (!tab || !isMasterTab(tab)) return '';
    const n = countLinkedPartCells(indices);
    if (!n) return '';
    return ' (' + pluralCells(n, 'linked part cell') + ' will ' + (verb || 'update') + ')';
  }

  function cellWrapAt(index) {
    return el.cellGrid ? el.cellGrid.querySelector('.cell-wrap[data-idx="' + index + '"]') ||
      (function () {
        const ta = el.cellGrid.querySelector('textarea.cell[data-idx="' + index + '"]');
        return ta && ta.closest ? ta.closest('.cell-wrap') : ta;
      })() : null;
  }

  function rowNumberAt(rowIndex) {
    return el.cellGrid ? el.cellGrid.querySelector('.row-controls[data-row="' + rowIndex + '"]') : null;
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
    appendRowToTab(tab);
    setStatus('Row added to ' + tab.title + ' (' + tab.cols + '×' + tab.rows + ')');
  }

  /** Tools '+ Row' core: one undo step, appends an empty row, re-renders. Returns new row index. */
  function appendRowToTab(tab) {
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
    return tab.rows - 1;
  }

  /**
   * Jump target rule: every cell in the row is empty on ALL of its cell pages
   * (current part page) and has no nest content. Column A counts (a row with
   * only a Col1 category is not empty). Rows hidden by the row/Values filter
   * are skipped.
   */
  function rowIsFullyEmpty(tab, rowIndex) {
    if (!rowIsEmpty(tab, rowIndex)) return false;
    ensureCellPagesArray(tab);
    for (let col = 0; col < tab.cols; col++) {
      const nests = tab.nestedCells[rowIndex * tab.cols + col] || [];
      for (let n = 0; n < nests.length; n++) {
        const np = (nests[n] && nests[n].pages) || [];
        for (let p = 0; p < np.length; p++) {
          if (String(np[p] || '').trim()) return false;
        }
      }
      const entry = tab.cellPages[rowIndex * tab.cols + col];
      if (!entry || !Array.isArray(entry.pages)) continue;
      for (let p = 0; p < entry.pages.length; p++) {
        if (String(entry.pages[p] || '').trim()) return false;
      }
    }
    return true;
  }

  /** Corner ⤓: instant scroll + caret to the next empty row (wraps to top). */
  function jumpToNextEmptyRow() {
    const tab = activeTab();
    const wrap = el.cellGrid && el.cellGrid.parentElement;
    if (!tab || !wrap || !tab.rows) return;
    // Start below the selected row; else at the first row visible under the header.
    let start = 0;
    if (focusedCell && focusedCell.tabId === tab.id && Number.isInteger(focusedCell.index) &&
        focusedCell.index >= 0 && focusedCell.index < tab.cells.length) {
      start = Math.floor(focusedCell.index / tab.cols) + 1;
    } else {
      const header = el.cellGrid.querySelector('.grid-corner');
      const visibleTop = wrap.getBoundingClientRect().top + (header ? header.getBoundingClientRect().height : 0);
      start = tab.rows;
      for (let r = 0; r < tab.rows; r++) {
        const rc = el.cellGrid.querySelector('.row-controls[data-row="' + r + '"]');
        if (rc && rc.getBoundingClientRect().top >= visibleTop - 2) {
          start = r;
          break;
        }
      }
    }
    function usable(r) {
      if (!rowMatchesFilter(tab, r)) return false;
      return rowIsFullyEmpty(tab, r);
    }
    let target = -1;
    let wrapped = false;
    for (let r = start; r < tab.rows; r++) {
      if (usable(r)) { target = r; break; }
    }
    if (target < 0) {
      for (let r = 0; r < Math.min(start, tab.rows); r++) {
        if (usable(r)) { target = r; wrapped = true; break; }
      }
    }
    let added = false;
    let filterNote = '';
    if (target < 0) {
      // No visible empty row: add one via the Tools '+ Row' path (single undo step).
      target = appendRowToTab(tab);
      added = true;
      // Make sure the new row is visible: Filled/Combined hide empty rows, and a
      // Values filter that excludes blank Column A hides it too.
      const prefs = getGridFilterPrefs(tab.id);
      const notes = [];
      if (prefs.rowFilter !== 'all' && !rowMatchesFilter(tab, target)) {
        prefs.rowFilter = 'all';
        syncRowFilterButtons();
        notes.push('row filter set to All');
      }
      if (prefs.valueFilter && !prefs.valueFilter.has('')) {
        prefs.valueFilter.add('');
        if (typeof syncCol1FilterControls === 'function') syncCol1FilterControls();
        notes.push('blank Column A added to Values filter');
      }
      if (notes.length) {
        filterNote = ' (' + notes.join(', ') + ' so it shows)';
        renderGrid();
      }
    }
    const rowEl = el.cellGrid.querySelector('.row-controls[data-row="' + target + '"]');
    const header = el.cellGrid.querySelector('.grid-corner');
    if (rowEl) {
      const headerH = header ? header.getBoundingClientRect().height : 0;
      const delta = rowEl.getBoundingClientRect().top - wrap.getBoundingClientRect().top - headerH - 4;
      wrap.scrollTop = Math.max(0, wrap.scrollTop + delta);
      wrap.scrollLeft = 0;
    }
    // First editable cell; on part tabs skip the validated Col1 (dropdown-only)
    // when another editable column exists, so typing isn't reverted.
    let focusIdx = -1;
    let col1Fallback = -1;
    for (let c = 0; c < tab.cols; c++) {
      const idx = target * tab.cols + c;
      if (isCellMasterEditBlocked(tab, idx)) continue;
      if (isPartTabCol1Cell(tab, idx)) {
        if (col1Fallback < 0) col1Fallback = idx;
        continue;
      }
      focusIdx = idx;
      break;
    }
    if (focusIdx < 0) focusIdx = col1Fallback;
    if (focusIdx >= 0) {
      const ta = el.cellGrid.querySelector('textarea.cell[data-idx="' + focusIdx + '"]');
      if (ta) {
        ta.focus({ preventScroll: true });
        try { ta.setSelectionRange(0, 0); } catch (err) { /* ignore */ }
      }
    }
    setStatus(added
      ? ('Added row ' + (target + 1) + filterNote)
      : ('Row ' + (target + 1) + ' is empty' + (wrapped ? ' (wrapped to top)' : '')), 'ok');
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
   * Extract row `from` and insert it at index `to` (0..rows-1); rows between
   * shift by one. Cells, nests, locks, cell pages, shades, rowHeights,
   * Combined links and (Master tab) part Master indices all follow the row.
   * No history push / render — callers do that (one undo step per gesture).
   */
  function reorderRowInTab(tab, from, to) {
    const rows = tab.rows;
    const cols = tab.cols;
    if (from === to || from < 0 || to < 0 || from >= rows || to >= rows) return false;
    const order = [];
    for (let r = 0; r < rows; r++) order.push(r);
    order.splice(from, 1);
    order.splice(to, 0, from);

    ensureNestedCells(tab);
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
    return true;
  }

  /** Drag-drop row move (hold ↑/↓ + drag): one undo step, same data path as ↑. */
  function moveRowTo(from, to) {
    const tab = activeTab();
    if (!tab) return false;
    if (from === to || from < 0 || to < 0 || from >= tab.rows || to >= tab.rows) return false;
    pushHistory();
    reorderRowInTab(tab, from, to);
    renderTabs();
    renderGrid();
    renderMasterLibrary();
    scheduleSave();
    setStatus('Row ' + (from + 1) + ' moved to row ' + (to + 1), 'ok');
    return true;
  }

  /* ── Row drag (press + hold on ↑/↓, then drag vertically) ───────────── */
  const ROW_DRAG_THRESHOLD_PX = 4;
  const ROW_DRAG_EDGE_PX = 44;
  let rowDrag = null;

  function rowDragScroller() {
    return el.cellGrid ? el.cellGrid.parentElement : null;
  }

  /** Visible row-control boxes (filtered rows skipped). */
  function rowDragVisibleRows() {
    if (!el.cellGrid) return [];
    const out = [];
    const list = el.cellGrid.querySelectorAll('.row-controls');
    for (let i = 0; i < list.length; i++) {
      const rc = list[i];
      if (rc.classList.contains('is-row-filtered')) continue;
      const rect = rc.getBoundingClientRect();
      if (!rect.height) continue;
      out.push({ row: parseInt(rc.dataset.row, 10), rect: rect });
    }
    return out;
  }

  /**
   * Slot under the pointer: insert-before row index (rows = after last).
   * Returns { slot, y } where y is the viewport y of the indicator line.
   */
  function rowDragSlotAt(clientY) {
    const rows = rowDragVisibleRows();
    if (!rows.length) return null;
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i].rect;
      if (clientY < r.top + r.height / 2) {
        return { slot: rows[i].row, y: r.top };
      }
    }
    const last = rows[rows.length - 1];
    return { slot: last.row + 1, y: last.rect.bottom };
  }

  function rowDragUpdateIndicator() {
    if (!rowDrag || !rowDrag.active) return;
    const hit = rowDragSlotAt(rowDrag.lastY);
    const ind = rowDrag.indicator;
    if (!hit) { ind.hidden = true; rowDrag.slot = null; return; }
    rowDrag.slot = hit.slot;
    const gridRect = el.cellGrid.getBoundingClientRect();
    const scroller = rowDragScroller();
    const sRect = scroller ? scroller.getBoundingClientRect() : gridRect;
    const left = Math.max(gridRect.left, sRect.left);
    const right = Math.min(gridRect.right, sRect.right);
    const y = Math.max(sRect.top, Math.min(sRect.bottom, hit.y));
    ind.hidden = false;
    ind.style.left = left + 'px';
    ind.style.width = Math.max(0, right - left) + 'px';
    ind.style.top = (y - 1) + 'px';
    const to = rowDragTargetIndex();
    ind.classList.toggle('is-noop', to === rowDrag.from);
    if (rowDrag.ghost) {
      rowDrag.ghost.style.top = (rowDrag.lastY - rowDrag.ghostOffset) + 'px';
    }
  }

  /** Final destination index after removing the source row. */
  function rowDragTargetIndex() {
    if (!rowDrag || rowDrag.slot == null) return rowDrag ? rowDrag.from : -1;
    const tab = activeTab();
    let to = rowDrag.slot > rowDrag.from ? rowDrag.slot - 1 : rowDrag.slot;
    if (tab) to = Math.max(0, Math.min(tab.rows - 1, to));
    return to;
  }

  function rowDragAutoScrollTick() {
    if (!rowDrag || !rowDrag.active) return;
    const scroller = rowDragScroller();
    if (scroller) {
      const sRect = scroller.getBoundingClientRect();
      const header = el.cellGrid.querySelector('.column-header');
      const topEdge = sRect.top + (header ? header.getBoundingClientRect().height : 0);
      const y = rowDrag.lastY;
      let dy = 0;
      if (y < topEdge + ROW_DRAG_EDGE_PX) {
        dy = -Math.ceil(((topEdge + ROW_DRAG_EDGE_PX) - y) / 3);
      } else if (y > sRect.bottom - ROW_DRAG_EDGE_PX) {
        dy = Math.ceil((y - (sRect.bottom - ROW_DRAG_EDGE_PX)) / 3);
      }
      if (dy) {
        dy = Math.max(-40, Math.min(40, dy));
        const before = scroller.scrollTop;
        scroller.scrollTop = before + dy;
        if (scroller.scrollTop !== before) rowDragUpdateIndicator();
      }
    }
    rowDrag.raf = requestAnimationFrame(rowDragAutoScrollTick);
  }

  function rowDragActivate() {
    const d = rowDrag;
    d.active = true;
    document.body.classList.add('row-dragging');
    const ind = document.createElement('div');
    ind.className = 'row-drag-indicator';
    ind.hidden = true;
    document.body.appendChild(ind);
    d.indicator = ind;
    // Faint ghost of the source row (row-number area + cells, as one strip).
    const rc = el.cellGrid.querySelector('.row-controls[data-row="' + d.from + '"]');
    if (rc) {
      const rRect = rc.getBoundingClientRect();
      const gridRect = el.cellGrid.getBoundingClientRect();
      const ghost = document.createElement('div');
      ghost.className = 'row-drag-ghost';
      ghost.style.left = gridRect.left + 'px';
      ghost.style.width = gridRect.width + 'px';
      ghost.style.height = rRect.height + 'px';
      d.ghostOffset = d.startY - rRect.top;
      ghost.style.top = rRect.top + 'px';
      const label = document.createElement('span');
      label.className = 'row-drag-ghost-label';
      label.textContent = 'Row ' + (d.from + 1);
      ghost.appendChild(label);
      document.body.appendChild(ghost);
      d.ghost = ghost;
    }
    el.cellGrid.querySelectorAll('.row-controls[data-row="' + d.from + '"], .cell-wrap[data-row="' + d.from + '"]')
      .forEach(function (n) { n.classList.add('row-drag-source'); });
    try { window.getSelection().removeAllRanges(); } catch (err) { /* no-op */ }
    rowDragUpdateIndicator();
    d.raf = requestAnimationFrame(rowDragAutoScrollTick);
  }

  function endRowDrag(commit) {
    const d = rowDrag;
    if (!d) return;
    rowDrag = null;
    document.removeEventListener('pointermove', onRowDragMove, true);
    document.removeEventListener('pointerup', onRowDragUp, true);
    document.removeEventListener('pointercancel', onRowDragCancel, true);
    document.removeEventListener('keydown', onRowDragKey, true);
    window.removeEventListener('blur', onRowDragCancel);
    if (d.btn && d.btn.releasePointerCapture && d.pointerId != null) {
      try { d.btn.releasePointerCapture(d.pointerId); } catch (err) { /* no-op */ }
    }
    if (d.raf) cancelAnimationFrame(d.raf);
    if (d.indicator) d.indicator.remove();
    if (d.ghost) d.ghost.remove();
    document.body.classList.remove('row-dragging');
    if (el.cellGrid) {
      el.cellGrid.querySelectorAll('.row-drag-source')
        .forEach(function (n) { n.classList.remove('row-drag-source'); });
    }
    if (!d.active) return; // plain click: let the button's click handler move 1 row
    // Swallow the synthetic click that follows pointerup on the ↑/↓ button.
    const swallow = function (ev) {
      document.removeEventListener('click', swallow, true);
      if (ev.target && ev.target.closest && ev.target.closest('.row-move')) {
        ev.preventDefault();
        ev.stopPropagation();
      }
    };
    document.addEventListener('click', swallow, true);
    window.setTimeout(function () { document.removeEventListener('click', swallow, true); }, 0);
    if (!commit) {
      setStatus('Row move cancelled');
      return;
    }
    const to = (function () {
      rowDrag = d; // temporarily for helper
      const v = rowDragTargetIndex();
      rowDrag = null;
      return v;
    })();
    if (to === d.from || to < 0) {
      setStatus('Row ' + (d.from + 1) + ' not moved');
      return;
    }
    moveRowTo(d.from, to);
  }

  function onRowDragMove(e) {
    const d = rowDrag;
    if (!d || e.pointerId !== d.pointerId) return;
    d.lastY = e.clientY;
    if (!d.active) {
      if (Math.abs(e.clientY - d.startY) < ROW_DRAG_THRESHOLD_PX) return;
      rowDragActivate();
    }
    e.preventDefault();
    rowDragUpdateIndicator();
  }

  function onRowDragUp(e) {
    if (!rowDrag || e.pointerId !== rowDrag.pointerId) return;
    if (rowDrag.active) e.preventDefault();
    endRowDrag(true);
  }

  function onRowDragCancel() {
    endRowDrag(false);
  }

  function onRowDragKey(e) {
    if (e.key !== 'Escape' || !rowDrag) return;
    e.preventDefault();
    e.stopPropagation();
    endRowDrag(false);
  }

  function beginRowDragFromArrow(e, row, btn) {
    if (e.button !== 0 || rowDrag) return;
    const tab = activeTab();
    if (!tab || row < 0 || row >= tab.rows) return;
    // No text selection / cell block / focus jump from this press.
    e.preventDefault();
    e.stopPropagation();
    rowDrag = {
      from: row,
      pointerId: e.pointerId,
      startY: e.clientY,
      lastY: e.clientY,
      active: false,
      btn: btn,
      slot: null,
      indicator: null,
      ghost: null,
      ghostOffset: 0,
      raf: 0
    };
    if (btn.setPointerCapture) {
      try { btn.setPointerCapture(e.pointerId); } catch (err) { /* no-op */ }
    }
    document.addEventListener('pointermove', onRowDragMove, true);
    document.addEventListener('pointerup', onRowDragUp, true);
    document.addEventListener('pointercancel', onRowDragCancel, true);
    document.addEventListener('keydown', onRowDragKey, true);
    window.addEventListener('blur', onRowDragCancel);
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
      reorderRowInTab(tab, rowIndex, destination);
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
      if (isMasterTab(tab)) remapMasterIndicesAfterPushDown(cols, rowIndex);
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
    if (isMasterTab(tab)) remapMasterIndicesAfterColumnAdd(oldCols, tab.cols);
    renderTabs();
    renderGrid();
    renderMasterLibrary();
    scheduleSave();
    setStatus('Column added to ' + tab.title + ' (' + tab.cols + '×' + tab.rows + ')');
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
    // v0.140: Clear only empties the Notes — checks (and their text) stay.
    const scope = currentPromptScope();
    if (!getNotes(scope)) {
      setStatus('Notes already empty — use Uncheck all to remove checked cells');
      return;
    }
    // Snapshot before clear for the dedicated "Undo last clear" button (and undo stack).
    lastClearSnapshot = cloneCurrentDocument();
    pushHistory();
    setNotes(scope, '');
    if (el.combinedNotes) el.combinedNotes.value = '';
    renderCombinedPrompt();
    scheduleSave();
    showUndoClearButton(true);
    setStatus('Cleared notes (checks kept)');
  }

  /** Uncheck every cell / nest in the current Combined scope (one undo step). */
  function uncheckAllCombined() {
    const scope = currentPromptScope();
    const count = state.confirmedLinks.filter(function (link) { return link.scope === scope; }).length;
    if (!count) {
      setStatus('Nothing is checked');
      return;
    }
    dismissUndoClear();
    pushHistory();
    dropLinksForScope(scope);
    refreshAfterConfirmedChange();
    setStatus('Unchecked all (' + count + ') — Ctrl+Z to undo', 'ok');
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
    if (!el.partSearchStatus) return;
    el.partSearchStatus.textContent = text || '';
    el.partSearchStatus.title = text || '';
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
          let from = 0;
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
    if (!isMasterTab(tab) && isCellMasterLocked(tab, hit.cellIndex)) return;
    nest.pages[hit.pageIndex] = text;
    if (isMasterTab(tab)) syncLockedPartNestsFromMaster(hit.cellIndex);
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
      // Plan first: unique fields (cell or nest page). Locked Master cells are
      // skipped (parent + nests), same as typing / single Replace.
      const keysDone = Object.create(null);
      const plan = [];
      const changeCells = Object.create(null);
      const lockedCells = Object.create(null);
      for (let h = 0; h < partSearchHits.length; h++) {
        const hit = partSearchHits[h];
        if (isCellMasterEditBlocked(tab, hit.cellIndex)) {
          lockedCells[hit.cellIndex] = true;
          continue;
        }
        const key = hit.cellIndex + ':' + (hit.nestIndex == null
          ? ('c:' + (hit.pageIndex == null ? 'x' : hit.pageIndex))
          : (hit.nestIndex + ':' + hit.pageIndex));
        if (keysDone[key]) continue;
        keysDone[key] = true;
        const before = getHitText(tab, hit);
        const result = replaceAllInText(before, query, replacement);
        if (!result.count || result.text === before) continue;
        plan.push({ hit: hit, text: result.text, count: result.count });
        changeCells[hit.cellIndex] = true;
      }
      const nLocked = Object.keys(lockedCells).length;
      const lockedNote = nLocked ? (' — ' + pluralCells(nLocked, 'locked cell') + ' skipped') : '';
      const changeIdx = Object.keys(changeCells).map(function (s) { return parseInt(s, 10); });
      if (!plan.length) {
        setPartSearchStatus('0');
        setStatus('Nothing replaced in ' + tab.title + lockedNote, 'err');
        return;
      }
      openActionConfirm({
        anchor: el.partReplaceInput,
        focusBack: el.partReplaceInput,
        message: 'Replace in ' + pluralCells(changeIdx.length, 'cell') + '?' +
          linkedPartSuffix(tab, changeIdx) + (nLocked ? (' ' + pluralCells(nLocked, 'locked cell') + ' skipped.') : ''),
        actionLabel: 'Replace all',
        cancelStatus: 'Replace all cancelled',
        ariaLabel: 'Confirm replace all',
        onConfirm: function () {
          if (currentPartSearchTab() !== tab) {
            setStatus('Replace all cancelled — tab changed', 'err');
            return;
          }
          applyReplaceAllPlan(tab, plan, lockedNote);
        }
      });
      return;
    }
    partReplaceSingle(tab, replacement);
  }

  function applyReplaceAllPlan(tab, plan, lockedNote) {
    pushHistory();
    let total = 0;
    const touchedCells = Object.create(null);
    for (let k = 0; k < plan.length; k++) {
      const hit = plan[k].hit;
      setHitText(tab, hit, plan[k].text);
      total += plan[k].count;
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
    setStatus((total ? ('Replaced ' + total + ' in ' + tab.title) : ('No matches in ' + tab.title)) + lockedNote,
      total ? 'ok' : 'err');
    if (partSearchHitIndex >= 0) revealPartSearchHit(partSearchHits[partSearchHitIndex], { silent: true });
  }

  function partReplaceSingle(tab, replacement) {
    const query = el.partFindInput ? el.partFindInput.value : '';
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
    if (isCellMasterEditBlocked(tab, hit.cellIndex)) {
      // Locked Master cell: never replace (it used to change the textarea but not
      // the saved cell). Move on to the next match instead.
      partFindNext({ silent: true });
      setStatus('Locked Master cell skipped — double-click to unlock before replacing', 'err');
      return;
    }
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
  el.btnAddRow.addEventListener('click', addRow);
  el.btnAddCol.addEventListener('click', addColumn);
  if (el.btnFilterAll) el.btnFilterAll.addEventListener('click', function () { setGridRowFilter('all'); });
  if (el.btnFilterNonempty) el.btnFilterNonempty.addEventListener('click', function () { setGridRowFilter('nonempty'); });
  if (el.btnFilterIncluded) el.btnFilterIncluded.addEventListener('click', function () { setGridRowFilter('included'); });
  if (el.btnFitRows) el.btnFitRows.addEventListener('click', autoFitAllRowHeights);
  document.addEventListener('pointerdown', function (e) {
    if (col1FilterMenuOpen) {
      const wrap = el.cellGrid && el.cellGrid.querySelector('.column-col1-filter');
      const menu = popoverMenus.grid;
      if (!(wrap && wrap.contains(e.target)) && !(menu && menu.contains(e.target))) closeCol1FilterMenu();
    }
    if (masterLibFilterMenuOpen) {
      const wrap = el.masterLibraryItems && el.masterLibraryItems.querySelector('.master-library-filter');
      const menu = popoverMenus.masterLib;
      if (!(wrap && wrap.contains(e.target)) && !(menu && menu.contains(e.target))) closeMasterLibFilterMenu();
    }
    if (tabIconPickerEl && !tabIconPickerEl.contains(e.target)) {
      closeTabIconPicker();
    }
  });
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape' && !e.defaultPrevented) {
      // v0.143: Esc clears the Col A filter when no editor/menu/selection is open.
      const t = e.target;
      const tag = t && t.tagName;
      const inCell = tag === 'TEXTAREA' && t.classList && t.classList.contains('cell');
      const inEditor = (tag === 'INPUT' || tag === 'SELECT' || (tag === 'TEXTAREA' && !inCell) ||
        (t && t.isContentEditable) || (inCell && t.classList.contains('col1-editing')));
      const busy = col1FilterMenuOpen || masterLibFilterMenuOpen || !!col1MenuEl ||
        !!stickyCellRange || isMasterCellEditSession() || masterSegmentDialogOpen ||
        partSearchHits.length || partSearchFindAll;
      if (!inEditor && !busy) clearCol1ValueFilter(inCell ? t : null);
    }
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
    el.masterLibrary.addEventListener('toggle', sizeMasterLibraryPane);
    if (typeof ResizeObserver === 'function' && el.partSection) {
      new ResizeObserver(scheduleMasterLibrarySize).observe(el.partSection);
    }
    window.addEventListener('resize', scheduleMasterLibrarySize);
  }
  if (el.masterLibrary) {
    // Capture scroll from the results pane (or legacy items scroll). Do not close
    // when the user scrolls inside the fixed Values dropdown itself — that was the
    // "scroll down Values glitch" (menu closed / jumped mid-scroll).
    el.masterLibrary.addEventListener('scroll', function (e) {
      if (suppressMasterLibMenuScrollClose) return;
      const t = e.target;
      if (!t || !t.classList) return;
      const menu = popoverMenus.masterLib;
      if (menu && (t === menu || menu.contains(t))) return;
      if (t.classList.contains('master-library-results') ||
          t.classList.contains('master-library-items') ||
          t === el.masterLibrary) {
        closeMasterLibFilterMenu();
      }
    }, { passive: true, capture: true });
  }
  el.btnCopy.addEventListener('click', copyCombined);
  el.btnClear.addEventListener('click', clearCombined);
  if (el.btnUncheckAll) el.btnUncheckAll.addEventListener('click', uncheckAllCombined);
  if (el.combinedNotes) {
    el.combinedNotes.addEventListener('input', function () {
      dismissUndoClear();
      pushHistory({ coalesce: true });
      setNotes(currentPromptScope(), el.combinedNotes.value);
      renderCombinedPrompt();
      scheduleSave();
    });
    el.combinedNotes.addEventListener('blur', function () {
      if (el.combinedNotes.value !== getNotes(currentPromptScope())) {
        el.combinedNotes.value = getNotes(currentPromptScope());
      }
    });
  }
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

  if (el.partFindInput) {
    el.partFindInput.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') {
        e.preventDefault();
        if (e.shiftKey) partFindAll();
        else partFindNext();
        // No Find button any more: keep typing focus here so Enter steps to the next match
        // (the current hit stays marked + scrolled into view in the grid).
        if (document.activeElement !== el.partFindInput) el.partFindInput.focus({ preventScroll: true });
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
  // v0.140: the generated Combined view is read-only — no Master segment
  // locking / unlocking here (Master locks live on the cells only).
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
  document.addEventListener('keydown', function () { blockDragSelGuard = null; }, true);
  document.addEventListener('pointerdown', function () { blockDragSelGuard = null; }, true);
  document.addEventListener('selectionchange', function () {
    // Invariant: a lit cell block and a text selection never show together.
    // A new user text selection (Ctrl+A, nest drag…) drops the block. But Chromium
    // also re-extends the drag-start textarea's selection after a block drag's
    // mouseup; dropping the block for THAT made Ctrl+C copy one cell and let the
    // system blue come back. That leftover is collapsed instead (block kept).
    if (stickyCellRange && el.cellGrid) {
      const a = document.activeElement;
      if (a && (a.tagName === 'TEXTAREA' || a.tagName === 'INPUT') && el.cellGrid.contains(a) &&
          typeof a.selectionStart === 'number' && a.selectionStart !== a.selectionEnd &&
          !(cellRangeDrag && cellRangeDrag.mode === 'multi')) {
        if (a === blockDragSelGuard) {
          collapseCellTextSelection(a);
        } else {
          clearStickyCellRange();
          clearKeyboardCellRange();
        }
      }
    }
    if (document.activeElement !== el.combined) return;
    rememberCombinedCaretFromDom();
  });

  // Press outside the grid clears the lit block (menus that act on it are exempt).
  document.addEventListener('pointerdown', function (e) {
    if (!stickyCellRange) return;
    const t = e.target;
    if (el.cellGrid && t && el.cellGrid.contains(t)) return;
    if (t && typeof t.closest === 'function' &&
        t.closest('.part-page-menu, .tab-icon-picker, .add-page-choice-menu, .paste-confirm, .modal-overlay, .col1-value-menu')) {
      return;
    }
    clearStickyCellRange();
    clearKeyboardCellRange();
  }, true);

  // Text never gets dragged/dropped between or into cells (no cut/move by drag).
  // In-cell drag-select still works (that is selection, not drag-and-drop).
  if (el.cellGrid) {
    el.cellGrid.addEventListener('dragstart', function (e) {
      const t = e.target;
      if (t && (t.tagName === 'TEXTAREA' || t.tagName === 'INPUT' || t.nodeType === 3)) e.preventDefault();
    }, true);
    const blockTextDrop = function (e) {
      const t = e.target;
      if (!t || (t.tagName !== 'TEXTAREA' && t.tagName !== 'INPUT')) return;
      e.preventDefault();
      if (e.dataTransfer) e.dataTransfer.dropEffect = 'none';
    };
    el.cellGrid.addEventListener('dragover', blockTextDrop, true);
    el.cellGrid.addEventListener('drop', blockTextDrop, true);
  }

  el.combined.addEventListener('keydown', function (event) {
    // Read-only view: Ctrl/Cmd+C copies Notes + generated text (what you see).
    const mod = event.ctrlKey || event.metaKey;
    if (mod && !event.altKey && (event.key === 'c' || event.key === 'C')) {
      const sel = window.getSelection();
      if (sel && !sel.isCollapsed && el.combined.contains(sel.anchorNode)) return;
      event.preventDefault();
      copyCombined();
    }
  });

  if (el.btnMasterOverwrite) {
    el.btnMasterOverwrite.addEventListener('click', function () {
      overwriteMasterFromSegmentEdit(false);
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
          ? 'Match source order on — Combined follows grid order'
          : 'Match source order off — Combined follows the order you checked',
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

    // Restore remembered grid scroll positions before the first render (no jump).
    if (window.click2copy && typeof window.click2copy.loadGridScroll === 'function') {
      try {
        seedGridScroll(await withTimeout(window.click2copy.loadGridScroll(), 2000, {}));
      } catch (err) {
        console.error(err);
      }
    }
    const gridWrapEl = el.cellGrid && el.cellGrid.parentElement;
    if (gridWrapEl) gridWrapEl.addEventListener('scroll', onGridScrolled, { passive: true });
    window.addEventListener('beforeunload', function () {
      captureLiveGridScroll();
      if (window.click2copy && typeof window.click2copy.saveGridScrollSync === 'function') {
        try { window.click2copy.saveGridScrollSync(gridScrollPayload()); } catch (err) { /* closing */ }
      }
    });

    const current = activeDocument();
    clearHistory();
    applyData(current.data);
    if (lastLockRepair && (lastLockRepair.resynced || lastLockRepair.repointed ||
        lastLockRepair.unlocked || lastLockRepair.relocked)) {
      const repairMsg = 'Repaired stale Master links on load: ' + reconcileSummary(lastLockRepair);
      pushEvent(repairMsg, 'ok');
      setStatus(repairMsg, 'ok');
    }
    if (lastCombinedMigration) {
      const m = lastCombinedMigration;
      pushEvent('Combined v0.140 migration: kept ' + m.keptChecks + ' checks, moved ' +
        m.movedChars + ' chars of typed text to Notes, dropped ' + m.droppedChecks + ' stale links', 'ok');
    }
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
    if (target && target.classList && (target.classList.contains('tab-rename-input') ||
      target.classList.contains('part-page-rename-input'))) return;

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
  /* ── Restart backup: Tools setting + Update-ready Restart flow ───────── */

  function renderBackupSettings(settings) {
    if (!settings) return;
    if (el.backupRemember) el.backupRemember.checked = settings.mode === 'remember';
    if (el.backupFolderPath) {
      el.backupFolderPath.textContent = settings.lastDir || 'No folder yet — Restart will ask';
      el.backupFolderPath.title = settings.lastDir || '';
    }
  }

  async function refreshBackupSettings() {
    if (!window.click2copy || typeof window.click2copy.getBackupSettings !== 'function') return null;
    try {
      const settings = await window.click2copy.getBackupSettings();
      renderBackupSettings(settings);
      return settings;
    } catch (err) {
      console.error(err);
      return null;
    }
  }

  async function chooseBackupFolder() {
    if (!window.click2copy || typeof window.click2copy.chooseBackupFolder !== 'function') return null;
    const res = await window.click2copy.chooseBackupFolder();
    if (res && res.settings) renderBackupSettings(res.settings);
    if (res && res.ok) setStatus('Restart backups → ' + res.settings.lastDir, 'ok');
    else if (res && res.error) setStatus(res.error, 'err');
    return res;
  }

  if (el.btnBackupFolder) {
    el.btnBackupFolder.addEventListener('click', function () {
      chooseBackupFolder().catch(function (err) {
        console.error(err);
        setStatus('Could not choose backup folder', 'err');
      });
    });
  }
  if (el.backupRemember) {
    el.backupRemember.addEventListener('change', async function () {
      if (!window.click2copy || typeof window.click2copy.setBackupSettings !== 'function') return;
      const wantRemember = el.backupRemember.checked;
      try {
        let settings = await window.click2copy.getBackupSettings();
        if (wantRemember && !settings.lastDir) {
          // Remember needs a folder — pick one now (cancel keeps Ask mode).
          const picked = await chooseBackupFolder();
          if (!picked || !picked.ok) {
            el.backupRemember.checked = false;
            return;
          }
        }
        const res = await window.click2copy.setBackupSettings({ mode: wantRemember ? 'remember' : 'ask' });
        if (res && res.settings) renderBackupSettings(res.settings);
        setStatus(wantRemember
          ? 'Restart backups save silently to the remembered folder'
          : 'Restart will ask where to save each backup', res && res.ok ? 'ok' : 'err');
      } catch (err) {
        console.error(err);
        setStatus('Could not save backup setting', 'err');
        refreshBackupSettings();
      }
    });
  }
  refreshBackupSettings();

  function withTimeout(promise, ms, fallback) {
    return Promise.race([
      promise,
      new Promise(function (resolve) { setTimeout(function () { resolve(fallback); }, ms); })
    ]);
  }

  function collectRendererPrefs() {
    const out = {};
    try {
      for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (key && key.indexOf('click2copy') === 0) out[key] = localStorage.getItem(key);
      }
    } catch (err) { /* ignore */ }
    return out;
  }

  let restartInFlight = false;

  async function restartWithBackup() {
    if (restartInFlight) return;
    if (!window.click2copy || typeof window.click2copy.relaunchApp !== 'function') {
      setStatus('Restart unavailable', 'err');
      return;
    }
    restartInFlight = true;
    if (el.btnUpdateRestart) el.btnUpdateRestart.disabled = true;
    // Remember grid scroll positions across the relaunch (bounded wait).
    await withTimeout(flushGridScroll(), 2000, false);
    try {
      if (typeof window.click2copy.backupAndRelaunch !== 'function') {
        setStatus('Restarting…', 'ok');
        await window.click2copy.relaunchApp();
        return;
      }
      // Flush the pending debounced save so the snapshot is current (bounded wait).
      setStatus('Saving session before backup…');
      clearTimeout(state.saveTimer);
      const persisted = await withTimeout(persist(), 8000, false);
      const extra = { rendererPrefs: collectRendererPrefs() };
      // Save failed / timed out → carry the live session in the backup itself.
      if (!persisted) extra.liveSession = JSON.parse(JSON.stringify(sessionSnapshot()));
      setStatus('Saving backup…');
      const res = await window.click2copy.backupAndRelaunch(extra);
      if (res && res.ok) {
        setStatus('Backup saved: ' + res.path + ' — restarting…', 'ok');
        return;
      }
      const why = res && res.canceled
        ? 'Backup was cancelled — no backup file was written.'
        : 'Backup failed: ' + ((res && res.error) || 'unknown error');
      const go = window.confirm(why + '\n\nRestart WITHOUT a backup?\n' +
        '(Your prompts are still saved in the app session.)\n\n' +
        'OK = restart without backup    Cancel = stay open');
      if (go) {
        setStatus('Restarting without backup…', 'ok');
        await window.click2copy.relaunchApp();
      } else {
        setStatus('Restart cancelled — nothing changed', 'err');
      }
    } catch (err) {
      console.error(err);
      setStatus('Restart failed', 'err');
    } finally {
      restartInFlight = false;
      if (el.btnUpdateRestart) el.btnUpdateRestart.disabled = false;
    }
  }

  if (el.btnUpdateRestart) {
    el.btnUpdateRestart.addEventListener('click', function () {
      restartWithBackup();
    });
  }
  if (el.btnUpdateDismiss) {
    el.btnUpdateDismiss.addEventListener('click', function () {
      if (el.updateBanner) el.updateBanner.hidden = true;
    });
  }

  init();

})();
