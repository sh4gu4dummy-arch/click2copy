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
    saveTimer: null
  };
  let initialized = false;
  const pendingDocumentPaths = [];

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
    btnAdd: document.getElementById('btn-add-tab'),
    btnRename: document.getElementById('btn-rename-tab'),
    btnDelete: document.getElementById('btn-delete-tab'),
    btnAddRow: document.getElementById('btn-add-row'),
    btnAddCol: document.getElementById('btn-add-col'),
    btnAppend: document.getElementById('btn-append'),
    btnCopy: document.getElementById('btn-copy'),
    btnClear: document.getElementById('btn-clear'),
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

  function makeTab(id, title) {
    return {
      id: id,
      title: title,
      cols: DEFAULT_COLS,
      rows: DEFAULT_ROWS,
      cells: emptyCells(DEFAULT_COLS, DEFAULT_ROWS)
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
      separators: Object.assign({}, DEFAULT_SEPARATORS)
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
      return { id: t.id, title: t.title, cols: DEFAULT_COLS, rows: DEFAULT_ROWS, cells: cells };
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
    const normalized = { id: t.id, title: t.title, cols: cols, rows: rows, cells: cells };
    if (columnWidths) normalized.columnWidths = columnWidths;
    return normalized;
  }

  function activeTab() {
    return state.tabs.find(function (t) {
      return t.id === state.activeTabId;
    }) || null;
  }

  function isMasterTab(tab) {
    return !!tab && state.tabs[0] === tab;
  }

  function snapshot() {
    return {
      tabs: state.tabs,
      activeTabId: state.activeTabId,
      combinedPrompt: state.combinedPrompt,
      globalCombined: state.globalCombined,
      partPrompts: Object.assign({}, state.partPrompts),
      separators: state.separators
    };
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

  function setStatus(msg, kind) {
    el.status.textContent = msg || '';
    el.status.className = 'status' + (kind ? ' ' + kind : '');
    if (msg) {
      clearTimeout(setStatus._t);
      setStatus._t = setTimeout(function () {
        el.status.textContent = '';
        el.status.className = 'status';
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

  function appendText(text, label) {
    if (!text) {
      setStatus('Nothing to append', 'err');
      return;
    }
    if (state.globalCombined) {
      state.combinedPrompt = appendWithPartSeparator(state.combinedPrompt, text);
      el.combined.value = state.combinedPrompt;
    } else {
      const tab = activeTab();
      if (!tab) return;
      state.partPrompts[tab.id] = appendWithPartSeparator(state.partPrompts[tab.id] || '', text);
      el.combined.value = state.partPrompts[tab.id];
    }
    scheduleSave();
    setStatus('Appended' + (label ? ' "' + label + '"' : ''), 'ok');
  }

  function appendWithPartSeparator(current, text) {
    const partSeparator = separatorValue(state.separators.part);
    if (current && partSeparator && !current.endsWith(partSeparator)) {
      return current + partSeparator + text;
    }
    return current + text;
  }

  function renderCombinedPrompt() {
    const tab = activeTab();
    el.combined.value = state.globalCombined
      ? state.combinedPrompt
      : (tab ? state.partPrompts[tab.id] || '' : '');
    el.globalCombined.checked = state.globalCombined;
  }

  function mergePartPrompts() {
    state.tabs.forEach(function (tab) {
      const prompt = state.partPrompts[tab.id];
      if (prompt) state.combinedPrompt = appendWithPartSeparator(state.combinedPrompt, prompt);
    });
    state.partPrompts = {};
  }

  function applyData(data) {
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
    state.activeTabId = data.activeTabId || tabs[0].id;
    if (!state.tabs.some(function (t) {
      return t.id === state.activeTabId;
    })) {
      state.activeTabId = state.tabs[0].id;
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
    applyData(next.data);
    renderDocuments();
    scheduleSave();
  }

  async function createDocument() {
    await flushCurrentDocument();
    const data = makeDefaultData();
    data.activeTabId = data.tabs[0].id;
    const document = {
      id: uid(),
      title: 'Untitled',
      filePath: null,
      data: data
    };
    state.documents.push(document);
    state.activeDocumentId = document.id;
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
      data.activeTabId = data.tabs[0].id;
      const replacement = { id: uid(), title: 'Untitled', filePath: null, data: data };
      state.documents.push(replacement);
      state.activeDocumentId = replacement.id;
      applyData(replacement.data);
      renderDocuments();
      scheduleSave();
      return;
    }
    if (state.activeDocumentId === id) {
      const next = state.documents[Math.min(index, state.documents.length - 1)];
      state.activeDocumentId = next.id;
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

    const tab = activeTab();
    el.partLabel.textContent = tab
      ? tab.title + ' (' + tab.cols + '×' + tab.rows + ')'
      : 'Part content';
    el.btnDelete.disabled = state.tabs.length <= 1 || isMasterTab(tab);
    el.btnRename.disabled = isMasterTab(tab);
    el.masterLibrary.hidden = isMasterTab(tab);
  }

  function renderMasterLibrary() {
    el.masterLibraryItems.innerHTML = '';
    const master = state.tabs.find(isMasterTab);
    const current = activeTab();
    if (!master || !current || isMasterTab(current)) return;

    for (let row = 0; row < master.rows; row++) {
      for (let col = 0; col < master.cols; col++) {
        const text = master.cells[row * master.cols + col];
        if (!text || !text.trim()) continue;
        const button = document.createElement('button');
        button.type = 'button';
        button.className = 'master-library-item';
        button.textContent = text;
        button.title = 'Add to next empty cell in ' + current.title;
        button.addEventListener('click', function () {
          insertMasterText(text);
        });
        el.masterLibraryItems.appendChild(button);
      }
    }
    if (el.masterLibraryItems.childElementCount === 0) {
      const empty = document.createElement('span');
      empty.className = 'master-library-empty';
      empty.textContent = 'Add reusable text in the Master part first.';
      el.masterLibraryItems.appendChild(empty);
    }
  }

  function insertMasterText(text) {
    const tab = activeTab();
    if (!tab || isMasterTab(tab)) return;
    let index = tab.cells.findIndex(function (cell) {
      return !cell || !cell.trim();
    });
    if (index === -1) {
      tab.cells.push.apply(tab.cells, emptyCells(tab.cols, 1));
      tab.rows++;
      index = tab.cells.length - tab.cols;
    }
    tab.cells[index] = text;
    renderTabs();
    renderGrid();
    renderMasterLibrary();
    scheduleSave();
    const cell = el.cellGrid.querySelector('[data-idx="' + index + '"]');
    if (cell) cell.focus();
    setStatus('Added Master text to ' + tab.title, 'ok');
  }

  function reorderTab(draggedId, targetId, afterTarget) {
    if (!draggedId || draggedId === targetId) return;
    const fromIndex = state.tabs.findIndex(function (tab) { return tab.id === draggedId; });
    const targetIndex = state.tabs.findIndex(function (tab) { return tab.id === targetId; });
    if (fromIndex <= 0 || targetIndex <= 0) return;

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
    if (Array.isArray(tab.columnWidths) && tab.columnWidths.length === tab.cols) {
      el.cellGrid.style.gridTemplateColumns = '40px ' + tab.columnWidths.map(function (width) {
        return width + 'px';
      }).join(' ');
      return;
    }
    // Until the first resize, let the browser distribute columns equally.
    el.cellGrid.style.gridTemplateColumns =
      '40px repeat(' + tab.cols + ', minmax(120px, 1fr))';
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
      rowBtn.title = 'Append row ' + (r + 1) + ' to prompt';
      rowBtn.setAttribute('aria-label', 'Append row ' + (r + 1));
      rowBtn.dataset.row = String(r);
      rowBtn.innerHTML = '<span class="row-append-mark" aria-hidden="true"></span>';
      rowBtn.addEventListener('click', function () {
        appendRow(r);
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
        const ta = document.createElement('textarea');
        ta.className = 'cell';
        ta.rows = 2;
        ta.spellcheck = false;
        ta.placeholder = 'R' + (r + 1) + 'C' + (c + 1);
        ta.value = tab.cells[idx] || '';
        ta.dataset.row = String(r);
        ta.dataset.col = String(c);
        ta.dataset.idx = String(idx);
        ta.setAttribute('aria-label', 'Row ' + (r + 1) + ' column ' + (c + 1));
        ta.addEventListener('input', onCellInput);
        el.cellGrid.appendChild(ta);
      }
    }
  }

  function resizeColumnByKeyboard(tab, col, delta) {
    const widths = Array.isArray(tab.columnWidths) && tab.columnWidths.length === tab.cols
      ? tab.columnWidths.slice()
      : currentColumnWidths(tab);
    if (!widths) return;
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

  function onCellInput(e) {
    const tab = activeTab();
    if (!tab) return;
    const idx = parseInt(e.target.dataset.idx, 10);
    if (Number.isNaN(idx) || idx < 0 || idx >= tab.cells.length) return;
    tab.cells[idx] = e.target.value;
    if (isMasterTab(tab)) renderMasterLibrary();
    scheduleSave();
  }

  function selectTab(id) {
    if (id === state.activeTabId) return;
    const tab = state.tabs.find(function (t) {
      return t.id === id;
    });
    if (!tab) return;
    state.activeTabId = id;
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
        tab.title = next;
        scheduleSave();
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
    const n = state.tabs.length + 1;
    const tab = makeTab(uid(), 'Part ' + n);
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
    tab.title = trimmed;
    renderTabs();
    scheduleSave();
  }

  function deleteTab(id) {
    const tab = state.tabs.find(function (item) { return item.id === id; });
    if (state.tabs.length <= 1 || isMasterTab(tab)) {
      setStatus('The Master part cannot be deleted', 'err');
      return;
    }
    if (!tab) return;

    if (tabHasContent(tab)) {
      const ok = window.confirm(
        'Delete "' + tab.title + '"? Its cell content and part-specific appended prompt will be lost.'
      );
      if (!ok) return;
    }

    const idx = state.tabs.findIndex(function (t) {
      return t.id === id;
    });
    state.tabs.splice(idx, 1);
    delete state.partPrompts[id];
    if (state.activeTabId === id) {
      const next = state.tabs[Math.min(idx, state.tabs.length - 1)];
      state.activeTabId = next.id;
    }
    renderTabs();
    renderGrid();
    renderMasterLibrary();
    scheduleSave();
    setStatus('Tab deleted');
  }

  function addRow() {
    const tab = activeTab();
    if (!tab) return;
    for (let c = 0; c < tab.cols; c++) {
      tab.cells.push('');
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

    if (direction < 0) {
      for (let col = 0; col < tab.cols; col++) {
        const sourceIndex = rowIndex * tab.cols + col;
        const destinationIndex = destination * tab.cols + col;
        const cell = tab.cells[destinationIndex];
        tab.cells[destinationIndex] = tab.cells[sourceIndex];
        tab.cells[sourceIndex] = cell;
      }
    } else {
      let emptyRow = rowIndex + 1;
      while (emptyRow < tab.rows && !rowIsEmpty(tab, emptyRow)) emptyRow++;

      if (emptyRow === tab.rows) {
        tab.cells.push.apply(tab.cells, emptyCells(tab.cols, 1));
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
    }

    renderTabs();
    renderGrid();
    scheduleSave();
    setStatus(direction < 0 ? 'Row moved up' : 'Row moved down');
  }

  function addColumn() {
    const tab = activeTab();
    if (!tab) return;
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
    renderTabs();
    renderGrid();
    renderMasterLibrary();
    scheduleSave();
    setStatus('Column added (' + tab.cols + '×' + tab.rows + ')');
  }

  function appendRow(rowIndex) {
    const tab = activeTab();
    if (!tab) return;
    if (rowIndex < 0 || rowIndex >= tab.rows) return;
    const text = rowText(tab, rowIndex);
    appendText(text, tab.title + ' row ' + (rowIndex + 1));
  }

  function appendAll() {
    const tab = activeTab();
    if (!tab) return;
    const text = partText(tab);
    appendText(text, tab.title);
  }

  async function copyCombined() {
    const text = el.combined.value;
    if (!text) {
      setStatus('Combined prompt is empty', 'err');
      return;
    }
    try {
      await navigator.clipboard.writeText(text);
      setStatus('Copied to clipboard', 'ok');
    } catch (err) {
      el.combined.focus();
      el.combined.select();
      try {
        document.execCommand('copy');
        setStatus('Copied to clipboard', 'ok');
      } catch (e2) {
        setStatus('Copy failed', 'err');
      }
    }
  }

  function clearCombined() {
    const prompt = state.globalCombined
      ? state.combinedPrompt
      : (state.partPrompts[state.activeTabId] || '');
    if (prompt && !window.confirm('Clear the combined prompt?')) {
      return;
    }
    if (state.globalCombined) state.combinedPrompt = '';
    else state.partPrompts[state.activeTabId] = '';
    renderCombinedPrompt();
    scheduleSave();
    setStatus('Cleared');
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
  el.btnAppend.addEventListener('click', appendAll);
  el.btnCopy.addEventListener('click', copyCombined);
  el.btnClear.addEventListener('click', clearCombined);
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
    const tab = activeTab();
    if (state.globalCombined) state.combinedPrompt = el.combined.value;
    else if (tab) state.partPrompts[tab.id] = el.combined.value;
    scheduleSave();
  });

  el.globalCombined.addEventListener('change', function () {
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
      if (!fallback.activeTabId) fallback.activeTabId = fallback.tabs[0].id;
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
      data.activeTabId = data.tabs[0].id;
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

  init();
})();
