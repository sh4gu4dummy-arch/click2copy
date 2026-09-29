(function () {
  'use strict';

  const DEFAULT_COLS = 3;
  const DEFAULT_ROWS = 8;

  const state = {
    tabs: [],
    activeTabId: null,
    combinedPrompt: '',
    saveTimer: null
  };

  const el = {
    tabBar: document.getElementById('tab-bar'),
    cellGrid: document.getElementById('cell-grid'),
    partLabel: document.getElementById('part-label'),
    combined: document.getElementById('combined-prompt'),
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
    btnImport: document.getElementById('btn-import')
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

    return { id: t.id, title: t.title, cols: cols, rows: rows, cells: cells };
  }

  function activeTab() {
    return state.tabs.find(function (t) {
      return t.id === state.activeTabId;
    }) || null;
  }

  function snapshot() {
    return {
      tabs: state.tabs,
      activeTabId: state.activeTabId,
      combinedPrompt: state.combinedPrompt
    };
  }

  function scheduleSave() {
    clearTimeout(state.saveTimer);
    state.saveTimer = setTimeout(persist, 250);
  }

  async function persist() {
    if (!window.click2copy) return;
    try {
      await window.click2copy.saveData(snapshot());
    } catch (err) {
      console.error(err);
      setStatus('Failed to save', 'err');
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

  /** Join non-empty cells of one row with " | ". */
  function rowText(tab, rowIndex) {
    const parts = [];
    for (let c = 0; c < tab.cols; c++) {
      const v = (tab.cells[rowIndex * tab.cols + c] || '').trim();
      if (v) parts.push(v);
    }
    return parts.join(' | ');
  }

  /** All non-empty cells, row-major, joined by newlines. */
  function partText(tab) {
    const lines = [];
    for (let r = 0; r < tab.rows; r++) {
      const line = rowText(tab, r);
      if (line) lines.push(line);
    }
    return lines.join('\n');
  }

  function tabHasContent(tab) {
    return tab.cells.some(function (c) {
      return (c || '').trim().length > 0;
    });
  }

  function appendText(text, label) {
    if (!text) {
      setStatus('Nothing to append', 'err');
      return;
    }
    if (state.combinedPrompt && !state.combinedPrompt.endsWith('\n')) {
      state.combinedPrompt += '\n';
    }
    state.combinedPrompt += text;
    el.combined.value = state.combinedPrompt;
    scheduleSave();
    setStatus('Appended' + (label ? ' "' + label + '"' : ''), 'ok');
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
    state.tabs = tabs;
    state.activeTabId = data.activeTabId || tabs[0].id;
    if (!state.tabs.some(function (t) {
      return t.id === state.activeTabId;
    })) {
      state.activeTabId = state.tabs[0].id;
    }
    state.combinedPrompt = typeof data.combinedPrompt === 'string' ? data.combinedPrompt : '';

    el.combined.value = state.combinedPrompt;
    renderTabs();
    renderGrid();
  }

  function renderTabs() {
    el.tabBar.innerHTML = '';
    state.tabs.forEach(function (tab) {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'tab' + (tab.id === state.activeTabId ? ' active' : '');
      btn.textContent = tab.title;
      btn.title = tab.title + ' (double-click to rename)';
      btn.setAttribute('role', 'tab');
      btn.setAttribute('aria-selected', tab.id === state.activeTabId ? 'true' : 'false');
      btn.dataset.id = tab.id;

      btn.addEventListener('click', function () {
        selectTab(tab.id);
      });
      btn.addEventListener('dblclick', function (e) {
        e.preventDefault();
        e.stopPropagation();
        startInlineRename(btn, tab);
      });
      btn.addEventListener('contextmenu', function (e) {
        e.preventDefault();
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
    el.btnDelete.disabled = state.tabs.length <= 1;
  }

  function renderGrid() {
    const tab = activeTab();
    el.cellGrid.innerHTML = '';
    if (!tab) return;

    // CSS grid: first column = row append button, then N cell columns
    el.cellGrid.style.gridTemplateColumns =
      '40px repeat(' + tab.cols + ', minmax(120px, 1fr))';

    for (let r = 0; r < tab.rows; r++) {
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
      el.cellGrid.appendChild(rowBtn);

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

  function onCellInput(e) {
    const tab = activeTab();
    if (!tab) return;
    const idx = parseInt(e.target.dataset.idx, 10);
    if (Number.isNaN(idx) || idx < 0 || idx >= tab.cells.length) return;
    tab.cells[idx] = e.target.value;
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
    scheduleSave();
    setStatus('Tab added');
  }

  function renameActiveTab() {
    const tab = activeTab();
    if (!tab) return;
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
    if (state.tabs.length <= 1) {
      setStatus('Keep at least one tab', 'err');
      return;
    }
    const tab = state.tabs.find(function (t) {
      return t.id === id;
    });
    if (!tab) return;

    if (tabHasContent(tab)) {
      const ok = window.confirm(
        'Delete "' + tab.title + '"? Its content is not empty and will be lost.'
      );
      if (!ok) return;
    }

    const idx = state.tabs.findIndex(function (t) {
      return t.id === id;
    });
    state.tabs.splice(idx, 1);
    if (state.activeTabId === id) {
      const next = state.tabs[Math.min(idx, state.tabs.length - 1)];
      state.activeTabId = next.id;
    }
    renderTabs();
    renderGrid();
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
    scheduleSave();
    setStatus('Row added (' + tab.cols + '×' + tab.rows + ')');
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
    tab.cols += 1;
    tab.cells = newCells;
    renderTabs();
    renderGrid();
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
    if (state.combinedPrompt && !window.confirm('Clear the combined prompt?')) {
      return;
    }
    state.combinedPrompt = '';
    el.combined.value = '';
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

  el.combined.addEventListener('input', function () {
    state.combinedPrompt = el.combined.value;
    scheduleSave();
  });

  async function init() {
    let data;
    if (window.click2copy) {
      try {
        data = await window.click2copy.loadData();
      } catch (err) {
        console.error(err);
      }
    }
    if (!data || !Array.isArray(data.tabs) || data.tabs.length === 0) {
      data = {
        tabs: [
          makeTab('tab-1', 'Part 1'),
          makeTab('tab-2', 'Part 2'),
          makeTab('tab-3', 'Part 3')
        ],
        activeTabId: 'tab-1',
        combinedPrompt: ''
      };
    }

    applyData(data);
  }

  init();
})();
