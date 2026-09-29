(function () {
  'use strict';

  const state = {
    tabs: [],
    activeTabId: null,
    combinedPrompt: '',
    saveTimer: null
  };

  const el = {
    tabBar: document.getElementById('tab-bar'),
    partContent: document.getElementById('part-content'),
    partLabel: document.getElementById('part-label'),
    combined: document.getElementById('combined-prompt'),
    status: document.getElementById('status'),
    btnAdd: document.getElementById('btn-add-tab'),
    btnRename: document.getElementById('btn-rename-tab'),
    btnDelete: document.getElementById('btn-delete-tab'),
    btnAppend: document.getElementById('btn-append'),
    btnCopy: document.getElementById('btn-copy'),
    btnClear: document.getElementById('btn-clear'),
    btnExport: document.getElementById('btn-export'),
    btnImport: document.getElementById('btn-import')
  };

  function uid() {
    return 'tab-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 7);
  }

  function activeTab() {
    return state.tabs.find((t) => t.id === state.activeTabId) || null;
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
      setStatus._t = setTimeout(() => {
        el.status.textContent = '';
        el.status.className = 'status';
      }, 2200);
    }
  }

  function applyData(data) {
    state.tabs = data.tabs;
    state.activeTabId = data.activeTabId || data.tabs[0].id;
    if (!state.tabs.some((t) => t.id === state.activeTabId)) {
      state.activeTabId = state.tabs[0].id;
    }
    state.combinedPrompt = typeof data.combinedPrompt === 'string' ? data.combinedPrompt : '';

    const tab = activeTab();
    el.partContent.value = tab ? tab.content : '';
    el.combined.value = state.combinedPrompt;
    renderTabs();
  }

  function renderTabs() {
    el.tabBar.innerHTML = '';
    state.tabs.forEach((tab) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'tab' + (tab.id === state.activeTabId ? ' active' : '');
      btn.textContent = tab.title;
      btn.title = tab.title + ' (double-click to rename)';
      btn.setAttribute('role', 'tab');
      btn.setAttribute('aria-selected', tab.id === state.activeTabId ? 'true' : 'false');
      btn.dataset.id = tab.id;

      btn.addEventListener('click', () => selectTab(tab.id));
      btn.addEventListener('dblclick', (e) => {
        e.preventDefault();
        e.stopPropagation();
        startInlineRename(btn, tab);
      });
      btn.addEventListener('contextmenu', (e) => {
        e.preventDefault();
        const action = window.prompt(
          'Tab actions — type: rename | delete',
          'rename'
        );
        if (!action) return;
        const a = action.trim().toLowerCase();
        if (a === 'rename') startInlineRename(btn, tab);
        else if (a === 'delete') deleteTab(tab.id);
      });

      el.tabBar.appendChild(btn);
    });

    const tab = activeTab();
    el.partLabel.textContent = tab ? tab.title : 'Part content';
    el.btnDelete.disabled = state.tabs.length <= 1;
  }

  function selectTab(id) {
    if (id === state.activeTabId) return;
    const tab = state.tabs.find((t) => t.id === id);
    if (!tab) return;
    state.activeTabId = id;
    el.partContent.value = tab.content;
    renderTabs();
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

    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        finish(true);
      } else if (e.key === 'Escape') {
        e.preventDefault();
        finish(false);
      }
    });
    input.addEventListener('blur', () => finish(true));
  }

  function addTab() {
    const n = state.tabs.length + 1;
    const tab = { id: uid(), title: 'Part ' + n, content: '' };
    state.tabs.push(tab);
    state.activeTabId = tab.id;
    el.partContent.value = '';
    renderTabs();
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
    const tab = state.tabs.find((t) => t.id === id);
    if (!tab) return;

    if (tab.content.trim()) {
      const ok = window.confirm(
        'Delete "' + tab.title + '"? Its content is not empty and will be lost.'
      );
      if (!ok) return;
    }

    const idx = state.tabs.findIndex((t) => t.id === id);
    state.tabs.splice(idx, 1);
    if (state.activeTabId === id) {
      const next = state.tabs[Math.min(idx, state.tabs.length - 1)];
      state.activeTabId = next.id;
      el.partContent.value = next.content;
    }
    renderTabs();
    scheduleSave();
    setStatus('Tab deleted');
  }

  function appendToPrompt() {
    const tab = activeTab();
    if (!tab) return;
    const text = el.partContent.value;
    if (!text) {
      setStatus('Nothing to append', 'err');
      return;
    }
    if (state.combinedPrompt && !state.combinedPrompt.endsWith('\n') && !text.startsWith('\n')) {
      state.combinedPrompt += '\n';
    }
    state.combinedPrompt += text;
    el.combined.value = state.combinedPrompt;
    scheduleSave();
    setStatus('Appended "' + tab.title + '"', 'ok');
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
      // Fallback for older/locked clipboard
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
  el.btnDelete.addEventListener('click', () => {
    if (state.activeTabId) deleteTab(state.activeTabId);
  });
  el.btnAppend.addEventListener('click', appendToPrompt);
  el.btnCopy.addEventListener('click', copyCombined);
  el.btnClear.addEventListener('click', clearCombined);
  el.btnExport.addEventListener('click', exportBackup);
  el.btnImport.addEventListener('click', importBackup);

  el.partContent.addEventListener('input', () => {
    const tab = activeTab();
    if (!tab) return;
    tab.content = el.partContent.value;
    scheduleSave();
  });

  el.combined.addEventListener('input', () => {
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
          { id: 'tab-1', title: 'Part 1', content: '' },
          { id: 'tab-2', title: 'Part 2', content: '' },
          { id: 'tab-3', title: 'Part 3', content: '' }
        ],
        activeTabId: 'tab-1',
        combinedPrompt: ''
      };
    }

    applyData(data);
  }

  init();
})();
