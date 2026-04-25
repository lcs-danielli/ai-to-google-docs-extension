document.addEventListener('DOMContentLoaded', async () => {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab) return;

  const headerSub  = document.getElementById('header-sub');
  const mainEl     = document.getElementById('main');
  const emptyEl    = document.getElementById('empty');
  const statusEl   = document.getElementById('status');
  const btnExport  = document.getElementById('btn-export');
  const btnLast    = document.getElementById('btn-last');
  const btnAll     = document.getElementById('btn-all');
  const lastRow    = document.getElementById('last-row');
  const lastNameEl = document.getElementById('last-name');
  const modeSelect = document.getElementById('mode-select');

  function setWorking() {
    statusEl.textContent = 'Working…';
    btnExport.disabled = btnLast.disabled = btnAll.disabled = true;
  }
  function resetButtons() {
    btnExport.disabled = btnLast.disabled = btnAll.disabled = false;
    statusEl.textContent = '';
  }

  function send(action) {
    setWorking();
    chrome.tabs.sendMessage(tab.id, { action }, () => {
      if (chrome.runtime.lastError) {
        statusEl.textContent = 'Error — try refreshing the page';
        resetButtons();
        return;
      }
      setTimeout(() => window.close(), 400);
    });
  }

  // Update primary button label to match the default mode
  function syncExportButton(mode) {
    if (mode === 'last')      btnExport.textContent = 'Export Last Response';
    else if (mode === 'full') btnExport.textContent = 'Export Full Conversation';
    else                      btnExport.textContent = 'Export to Docs ▾';
  }

  // Load stored default mode
  chrome.storage.local.get('defaultExportMode', (data) => {
    const mode = data.defaultExportMode || 'select';
    modeSelect.value = mode;
    syncExportButton(mode);
  });

  modeSelect.addEventListener('change', () => {
    const mode = modeSelect.value;
    chrome.storage.local.set({ defaultExportMode: mode });
    syncExportButton(mode);
  });

  // ── Folder picker ──
  const folderBtn      = document.getElementById('folder-btn');
  const folderDropdown = document.getElementById('folder-dropdown');

  // Show current folder name
  chrome.storage.local.get('customFolderName', (d) => {
    if (d.customFolderName) {
      const n = d.customFolderName;
      folderBtn.textContent = (n.length > 20 ? n.slice(0, 18) + '…' : n) + ' ▾';
    }
  });

  folderBtn.addEventListener('click', () => {
    if (folderDropdown.style.display !== 'none') {
      folderDropdown.style.display = 'none';
      return;
    }
    folderDropdown.innerHTML = '<div style="padding:7px 12px;font-size:11px;color:#999">Loading…</div>';
    folderDropdown.style.display = 'block';

    chrome.runtime.sendMessage({ action: 'listFolders' }, (resp) => {
      folderDropdown.innerHTML = '';
      chrome.storage.local.get('customFolderId', (d) => {
        const activeFolderId = d.customFolderId || null;

        // Default option
        const defItem = document.createElement('div');
        defItem.className = 'folder-item folder-item-default' + (!activeFolderId ? ' folder-active' : '');
        defItem.textContent = (!activeFolderId ? '✓ ' : '') + 'AI Chat Exports (default)';
        defItem.addEventListener('click', () => {
          chrome.storage.local.remove(['customFolderId', 'customFolderName']);
          folderBtn.textContent = 'AI Chat Exports ▾';
          folderDropdown.style.display = 'none';
        });
        folderDropdown.appendChild(defItem);

        if (!resp || !resp.success) {
          const err = document.createElement('div');
          err.style.cssText = 'padding:5px 12px;font-size:11px;color:#999';
          err.textContent = 'Sign in to see folders';
          folderDropdown.appendChild(err);
          return;
        }

        for (const f of resp.folders) {
          const item = document.createElement('div');
          const isActive = f.id === activeFolderId;
          item.className = 'folder-item' + (isActive ? ' folder-active' : '');
          item.textContent = (isActive ? '✓ ' : '') + f.name;
          item.title = f.name;
          item.addEventListener('click', () => {
            chrome.storage.local.set({ customFolderId: f.id, customFolderName: f.name });
            const n = f.name;
            folderBtn.textContent = (n.length > 20 ? n.slice(0, 18) + '…' : n) + ' ▾';
            folderDropdown.style.display = 'none';
          });
          folderDropdown.appendChild(item);
        }

        if (resp.folders.length === 0) {
          const empty = document.createElement('div');
          empty.style.cssText = 'padding:5px 12px;font-size:11px;color:#999';
          empty.textContent = 'No folders in Drive root';
          folderDropdown.appendChild(empty);
        }
      });
    });
  });

  // Close dropdown when clicking outside
  document.addEventListener('click', (e) => {
    if (!folderBtn.contains(e.target) && !folderDropdown.contains(e.target)) {
      folderDropdown.style.display = 'none';
    }
  });

  // Shortcut customize link
  document.getElementById('shortcut-customize')?.addEventListener('click', (e) => {
    e.preventDefault();
    chrome.tabs.create({ url: 'chrome://extensions/shortcuts' });
    window.close();
  });

  // Ask content script for platform + response count
  chrome.tabs.sendMessage(tab.id, { action: 'getPlatform' }, (resp) => {
    if (chrome.runtime.lastError || !resp?.platform) {
      emptyEl.style.display = 'block';
      return;
    }

    const count = resp.responseCount || 0;
    const countLabel = count === 1 ? '1 response' : `${count} responses`;
    headerSub.textContent = `on ${resp.platform} · ${countLabel}`;
    mainEl.style.display = 'block';

    // Load last export shortcut for this conversation
    try {
      const { hostname, pathname } = new URL(tab.url);
      const convKey = hostname + pathname;
      chrome.storage.local.get('lastExports', (d) => {
        const raw = (d.lastExports || {})[convKey];
        const last = Array.isArray(raw) ? raw[0] : raw;
        if (last) {
          const display = last.fileName.length > 24
            ? last.fileName.slice(0, 22) + '…'
            : last.fileName;
          lastNameEl.textContent = display;
          lastRow.href = last.url || `https://docs.google.com/document/d/${last.fileId}/edit`;
          lastRow.style.display = 'flex';
        }
      });
    } catch (_) {}
  });

  // Primary button: action matches current default mode
  btnExport.addEventListener('click', () => {
    const mode = modeSelect.value;
    if      (mode === 'last') send('exportLast');
    else if (mode === 'full') send('exportFull');
    else                      send('openPanel');
  });

  btnLast.addEventListener('click', () => send('exportLast'));
  btnAll.addEventListener('click',  () => send('exportFull'));
});
