document.addEventListener('DOMContentLoaded', async () => {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab) return;

  const headerSub      = document.getElementById('header-sub');
  const mainEl         = document.getElementById('main');
  const emptyEl        = document.getElementById('empty');
  const statusEl       = document.getElementById('status');
  const btnExport      = document.getElementById('btn-export');
  const recentSection  = document.getElementById('recent-section');
  const recentList     = document.getElementById('recent-list');
  const optDrive       = document.getElementById('opt-drive');
  const optLocal       = document.getElementById('opt-local');
  const driveDetail    = document.getElementById('drive-detail');
  const changeFolderBtn = document.getElementById('change-folder');
  const gearBtn        = document.getElementById('gear-btn');
  const settingsPanel  = document.getElementById('settings-panel');

  let currentDest = 'drive';
  let currentMode = 'last';

  // ── Load saved settings ──
  chrome.storage.local.get(['exportDest', 'defaultExportMode', 'customFolderName'], (d) => {
    currentDest = d.exportDest || 'drive';
    currentMode = d.defaultExportMode || 'last';
    applyDest();
    applyMode();
    if (d.customFolderName) driveDetail.textContent = d.customFolderName;
  });

  function applyDest() {
    optDrive.classList.toggle('active', currentDest === 'drive');
    optLocal.classList.toggle('active', currentDest === 'local');
    updateExportLabel();
    updateRecent();
  }

  function applyMode() {
    document.querySelectorAll('.mode-tab').forEach(b => {
      b.classList.toggle('active', b.dataset.mode === currentMode);
    });
    updateExportLabel();
  }

  function updateExportLabel() {
    const drive = currentDest === 'drive';
    if (currentMode === 'last')
      btnExport.textContent = drive ? '↩ Export Last' : '↩ Save Last';
    else if (currentMode === 'full')
      btnExport.textContent = drive ? '≡ Export Full' : '≡ Save Full';
    else
      btnExport.textContent = drive ? '☑ Pick & Export' : '☑ Pick & Save';
  }

  // ── Destination selection ──
  [optDrive, optLocal].forEach(opt => {
    opt.addEventListener('click', (e) => {
      if (changeFolderBtn.contains(e.target)) return;
      currentDest = opt.dataset.dest;
      chrome.storage.local.set({ exportDest: currentDest });
      applyDest();
    });
  });

  // ── Change folder (›) ──
  changeFolderBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    currentDest = 'drive';
    chrome.storage.local.set({ exportDest: 'drive' });
    applyDest();
    chrome.windows.create({
      url: chrome.runtime.getURL('picker-host.html'),
      type: 'popup', width: 640, height: 540
    });
    window.close();
  });

  // ── Mode tabs ──
  document.querySelectorAll('.mode-tab').forEach(b => {
    b.addEventListener('click', () => {
      currentMode = b.dataset.mode;
      chrome.storage.local.set({ defaultExportMode: currentMode });
      applyMode();
    });
  });

  // ── Gear / settings ──
  gearBtn.addEventListener('click', () => {
    const open = settingsPanel.style.display !== 'none';
    settingsPanel.style.display = open ? 'none' : 'block';
    gearBtn.classList.toggle('on', !open);
  });

  // ── Shortcuts ──
  document.getElementById('shortcut-customize').addEventListener('click', (e) => {
    e.preventDefault();
    chrome.tabs.create({ url: 'chrome://extensions/shortcuts' });
    window.close();
  });
  document.getElementById('customize-shortcut').addEventListener('click', () => {
    chrome.tabs.create({ url: 'chrome://extensions/shortcuts' });
    window.close();
  });

  // ── Export ──
  function setWorking() {
    btnExport.disabled = true;
    statusEl.textContent = 'Working…';
  }

  function send(action) {
    setWorking();
    chrome.tabs.sendMessage(tab.id, { action, dest: currentDest }, () => {
      if (chrome.runtime.lastError) {
        statusEl.textContent = 'Error — try refreshing the page';
        btnExport.disabled = false;
        return;
      }
      setTimeout(() => window.close(), 400);
    });
  }

  btnExport.addEventListener('click', () => {
    if (currentMode === 'last')      send('exportLast');
    else if (currentMode === 'full') send('exportFull');
    else                             send('openPanel');
  });

  // ── Recent exports (Drive mode only) ──
  function updateRecent() {
    if (currentDest !== 'drive') { recentSection.style.display = 'none'; return; }
    try {
      const { hostname, pathname } = new URL(tab.url);
      const convKey = hostname + pathname;
      chrome.storage.local.get(['lastExports', 'globalRecentDocs'], (d) => {
        const raw = (d.lastExports || {})[convKey] || null;
        const convHistory = Array.isArray(raw) ? raw : (raw ? [raw] : []);
        const convIds = new Set(convHistory.map(e => e.fileId));
        const globalExtra = (Array.isArray(d.globalRecentDocs) ? d.globalRecentDocs : [])
          .filter(e => !convIds.has(e.fileId));
        const all = [...convHistory, ...globalExtra].slice(0, 5);

        if (all.length === 0) { recentSection.style.display = 'none'; return; }
        recentSection.style.display = 'block';
        recentList.innerHTML = '';
        for (const exp of all) {
          const a = document.createElement('a');
          a.className = 'recent-chip';
          a.href = exp.url || `https://docs.google.com/document/d/${exp.fileId}/edit`;
          a.target = '_blank';
          const name = exp.fileName || 'Untitled';
          const short = name.length > 26 ? name.slice(0, 24) + '…' : name;
          a.innerHTML = `<span class="r-icon">↩</span><span class="r-name">${short}</span><span class="r-arrow">↗</span>`;
          a.title = name;
          recentList.appendChild(a);
        }
      });
    } catch (_) {
      recentSection.style.display = 'none';
    }
  }

  // ── Platform detection ──
  chrome.tabs.sendMessage(tab.id, { action: 'getPlatform' }, (resp) => {
    if (chrome.runtime.lastError || !resp?.platform) {
      emptyEl.style.display = 'block';
      return;
    }
    const count = resp.responseCount || 0;
    headerSub.textContent = `${resp.platform} · ${count} response${count !== 1 ? 's' : ''}`;
    mainEl.style.display = 'block';
    updateRecent();
  });
});
