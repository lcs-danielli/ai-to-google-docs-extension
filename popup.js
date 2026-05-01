document.addEventListener('DOMContentLoaded', () => {
  const optDrive       = document.getElementById('opt-drive');
  const optLocal       = document.getElementById('opt-local');
  const driveDetail    = document.getElementById('drive-detail');
  const changeFolderBtn = document.getElementById('change-folder');
  const modeHint       = document.getElementById('mode-hint');
  const shortcutDisplay = document.getElementById('shortcut-display');
  const shortcutCustomize = document.getElementById('shortcut-customize');

  let currentDest = 'drive';
  let currentMode = 'last';

  const MODE_HINTS = {
    last:   'Exports the last AI response',
    full:   'Exports the full conversation',
    select: 'Opens panel to pick responses'
  };

  // ── Load saved settings ──
  chrome.storage.local.get(['exportDest', 'defaultExportMode', 'customFolderName'], (d) => {
    currentDest = d.exportDest || 'drive';
    currentMode = d.defaultExportMode || 'last';
    applyDest();
    applyMode();
    if (d.customFolderName) driveDetail.textContent = d.customFolderName;
  });

  // ── Read real keyboard shortcut ──
  chrome.commands.getAll((commands) => {
    const cmd = commands.find(c => c.name === 'trigger-export');
    if (cmd && cmd.shortcut) shortcutDisplay.textContent = cmd.shortcut;
    else shortcutDisplay.textContent = 'Not set';
  });

  function applyDest() {
    optDrive.classList.toggle('active', currentDest === 'drive');
    optLocal.classList.toggle('active', currentDest === 'local');
  }

  function applyMode() {
    document.querySelectorAll('.mode-tab').forEach(b => {
      b.classList.toggle('active', b.dataset.mode === currentMode);
    });
    modeHint.textContent = MODE_HINTS[currentMode] || '';
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
      type: 'popup', width: 420, height: 440
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

  // ── Keyboard shortcut customize ──
  shortcutCustomize.addEventListener('click', () => {
    chrome.tabs.create({ url: 'chrome://extensions/shortcuts' });
    window.close();
  });
});
