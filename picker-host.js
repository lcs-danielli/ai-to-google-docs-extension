document.addEventListener('DOMContentLoaded', () => {
  const listEl     = document.getElementById('list');
  const qInput     = document.getElementById('q');
  const btnOk      = document.getElementById('btnOk');
  const btnCancel  = document.getElementById('btnCancel');
  const newArea    = document.getElementById('newArea');
  const newInput   = document.getElementById('newInput');
  const btnCreate  = document.getElementById('btnCreate');

  let allFolders = [];
  let selId   = '__default__';
  let selName = 'AI Chat Exports';
  let token   = null;

  // Pre-select whatever is currently saved
  chrome.storage.local.get(['customFolderId', 'customFolderName'], (d) => {
    if (d.customFolderId) { selId = d.customFolderId; selName = d.customFolderName || ''; }
    btnOk.disabled = false;
  });

  loadFolders();

  // ── Auth (direct — extension pages can call chrome.identity) ──
  async function getToken(interactive) {
    return new Promise((resolve) => {
      chrome.identity.getAuthToken({ interactive }, (t) => {
        resolve((chrome.runtime.lastError || !t) ? null : t);
      });
    });
  }

  async function loadFolders() {
    setStatus('Loading…');

    token = await getToken(false);
    if (!token) token = await getToken(true);
    if (!token) { showSignIn(); return; }

    // List ALL Drive folders (no root restriction — user may have nested folders)
    try {
      const q = encodeURIComponent(
        "mimeType='application/vnd.google-apps.folder' and trashed=false"
      );
      const res = await fetch(
        `https://www.googleapis.com/drive/v3/files?q=${q}&fields=files(id,name)&orderBy=name&pageSize=100`,
        { headers: { Authorization: 'Bearer ' + token } }
      );
      if (res.status === 401) {
        await new Promise(r => chrome.identity.removeCachedAuthToken({ token }, r));
        token = null; loadFolders(); return;
      }
      if (!res.ok) throw new Error(`Drive API error (${res.status})`);
      const data = await res.json();
      allFolders = data.files || [];
      render('');
    } catch (e) {
      setStatus('Could not load folders: ' + e.message);
    }
  }

  function showSignIn() {
    listEl.innerHTML = '';
    const wrap = document.createElement('div');
    wrap.style.cssText = 'padding:20px;text-align:center';
    const msg = document.createElement('p');
    msg.style.cssText = 'font-size:12px;color:#999;margin-bottom:12px';
    msg.textContent = 'Sign in to Google to pick a folder.';
    const btn = document.createElement('button');
    btn.className = 'btn btn-ok'; btn.textContent = 'Sign in to Google';
    btn.addEventListener('click', loadFolders);
    wrap.appendChild(msg); wrap.appendChild(btn);
    listEl.appendChild(wrap);
  }

  function setStatus(text) {
    listEl.innerHTML = '';
    const d = document.createElement('div');
    d.className = 'status'; d.textContent = text;
    listEl.appendChild(d);
  }

  // ── Render folder list ─────────────────────────────────────────
  function render(q) {
    listEl.innerHTML = '';

    // Default option always at top
    listEl.appendChild(makeItem('📂', 'AI Chat Exports', 'Default — auto-created by extension', '__default__'));
    const div = document.createElement('div'); div.className = 'divider';
    listEl.appendChild(div);

    const matches = q
      ? allFolders.filter(f => f.name.toLowerCase().includes(q.toLowerCase()))
      : allFolders;

    if (allFolders.length === 0) {
      listEl.appendChild(statusEl('No folders in your Drive yet.'));
    } else if (matches.length === 0) {
      listEl.appendChild(statusEl('No matching folders.'));
    } else {
      matches.forEach(f => listEl.appendChild(makeItem('📁', f.name, '', f.id)));
    }
  }

  function statusEl(text) {
    const d = document.createElement('div'); d.className = 'status'; d.textContent = text; return d;
  }

  function makeItem(icon, name, sub, id) {
    const div = document.createElement('div');
    div.className = 'item' + (selId === id ? ' sel' : '');

    const ic = document.createElement('span'); ic.className = 'item-icon'; ic.textContent = icon;
    const tx = document.createElement('div'); tx.style.cssText = 'flex:1;min-width:0';
    const nm = document.createElement('div'); nm.className = 'item-name'; nm.textContent = name;
    tx.appendChild(nm);
    if (sub) { const sb = document.createElement('div'); sb.className = 'item-sub'; sb.textContent = sub; tx.appendChild(sb); }

    div.appendChild(ic); div.appendChild(tx);
    div.addEventListener('click', () => {
      selId = id; selName = name;
      listEl.querySelectorAll('.item').forEach(el => el.classList.remove('sel'));
      div.classList.add('sel');
      btnOk.disabled = false;
    });
    return div;
  }

  qInput.addEventListener('input', () => render(qInput.value.trim()));

  // ── Create new folder ──────────────────────────────────────────
  document.getElementById('btnNewFolder').addEventListener('click', () => {
    newArea.style.display = newArea.style.display === 'none' ? 'flex' : 'none';
    if (newArea.style.display === 'flex') newInput.focus();
  });

  btnCreate.addEventListener('click', createFolder);
  newInput.addEventListener('keydown', e => { if (e.key === 'Enter') createFolder(); });

  async function createFolder() {
    const name = newInput.value.trim();
    if (!name) return;
    btnCreate.disabled = true; btnCreate.textContent = 'Creating…';
    try {
      if (!token) token = await getToken(true);
      const res = await fetch('https://www.googleapis.com/drive/v3/files?fields=id,name', {
        method: 'POST',
        headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
        body: JSON.stringify({ name, mimeType: 'application/vnd.google-apps.folder' })
      });
      if (!res.ok) throw new Error(`Drive API error (${res.status})`);
      const folder = await res.json();
      allFolders = [folder, ...allFolders];
      selId = folder.id; selName = folder.name;
      newInput.value = ''; newArea.style.display = 'none';
      render(qInput.value.trim());
      btnOk.disabled = false;
    } catch (e) {
      alert('Could not create folder: ' + e.message);
    }
    btnCreate.disabled = false; btnCreate.textContent = 'Create';
  }

  // ── Select / Cancel ────────────────────────────────────────────
  btnOk.addEventListener('click', () => {
    if (selId === '__default__') {
      chrome.storage.local.remove('customFolderId', () =>
        chrome.storage.local.set({ customFolderName: 'AI Chat Exports', pickerState: 'done' }, () => window.close())
      );
    } else {
      chrome.storage.local.set({ customFolderId: selId, customFolderName: selName, pickerState: 'done' }, () => window.close());
    }
  });

  btnCancel.addEventListener('click', () =>
    chrome.storage.local.set({ pickerState: 'cancelled' }, () => window.close())
  );
});
