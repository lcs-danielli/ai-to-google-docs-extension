document.addEventListener('DOMContentLoaded', () => {
  const listEl    = document.getElementById('list');
  const qInput    = document.getElementById('q');
  const btnOk     = document.getElementById('btnOk');
  const btnCancel = document.getElementById('btnCancel');

  let allFolders = [];
  let selId   = '__default__';
  let selName = 'AI Chat Exports';

  // Pre-select whatever is currently saved
  chrome.storage.local.get(['customFolderId', 'customFolderName'], (d) => {
    if (d.customFolderId) {
      selId   = d.customFolderId;
      selName = d.customFolderName || '';
    }
    btnOk.disabled = false;
  });

  // Fetch Drive root folders via background
  chrome.runtime.sendMessage({ action: 'listFolders' }, (result) => {
    if (!result?.success) {
      listEl.innerHTML = '';
      const s = document.createElement('div');
      s.className = 'status';
      s.textContent = result?.error || 'Could not load folders. Make sure you are signed in to Google.';
      listEl.appendChild(s);
      return;
    }
    allFolders = result.folders || [];
    render('');
  });

  function render(q) {
    listEl.innerHTML = '';

    // Default option (always first)
    listEl.appendChild(makeItem('📂', 'AI Chat Exports', 'Default — auto-created by extension', '__default__'));

    const divider = document.createElement('div');
    divider.className = 'divider';
    listEl.appendChild(divider);

    const filtered = q
      ? allFolders.filter(f => f.name.toLowerCase().includes(q.toLowerCase()))
      : allFolders;

    if (allFolders.length === 0) {
      const s = document.createElement('div');
      s.className = 'status';
      s.textContent = 'No folders found in Drive root.';
      listEl.appendChild(s);
    } else if (filtered.length === 0) {
      const s = document.createElement('div');
      s.className = 'status';
      s.textContent = 'No matching folders.';
      listEl.appendChild(s);
    } else {
      filtered.forEach(f => listEl.appendChild(makeItem('📁', f.name, '', f.id)));
    }
  }

  function makeItem(icon, name, sub, id) {
    const div = document.createElement('div');
    div.className = 'item' + (selId === id ? ' sel' : '');

    const iconSpan = document.createElement('span');
    iconSpan.className = 'item-icon';
    iconSpan.textContent = icon;

    const textDiv = document.createElement('div');
    textDiv.style.cssText = 'flex:1;min-width:0';

    const nameDiv = document.createElement('div');
    nameDiv.className = 'item-name';
    nameDiv.textContent = name;
    textDiv.appendChild(nameDiv);

    if (sub) {
      const subDiv = document.createElement('div');
      subDiv.className = 'item-sub';
      subDiv.textContent = sub;
      textDiv.appendChild(subDiv);
    }

    div.appendChild(iconSpan);
    div.appendChild(textDiv);

    div.addEventListener('click', () => {
      selId   = id;
      selName = name;
      listEl.querySelectorAll('.item').forEach(el => el.classList.remove('sel'));
      div.classList.add('sel');
      btnOk.disabled = false;
    });

    return div;
  }

  qInput.addEventListener('input', () => render(qInput.value.trim()));

  btnOk.addEventListener('click', () => {
    const isDefault = selId === '__default__';
    if (isDefault) {
      chrome.storage.local.remove('customFolderId', () => {
        chrome.storage.local.set({ customFolderName: 'AI Chat Exports', pickerState: 'done' }, () => window.close());
      });
    } else {
      chrome.storage.local.set({
        customFolderId: selId,
        customFolderName: selName,
        pickerState: 'done'
      }, () => window.close());
    }
  });

  btnCancel.addEventListener('click', () => {
    chrome.storage.local.set({ pickerState: 'cancelled' }, () => window.close());
  });
});
