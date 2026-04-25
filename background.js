/**
 * Background Service Worker
 * Handles Google OAuth, Drive upload, Docs append, and keyboard shortcuts
 */

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'uploadToDrive') {
    handleUpload(request.docxBase64, request.filename)
      .then(result => sendResponse(result))
      .catch(err => sendResponse({ success: false, error: err.message }));
    return true;
  }

  if (request.action === 'appendToDoc') {
    appendContent(request.fileId, request.text)
      .then(result => sendResponse(result))
      .catch(err => sendResponse({ success: false, error: err.message }));
    return true;
  }

  if (request.action === 'checkAuth') {
    getAuthToken(false)
      .then(token => sendResponse({ authenticated: !!token }))
      .catch(() => sendResponse({ authenticated: false }));
    return true;
  }

  if (request.action === 'listFolders') {
    listDriveFolders()
      .then(result => sendResponse(result))
      .catch(err => sendResponse({ success: false, error: err.message }));
    return true;
  }

  if (request.action === 'fetchImage') {
    fetchImageAsBase64(request.url)
      .then(result => sendResponse(result))
      .catch(() => sendResponse({ success: false }));
    return true;
  }
});

// ── Keyboard shortcut: forward to active tab's content script ──
chrome.commands.onCommand.addListener((command) => {
  if (command === 'trigger-export') {
    chrome.tabs.query({ active: true, currentWindow: true }, ([tab]) => {
      if (tab) chrome.tabs.sendMessage(tab.id, { action: 'triggerDefault' }, () => {
        void chrome.runtime.lastError; // suppress "no receiver" error if not on AI page
      });
    });
  }
});

// ── Auth helpers ──
function getAuthToken(interactive = true) {
  return new Promise((resolve, reject) => {
    chrome.identity.getAuthToken({ interactive }, (token) => {
      if (chrome.runtime.lastError) reject(new Error(chrome.runtime.lastError.message));
      else resolve(token);
    });
  });
}

function removeCachedToken(token) {
  return new Promise((resolve) => {
    chrome.identity.removeCachedAuthToken({ token }, resolve);
  });
}

// ── List top-level Drive folders ──
async function listDriveFolders() {
  let token;
  try { token = await getAuthToken(true); }
  catch (e) { return { success: false, error: 'Not signed in' }; }
  if (!token) return { success: false, error: 'Not signed in' };

  const q = encodeURIComponent("mimeType='application/vnd.google-apps.folder' and 'root' in parents and trashed=false");
  const res = await fetch(
    `https://www.googleapis.com/drive/v3/files?q=${q}&fields=files(id,name)&orderBy=name&pageSize=50`,
    { headers: { 'Authorization': 'Bearer ' + token } }
  );
  if (!res.ok) return { success: false, error: 'Drive API error' };
  const data = await res.json();
  return { success: true, folders: data.files || [] };
}

// ── Fetch image via background worker (bypasses content script CORS) ──
async function fetchImageAsBase64(url) {
  try {
    // Request optional permission for this origin at runtime
    const origin = new URL(url).origin + '/*';
    const hasIt = await chrome.permissions.contains({ origins: [origin] });
    if (!hasIt) {
      const granted = await chrome.permissions.request({ origins: [origin] });
      if (!granted) return { success: false, denied: true };
    }
    const resp = await fetch(url);
    if (!resp.ok) return { success: false };
    const blob = await resp.blob();
    const bitmap = await createImageBitmap(blob);
    const maxPx = 1200;
    let w = bitmap.width, h = bitmap.height;
    if (!w || !h) { bitmap.close(); return { success: false }; }
    if (w > maxPx) { h = Math.round(h * maxPx / w); w = maxPx; }
    const canvas = new OffscreenCanvas(w, h);
    canvas.getContext('2d').drawImage(bitmap, 0, 0, w, h);
    bitmap.close();
    const pngBlob = await canvas.convertToBlob({ type: 'image/png' });
    const buffer = await pngBlob.arrayBuffer();
    const bytes = new Uint8Array(buffer);
    let binary = '';
    for (let i = 0; i < bytes.length; i += 8192) {
      binary += String.fromCharCode(...bytes.subarray(i, Math.min(i + 8192, bytes.length)));
    }
    return { success: true, base64: btoa(binary), w, h };
  } catch (_) {
    return { success: false };
  }
}

// ── Auto-create "AI Chat Exports" folder in Drive ──
async function getOrCreateExportFolder(token) {
  try {
    // User-selected folder takes priority
    const custom = await chrome.storage.local.get('customFolderId');
    if (custom.customFolderId) {
      const check = await fetch(
        `https://www.googleapis.com/drive/v3/files/${custom.customFolderId}?fields=id,trashed`,
        { headers: { 'Authorization': 'Bearer ' + token } }
      );
      if (check.ok) {
        const data = await check.json();
        if (!data.trashed) return custom.customFolderId;
      }
      await chrome.storage.local.remove(['customFolderId', 'customFolderName']);
    }

    const stored = await chrome.storage.local.get('exportFolderId');
    const folderId = stored.exportFolderId;

    if (folderId) {
      // Verify folder still exists and isn't trashed
      const check = await fetch(
        `https://www.googleapis.com/drive/v3/files/${folderId}?fields=id,trashed`,
        { headers: { 'Authorization': 'Bearer ' + token } }
      );
      if (check.ok) {
        const data = await check.json();
        if (!data.trashed) return folderId;
      }
      await chrome.storage.local.remove('exportFolderId');
    }

    // Create the folder
    const res = await fetch('https://www.googleapis.com/drive/v3/files?fields=id', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        name: 'AI Chat Exports',
        mimeType: 'application/vnd.google-apps.folder'
      })
    });

    if (!res.ok) return null; // silently fall back to Drive root
    const { id } = await res.json();
    await chrome.storage.local.set({ exportFolderId: id });
    return id;
  } catch {
    return null; // never block an export over folder issues
  }
}

// ── Drive upload ──
async function handleUpload(docxBase64, filename) {
  let token;
  try {
    token = await getAuthToken(true);
  } catch (e) {
    throw new Error('Google sign-in failed: ' + e.message + '. Make sure you set up OAuth (see SETUP_GUIDE.md).');
  }

  if (!token) throw new Error('No auth token received. Please try again.');

  const binaryString = atob(docxBase64);
  const bytes = new Uint8Array(binaryString.length);
  for (let i = 0; i < binaryString.length; i++) bytes[i] = binaryString.charCodeAt(i);

  const MAX_BYTES = 5 * 1024 * 1024;
  if (bytes.length > MAX_BYTES) {
    throw new Error(
      `Export too large (${(bytes.length / 1024 / 1024).toFixed(1)} MB). ` +
      'Google Drive multipart uploads are limited to 5 MB. Try exporting a shorter message.'
    );
  }

  // Get or create "AI Chat Exports" folder
  const folderId = await getOrCreateExportFolder(token);

  const metadata = {
    name: filename.replace('.docx', ''),
    mimeType: 'application/vnd.google-apps.document',
    ...(folderId ? { parents: [folderId] } : {})
  };

  const boundary = 'chatgpt_export_boundary_' + Date.now();
  const body = buildMultipartBody(boundary, metadata, bytes);

  let response;
  try {
    response = await fetch(
      'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,webViewLink',
      {
        method: 'POST',
        headers: {
          'Authorization': 'Bearer ' + token,
          'Content-Type': 'multipart/related; boundary=' + boundary
        },
        body
      }
    );
  } catch (e) {
    throw new Error('Network error uploading to Drive: ' + e.message);
  }

  if (response.status === 401) {
    await removeCachedToken(token);
    token = await getAuthToken(true);
    response = await fetch(
      'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,webViewLink',
      {
        method: 'POST',
        headers: {
          'Authorization': 'Bearer ' + token,
          'Content-Type': 'multipart/related; boundary=' + boundary
        },
        body
      }
    );
  }

  if (!response.ok) {
    const errText = await response.text();
    let reason = '';
    try { reason = JSON.parse(errText)?.error?.errors?.[0]?.reason || ''; } catch {}
    if (reason === 'storageQuotaExceeded') {
      throw new Error('Your Google Drive storage is full. Free up space and try again.');
    } else if (reason === 'userRateLimitExceeded' || reason === 'rateLimitExceeded') {
      throw new Error('Google Drive rate limit reached. Please wait a minute and try again.');
    } else {
      let msg = '';
      try { msg = JSON.parse(errText)?.error?.message || errText; } catch { msg = errText; }
      throw new Error(`Drive API error (${response.status}): ${msg}`);
    }
  }

  const result = await response.json();
  return {
    success: true,
    fileId: result.id,
    fileName: result.name,
    url: result.webViewLink || ('https://docs.google.com/document/d/' + result.id + '/edit')
  };
}

// ── Docs API: append text to existing Google Doc ──
async function appendContent(fileId, text) {
  let token;
  try { token = await getAuthToken(true); }
  catch (e) { throw new Error('Sign-in failed: ' + e.message); }

  const now = new Date();
  const dateStr = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}-${String(now.getDate()).padStart(2,'0')}`;
  const separator = '\n\n────────────────────────────────────────\n\n';
  const insertText = separator + 'Added ' + dateStr + '\n\n' + text;

  const docsRequest = async (tok) => fetch(
    `https://docs.googleapis.com/v1/documents/${fileId}:batchUpdate`,
    {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + tok, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        requests: [{ insertText: { endOfSegmentLocation: { segmentId: '' }, text: insertText } }]
      })
    }
  );

  let response = await docsRequest(token);

  if (response.status === 401) {
    await removeCachedToken(token);
    token = await getAuthToken(true);
    response = await docsRequest(token);
  }

  if (!response.ok) {
    const errText = await response.text();
    if (response.status === 403) {
      throw new Error('Google Docs API not enabled. Go to console.cloud.google.com → APIs → enable "Google Docs API".');
    }
    let msg = errText;
    try { msg = JSON.parse(errText)?.error?.message || errText; } catch {}
    throw new Error(`Docs API error (${response.status}): ${msg}`);
  }

  return { success: true };
}

// ── Build multipart body for Drive upload ──
function buildMultipartBody(boundary, metadata, fileBytes) {
  const encoder = new TextEncoder();
  const metadataPart = encoder.encode(
    '--' + boundary + '\r\n' +
    'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
    JSON.stringify(metadata) + '\r\n'
  );
  const fileHeader = encoder.encode(
    '--' + boundary + '\r\n' +
    'Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document\r\n' +
    'Content-Transfer-Encoding: binary\r\n\r\n'
  );
  const fileFooter = encoder.encode('\r\n--' + boundary + '--');

  const totalLength = metadataPart.length + fileHeader.length + fileBytes.length + fileFooter.length;
  const body = new Uint8Array(totalLength);
  let offset = 0;
  body.set(metadataPart, offset); offset += metadataPart.length;
  body.set(fileHeader, offset);   offset += fileHeader.length;
  body.set(fileBytes, offset);    offset += fileBytes.length;
  body.set(fileFooter, offset);
  return body;
}

