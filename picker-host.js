document.addEventListener('DOMContentLoaded', () => {
  const frame = document.getElementById('picker-frame');
  let token = null;
  let frameReady = false;

  function trySend() {
    if (token !== null && frameReady) {
      frame.contentWindow.postMessage({ type: 'TOKEN', token }, '*');
    }
  }

  // Get OAuth token from background service worker
  chrome.runtime.sendMessage({ action: 'getPickerToken' }, (resp) => {
    token = resp?.token || null;
    trySend();
  });

  // Wait for sandboxed iframe to finish loading before sending token
  frame.addEventListener('load', () => {
    frameReady = true;
    trySend();
  });

  // Receive folder selection or cancel from sandboxed picker
  window.addEventListener('message', (e) => {
    if (e.data?.type === 'FOLDER_SELECTED') {
      chrome.storage.local.set({
        customFolderId: e.data.folderId,
        customFolderName: e.data.folderName
      }, () => window.close());
    } else if (e.data?.type === 'PICKER_CANCEL') {
      window.close();
    }
  });
});
