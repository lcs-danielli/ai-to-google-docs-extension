// Sandboxed page — no chrome.* APIs. Communicates via postMessage.
const PICKER_API_KEY = 'AIzaSyCGZjDDRW6e6ogByzjwduNYC3XiQVBEpjY';
let _token = null;
let _pickerReady = false;

// Receive OAuth token from picker-host.js
window.addEventListener('message', (e) => {
  if (e.data?.type === 'TOKEN') {
    _token = e.data.token;
    if (_pickerReady) openPicker(_token);
  }
});

function gapiLoaded() {
  gapi.load('picker', () => {
    _pickerReady = true;
    if (_token !== null) openPicker(_token);
  });
}

function openPicker(token) {
  if (!token) {
    document.getElementById('status').textContent = 'Sign-in failed. Close and try again.';
    return;
  }
  const view = new google.picker.DocsView(google.picker.ViewId.FOLDERS)
    .setSelectFolderEnabled(true)
    .setMode(google.picker.DocsViewMode.LIST);

  const picker = new google.picker.PickerBuilder()
    .addView(view)
    .setOAuthToken(token)
    .setDeveloperKey(PICKER_API_KEY)
    .setCallback(pickerCallback)
    .setTitle('Choose a folder for exports')
    .build();

  picker.setVisible(true);
  document.getElementById('status').textContent = '';
}

function pickerCallback(data) {
  if (data.action === google.picker.Action.PICKED) {
    const folder = data.docs[0];
    window.parent.postMessage({
      type: 'FOLDER_SELECTED',
      folderId: folder.id,
      folderName: folder.name
    }, '*');
  } else if (data.action === google.picker.Action.CANCEL) {
    window.parent.postMessage({ type: 'PICKER_CANCEL' }, '*');
  }
}
