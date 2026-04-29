const PICKER_API_KEY = 'AIzaSyCGZjDDRW6e6ogByzjwduNYC3XiQVBEpjY';

function gapiLoaded() {
  document.getElementById('status').textContent = 'Signing in…';
  chrome.runtime.sendMessage({ action: 'getPickerToken' }, (resp) => {
    if (chrome.runtime.lastError || !resp?.token) {
      document.getElementById('status').textContent = 'Sign-in failed. Close and try again.';
      return;
    }
    document.getElementById('status').textContent = 'Opening folder browser…';
    gapi.load('picker', () => openPicker(resp.token));
  });
}

function openPicker(token) {
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
    chrome.storage.local.set(
      { customFolderId: folder.id, customFolderName: folder.name },
      () => window.close()
    );
  } else if (data.action === google.picker.Action.CANCEL) {
    window.close();
  }
}
