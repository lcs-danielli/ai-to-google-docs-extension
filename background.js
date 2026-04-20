/**
 * Background Service Worker
 * Handles Google OAuth and Drive upload
 */

// Listen for messages from content script
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'uploadToDrive') {
    handleUpload(request.docxBase64, request.filename)
      .then(result => sendResponse(result))
      .catch(err => sendResponse({ success: false, error: err.message }));
    return true; // Keep message channel open for async response
  }

  if (request.action === 'checkAuth') {
    getAuthToken(false)
      .then(token => sendResponse({ authenticated: !!token }))
      .catch(() => sendResponse({ authenticated: false }));
    return true;
  }
});

// Get OAuth token (interactive = show login popup)
function getAuthToken(interactive = true) {
  return new Promise((resolve, reject) => {
    chrome.identity.getAuthToken({ interactive }, (token) => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
      } else {
        resolve(token);
      }
    });
  });
}

// Remove cached token (for re-auth)
function removeCachedToken(token) {
  return new Promise((resolve) => {
    chrome.identity.removeCachedAuthToken({ token }, resolve);
  });
}

// Main upload flow
async function handleUpload(docxBase64, filename) {
  // 1. Get auth token
  let token;
  try {
    token = await getAuthToken(true);
  } catch (e) {
    throw new Error('Google sign-in failed: ' + e.message + '. Make sure you set up OAuth (see SETUP_GUIDE.md).');
  }

  if (!token) {
    throw new Error('No auth token received. Please try again.');
  }

  // 2. Convert base64 to binary
  const binaryString = atob(docxBase64);
  const bytes = new Uint8Array(binaryString.length);
  for (let i = 0; i < binaryString.length; i++) {
    bytes[i] = binaryString.charCodeAt(i);
  }

  // 3. Upload to Google Drive with conversion to Google Docs format
  const metadata = {
    name: filename.replace('.docx', ''),
    mimeType: 'application/vnd.google-apps.document' // Convert to Google Docs
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
        body: body
      }
    );
  } catch (e) {
    throw new Error('Network error uploading to Drive: ' + e.message);
  }

  // Handle token expiration
  if (response.status === 401) {
    await removeCachedToken(token);
    // Retry once with fresh token
    token = await getAuthToken(true);
    response = await fetch(
      'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,webViewLink',
      {
        method: 'POST',
        headers: {
          'Authorization': 'Bearer ' + token,
          'Content-Type': 'multipart/related; boundary=' + boundary
        },
        body: body
      }
    );
  }

  if (!response.ok) {
    const errText = await response.text();
    throw new Error('Drive API error (' + response.status + '): ' + errText);
  }

  const result = await response.json();

  return {
    success: true,
    fileId: result.id,
    fileName: result.name,
    url: result.webViewLink || ('https://docs.google.com/document/d/' + result.id + '/edit')
  };
}

// Build multipart request body for Drive API
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

  // Combine all parts
  const totalLength = metadataPart.length + fileHeader.length + fileBytes.length + fileFooter.length;
  const body = new Uint8Array(totalLength);
  let offset = 0;

  body.set(metadataPart, offset); offset += metadataPart.length;
  body.set(fileHeader, offset); offset += fileHeader.length;
  body.set(fileBytes, offset); offset += fileBytes.length;
  body.set(fileFooter, offset);

  return body;
}

chrome.runtime.onInstalled.addListener(() => {
  console.log('ChatGPT to Google Docs extension v2 installed');
});
