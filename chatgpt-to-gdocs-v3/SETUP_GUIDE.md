# ChatGPT → Google Docs Exporter — Setup Guide

## What This Does

Click **"Export to Google Docs"** on any ChatGPT response → it instantly appears as a **Google Doc** with native, editable math equations. No .docx files, no dragging, no $20/month extension.

---

## Setup (10 minutes, one time only)

### Part 1: Load the Extension (2 min)

1. Unzip this folder to a permanent location (e.g., `Documents/chatgpt-to-gdocs-extension`)
2. Open Chrome → go to `chrome://extensions/`
3. Turn on **Developer mode** (top-right toggle)
4. Click **"Load unpacked"** → select the extension folder
5. **Copy your Extension ID** — it's the long string under the extension name (like `abcdefghijklmnopabcdefghijklmnop`). You'll need this in Part 2.

### Part 2: Create Google OAuth Credentials (8 min)

This lets the extension upload to YOUR Google Drive. It's free and your data never touches any server.

#### Step 1: Create a Google Cloud Project
1. Go to [console.cloud.google.com](https://console.cloud.google.com/)
2. Sign in with the same Google account you use for Google Drive
3. Click the project dropdown (top-left, next to "Google Cloud") → **New Project**
4. Name it anything (e.g., "ChatGPT Exporter") → **Create**
5. Make sure it's selected as the active project

#### Step 2: Enable the Google Drive API
1. Go to [console.cloud.google.com/apis/library/drive.googleapis.com](https://console.cloud.google.com/apis/library/drive.googleapis.com)
2. Click **Enable**

#### Step 3: Configure OAuth Consent Screen
1. Go to [console.cloud.google.com/apis/credentials/consent](https://console.cloud.google.com/apis/credentials/consent)
2. Select **External** → **Create**
3. Fill in:
   - App name: `ChatGPT Exporter` (anything you want)
   - User support email: your email
   - Developer contact email: your email
4. Click **Save and Continue**
5. On "Scopes" page → **Add or Remove Scopes** → search for `drive.file` → check the box for `../auth/drive.file` → **Update** → **Save and Continue**
6. On "Test users" page → **Add Users** → add your own email → **Save and Continue**
7. Click **Back to Dashboard**

#### Step 4: Create OAuth Client ID
1. Go to [console.cloud.google.com/apis/credentials](https://console.cloud.google.com/apis/credentials)
2. Click **+ Create Credentials** → **OAuth client ID**
3. Application type: **Chrome extension**
4. Name: `ChatGPT Exporter` (anything)
5. Item ID: paste your **Extension ID** from Part 1
6. Click **Create**
7. **Copy the Client ID** (looks like `123456789-abcdefg.apps.googleusercontent.com`)

#### Step 5: Put the Client ID in the Extension
1. Open the extension folder
2. Edit `manifest.json` in any text editor (Notepad, TextEdit, VS Code, etc.)
3. Find this line:
   ```
   "client_id": "YOUR_CLIENT_ID_HERE.apps.googleusercontent.com",
   ```
4. Replace `YOUR_CLIENT_ID_HERE.apps.googleusercontent.com` with your actual Client ID
5. Save the file

#### Step 6: Reload the Extension
1. Go back to `chrome://extensions/`
2. Click the **refresh icon** (🔄) on your extension
3. Done!

---

## How to Use

1. Go to **chatgpt.com** and ask a question
2. Click **"📄 Export to Google Docs"** (blue button, bottom-right)
3. First time: Google will ask you to sign in and grant permission → click Allow
4. The response is uploaded to your Google Drive and **opens automatically as a Google Doc**
5. All math equations are **native and editable** ✅

---

## Troubleshooting

### "Google Drive not set up yet" message
You haven't completed Part 2 above. The extension falls back to downloading a .docx file instead.

### "OAuth2 error" or "invalid_client"
- Double-check that the Client ID in `manifest.json` matches exactly what Google gave you
- Make sure the Extension ID in Google Cloud Console matches your extension (check `chrome://extensions/`)
- Make sure you added yourself as a test user in the OAuth consent screen

### "Access blocked: This app's request is invalid"
- Verify the Extension ID in your OAuth credentials matches the one shown in `chrome://extensions/`
- After changing the manifest, always click the refresh button on the extension

### "This app isn't verified" warning
This is normal for personal-use apps. Click **Advanced** → **Go to ChatGPT Exporter (unsafe)** → **Allow**. It's your own app, running locally, only accessing your own Drive.

### Equations show as text in Google Docs
Make sure the file opens as a Google Doc, not in Word preview mode. Right-click in Drive → **Open with → Google Docs**.

### Extension not showing buttons on ChatGPT
- Make sure you're on `chatgpt.com` (not `chat.openai.com`)
- Try refreshing the ChatGPT page
- Check the extension is enabled in `chrome://extensions/`

---

## How It Works (Technical)

1. Reads the ChatGPT response from the page DOM
2. Parses markdown + LaTeX into structured blocks
3. Converts LaTeX → OMML (Office Math Markup Language) — the native equation format for Word/Google Docs
4. Builds a valid .docx file as a ZIP (entirely in the browser, zero server calls)
5. Uploads the .docx to Google Drive using the Drive API with `mimeType: application/vnd.google-apps.document` which tells Google to auto-convert it to a native Google Doc
6. Opens the resulting Google Doc in a new tab

All processing happens locally in your browser. The only network request is the upload to YOUR Google Drive.

---

## Without Google Drive Setup

If you skip Part 2, the extension still works! It just downloads the .docx file and opens Google Drive so you can drag it in. You can always complete the setup later.
