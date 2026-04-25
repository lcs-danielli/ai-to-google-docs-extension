# AI Chat Exporter for Google Docs

A Chrome extension that adds an **Export to Docs** button directly on ChatGPT, Gemini, and Claude. One click exports any AI response — with full formatting — to Google Docs, saved straight to your Google Drive.

## Supported Platforms

- ChatGPT (chatgpt.com)
- Google Gemini (gemini.google.com)
- Claude (claude.ai)

## Features

- **One-click export** — button appears next to every AI response
- **Selection panel** — choose individual responses or export the full conversation
- **Append to existing docs** — add new content to a previously exported document, with source metadata (platform, title, date, URL)
- **Folder picker** — choose which Google Drive folder to save to
- **Formatting preserved** — headings, bold/italic, code blocks, tables, math equations, lists
- **CSV export** — download any table in a response as a `.csv` file
- **Keyboard shortcut** — `Cmd+Shift+E` / `Ctrl+Shift+E`
- **Dark mode** support

## Demo

<img width="373" height="85" alt="image" src="https://github.com/user-attachments/assets/b74e0e3f-415d-4437-b2ce-b34139d7f26e" />
<img width="378" height="100" alt="image" src="https://github.com/user-attachments/assets/524279d8-e112-49fc-971b-be07bcd9ea67" />

## Installation (Development)

1. Clone this repository
2. Open `chrome://extensions` in Chrome
3. Enable **Developer Mode**
4. Click **Load unpacked** and select this folder
5. Sign in with your Google account when prompted

## Tech Stack

- JavaScript (Manifest V3)
- Chrome Extension APIs: `identity`, `storage`, `commands`
- Google Drive API v3 / Google Docs API v1
- Office Open XML (OOXML) for `.docx` generation

## Privacy

This extension processes all data locally in your browser. No data is sent to any developer-owned server. See the full [Privacy Policy](https://daniel-zeyuli.github.io/ai-to-google-docs-extension/privacy.html).

## Author

Daniel Li  
GitHub: https://github.com/daniel-zeyuli
