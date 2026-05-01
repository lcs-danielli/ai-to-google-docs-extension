/**
 * AI Chat Exporter — Content Script
 * Supports ChatGPT, Gemini, and Claude.
 */

(function() {
  'use strict';

  const BUTTON_CLASS = 'cgd-export-btn';
  const BANNER_CLASS = 'cgd-toast';

  // Detect platform
  const isGemini = location.hostname.includes('gemini.google.com');
  const isChatGPT = location.hostname.includes('chatgpt.com') || location.hostname.includes('chat.openai.com');
  const isClaude = location.hostname.includes('claude.ai');

  let exportDest = 'drive';
  chrome.storage.local.get('exportDest', d => { exportDest = d.exportDest || 'drive'; });

  // Update all injected export buttons when destination or folder changes
  chrome.storage.onChanged.addListener((changes) => {
    if (changes.exportDest || changes.customFolderName) {
      if (changes.exportDest) exportDest = changes.exportDest.newValue || 'drive';
      document.querySelectorAll('.' + BUTTON_CLASS).forEach(btn => _updateExportBtnContent(btn));
    }
  });

  function isDarkMode() {
    const root = document.documentElement;
    const body = document.body;
    if (root.classList.contains('dark') || body.classList.contains('dark')) return true;
    if (root.getAttribute('data-theme') === 'dark' || body.getAttribute('data-theme') === 'dark') return true;
    if (root.getAttribute('data-color-scheme') === 'dark' || root.getAttribute('data-color-mode') === 'dark') return true;
    if (window.matchMedia('(prefers-color-scheme: dark)').matches) return true;
    // Fallback: measure body background luminance
    const bg = window.getComputedStyle(body).backgroundColor;
    const rgb = bg.match(/\d+/g);
    if (rgb && rgb.length >= 3) {
      const lum = (0.299 * +rgb[0] + 0.587 * +rgb[1] + 0.114 * +rgb[2]) / 255;
      return lum < 0.4;
    }
    return false;
  }

  function escHtml(str) {
    return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  }

  function showToast(message, isError = false, duration = 4000) {
    const existing = document.querySelector('.' + BANNER_CLASS);
    if (existing) existing.remove();
    const toast = document.createElement('div');
    toast.className = BANNER_CLASS;
    toast.innerHTML = message;
    toast.style.cssText = `position:fixed;bottom:24px;left:50%;transform:translateX(-50%);padding:12px 24px;border-radius:8px;z-index:99999;font-size:14px;font-family:-apple-system,sans-serif;color:white;box-shadow:0 4px 12px rgba(0,0,0,0.3);background:${isError?'#d93025':'#1a7f37'};transition:opacity 0.3s;max-width:600px;text-align:center;`;
    document.body.appendChild(toast);
    setTimeout(() => { toast.style.opacity = '0'; setTimeout(() => toast.remove(), 300); }, duration);
  }

  // ═══════════════════════════════════════════════════════════════
  //  IMAGE CAPTURE (canvas-based, best-effort)
  // ═══════════════════════════════════════════════════════════════

  const _imgCaptures = [];
  let _imgIdx = 0;
  function _resetImgCaptures() { _imgCaptures.length = 0; _imgIdx = 0; }

  async function _captureImages() {
    const map = {};
    let shownPermissionToast = false;
    for (const { idx, el, alt } of _imgCaptures) {
      // Strategy 1: canvas (works if same-origin or CORS-permissive)
      if (el && el.naturalWidth && el.naturalHeight) {
        try {
          const maxPx = 1200;
          let w = el.naturalWidth, h = el.naturalHeight;
          if (w > maxPx) { h = Math.round(h * maxPx / w); w = maxPx; }
          const c = document.createElement('canvas');
          c.width = w; c.height = h;
          c.getContext('2d').drawImage(el, 0, 0, w, h);
          const b64 = c.toDataURL('image/png').split(',')[1];
          if (b64) { map[idx] = { data: b64, w, h, alt: alt || '' }; continue; }
        } catch (_) { /* tainted canvas — try fetch */ }
      }
      // Strategy 2: background worker fetch → OffscreenCanvas → PNG
      const src = (el && (el.src || el.getAttribute('src'))) || '';
      if (!src || src.startsWith('data:')) continue;
      try {
        const result = await new Promise(resolve => {
          chrome.runtime.sendMessage({ action: 'fetchImage', url: src }, resp => {
            resolve(resp || { success: false });
          });
        });
        if (result.success && result.base64 && result.w && result.h) {
          map[idx] = { data: result.base64, w: result.w, h: result.h, alt: alt || '' };
        } else if (result.denied && !shownPermissionToast) {
          shownPermissionToast = true;
          showToast('⚠️ Image embedding needs permission. Click "Allow" when Chrome asks to enable image export.', false, 6000);
        }
      } catch (_) { /* background fetch failed — will show [Image] placeholder */ }
    }
    return map;
  }

  // ═══════════════════════════════════════════════════════════════
  //  MARKDOWN EXTRACTION (shared logic)
  // ═══════════════════════════════════════════════════════════════

  function processNode(node) {
    try {
      if (!node) return '';
      if (node.nodeType === Node.TEXT_NODE) return node.textContent || '';
      if (node.nodeType !== Node.ELEMENT_NODE) return '';
      const tag = node.tagName ? node.tagName.toLowerCase() : '';
      if (!tag) return '';

      // === MATH DETECTION (multiple strategies) ===

      // Strategy 0: Gemini's math-inline / math-block with data-math attribute
      if (node.classList && (node.classList.contains('math-inline') || node.classList.contains('math-block'))) {
        const tex = node.getAttribute('data-math');
        if (tex) {
          const isDisplay = node.classList.contains('math-block');
          return isDisplay ? `\n$$${tex}$$\n` : `$${tex}$`;
        }
      }

      // Strategy 1: KaTeX display math wrapper
      if (node.classList && node.classList.contains('katex-display')) {
        const tex = extractTeX(node);
        if (tex) return `\n$$${tex}$$\n`;
      }

      // Strategy 2: KaTeX inline math
      if (node.classList && node.classList.contains('katex')) {
        const tex = extractTeX(node);
        if (tex) {
          const isDisplay = node.closest('.katex-display');
          return isDisplay ? `\n$$${tex}$$\n` : `$${tex}$`;
        }
      }

      // Strategy 3: wrapper span/div that directly contains .katex-mathml as a child
      // (Strategies 1/2 handle .katex-display and .katex; this catches any remaining wrapper)
      if (node.querySelector && node.querySelector(':scope > .katex-mathml annotation[encoding="application/x-tex"]')) {
        const tex = extractTeX(node);
        if (tex) {
          const isDisplay = !!node.closest('.katex-display');
          return isDisplay ? `\n$$${tex}$$\n` : `$${tex}$`;
        }
      }

      // Strategy 4: MathJax v3 (Gemini)
      if (tag === 'mjx-container') {
        const tex = node.getAttribute('data-formula') || node.getAttribute('aria-label') || '';
        if (tex) {
          const isDisplay = node.hasAttribute('display');
          return isDisplay ? `\n$$${tex}$$\n` : `$${tex}$`;
        }
      }

      // Strategy 5: <math> element directly
      if (tag === 'math') {
        const ann = node.querySelector('annotation[encoding="application/x-tex"]');
        if (ann) return `$${ann.textContent.trim()}$`;
      }

      // Strategy 6: Any span/div with a <math> descendant containing annotation
      // But DON'T process if it's a large container — only small math wrappers
      if ((tag === 'span' || tag === 'div') && node.childNodes.length <= 5) {
        const ann = node.querySelector('annotation[encoding="application/x-tex"]');
        if (ann && !node.querySelector('p') && !node.querySelector('li')) {
          const isDisplay = node.classList.contains('katex-display') || node.closest('.katex-display');
          const tex = ann.textContent.trim();
          return isDisplay ? `\n$$${tex}$$\n` : `$${tex}$`;
        }
      }

      // Skip katex-html (the visible rendered version) — we only want the annotation
      if (node.classList && node.classList.contains('katex-html')) return '';
      if (node.classList && node.classList.contains('katex-mathml')) return '';

      // === STANDARD HTML ELEMENTS ===

      if (/^h[1-6]$/.test(tag)) return `\n${'#'.repeat(parseInt(tag[1]))} ${getInner(node)}\n`;
      if (tag === 'p') return `\n${getInner(node)}\n`;
      if (tag === 'ol') { let r = '\n', n = 1; for (const li of node.querySelectorAll(':scope > li')) { r += `${n}. ${getInner(li).replace(/\n+/g, ' ').trim()}\n`; n++; } return r; }
      if (tag === 'ul') { let r = '\n'; for (const li of node.querySelectorAll(':scope > li')) r += `- ${getInner(li).replace(/\n+/g, ' ').trim()}\n`; return r; }
      if (tag === 'pre') { const code = node.querySelector('code'); return `\n\`\`\`\n${(code || node).textContent}\n\`\`\`\n`; }
      if (tag === 'code' && !node.closest('pre')) return `\`${node.textContent}\``;
      if (tag === 'img') {
        const alt = node.getAttribute('alt') || node.getAttribute('aria-label') || '';
        const src = node.src || node.getAttribute('src') || '';
        if (src && !src.startsWith('data:image/svg')) {
          _imgCaptures.push({ idx: _imgIdx, el: node, alt });
          return `[[IMG:${_imgIdx++}]]`;
        }
        return alt || '[Image]';
      }
      if (tag === 'br') return '\n';
      if (tag === 'hr') return '\n---\n';
      if (tag === 'strong' || tag === 'b') return `**${getInner(node)}**`;
      if (tag === 'em' || tag === 'i') return `*${getInner(node)}*`;
      if (tag === 'sub') return `~${getInner(node)}~`;
      if (tag === 'sup') return `^${getInner(node)}^`;
      if (tag === 'table') return processTable(node);
      return getInner(node);
    } catch(e) { return node.textContent || ''; }
  }

  // Extract TeX from any element that might contain KaTeX/MathJax annotations
  function extractTeX(el) {
    // Try data-math attribute (Gemini)
    const dataMath = el.getAttribute('data-math');
    if (dataMath) return dataMath.trim();
    // Try annotation element (ChatGPT KaTeX)
    const ann = el.querySelector('annotation[encoding="application/x-tex"]');
    if (ann) return ann.textContent.trim();
    // Try MathJax script
    const script = el.querySelector('script[type="math/tex"], script[type="math/tex; mode=display"]');
    if (script) return script.textContent.trim();
    // Try other data attributes
    const formula = el.getAttribute('data-formula') || el.getAttribute('aria-label');
    if (formula) return formula.trim();
    return '';
  }

  function getInner(node) { let r = ''; for (const c of node.childNodes) r += processNode(c); return r; }

  function processTable(table) {
    let md = '\n';
    const rows = table.querySelectorAll('tr');
    rows.forEach((row, idx) => {
      const cells = Array.from(row.querySelectorAll('th, td')).map(c => getInner(c).trim());
      md += '| ' + cells.join(' | ') + ' |\n';
      if (idx === 0) md += '| ' + cells.map(() => '---').join(' | ') + ' |\n';
    });
    return md;
  }

  function extractMarkdown(messageEl) {
    try {
      let contentDiv;
      if (isChatGPT) {
        contentDiv = messageEl.querySelector('.markdown') || messageEl;
      } else if (isGemini) {
        contentDiv = messageEl.querySelector('.markdown-main-panel') ||
                     messageEl.querySelector('.model-response-text') ||
                     messageEl.querySelector('.response-content') ||
                     messageEl;
      } else if (isClaude) {
        contentDiv = messageEl.querySelector('.font-claude-response') ||
                     messageEl.querySelector('[class*="font-claude-response"]') ||
                     messageEl;
      } else {
        contentDiv = messageEl;
      }
      let md = '';
      for (const child of contentDiv.childNodes) md += processNode(child);

      // VALIDATION: Check if math annotations exist but weren't captured
      const annotations = contentDiv.querySelectorAll('annotation[encoding="application/x-tex"], .math-inline[data-math], .math-block[data-math]');
      if (annotations.length > 0) {
        const hasMath = md.includes('$');
        if (!hasMath) md = directExtractWithMath(contentDiv);
      }

      return md.trim();
    } catch(e) {
      return messageEl.textContent || messageEl.innerText || '';
    }
  }

  // Direct extraction: walk the DOM more carefully, explicitly finding all math
  function directExtractWithMath(container) {
    let md = '';
    const walker = document.createTreeWalker(container, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT);
    const processedMathRoots = new Set();
    let node;

    while (node = walker.nextNode()) {
      // Skip descendants of already-processed math subtrees
      let el = node.parentElement;
      let insideMath = false;
      while (el && el !== container) {
        if (processedMathRoots.has(el)) { insideMath = true; break; }
        el = el.parentElement;
      }
      if (insideMath) continue;

      if (node.nodeType === Node.TEXT_NODE) {
        const parent = node.parentElement;
        if (parent && (
          parent.closest('.katex-html') ||
          parent.closest('.katex-mathml') ||
          parent.closest('annotation') ||
          parent.tagName === 'ANNOTATION'
        )) continue;
        md += node.textContent;
      } else {
        if (node.classList && (node.classList.contains('katex-display') || node.classList.contains('katex'))) {
          processedMathRoots.add(node);
          const tex = extractTeX(node);
          if (tex) {
            const isDisplay = node.classList.contains('katex-display') || !!node.closest('.katex-display');
            md += isDisplay ? `\n$$${tex}$$\n` : `$${tex}$`;
          }
          continue;
        }
        const blockTag = node.tagName.toLowerCase();
        if (['p', 'div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'li', 'br'].includes(blockTag)) {
          if (blockTag === 'br') md += '\n';
          else if (/^h[1-6]$/.test(blockTag)) md += '\n' + '#'.repeat(parseInt(blockTag[1])) + ' ';
          else if (blockTag === 'li') md += '\n- ';
          else md += '\n';
        }
        if (blockTag === 'strong' || blockTag === 'b') md += '**';
      }
    }
    return md;
  }

  // ═══════════════════════════════════════════════════════════════
  //  CONVERSATION TITLE EXTRACTION
  // ═══════════════════════════════════════════════════════════════

  function getConversationTitle() {
    let title = '';

    if (isChatGPT) {
      // Active conversation in sidebar
      const el = document.querySelector('nav [aria-current="page"] [class*="truncate"]') ||
                 document.querySelector('nav li.active [class*="truncate"]') ||
                 document.querySelector('nav [data-active="true"] [class*="truncate"]');
      if (el) title = el.textContent.trim();
    } else if (isGemini) {
      // Active conversation in sidebar
      const el = document.querySelector('.conversation-title[aria-selected="true"]') ||
                 document.querySelector('[class*="conversation-title"][class*="selected"]') ||
                 document.querySelector('chat-window-title-bar');
      if (el) title = el.textContent.trim();
    } else if (isClaude) {
      // Active conversation in sidebar
      const el = document.querySelector('nav [aria-current="page"]') ||
                 document.querySelector('[class*="ConversationTitle"]') ||
                 document.querySelector('nav a[class*="active"] [class*="truncate"]');
      if (el) title = el.textContent.trim();
    }

    // Fallback: strip platform suffix from document.title
    if (!title) {
      title = document.title
        .replace(/\s*[-|–]\s*(Google\s+)?(ChatGPT|Claude|Gemini)\s*$/i, '')
        .replace(/^(Google\s+)?(ChatGPT|Claude|Gemini)\s*[-|–]?\s*/i, '')
        .trim();
    }

    // Sanitize: remove filename-unsafe chars, collapse spaces, cap at 50 chars
    if (title) {
      title = title.replace(/[\\/:*?"<>|]/g, '').replace(/\s+/g, '_').slice(0, 50);
    }

    return title || null;
  }

  // ═══════════════════════════════════════════════════════════════
  //  BLOB TO BASE64
  // ═══════════════════════════════════════════════════════════════

  function blobToBase64(blob) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result.split(',')[1]);
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  }

  // ═══════════════════════════════════════════════════════════════
  //  CORE EXPORT PIPELINE  (markdown string → Drive / download)
  // ═══════════════════════════════════════════════════════════════

  function markdownToPlainText(md) {
    return md
      .replace(/\[\[IMG:\d+\]\]/g, '[Image]')
      .replace(/^#{1,6}\s+/gm, '')
      .replace(/\*\*\*([\s\S]+?)\*\*\*/g, '$1')
      .replace(/\*\*([\s\S]+?)\*\*/g, '$1')
      .replace(/\*([\s\S]+?)\*/g, '$1')
      .replace(/~([^~]+)~/g, '$1')
      .replace(/\^([^^]+)\^/g, '$1')
      .replace(/```[\w]*\n?([\s\S]*?)```/g, '$1')
      .replace(/`([^`]+)`/g, '$1')
      .replace(/^-{3,}$/gm, '────────────────────')
      .replace(/^\s*[-*+]\s/gm, '• ')
      .replace(/\n{3,}/g, '\n\n')
      .trim();
  }

  async function exportMarkdown(markdown, suffix, imageMap = {}) {
    suffix = suffix || '';
    // Replace image markers that couldn't be captured with text fallbacks
    markdown = markdown.replace(/\[\[IMG:(\d+)\]\]/g, (match, raw) => {
      const idx = parseInt(raw);
      if (imageMap[idx]) return match;
      const cap = _imgCaptures.find(c => c.idx === idx);
      return cap?.alt ? `[Image: ${cap.alt}]` : '[Image]';
    });

    // Prepend source metadata header (italic line at top of every export)
    const now = new Date();
    const dateStr = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}-${String(now.getDate()).padStart(2,'0')}`;
    const platformName = isGemini ? 'Gemini' : isClaude ? 'Claude' : 'ChatGPT';
    const convTitle = getConversationTitle();
    const metaParts = [platformName, dateStr, ...(convTitle ? [convTitle] : [])];
    const sourceUrl = location.origin + location.pathname;
    markdown = `*${metaParts.join(' · ')}*\n*${sourceUrl}*\n\n` + markdown;

    let blob;
    try {
      blob = window.convertChatGPTToDocx(markdown, imageMap);
    } catch(e) {
      showToast('❌ Error generating document: ' + e.message, true);
      return;
    }

    const timestamp = `${now.getFullYear()}${(now.getMonth()+1).toString().padStart(2,'0')}${now.getDate().toString().padStart(2,'0')}_${now.getHours().toString().padStart(2,'0')}${now.getMinutes().toString().padStart(2,'0')}`;
    const platform = platformName;
    const filename = convTitle
      ? `${convTitle}${suffix}.docx`
      : `${platform}_Export${suffix}_${timestamp}.docx`;

    if (!chrome.runtime || !chrome.runtime.sendMessage) {
      showToast('❌ Extension reloaded. Please refresh this page.', true);
      return;
    }
    const base64 = await blobToBase64(blob);

    if (exportDest === 'local') {
      showToast('⏳ Preparing download...');
      try {
        await new Promise((resolve, reject) => {
          chrome.runtime.sendMessage(
            { action: 'downloadLocal', docxBase64: base64, filename },
            (resp) => {
              if (chrome.runtime.lastError) reject(new Error(chrome.runtime.lastError.message));
              else if (resp?.success) resolve();
              else reject(new Error(resp?.error || 'Download failed'));
            }
          );
        });
        showToast('✅ Saved as .docx!', false, 4000);
      } catch(e) {
        showToast('❌ Save failed: ' + e.message, true);
      }
    } else {
      showToast('⏳ Uploading to Google Drive...');
      try {
        const result = await new Promise((resolve, reject) => {
          chrome.runtime.sendMessage(
            { action: 'uploadToDrive', docxBase64: base64, filename },
            (response) => {
              if (chrome.runtime.lastError) reject(new Error(chrome.runtime.lastError.message));
              else if (response && response.success) resolve(response);
              else reject(new Error(response ? response.error : 'Unknown error'));
            }
          );
        });

        showToast(`✅ Created "<b>${escHtml(result.fileName)}</b>" in Google Drive! Opening...`, false, 5000);
        setTimeout(() => window.open(result.url, '_blank'), 500);
        const convKey = location.hostname + location.pathname;
        chrome.storage.local.get('lastExports', (d) => {
          const allExports = d.lastExports || {};
          const history = Array.isArray(allExports[convKey]) ? allExports[convKey] : (allExports[convKey] ? [allExports[convKey]] : []);
          const newEntry = { fileName: result.fileName, url: result.url, fileId: result.fileId };
          const filtered = history.filter(e => e.fileId !== newEntry.fileId);
          allExports[convKey] = [newEntry, ...filtered].slice(0, 3);
          chrome.storage.local.set({ lastExports: allExports });
          chrome.storage.local.get('globalRecentDocs', (gd) => {
            const global = Array.isArray(gd.globalRecentDocs) ? gd.globalRecentDocs : [];
            const gFiltered = global.filter(e => e.fileId !== newEntry.fileId);
            chrome.storage.local.set({ globalRecentDocs: [newEntry, ...gFiltered].slice(0, 5) });
          });
        });

      } catch(e) {
        const msg = e.message || '';
        if (msg.includes('not signed in') || msg.includes('Not signed in')) {
          showToast('⚠️ Sign in to Chrome with your Google account to use Drive export. Downloading .docx instead.', true, 7000);
        } else if (msg.includes('denied') || msg.includes('not granted')) {
          showToast('⚠️ Drive access denied. Please allow access when prompted. Downloading .docx instead.', true, 7000);
        } else if (msg.includes('invalid_client') || msg.includes('client_id')) {
          showToast('⚠️ Google Drive not set up yet. Downloading .docx instead.<br><small>See SETUP_GUIDE.md to enable one-click export.</small>', true, 6000);
        } else if (msg.includes('sign-in') || msg.includes('OAuth2')) {
          showToast('⚠️ Could not sign in to Google. Downloading .docx instead.', true, 6000);
        } else {
          showToast('⚠️ Drive upload failed. Downloading .docx instead.', true);
        }

        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url; a.download = filename;
        document.body.appendChild(a); a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      }
    }
  }

  async function exportMessage(messageEl) {
    showToast('⏳ Generating document...');
    _resetImgCaptures();
    const markdown = extractMarkdown(messageEl);
    if (!markdown || !markdown.trim()) {
      showToast('❌ Could not extract content from this message', true);
      return;
    }
    const imageMap = await _captureImages();
    await exportMarkdown(markdown, '', imageMap);
  }

  // ═══════════════════════════════════════════════════════════════
  //  POPUP: GET LAST AI MESSAGE
  // ═══════════════════════════════════════════════════════════════

  function getLastAIMessage() {
    if (isChatGPT) {
      const msgs = document.querySelectorAll('[data-message-author-role="assistant"]');
      const last = msgs[msgs.length - 1];
      return last ? (last.querySelector('.markdown') || last) : null;
    }
    if (isGemini) {
      const msgs = getAllAIMessages();
      return msgs[msgs.length - 1] || null;
    }
    if (isClaude) {
      const responses = Array.from(document.querySelectorAll('[class*="font-claude-response"]'))
        .filter(el => !el.parentElement?.closest('[class*="font-claude-response"]'));
      return responses[responses.length - 1] || null;
    }
    return null;
  }

  // ═══════════════════════════════════════════════════════════════
  //  POPUP: EXPORT FULL CONVERSATION
  // ═══════════════════════════════════════════════════════════════

  async function exportFullConversation() {
    showToast('⏳ Collecting conversation...');
    _resetImgCaptures();
    const turns = [];

    if (isChatGPT) {
      const allMsgs = document.querySelectorAll('[data-message-author-role]');
      for (const msg of allMsgs) {
        const role = msg.getAttribute('data-message-author-role');
        const contentEl = role === 'assistant' ? (msg.querySelector('.markdown') || msg) : msg;
        const text = extractMarkdown(contentEl).trim();
        if (text) turns.push({ role: role === 'user' ? 'You' : 'ChatGPT', text });
      }
    } else if (isGemini) {
      // User queries: query the top-level custom element directly to avoid
      // duplicate child matches. Fall back to message-content if not found.
      let userEls = Array.from(document.querySelectorAll('user-query'));
      if (userEls.length === 0) {
        userEls = Array.from(document.querySelectorAll(
          'message-content[data-content-type="user"]'
        ));
      }
      const aiEls = getAllAIMessages();
      const all = [
        ...userEls.map(el => ({ el, role: 'You' })),
        ...aiEls.map(el => ({ el, role: 'Gemini' }))
      ].sort((a, b) => a.el.compareDocumentPosition(b.el) & Node.DOCUMENT_POSITION_FOLLOWING ? -1 : 1);
      for (const { el, role } of all) {
        const text = extractMarkdown(el).trim();
        if (text) turns.push({ role, text });
      }
    } else if (isClaude) {
      const userEls = Array.from(document.querySelectorAll('[class*="font-user-message"]'))
        .filter(el => !el.parentElement?.closest('[class*="font-user-message"]'));
      const aiEls = Array.from(document.querySelectorAll('[class*="font-claude-response"]'))
        .filter(el => !el.parentElement?.closest('[class*="font-claude-response"]'));
      const all = [
        ...userEls.map(el => ({ el, role: 'You' })),
        ...aiEls.map(el => ({ el, role: 'Claude' }))
      ].sort((a, b) => a.el.compareDocumentPosition(b.el) & Node.DOCUMENT_POSITION_FOLLOWING ? -1 : 1);
      for (const { el, role } of all) {
        const text = extractMarkdown(el).trim();
        if (text) turns.push({ role, text });
      }
    }

    if (turns.length === 0) {
      showToast('❌ No conversation content found', true);
      return;
    }

    const markdown = turns.map(t => `## ${t.role}\n\n${t.text}`).join('\n\n---\n\n');
    const imageMap = await _captureImages();
    await exportMarkdown(markdown, '_full', imageMap);
  }

  // ═══════════════════════════════════════════════════════════════
  //  DEFAULT-MODE CLICK HANDLER
  // ═══════════════════════════════════════════════════════════════

  function handleExportClick(e, messageEl) {
    e.preventDefault();
    e.stopPropagation();
    chrome.storage.local.get(['defaultExportMode', 'exportDest'], (data) => {
      exportDest = data.exportDest || 'drive';
      const mode = data.defaultExportMode || 'select';
      if (mode === 'last') {
        const el = getLastAIMessage();
        if (el) exportMessage(el);
        else showToast('❌ No AI response found', true);
      } else if (mode === 'full') {
        exportFullConversation();
      } else {
        showSelectPanel(messageEl);
      }
    });
  }

  // ═══════════════════════════════════════════════════════════════
  //  CSV TABLE EXPORT
  // ═══════════════════════════════════════════════════════════════

  function tableToCSV(tableEl) {
    return Array.from(tableEl.querySelectorAll('tr')).map(row =>
      Array.from(row.querySelectorAll('th, td'))
        .map(c => '"' + c.textContent.replace(/"/g, '""').replace(/\s+/g, ' ').trim() + '"')
        .join(',')
    ).join('\n');
  }

  function downloadCSV(tableEl, msgIndex, tableIndex) {
    const csv = tableToCSV(tableEl);
    const base = getConversationTitle() || 'table';
    const suffix = tableIndex > 0 ? `_table${tableIndex + 1}` : '_table';
    const filename = `${base}${suffix}.csv`;
    const blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  // ═══════════════════════════════════════════════════════════════
  //  SELECTION PANEL
  // ═══════════════════════════════════════════════════════════════

  function getAllAIMessages() {
    if (isChatGPT) {
      return Array.from(document.querySelectorAll('[data-message-author-role="assistant"]'))
        .map(el => el.querySelector('.markdown') || el);
    }
    if (isGemini) {
      // Use only model-response (top-level) to avoid duplicates with child selectors
      const els = Array.from(document.querySelectorAll('model-response'));
      if (els.length > 0) return els.filter(el => !el.parentElement?.closest('model-response'));
      // Fallback for alternate Gemini DOM
      return Array.from(document.querySelectorAll('message-content[data-content-type="model"]'))
        .filter(el => !el.parentElement?.closest('message-content[data-content-type="model"]'));
    }
    if (isClaude) {
      return Array.from(document.querySelectorAll('[class*="font-claude-response"]'))
        .filter(el => !el.parentElement?.closest('[class*="font-claude-response"]'));
    }
    return [];
  }

  function getCleanPreview(msgEl) {
    const clone = msgEl.cloneNode(true);
    // Remove our injected buttons and any native UI buttons
    clone.querySelectorAll('.' + BUTTON_CLASS + ', button, [role="button"], svg, style, script').forEach(el => el.remove());
    // Remove external image attribution links (e.g. "Opens in a new window · stockcake.com")
    clone.querySelectorAll('a[target="_blank"]').forEach(el => el.remove());
    const text = (clone.textContent || '').replace(/\s+/g, ' ').trim();
    const firstMeaningful = text.split(/\.|\n/).find(s => s.trim().length > 12) || text;
    return firstMeaningful.trim().slice(0, 85);
  }

  // Open the Drive folder picker (via background worker) and wait for the user
  // to select a folder or cancel. Returns 'done', 'cancelled', or 'timeout'.
  function _pickDriveFolder() {
    return new Promise((resolve) => {
      chrome.storage.local.remove('pickerState', () => {
        // Auth must succeed in background first; background's getAuthToken(true)
        // reliably shows the OAuth UI, unlike calling it from a programmatic window.
        chrome.runtime.sendMessage({ action: 'ensureAuth' }, (resp) => {
          if (!resp?.ok) { resolve('cancelled'); return; }
          chrome.runtime.sendMessage({ action: 'openPickerWindow' }, () => {});
          const handler = (changes) => {
            if (changes.pickerState) {
              chrome.storage.onChanged.removeListener(handler);
              clearTimeout(tid);
              resolve(changes.pickerState.newValue);
            }
          };
          chrome.storage.onChanged.addListener(handler);
          const tid = setTimeout(() => {
            chrome.storage.onChanged.removeListener(handler);
            resolve('timeout');
          }, 180000);
        });
      });
    });
  }

  function _refreshDriveBtn(btn) {
    chrome.storage.local.get('customFolderName', (d) => {
      const n = d.customFolderName || 'AI Chat Exports';
      const s = n.length > 18 ? n.slice(0, 16) + '…' : n;
      btn.textContent = `☁️ ${s}`;
    });
  }

  function showSelectPanel(thisMessageEl) {
    const existing = document.querySelector('.cgd-panel');
    if (existing) { existing.remove(); return; }

    const messages = getAllAIMessages();
    if (messages.length === 0) { showToast('❌ No AI responses found', true); return; }

    chrome.storage.local.get(['customFolderName', 'lastExports', 'globalRecentDocs'], (storageData) => {
      _buildSelectPanel(messages, thisMessageEl, storageData);
    });
  }

  function _buildSelectPanel(messages, thisMessageEl, storageData = {}) {
    const platform = isGemini ? 'Gemini' : isClaude ? 'Claude' : 'ChatGPT';
    const dark = isDarkMode();

    const thisIdx = thisMessageEl
      ? messages.findIndex(m => m === thisMessageEl || m.contains(thisMessageEl) || thisMessageEl.contains(m))
      : -1;

    const panel = document.createElement('div');
    panel.className = 'cgd-panel' + (dark ? ' cgd-dark' : '');

    // ── Header ──
    const header = document.createElement('div');
    header.className = 'cgd-panel-header';
    header.innerHTML = `<span class="cgd-panel-title">Export · ${platform}</span><button class="cgd-panel-close" title="Close">✕</button>`;

    header.addEventListener('mousedown', (e) => {
      if (e.target.closest('.cgd-panel-close')) return;
      const startX = e.clientX, startY = e.clientY;
      const rect = panel.getBoundingClientRect();
      const origLeft = rect.left, origTop = rect.top;
      document.removeEventListener('click', outsideClickHandler);
      function onMove(me) {
        panel.style.left = (origLeft + me.clientX - startX) + 'px';
        panel.style.top  = (origTop  + me.clientY - startY) + 'px';
        panel.style.right = 'auto';
        panel.style.transform = 'none';
      }
      function onUp() {
        document.removeEventListener('mousemove', onMove);
        document.removeEventListener('mouseup', onUp);
        setTimeout(() => document.addEventListener('click', outsideClickHandler), 0);
      }
      document.addEventListener('mousemove', onMove);
      document.addEventListener('mouseup', onUp);
      e.preventDefault();
    });

    // ── Dest row: Drive / Local ──
    const folderName = storageData.customFolderName || 'AI Chat Exports';
    const folderShort = folderName.length > 18 ? folderName.slice(0, 16) + '…' : folderName;

    const destRow = document.createElement('div');
    destRow.className = 'cgd-dest-row';

    const btnDrive = document.createElement('button');
    btnDrive.className = 'cgd-dest-btn';
    btnDrive.textContent = `☁️ ${folderShort}`;

    const btnLocal = document.createElement('button');
    btnLocal.className = 'cgd-dest-btn';
    btnLocal.textContent = '💾 Local .docx';

    destRow.appendChild(btnDrive);
    destRow.appendChild(btnLocal);

    // ── Recent exports row (Drive only, up to 2 chips) ──
    const recentRow = document.createElement('div');
    recentRow.className = 'cgd-recent-row';

    const convKey = location.hostname + location.pathname;
    const rawExp = (storageData.lastExports || {})[convKey] || null;
    const convHistory = Array.isArray(rawExp) ? rawExp : (rawExp ? [rawExp] : []);
    const convIds = new Set(convHistory.map(e => e.fileId));
    const globalExtra = (Array.isArray(storageData.globalRecentDocs) ? storageData.globalRecentDocs : [])
      .filter(e => !convIds.has(e.fileId));
    const recents = [...convHistory, ...globalExtra].slice(0, 2);

    function buildRecentChips() {
      recentRow.innerHTML = '';
      if (exportDest !== 'drive' || recents.length === 0) {
        recentRow.style.display = 'none';
        return;
      }
      recentRow.style.display = 'flex';
      recents.forEach(exp => {
        const chip = document.createElement('a');
        chip.className = 'cgd-recent-chip';
        chip.href = exp.url || `https://docs.google.com/document/d/${exp.fileId}/edit`;
        chip.target = '_blank';
        const name = exp.fileName || 'Untitled';
        const short = name.length > 28 ? name.slice(0, 26) + '…' : name;
        chip.innerHTML = `<span class="cgd-rc-icon">↩</span><span class="cgd-rc-name">${short}</span><span class="cgd-rc-arrow">↗</span>`;
        chip.title = name;
        recentRow.appendChild(chip);
      });
    }
    buildRecentChips();

    function applyDestUI() {
      btnDrive.classList.toggle('cgd-dest-active', exportDest === 'drive');
      btnLocal.classList.toggle('cgd-dest-active', exportDest === 'local');
      buildRecentChips();
    }
    applyDestUI();

    btnDrive.addEventListener('click', async () => {
      exportDest = 'drive';
      chrome.storage.local.set({ exportDest: 'drive' });
      applyDestUI();
      exportBtn.textContent = 'Export to Docs →';
      // Open folder picker to change destination folder
      const result = await _pickDriveFolder();
      if (result === 'done') _refreshDriveBtn(btnDrive);
    });

    btnLocal.addEventListener('click', () => {
      exportDest = 'local';
      chrome.storage.local.set({ exportDest: 'local' });
      applyDestUI();
      exportBtn.textContent = 'Save .docx →';
    });

    // ── Action row: Last / Full / Pick ──
    const actionRow = document.createElement('div');
    actionRow.className = 'cgd-action-row';

    const btnLast = document.createElement('button');
    btnLast.className = 'cgd-action-btn';
    btnLast.textContent = '↩ Last';
    btnLast.title = 'Export last AI response';

    const btnFull = document.createElement('button');
    btnFull.className = 'cgd-action-btn';
    btnFull.textContent = '≡ Full';
    btnFull.title = 'Export full conversation';

    const btnPick = document.createElement('button');
    btnPick.className = 'cgd-action-btn';
    btnPick.textContent = '☑ Pick';
    btnPick.title = 'Select specific responses';

    actionRow.appendChild(btnLast);
    actionRow.appendChild(btnFull);
    actionRow.appendChild(btnPick);

    // ── Pick area (collapsed by default) ──
    const pickArea = document.createElement('div');
    pickArea.className = 'cgd-pick-area';
    pickArea.style.display = 'none';

    // ── Select all / none ──
    const controls = document.createElement('div');
    controls.className = 'cgd-panel-controls';
    controls.innerHTML = `<button class="cgd-ctrl-btn" id="cgd-sa">Select all</button><button class="cgd-ctrl-btn" id="cgd-sn">Deselect all</button>`;

    // ── Message list ──
    const list = document.createElement('div');
    list.className = 'cgd-panel-list';
    const checkboxes = [];
    const rowEls = [];

    messages.forEach((msgEl, i) => {
      const preview = getCleanPreview(msgEl);

      const row = document.createElement('label');
      row.className = 'cgd-msg-row';

      if (i === thisIdx) {
        row.style.background = dark ? 'rgba(138,180,248,0.12)' : '#e8f0fe';
        row.style.borderRadius = '8px';
      }

      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.checked = true;
      checkboxes.push(cb);
      cb.addEventListener('change', updateCount);

      const textWrap = document.createElement('div');
      textWrap.style.flex = '1';
      textWrap.style.cursor = 'pointer';
      textWrap.title = 'Click to jump to this response';

      const numDiv = document.createElement('div');
      numDiv.className = 'cgd-msg-num';

      const numText = document.createElement('span');
      numText.textContent = `Response ${i + 1}`;

      // Small jump button — separate from label so it doesn't block checkbox toggle
      const jumpBtn = document.createElement('button');
      jumpBtn.textContent = '↗';
      jumpBtn.title = 'Jump to this response';
      jumpBtn.style.cssText = 'background:none;border:none;cursor:pointer;font-size:10px;color:inherit;padding:0 2px;opacity:0.6;';
      jumpBtn.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        msgEl.scrollIntoView({ behavior: 'smooth', block: 'center' });
        msgEl.style.outline = '2px solid #1a73e8';
        msgEl.style.borderRadius = '6px';
        setTimeout(() => { msgEl.style.outline = ''; msgEl.style.borderRadius = ''; }, 1800);
      });
      numDiv.appendChild(numText);
      numDiv.appendChild(jumpBtn);

      const prevDiv = document.createElement('div');
      prevDiv.className = 'cgd-msg-preview';
      prevDiv.textContent = preview;
      textWrap.appendChild(numDiv);
      textWrap.appendChild(prevDiv);

      row.appendChild(cb);
      row.appendChild(textWrap);

      const tables = Array.from(msgEl.querySelectorAll('table'))
        .filter(t => !t.closest('pre') && !t.closest('code'));
      if (tables.length > 0) {
        const csvBtn = document.createElement('button');
        csvBtn.className = 'cgd-csv-btn';
        csvBtn.textContent = tables.length > 1 ? `📊 ${tables.length} CSV` : '📊 CSV';
        csvBtn.title = tables.length > 1 ? `Download ${tables.length} tables as CSV` : 'Download table as CSV';
        csvBtn.addEventListener('click', (e) => {
          e.preventDefault(); e.stopPropagation();
          tables.forEach((tbl, tIdx) => downloadCSV(tbl, i, tIdx));
        });
        row.appendChild(csvBtn);
      }

      list.appendChild(row);
      rowEls.push(row);
    });

    // ── Footer ──
    const footer = document.createElement('div');
    footer.className = 'cgd-panel-footer';
    const footerMain = document.createElement('div');
    footerMain.className = 'cgd-footer-main';

    const countLabel = document.createElement('span');
    countLabel.className = 'cgd-count-label';
    countLabel.textContent = messages.length + ' selected';

    const exportBtn = document.createElement('button');
    exportBtn.className = 'cgd-export-sel-btn';
    exportBtn.textContent = exportDest === 'local' ? 'Save .docx →' : 'Export to Docs →';

    footerMain.appendChild(countLabel);
    footerMain.appendChild(exportBtn);
    footer.appendChild(footerMain);

    pickArea.appendChild(controls);
    pickArea.appendChild(list);
    pickArea.appendChild(footer);

    // ── Assemble ──
    panel.appendChild(header);
    panel.appendChild(destRow);
    panel.appendChild(recentRow);
    panel.appendChild(actionRow);
    panel.appendChild(pickArea);

    // ── Shared logic ──
    function close() {
      darkWatcher.disconnect();
      document.removeEventListener('click', outsideClickHandler);
      panel.remove();
    }

    function outsideClickHandler(e) {
      if (!panel.contains(e.target)) close();
    }

    const darkWatcher = new MutationObserver(() => {
      panel.classList.toggle('cgd-dark', isDarkMode());
    });
    darkWatcher.observe(document.documentElement, { attributes: true, attributeFilter: ['class', 'data-theme', 'data-color-scheme', 'data-color-mode', 'style'] });
    darkWatcher.observe(document.body, { attributes: true, attributeFilter: ['class', 'data-theme', 'style'] });

    function updateCount() {
      const n = checkboxes.filter(cb => cb.checked).length;
      countLabel.textContent = n + ' selected';
      exportBtn.disabled = n === 0;
    }

    function getSelectedIndices() {
      return checkboxes.reduce((acc, cb, i) => { if (cb.checked) acc.push(i); return acc; }, []);
    }

    async function exportSelected() {
      const selectedIndices = getSelectedIndices();
      if (selectedIndices.length === 0) return;

      // Panel stays open — user can pick again for another export
      exportBtn.disabled = true;
      exportBtn.textContent = 'Exporting…';

      if (selectedIndices.length === 1) {
        await exportMessage(messages[selectedIndices[0]]);
      } else {
        _resetImgCaptures();
        showToast('⏳ Generating document...');
        const parts = selectedIndices.map(origIdx => {
          const text = extractMarkdown(messages[origIdx]).trim();
          return text ? `## ${platform} (Response ${origIdx + 1})\n\n${text}` : '';
        }).filter(Boolean);
        const imageMap = await _captureImages();
        if (parts.length) await exportMarkdown(parts.join('\n\n---\n\n'), '_selected', imageMap);
        else showToast('❌ Could not extract content', true);
      }

      exportBtn.disabled = false;
      exportBtn.textContent = exportDest === 'local' ? 'Save .docx →' : 'Export to Docs →';
    }

    header.querySelector('.cgd-panel-close').addEventListener('click', close);

    btnLast.addEventListener('click', () => {
      close();
      const el = getLastAIMessage();
      if (el) exportMessage(el); else showToast('❌ No AI response found', true);
    });

    btnFull.addEventListener('click', () => {
      close();
      exportFullConversation();
    });

    btnPick.addEventListener('click', () => {
      const open = pickArea.style.display !== 'none';
      pickArea.style.display = open ? 'none' : 'flex';
      btnPick.classList.toggle('cgd-action-btn-active', !open);
      if (!open) {
        setTimeout(() => {
          if (thisIdx !== -1 && rowEls[thisIdx]) rowEls[thisIdx].scrollIntoView({ block: 'nearest' });
          else list.scrollTop = list.scrollHeight;
        }, 30);
      }
    });

    controls.querySelector('#cgd-sa').addEventListener('click', () => { checkboxes.forEach(cb => cb.checked = true); updateCount(); });
    controls.querySelector('#cgd-sn').addEventListener('click', () => { checkboxes.forEach(cb => cb.checked = false); updateCount(); });

    exportBtn.addEventListener('click', exportSelected);

    document.body.appendChild(panel);
    setTimeout(() => document.addEventListener('click', outsideClickHandler), 0);
  }

  // ═══════════════════════════════════════════════════════════════
  //  CHATGPT: INJECT BUTTONS
  // ═══════════════════════════════════════════════════════════════

  function addChatGPTButtons() {
    const messages = document.querySelectorAll('[data-message-author-role="assistant"]');
    for (const msg of messages) {
      const container = msg.closest('.group\\/conversation-turn') || msg.closest('[data-testid^="conversation-turn"]');
      if (!container || container.querySelector('.' + BUTTON_CLASS)) continue;
      const actionArea = findChatGPTActionBar(container);
      if (!actionArea) continue;

      const btn = createExportButton();
      btn.addEventListener('click', (e) => handleExportClick(e, msg.querySelector('.markdown') || msg));
      insertBeforeMoreButton(actionArea, btn);
    }
  }

  function findChatGPTActionBar(container) {
    // Text responses: copy button has a specific data-testid
    const copyBtn = container.querySelector('button[data-testid="copy-turn-action-button"]');
    if (copyBtn) {
      let bar = copyBtn.parentElement;
      for (let i = 0; i < 3 && bar; i++) {
        if (bar.querySelector('button[data-testid*="more"]')) return bar;
        bar = bar.parentElement;
      }
      return copyBtn.parentElement;
    }
    // Image-only responses: action bar has thumbs-up/down buttons (no copy-turn-action-button)
    const thumbBtn = container.querySelector(
      'button[data-testid="thumbs-up-button"], button[data-testid="thumbs-down-button"], ' +
      'button[aria-label="Good response"], button[aria-label="Bad response"]'
    );
    if (thumbBtn) {
      let bar = thumbBtn.parentElement;
      for (let i = 0; i < 3 && bar; i++) {
        if (bar.querySelectorAll('button').length >= 2 && bar.offsetHeight < 60) return bar;
        bar = bar.parentElement;
      }
      return thumbBtn.parentElement;
    }
    const allDivs = container.querySelectorAll('div.flex');
    for (const div of allDivs) {
      if (div.querySelectorAll('button').length >= 2 && div.offsetHeight < 50) return div;
    }
    return null;
  }

  // ═══════════════════════════════════════════════════════════════
  //  GEMINI: INJECT BUTTONS
  // ═══════════════════════════════════════════════════════════════

  function addGeminiButtons() {
    // Gemini response containers
    const responses = document.querySelectorAll(
      'model-response, .model-response-text, .response-container, message-content[data-content-type="model"]'
    );

    for (const resp of responses) {
      // Walk up to find the message turn container
      const turnContainer = resp.closest('.conversation-turn') ||
                            resp.closest('message-content') ||
                            resp.closest('.response-turn') ||
                            resp;

      if (turnContainer.querySelector('.' + BUTTON_CLASS)) continue;

      // Find action bar (Gemini has copy, thumbs up/down buttons)
      const actionArea = findGeminiActionBar(turnContainer) || findGeminiActionBar(resp);

      if (actionArea) {
        const btn = createExportButton();
        btn.addEventListener('click', (e) => handleExportClick(e, resp));
        insertBeforeMoreButton(actionArea, btn);
      }
    }

    // Also try to find responses by looking for the "copy" button in Gemini
    const copyButtons = document.querySelectorAll('button[aria-label="Copy"], button[data-tooltip="Copy"]');
    for (const copyBtn of copyButtons) {
      // Walk up to find the full action bar including three-dots button
      let actionBar = copyBtn.parentElement;
      for (let i = 0; i < 3 && actionBar; i++) {
        const hasMore = actionBar.querySelector('button[aria-label*="more" i], button[data-tooltip*="more" i]');
        if (hasMore) break;
        actionBar = actionBar.parentElement;
      }
      if (!actionBar || actionBar.querySelector('.' + BUTTON_CLASS)) continue;

      // Skip image overlay buttons — those only have copy+download, no thumbs/share.
      // Real response action bars always contain thumbs-up or thumbs-down buttons.
      const isResponseBar = actionBar.querySelector(
        'button[aria-label*="thumb" i], button[aria-label*="like" i], ' +
        'button[aria-label*="dislike" i], button[aria-label*="good" i], ' +
        'button[aria-label*="bad" i], button[data-tooltip*="thumb" i]'
      );
      if (!isResponseBar) continue;

      // Find the associated response content
      const turnContainer = copyBtn.closest('.conversation-turn') ||
                            copyBtn.closest('message-content') ||
                            copyBtn.closest('.response-container') ||
                            copyBtn.closest('div[class*="response"]');

      if (!turnContainer) continue;

      const contentEl = turnContainer.querySelector('.markdown-main-panel') ||
                        turnContainer.querySelector('.model-response-text') ||
                        turnContainer.querySelector('.response-content') ||
                        turnContainer;

      const btn = createExportButton();
      btn.addEventListener('click', (e) => handleExportClick(e, contentEl));
      insertBeforeMoreButton(actionBar, btn);
    }
  }

  function findGeminiActionBar(container) {
    const buttons = container.querySelectorAll('button');
    for (const btn of buttons) {
      const label = (btn.getAttribute('aria-label') || btn.getAttribute('data-tooltip') || '').toLowerCase();
      if (label.includes('copy') || label.includes('share') || label.includes('thumb')) {
        // Walk up to find the bar that also contains the three-dots / more-options button
        let bar = btn.parentElement;
        for (let i = 0; i < 3 && bar; i++) {
          const hasMore = bar.querySelector('button[aria-label*="more" i], button[aria-label*="option" i], button[data-tooltip*="more" i]');
          if (hasMore) return bar;
          bar = bar.parentElement;
        }
        return btn.parentElement;
      }
    }
    return container.querySelector('.action-buttons') ||
           container.querySelector('.response-actions') ||
           container.querySelector('[class*="action"]');
  }

  // ═══════════════════════════════════════════════════════════════
  //  SHARED: CREATE BUTTON
  // ═══════════════════════════════════════════════════════════════

  function insertBeforeMoreButton(container, el) {
    const moreBtn = container.querySelector(
      'button[aria-label*="more" i], button[aria-label*="option" i], ' +
      'button[data-tooltip*="more" i], button[data-testid*="more"]'
    );
    if (!moreBtn) { container.appendChild(el); return; }
    let anchor = moreBtn;
    while (anchor.parentElement !== container) anchor = anchor.parentElement;
    container.insertBefore(el, anchor);
  }

  function _updateExportBtnContent(btn) {
    chrome.storage.local.get('customFolderName', (d) => {
      const folderName = d.customFolderName || 'AI Chat Exports';
      const pathSpan  = btn.querySelector('.cgd-btn-path');
      const labelSpan = btn.querySelector('.cgd-btn-label');
      if (!pathSpan || !labelSpan) return;
      if (exportDest === 'local') {
        pathSpan.textContent = '💾 Local';
        labelSpan.textContent = 'Save .docx';
      } else {
        const short = folderName.length > 18 ? folderName.slice(0, 16) + '…' : folderName;
        pathSpan.textContent = `☁️ ${short}`;
        labelSpan.textContent = 'Export to Docs';
      }
    });
  }

  function createExportButton() {
    const btn = document.createElement('button');
    btn.className = BUTTON_CLASS;

    const svgFile = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>`;
    const svgChevron = `<svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>`;

    const pathSpan = document.createElement('span');
    pathSpan.className = 'cgd-btn-path';

    const labelSpan = document.createElement('span');
    labelSpan.className = 'cgd-btn-label';

    btn.insertAdjacentHTML('beforeend', svgFile);
    btn.appendChild(pathSpan);
    btn.appendChild(labelSpan);
    btn.insertAdjacentHTML('beforeend', svgChevron);

    btn.title = 'Export options';
    _updateExportBtnContent(btn);
    return btn;
  }

  // ═══════════════════════════════════════════════════════════════
  //  CLAUDE: INJECT BUTTONS
  // ═══════════════════════════════════════════════════════════════

  function addClaudeButtons() {
    const copyButtons = document.querySelectorAll('button[aria-label="Copy"]');

    for (const copyBtn of copyButtons) {
      // Exclude code-block copy buttons — they live inside <pre> or a code toolbar
      if (copyBtn.closest('pre') || copyBtn.closest('[data-code-block]') || copyBtn.closest('.code-block')) continue;

      // Find the message-level action bar (try multiple class names Claude has used)
      const actionBar = copyBtn.closest('.text-text-300') ||
                        copyBtn.closest('[class*="message-actions"]') ||
                        copyBtn.closest('[class*="action-bar"]') ||
                        copyBtn.parentElement?.parentElement;

      if (!actionBar) continue;
      if (actionBar.querySelector('.' + BUTTON_CLASS)) continue;

      // Walk up from action bar to find the first container that has a
      // font-claude-response descendant (the copy button is a sibling of the
      // response content, not a descendant, so we search from the container level)
      let responseContainer = actionBar.parentElement;
      for (let i = 0; i < 10 && responseContainer && responseContainer !== document.body; i++) {
        if (responseContainer.querySelector('[class*="font-claude-response"]')) break;
        responseContainer = responseContainer.parentElement;
      }
      if (!responseContainer || responseContainer === document.body) continue;

      // Skip user message containers
      if (responseContainer.querySelector('[class*="font-user-message"]')) continue;

      // Find the content element with multiple fallbacks
      const contentEl =
        responseContainer.querySelector('[class*="font-claude-response"]') ||
        responseContainer.querySelector('.prose') ||
        responseContainer.querySelector('[data-is-streaming]') ||
        responseContainer;

      const btn = createExportButton();
      btn.addEventListener('click', (e) => handleExportClick(e, contentEl));

      const wrapper = document.createElement('div');
      wrapper.className = 'w-fit';
      wrapper.appendChild(btn);
      actionBar.appendChild(wrapper);
    }
  }

  // ═══════════════════════════════════════════════════════════════
  //  INIT
  // ═══════════════════════════════════════════════════════════════

  function addButtons() {
    if (isChatGPT) addChatGPTButtons();
    if (isGemini) addGeminiButtons();
    if (isClaude) addClaudeButtons();
  }

  function init() { addButtons(); }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();

  const observer = new MutationObserver(() => {
    clearTimeout(observer._t);
    observer._t = setTimeout(addButtons, 200);
  });
  observer.observe(document.body, { childList: true, subtree: true });

  // ═══════════════════════════════════════════════════════════════
  //  POPUP MESSAGE LISTENER
  // ═══════════════════════════════════════════════════════════════

  chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.action === 'getPlatform') {
      const platform = isChatGPT ? 'ChatGPT' : isGemini ? 'Gemini' : isClaude ? 'Claude' : null;
      const responseCount = platform ? getAllAIMessages().length : 0;
      sendResponse({ platform, responseCount });
      return;
    }
    if (request.action === 'exportLast') {
      if (request.dest) exportDest = request.dest;
      sendResponse({ ok: true });
      const el = getLastAIMessage();
      if (!el) { showToast('❌ No AI response found on this page', true); return; }
      exportMessage(el);
      return;
    }
    if (request.action === 'exportFull') {
      if (request.dest) exportDest = request.dest;
      sendResponse({ ok: true });
      exportFullConversation();
      return;
    }
    if (request.action === 'openPanel') {
      if (request.dest) exportDest = request.dest;
      sendResponse({ ok: true });
      showSelectPanel(null);
      return;
    }
    if (request.action === 'triggerDefault') {
      sendResponse({ ok: true });
      chrome.storage.local.get(['defaultExportMode', 'exportDest'], (data) => {
        exportDest = data.exportDest || 'drive';
        const mode = data.defaultExportMode || 'select';
        if (mode === 'last') {
          const el = getLastAIMessage();
          if (el) exportMessage(el); else showToast('❌ No AI response found', true);
        } else if (mode === 'full') {
          exportFullConversation();
        } else {
          showSelectPanel(null);
        }
      });
      return;
    }
  });
})();
