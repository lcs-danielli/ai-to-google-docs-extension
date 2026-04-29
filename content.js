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
        if (e.message.includes('OAuth2') || e.message.includes('client_id') || e.message.includes('invalid_client') || e.message.includes('sign-in')) {
          showToast('⚠️ Google Drive not set up yet. Downloading .docx instead.<br><small>See SETUP_GUIDE.md to enable one-click export.</small>', true, 6000);
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

  function showSelectPanel(thisMessageEl) {
    // Toggle: close if already open
    const existing = document.querySelector('.cgd-panel');
    if (existing) { existing.remove(); return; }

    const messages = getAllAIMessages();
    if (messages.length === 0) { showToast('❌ No AI responses found', true); return; }

    const convKey = location.hostname + location.pathname;
    chrome.storage.local.get(['lastExports', 'globalRecentDocs', 'customFolderName'], (storageData) => {
      const raw = (storageData.lastExports || {})[convKey] || null;
      const convHistory = Array.isArray(raw) ? raw : (raw ? [raw] : []);
      const convIds = new Set(convHistory.map(e => e.fileId));
      const globalExtra = (Array.isArray(storageData.globalRecentDocs) ? storageData.globalRecentDocs : [])
        .filter(e => !convIds.has(e.fileId));
      const exportHistory = [...convHistory, ...globalExtra].slice(0, 3);
      const currentFolderName = storageData.customFolderName || 'AI Chat Exports';
      _buildSelectPanel(messages, exportHistory, thisMessageEl, currentFolderName);
    });
  }

  function _buildSelectPanel(messages, exportHistory, thisMessageEl, currentFolderName = 'AI Chat Exports') {
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
    header.innerHTML = `<span class="cgd-panel-title">Export Responses</span><button class="cgd-panel-close" title="Close">✕</button>`;

    // Make panel draggable via header
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

    // ── Destination selector ──
    // "New Doc" always first; up to 3 previous exports shown as selectable pills
    let destMode = 'new';
    let selectedExport = null; // which export entry is active when destMode === 'append'

    // ── Folder row ──
    const folderRow = document.createElement('div');
    folderRow.className = 'cgd-folder-row';
    const folderLabel = document.createElement('span');
    folderLabel.className = 'cgd-folder-label';
    folderLabel.textContent = 'Save to';
    const folderPickBtn = document.createElement('button');
    folderPickBtn.className = 'cgd-folder-pick-btn';
    const folderShort = currentFolderName.length > 22 ? currentFolderName.slice(0, 20) + '…' : currentFolderName;
    folderPickBtn.textContent = folderShort + ' ▾';
    folderRow.appendChild(folderLabel);
    folderRow.appendChild(folderPickBtn);

    // Inline folder dropdown (toggled on click)
    const folderDropEl = document.createElement('div');
    folderDropEl.className = 'cgd-folder-drop' + (dark ? ' cgd-dark' : '');
    folderDropEl.style.display = 'none';

    folderPickBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      if (folderDropEl.style.display !== 'none') { folderDropEl.style.display = 'none'; return; }
      folderDropEl.innerHTML = `<div class="cgd-folder-drop-item cgd-folder-loading">Loading…</div>`;
      folderDropEl.style.display = 'block';
      chrome.runtime.sendMessage({ action: 'listFolders' }, (resp) => {
        folderDropEl.innerHTML = '';
        chrome.storage.local.get('customFolderId', (d) => {
          const activeId = d.customFolderId || null;
          const defItem = document.createElement('div');
          defItem.className = 'cgd-folder-drop-item' + (!activeId ? ' cgd-folder-drop-active' : '');
          defItem.textContent = (!activeId ? '✓ ' : '') + 'AI Chat Exports (default)';
          defItem.addEventListener('click', (e) => {
            e.stopPropagation();
            chrome.storage.local.remove(['customFolderId', 'customFolderName']);
            folderPickBtn.textContent = 'AI Chat Exports ▾';
            folderDropEl.style.display = 'none';
          });
          folderDropEl.appendChild(defItem);

          if (!resp || !resp.success) {
            const errItem = document.createElement('div');
            errItem.className = 'cgd-folder-drop-item cgd-folder-loading';
            errItem.textContent = resp?.error || 'Could not load folders';
            folderDropEl.appendChild(errItem);
            return;
          }
          for (const f of (resp.folders || [])) {
            const item = document.createElement('div');
            const isActive = f.id === activeId;
            item.className = 'cgd-folder-drop-item' + (isActive ? ' cgd-folder-drop-active' : '');
            item.textContent = (isActive ? '✓ ' : '') + f.name;
            item.title = f.name;
            item.addEventListener('click', (e) => {
              e.stopPropagation();
              chrome.storage.local.set({ customFolderId: f.id, customFolderName: f.name });
              const n = f.name.length > 22 ? f.name.slice(0, 20) + '…' : f.name;
              folderPickBtn.textContent = n + ' ▾';
              folderDropEl.style.display = 'none';
            });
            folderDropEl.appendChild(item);
          }
          if (!resp.folders?.length) {
            const empty = document.createElement('div');
            empty.className = 'cgd-folder-drop-item cgd-folder-loading';
            empty.textContent = 'No folders in Drive root';
            folderDropEl.appendChild(empty);
          }
        });
      });
    });

    const destRow = document.createElement('div');
    destRow.className = 'cgd-dest-row';

    const destNewBtn = document.createElement('button');
    destNewBtn.className = 'cgd-dest-btn cgd-dest-active';
    destNewBtn.textContent = '+ New Doc';
    destRow.appendChild(destNewBtn);

    const appendBtns = [];
    for (const exp of exportHistory) {
      if (!exp || !exp.fileId) continue;
      const shortName = exp.fileName.length > 16 ? exp.fileName.slice(0, 14) + '…' : exp.fileName;
      const btn = document.createElement('button');
      btn.className = 'cgd-dest-btn';
      btn.textContent = `→ ${shortName}`;
      btn.title = `Add to: ${exp.fileName} · Double-click to open`;
      const docUrl = exp.url || `https://docs.google.com/document/d/${exp.fileId}/edit`;

      btn.addEventListener('click', () => {
        destMode = 'append';
        selectedExport = exp;
        destNewBtn.classList.remove('cgd-dest-active');
        appendBtns.forEach(b => b.classList.remove('cgd-dest-active'));
        btn.classList.add('cgd-dest-active');
        exportBtn.textContent = 'Append →';
      });
      btn.addEventListener('dblclick', (e) => { e.stopPropagation(); window.open(docUrl, '_blank'); });
      appendBtns.push(btn);
      destRow.appendChild(btn);
    }

    destNewBtn.addEventListener('click', () => {
      destMode = 'new';
      selectedExport = null;
      destNewBtn.classList.add('cgd-dest-active');
      appendBtns.forEach(b => b.classList.remove('cgd-dest-active'));
      exportBtn.textContent = 'Export →';
    });

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
      cb.checked = thisIdx === -1 || i === thisIdx;
      checkboxes.push(cb);
      cb.addEventListener('change', updateCount);

      const textWrap = document.createElement('div');
      textWrap.style.flex = '1';
      textWrap.style.cursor = 'pointer';
      textWrap.title = 'Click to jump to this response';

      const numDiv = document.createElement('div');
      numDiv.className = 'cgd-msg-num';
      numDiv.textContent = `Response ${i + 1}  ↗`;

      const prevDiv = document.createElement('div');
      prevDiv.className = 'cgd-msg-preview';
      prevDiv.textContent = preview;
      textWrap.appendChild(numDiv);
      textWrap.appendChild(prevDiv);

      textWrap.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        msgEl.scrollIntoView({ behavior: 'smooth', block: 'center' });
        msgEl.style.outline = '2px solid #1a73e8';
        msgEl.style.borderRadius = '6px';
        setTimeout(() => { msgEl.style.outline = ''; msgEl.style.borderRadius = ''; }, 1800);
      });

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
    countLabel.textContent = (thisIdx !== -1 ? 1 : messages.length) + ' selected';

    const exportBtn = document.createElement('button');
    exportBtn.className = 'cgd-export-sel-btn';
    exportBtn.textContent = 'Export →';

    footerMain.appendChild(countLabel);
    footerMain.appendChild(exportBtn);
    footer.appendChild(footerMain);

    // ── Dest row: wrap with scroll arrows if there are append options ──
    let destMount = destRow;
    if (appendBtns.length > 0) {
      const destWrap = document.createElement('div');
      destWrap.style.cssText = `display:flex;align-items:center;border-bottom:1px solid ${dark?'rgba(255,255,255,0.08)':'#eee'};padding:0 10px;`;

      const arrowBtn = (char, dir) => {
        const b = document.createElement('button');
        b.textContent = char;
        b.style.cssText = `flex-shrink:0;background:none;border:none;width:20px;font-size:18px;line-height:1;color:${dark?'#5f6368':'#bbb'};cursor:pointer;padding:0;`;
        b.addEventListener('click', (e) => { e.stopPropagation(); destRow.scrollBy({ left: dir * 110, behavior: 'smooth' }); });
        return b;
      };
      destRow.style.borderBottom = 'none';
      destWrap.appendChild(arrowBtn('‹', -1));
      destWrap.appendChild(destRow);
      destWrap.appendChild(arrowBtn('›', 1));
      destMount = destWrap;
    }

    // ── Assemble ──
    panel.appendChild(header);
    panel.appendChild(folderRow);
    panel.appendChild(folderDropEl);
    panel.appendChild(destMount);
    panel.appendChild(controls);
    panel.appendChild(list);
    panel.appendChild(footer);

    // ── Event listeners ──
    function close() {
      document.removeEventListener('click', outsideClickHandler);
      panel.remove();
    }

    function outsideClickHandler(e) {
      if (!panel.contains(e.target)) close();
    }

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
      close();
      if (selectedIndices.length === 0) return;
      if (selectedIndices.length === 1) { exportMessage(messages[selectedIndices[0]]); return; }
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

    function appendSelected() {
      if (!selectedExport) { showToast('❌ Select a destination doc first', true); return; }
      const selectedIndices = getSelectedIndices();
      if (selectedIndices.length === 0) { showToast('❌ Nothing selected', true); return; }
      close();
      showToast('⏳ Appending to Doc...');
      _resetImgCaptures();
      const now = new Date();
      const dateStr = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}-${String(now.getDate()).padStart(2,'0')}`;
      const convTitle = getConversationTitle();
      const sourceUrl = location.origin + location.pathname;
      const sourceHeader = `${platform}${convTitle ? ' · ' + convTitle : ''} · ${dateStr}\n${sourceUrl}`;
      const rawParts = selectedIndices.map(origIdx => {
        const text = extractMarkdown(messages[origIdx]).trim();
        return text ? `${platform} Response ${origIdx + 1}\n\n${text}` : '';
      }).filter(Boolean);
      const plainText = sourceHeader + '\n\n' + markdownToPlainText(rawParts.join('\n\n---\n\n'));
      const docUrl = selectedExport.url || `https://docs.google.com/document/d/${selectedExport.fileId}/edit`;
      chrome.runtime.sendMessage(
        { action: 'appendToDoc', fileId: selectedExport.fileId, text: plainText },
        (resp) => {
          if (chrome.runtime.lastError) {
            showToast('❌ Append failed: ' + chrome.runtime.lastError.message, true);
            return;
          }
          if (resp && resp.success) {
            showToast(
              `✅ Added to "<b>${escHtml(selectedExport.fileName)}</b>"! <a href="${docUrl}" target="_blank" style="color:#fff;text-decoration:underline">Open ↗</a>`,
              false, 5000
            );
          } else {
            showToast('❌ Append failed: ' + (resp?.error || 'Unknown error'), true);
          }
        }
      );
    }

    header.querySelector('.cgd-panel-close').addEventListener('click', close);

    controls.querySelector('#cgd-sa').addEventListener('click', () => { checkboxes.forEach(cb => cb.checked = true); updateCount(); });
    controls.querySelector('#cgd-sn').addEventListener('click', () => { checkboxes.forEach(cb => cb.checked = false); updateCount(); });

    exportBtn.addEventListener('click', () => {
      if (destMode === 'append' && selectedExport?.fileId) appendSelected();
      else exportSelected();
    });

    document.body.appendChild(panel);
    // Delay so the click that opened the panel doesn't immediately close it
    setTimeout(() => document.addEventListener('click', outsideClickHandler), 0);

    if (thisIdx !== -1 && rowEls[thisIdx]) {
      setTimeout(() => rowEls[thisIdx].scrollIntoView({ block: 'nearest' }), 50);
    } else {
      list.scrollTop = list.scrollHeight;
    }
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

  function createExportButton() {
    const btn = document.createElement('button');
    btn.className = BUTTON_CLASS;
    btn.innerHTML = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg> Export to Docs <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>`;
    btn.title = 'Export options';
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
