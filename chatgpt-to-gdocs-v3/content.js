/**
 * Content Script v3: Works on ChatGPT AND Gemini
 * Detects which platform, uses appropriate selectors
 */

(function() {
  'use strict';

  const BUTTON_CLASS = 'cgd-export-btn';
  const BANNER_CLASS = 'cgd-toast';

  // Detect platform
  const isGemini = location.hostname.includes('gemini.google.com');
  const isChatGPT = location.hostname.includes('chatgpt.com') || location.hostname.includes('chat.openai.com');
  const isClaude = location.hostname.includes('claude.ai');

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

      // Strategy 3: Any element containing a katex-mathml child (catches wrapper spans)
      if (node.querySelector && node.querySelector('.katex-mathml annotation[encoding="application/x-tex"]')) {
        // Only if this node IS the katex container (not a large parent)
        const katexEl = node.querySelector('.katex');
        if (katexEl === null || node.classList.contains('katex') || node.querySelector(':scope > .katex-mathml')) {
          const tex = extractTeX(node);
          if (tex) {
            const isDisplay = node.closest('.katex-display') || node.classList.contains('katex-display');
            return isDisplay ? `\n$$${tex}$$\n` : `$${tex}$`;
          }
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
      if (tag === 'strong' || tag === 'b') return `**${getInner(node)}**`;
      if (tag === 'em' || tag === 'i') return `*${getInner(node)}*`;
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
        // Check if any LaTeX made it into the markdown
        const hasMath = md.includes('$');
        if (!hasMath) {
          console.warn('Math annotations found but not captured, trying direct extraction...');
          md = directExtractWithMath(contentDiv);
        }
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
    let node;
    const processedMathNodes = new Set();

    while (node = walker.nextNode()) {
      if (node.nodeType === Node.TEXT_NODE) {
        // Skip text inside katex rendered output
        const parent = node.parentElement;
        if (parent && (
          parent.closest('.katex-html') ||
          parent.closest('.katex-mathml') ||
          parent.closest('annotation') ||
          parent.tagName === 'ANNOTATION'
        )) continue;

        // Skip if inside a katex span (we handle it separately)
        if (parent && parent.closest('.katex')) continue;

        md += node.textContent;
      } else if (node.nodeType === Node.ELEMENT_NODE) {
        // Check for math elements
        if (node.classList && (node.classList.contains('katex') || node.classList.contains('katex-display'))) {
          if (processedMathNodes.has(node)) continue;
          processedMathNodes.add(node);
          const tex = extractTeX(node);
          if (tex) {
            const isDisplay = node.classList.contains('katex-display') || node.closest('.katex-display');
            md += isDisplay ? `\n$$${tex}$$\n` : `$${tex}$`;
          }
          // Skip all children
          let skip = node;
          while (walker.nextNode()) {
            if (!node.contains(walker.currentNode)) { walker.previousNode(); break; }
          }
        }
        // Add line breaks for block elements
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
  //  CLIPBOARD EXTRACTION (ChatGPT only - uses copy button)
  // ═══════════════════════════════════════════════════════════════

  async function extractViaClipboard(messageEl) {
    if (isGemini) return null; // Gemini clipboard doesn't include LaTeX
    try {
      let container, copyBtn;
      
      if (isChatGPT) {
        container = messageEl.closest('[data-message-id]') || messageEl.closest('.group\\/conversation-turn');
        if (!container) return null;
        copyBtn = container.querySelector('button[data-testid="copy-turn-action-button"]');
      } else if (isClaude) {
        // Walk up to find the Copy button near this content
        container = messageEl;
        for (let i = 0; i < 8 && container; i++) {
          copyBtn = container.querySelector('button[aria-label="Copy"]');
          if (copyBtn) break;
          container = container.parentElement;
        }
      }
      
      if (!copyBtn) return null;
      copyBtn.click();
      await new Promise(r => setTimeout(r, 400));
      return await navigator.clipboard.readText();
    } catch(e) { return null; }
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
  //  MAIN EXPORT FUNCTION
  // ═══════════════════════════════════════════════════════════════

  async function exportMessage(messageEl) {
    showToast('⏳ Generating document...');

    let markdown = null;

    // Check if the content has math (KaTeX elements)
    const contentDiv = messageEl.querySelector('.markdown') || messageEl;
    const hasMathInDOM = contentDiv.querySelector('annotation[encoding="application/x-tex"], .katex, .math-inline, .math-block, mjx-container, [data-math]');

    // Try clipboard first
    try { markdown = await extractViaClipboard(messageEl); } catch(e) {}

    // If content has math but clipboard doesn't have $ delimiters, clipboard lost the math
    if (markdown && hasMathInDOM && !markdown.includes('$') && !markdown.includes('\\(') && !markdown.includes('\\[')) {
      console.log('Clipboard lost math, falling back to DOM extraction');
      markdown = null;
    }

    if (!markdown || markdown.trim().length < 10) markdown = extractMarkdown(messageEl);
    if (!markdown || markdown.trim().length === 0) {
      showToast('❌ Could not extract content from this message', true);
      return;
    }

    // Debug: log the extracted markdown to console
    console.log('=== EXTRACTED MARKDOWN ===');
    console.log(markdown);
    console.log('=== END MARKDOWN ===');

    let blob;
    try {
      blob = window.convertChatGPTToDocx(markdown);
    } catch(e) {
      console.error('Converter error:', e, '\nMarkdown:', markdown);
      showToast('❌ Error generating document: ' + e.message, true);
      return;
    }

    const now = new Date();
    const timestamp = `${now.getFullYear()}${(now.getMonth()+1).toString().padStart(2,'0')}${now.getDate().toString().padStart(2,'0')}_${now.getHours().toString().padStart(2,'0')}${now.getMinutes().toString().padStart(2,'0')}`;
    const platform = isGemini ? 'Gemini' : isClaude ? 'Claude' : 'ChatGPT';
    const filename = `${platform}_Export_${timestamp}.docx`;

    showToast('⏳ Uploading to Google Drive...');

    try {
      if (!chrome.runtime || !chrome.runtime.sendMessage) {
        throw new Error('Extension reloaded. Please refresh this page and try again.');
      }
      const base64 = await blobToBase64(blob);
      const result = await new Promise((resolve, reject) => {
        chrome.runtime.sendMessage(
          { action: 'uploadToDrive', docxBase64: base64, filename: filename },
          (response) => {
            if (chrome.runtime.lastError) reject(new Error(chrome.runtime.lastError.message));
            else if (response && response.success) resolve(response);
            else reject(new Error(response ? response.error : 'Unknown error'));
          }
        );
      });

      showToast(`✅ Created "<b>${result.fileName}</b>" in Google Drive! Opening...`, false, 5000);
      setTimeout(() => window.open(result.url, '_blank'), 500);

    } catch(e) {
      console.warn('Drive upload failed, falling back to download:', e.message);

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

      setTimeout(() => window.open('https://drive.google.com/drive/my-drive', '_blank'), 800);
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
      btn.addEventListener('click', (e) => {
        e.preventDefault(); e.stopPropagation();
        exportMessage(msg.querySelector('.markdown') || msg);
      });
      actionArea.appendChild(btn);
    }
  }

  function findChatGPTActionBar(container) {
    const copyBtn = container.querySelector('button[data-testid="copy-turn-action-button"]');
    if (copyBtn) return copyBtn.parentElement;
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
        btn.addEventListener('click', (e) => {
          e.preventDefault(); e.stopPropagation();
          exportMessage(resp);
        });
        actionArea.appendChild(btn);
      }
    }

    // Also try to find responses by looking for the "copy" button in Gemini
    const copyButtons = document.querySelectorAll('button[aria-label="Copy"], button[data-tooltip="Copy"]');
    for (const copyBtn of copyButtons) {
      const actionBar = copyBtn.parentElement;
      if (!actionBar || actionBar.querySelector('.' + BUTTON_CLASS)) continue;

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
      btn.addEventListener('click', (e) => {
        e.preventDefault(); e.stopPropagation();
        exportMessage(contentEl);
      });
      actionBar.appendChild(btn);
    }
  }

  function findGeminiActionBar(container) {
    // Look for button groups containing copy/thumbs
    const buttons = container.querySelectorAll('button');
    for (const btn of buttons) {
      const label = (btn.getAttribute('aria-label') || btn.getAttribute('data-tooltip') || '').toLowerCase();
      if (label.includes('copy') || label.includes('share') || label.includes('thumb')) {
        return btn.parentElement;
      }
    }
    // Look for action containers
    const actionDiv = container.querySelector('.action-buttons') ||
                      container.querySelector('.response-actions') ||
                      container.querySelector('[class*="action"]');
    return actionDiv;
  }

  // ═══════════════════════════════════════════════════════════════
  //  SHARED: CREATE BUTTON
  // ═══════════════════════════════════════════════════════════════

  function createExportButton() {
    const btn = document.createElement('button');
    btn.className = BUTTON_CLASS;
    btn.innerHTML = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg> Export to Docs`;
    btn.title = 'Export to Google Docs with editable math equations';
    return btn;
  }

  // ═══════════════════════════════════════════════════════════════
  //  CLAUDE: INJECT BUTTONS
  // ═══════════════════════════════════════════════════════════════

  function addClaudeButtons() {
    // Find all Copy buttons (aria-label="Copy")
    const copyButtons = document.querySelectorAll('button[aria-label="Copy"]');

    for (const copyBtn of copyButtons) {
      // The grandparent is the action bar: "text-text-300 flex items-stretch justify-between"
      const actionBar = copyBtn.closest('.text-text-300');
      if (!actionBar || actionBar.querySelector('.' + BUTTON_CLASS)) continue;

      // Walk up from the action bar to find the full response
      // Claude structure: response container > content area > action bar
      // We need to find the div that contains BOTH the content and the action bar
      let responseContainer = actionBar.parentElement;
      
      // Keep walking up until we find the container with the actual response text
      for (let i = 0; i < 5 && responseContainer; i++) {
        const hasContent = responseContainer.querySelector('.font-claude-response') ||
                           responseContainer.querySelector('[class*="font-claude-response-body"]');
        if (hasContent) break;
        responseContainer = responseContainer.parentElement;
      }

      if (!responseContainer) continue;

      // Find the content div - try multiple selectors
      const contentEl = responseContainer.querySelector('[class*="font-claude-response"]:not(.text-text-300)') ||
                        responseContainer.querySelector('.contents') ||
                        responseContainer;

      // Skip user messages
      if (responseContainer.querySelector('.font-user-message')) continue;

      const btn = createExportButton();
      btn.addEventListener('click', (e) => {
        e.preventDefault(); e.stopPropagation();
        // Log what we're exporting for debugging
        console.log('Claude export - contentEl class:', contentEl.className);
        console.log('Claude export - contentEl children:', contentEl.children.length);
        console.log('Claude export - text length:', contentEl.textContent.length);
        exportMessage(contentEl);
      });

      // Insert next to the Copy button's parent (w-fit)
      const copyParent = copyBtn.parentElement;
      if (copyParent) {
        const wrapper = document.createElement('div');
        wrapper.className = 'w-fit';
        wrapper.appendChild(btn);
        copyParent.after(wrapper);
      } else {
        actionBar.appendChild(btn);
      }
    }
  }

  function findClaudeActionBar(container) {
    return container.querySelector('.text-text-300');
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
    observer._t = setTimeout(addButtons, 500);
  });
  observer.observe(document.body, { childList: true, subtree: true });
})();
