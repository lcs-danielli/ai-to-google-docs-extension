/**
 * ChatGPT/Gemini → Google Docs Converter v3
 * Zero deps. Markdown+LaTeX → .docx with native editable equations.
 * Each ordered list restarts numbering at 1.
 */

function tokenizeLatex(latex) {
  const tokens = []; let i = 0; latex = latex.trim();
  while (i < latex.length) {
    const ch = latex[i];
    if (ch === ' ' || ch === '\t') { i++; continue; }
    if (ch === '{') { tokens.push({ type: 'LBRACE' }); i++; continue; }
    if (ch === '}') { tokens.push({ type: 'RBRACE' }); i++; continue; }
    if (ch === '^') { tokens.push({ type: 'CARET' }); i++; continue; }
    if (ch === '_') { tokens.push({ type: 'UNDERSCORE' }); i++; continue; }
    if ('()[]&'.includes(ch)) { tokens.push({ type: 'TEXT', value: ch }); i++; continue; }
    if (ch === '\\') {
      let cmd = '\\'; i++;
      if (i < latex.length && /[a-zA-Z]/.test(latex[i])) {
        while (i < latex.length && /[a-zA-Z]/.test(latex[i])) { cmd += latex[i]; i++; }
        if (cmd === '\\begin' && i < latex.length && latex[i] === '{') {
          i++; // consume {
          let envName = '';
          while (i < latex.length && latex[i] !== '}') { envName += latex[i]; i++; }
          if (i < latex.length) i++; // consume }
          const endMarker = `\\end{${envName}}`;
          const endIdx = latex.indexOf(endMarker, i);
          const content = endIdx >= 0 ? latex.substring(i, endIdx).trim() : latex.substring(i).trim();
          if (endIdx >= 0) i = endIdx + endMarker.length; else i = latex.length;
          tokens.push({ type: 'ENV', env: envName, content });
          continue;
        }
      } else if (i < latex.length) { cmd += latex[i]; i++; }
      tokens.push({ type: 'COMMAND', value: cmd }); continue;
    }
    if (/[0-9]/.test(ch)) {
      let num = ''; while (i < latex.length && /[0-9\.]/.test(latex[i])) { num += latex[i]; i++; }
      tokens.push({ type: 'TEXT', value: num }); continue;
    }
    tokens.push({ type: 'TEXT', value: ch }); i++;
  }
  return tokens;
}

function parseLatex(tokens) {
  let pos = 0;
  function peek() { return pos < tokens.length ? tokens[pos] : null; }
  function advance() { return tokens[pos++]; }
  const SYMBOLS = {
    '\\pm':'\u00B1','\\mp':'\u2213','\\times':'\u00D7','\\div':'\u00F7','\\cdot':'\u00B7',
    '\\leq':'\u2264','\\le':'\u2264','\\geq':'\u2265','\\ge':'\u2265','\\neq':'\u2260','\\ne':'\u2260',
    '\\approx':'\u2248','\\equiv':'\u2261','\\infty':'\u221E',
    '\\alpha':'\u03B1','\\beta':'\u03B2','\\gamma':'\u03B3','\\delta':'\u03B4',
    '\\epsilon':'\u03B5','\\varepsilon':'\u03B5','\\theta':'\u03B8',
    '\\lambda':'\u03BB','\\mu':'\u03BC','\\pi':'\u03C0','\\sigma':'\u03C3',
    '\\phi':'\u03C6','\\omega':'\u03C9','\\tau':'\u03C4','\\rho':'\u03C1',
    '\\Delta':'\u0394','\\Sigma':'\u03A3','\\Omega':'\u03A9','\\Pi':'\u03A0',
    '\\Gamma':'\u0393','\\Lambda':'\u039B','\\Phi':'\u03A6','\\Theta':'\u0398',
    '\\to':'\u2192','\\rightarrow':'\u2192','\\leftarrow':'\u2190',
    '\\Rightarrow':'\u21D2','\\Leftarrow':'\u21D0','\\leftrightarrow':'\u2194',
    '\\in':'\u2208','\\notin':'\u2209','\\subset':'\u2282','\\supset':'\u2283',
    '\\subseteq':'\u2286','\\supseteq':'\u2287',
    '\\cup':'\u222A','\\cap':'\u2229','\\forall':'\u2200','\\exists':'\u2203',
    '\\partial':'\u2202','\\nabla':'\u2207','\\angle':'\u2220',
    '\\quad':'  ','\\qquad':'    ','\\,':' ','\\;':' ','\\!':'',
    '\\ldots':'\u2026','\\cdots':'\u22EF','\\vdots':'\u22EE','\\dots':'\u2026',
    '\\left':'','\\right':'','\\Big':'','\\big':'','\\bigg':'','\\Bigg':'',
    '\\bigl':'','\\bigr':'','\\biggl':'','\\biggr':'',
    '\\langle':'\u27E8','\\rangle':'\u27E9',
    '\\lfloor':'\u230A','\\rfloor':'\u230B','\\lceil':'\u2308','\\rceil':'\u2309',
    '\\neg':'\u00AC','\\wedge':'\u2227','\\vee':'\u2228',
    '\\sum':'\u2211','\\prod':'\u220F','\\int':'\u222B','\\iint':'\u222C','\\iiint':'\u222D',
    '\\prime':'\u2032','\\propto':'\u221D','\\perp':'\u22A5',
  };
  const FUNCS = ['\\log','\\ln','\\sin','\\cos','\\tan','\\sec','\\csc','\\cot','\\arcsin','\\arccos','\\arctan','\\sinh','\\cosh','\\tanh','\\lim','\\max','\\min','\\sup','\\inf','\\det','\\gcd','\\exp'];
  function parseGroup() {
    const t = peek(); if (!t) return {type:'text',value:''};
    if (t.type==='LBRACE') { advance(); const items=[]; while(peek()&&peek().type!=='RBRACE') items.push(parseExpr()); if(peek()&&peek().type==='RBRACE') advance(); return items.length===1?items[0]:{type:'group',children:items}; }
    return parseAtom();
  }
  function parseAtom() {
    const t = peek(); if (!t) return {type:'text',value:''};
    if (t.type==='ENV') {
      advance();
      const envRows = t.content.split(/\\\\/)
        .filter(row => row.trim())
        .map(row => row.split('&').map(cell => parseLatex(tokenizeLatex(cell.trim()))));
      return { type: 'matrix', env: t.env, rows: envRows };
    }
    if (t.type==='COMMAND') { advance(); const cmd=t.value;
      if (SYMBOLS[cmd]!==undefined) return {type:'text',value:SYMBOLS[cmd]};
      if (cmd==='\\frac'||cmd==='\\dfrac'||cmd==='\\tfrac') return {type:'fraction',numerator:parseGroup(),denominator:parseGroup()};
      if (cmd==='\\sqrt') {
        if (peek()&&peek().type==='TEXT'&&peek().value==='[') { advance(); let d=''; while(peek()&&!(peek().type==='TEXT'&&peek().value===']')){if(peek().type==='TEXT')d+=peek().value;advance();} if(peek())advance(); return {type:'radical',degree:d,content:parseGroup()}; }
        return {type:'radical',content:parseGroup()};
      }
      if (cmd==='\\text'||cmd==='\\mathrm'||cmd==='\\textrm'||cmd==='\\textbf'||cmd==='\\mathbf') return {type:'textmode',content:parseGroup(),bold:cmd==='\\textbf'||cmd==='\\mathbf'};
      if (cmd==='\\overline'||cmd==='\\bar'||cmd==='\\hat'||cmd==='\\tilde'||cmd==='\\vec') { const ac={'\\overline':'\u0305','\\bar':'\u0304','\\hat':'\u0302','\\tilde':'\u0303','\\vec':'\u20D7'}; return {type:'accent',accent:ac[cmd]||'',content:parseGroup()}; }
      if (cmd==='\\underbrace'||cmd==='\\overbrace') return {type:'group',children:[parseGroup()]};
      if (FUNCS.includes(cmd)) return {type:'funcname',value:cmd.substring(1)};
      if (cmd==='\\binom'){const n=parseGroup(),k=parseGroup();return{type:'group',children:[{type:'text',value:'('},{type:'fraction',numerator:n,denominator:k},{type:'text',value:')'}]};}
      return {type:'text',value:cmd.substring(1)};
    }
    if (t.type==='TEXT'){advance();return{type:'text',value:t.value};}
    if (t.type==='LBRACE') return parseGroup();
    advance(); return {type:'text',value:''};
  }
  function parseExpr() {
    let node = parseAtom();
    while (peek()&&(peek().type==='CARET'||peek().type==='UNDERSCORE')) {
      if (peek().type==='CARET'){advance();const sup=parseGroup();if(peek()&&peek().type==='UNDERSCORE'){advance();const sub=parseGroup();node={type:'subsup',base:node,superscript:sup,subscript:sub};}else node={type:'superscript',base:node,superscript:sup};}
      else{advance();const sub=parseGroup();if(peek()&&peek().type==='CARET'){advance();const sup=parseGroup();node={type:'subsup',base:node,superscript:sup,subscript:sub};}else node={type:'subscript',base:node,subscript:sub};}
    }
    return node;
  }
  const result=[]; while(pos<tokens.length){if(peek()&&peek().type==='RBRACE')break;result.push(parseExpr());}
  return result.length===1?result[0]:{type:'group',children:result};
}

function escapeXml(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
function flattenText(n){if(!n)return'';if(n.type==='text')return n.value||'';if(n.type==='group')return(n.children||[]).map(flattenText).join('');return'';}

function astToOmml(node) {
  if(!node)return'<m:r><m:t></m:t></m:r>';
  switch(node.type){
    case'text':return`<m:r><m:t>${escapeXml(node.value||'')}</m:t></m:r>`;
    case'funcname':return`<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>${escapeXml(node.value)}</m:t></m:r>`;
    case'textmode':{const sty=node.bold?'b':'p';return`<m:r><m:rPr><m:sty m:val="${sty}"/></m:rPr><m:t>${escapeXml(flattenText(node.content))}</m:t></m:r>`;}
    case'group':return(node.children||[]).map(astToOmml).join('');
    case'fraction':return`<m:f><m:fPr><m:ctrlPr/></m:fPr><m:num>${astToOmml(node.numerator)}</m:num><m:den>${astToOmml(node.denominator)}</m:den></m:f>`;
    case'superscript':return`<m:sSup><m:sSupPr/><m:e>${astToOmml(node.base)}</m:e><m:sup>${astToOmml(node.superscript)}</m:sup></m:sSup>`;
    case'subscript':return`<m:sSub><m:sSubPr/><m:e>${astToOmml(node.base)}</m:e><m:sub>${astToOmml(node.subscript)}</m:sub></m:sSub>`;
    case'subsup':return`<m:sSubSup><m:sSubSupPr/><m:e>${astToOmml(node.base)}</m:e><m:sub>${astToOmml(node.subscript)}</m:sub><m:sup>${astToOmml(node.superscript)}</m:sup></m:sSubSup>`;
    case'radical':if(node.degree)return`<m:rad><m:radPr/><m:deg>${astToOmml({type:'text',value:node.degree})}</m:deg><m:e>${astToOmml(node.content)}</m:e></m:rad>`;return`<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${astToOmml(node.content)}</m:e></m:rad>`;
    case'accent':return`<m:acc><m:accPr><m:chr m:val="${escapeXml(node.accent)}"/></m:accPr><m:e>${astToOmml(node.content)}</m:e></m:acc>`;
    case'matrix':{
      const DELIMS={pmatrix:['(',')'],bmatrix:['[',']'],vmatrix:['|','|'],Vmatrix:['\u2016','\u2016'],Bmatrix:['{','}'],cases:['{','']};
      const colCount=Math.max(1,...node.rows.map(r=>r.length));
      const jc=node.env==='cases'?'left':'center';
      let mXml=`<m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val="${colCount}"/><m:mcJc m:val="${jc}"/></m:mcPr></m:mc></m:mcs><m:ctrlPr/></m:mPr>`;
      for(const row of node.rows){mXml+='<m:mr>';for(const cell of row)mXml+=`<m:e>${astToOmml(cell)}</m:e>`;mXml+='</m:mr>';}
      mXml+='</m:m>';
      const d=DELIMS[node.env];
      if(d)return`<m:d><m:dPr><m:begChr m:val="${escapeXml(d[0])}"/><m:endChr m:val="${escapeXml(d[1])}"/><m:ctrlPr/></m:dPr><m:e>${mXml}</m:e></m:d>`;
      return mXml;
    }
    default:return'<m:r><m:t></m:t></m:r>';
  }
}

function latexToOmml(latex) {
  try {
    return astToOmml(parseLatex(tokenizeLatex(latex)));
  } catch(e) {
    return `<m:r><m:t>${escapeXml(latex)}</m:t></m:r>`;
  }
}

// ── MARKDOWN PARSER ──

function parseChatGPTMarkdown(text) {
  text = text.replace(/\r\n/g,'\n').replace(/\r/g,'\n');
  text = mergeListsWithMath(text);
  const lines=text.split('\n'), blocks=[]; let i=0;
  while(i<lines.length){
    const line=lines[i];
    if(line.trim()===''){i++;continue;}
    if(line.trim().startsWith('$$')){const r=consumeDisplayMath(lines,i);blocks.push({type:'displayMath',latex:r.latex});i=r.next;continue;}
    if(line.trim().startsWith('\\[')){const r=consumeBracketMath(lines,i);blocks.push({type:'displayMath',latex:r.latex});i=r.next;continue;}
    const hm=line.match(/^(#{1,6})\s+(.+)/);
    if(hm){blocks.push({type:'heading',level:hm[1].length,text:hm[2]});i++;continue;}
    if(line.trim().startsWith('```')){const cl=[];i++;while(i<lines.length&&!lines[i].trim().startsWith('```')){cl.push(lines[i]);i++;}if(i<lines.length)i++;blocks.push({type:'code',text:cl.join('\n')});continue;}
    if(line.match(/^\s*\d+[\.\)]\s/)){
      const items=[];
      while(i<lines.length){
        const m=lines[i].match(/^\s*\d+[\.\)]\s+(.*)/);
        if(m){items.push(m[1]||'');i++;}
        else if(lines[i].trim()===''){let a=i+1;while(a<lines.length&&lines[a].trim()==='')a++;if(a<lines.length&&lines[a].match(/^\s*\d+[\.\)]\s/)){i=a;}else break;}
        else break;
      }
      blocks.push({type:'orderedList',items});continue;
    }
    if(line.match(/^\s*[-\*\+]\s+/)){
      const items=[];
      while(i<lines.length){
        const m=lines[i].match(/^\s*[-\*\+]\s+(.*)/);
        if(m){items.push(m[1]||'');i++;}
        else if(lines[i].trim()===''){let a=i+1;while(a<lines.length&&lines[a].trim()==='')a++;if(a<lines.length&&lines[a].match(/^\s*[-\*\+]\s+/)){i=a;}else break;}
        else break;
      }
      blocks.push({type:'unorderedList',items});continue;
    }
    if(line.trim().startsWith('|')){
      const rows=[];
      while(i<lines.length&&lines[i].trim().startsWith('|')){
        const row=lines[i].trim().replace(/^\||\|$/g,'').split('|').map(c=>c.trim());
        if(!row.every(c=>/^[-:\s]+$/.test(c)))rows.push(row);
        i++;
      }
      if(rows.length>0)blocks.push({type:'table',rows});continue;
    }
    const pl=[];
    while(i<lines.length&&lines[i].trim()!==''&&!lines[i].match(/^#{1,6}\s/)&&!lines[i].trim().startsWith('$$')&&!lines[i].trim().startsWith('\\[')&&!lines[i].trim().startsWith('```')&&!lines[i].match(/^\s*\d+[\.\)]\s/)&&!lines[i].match(/^\s*[-\*\+]\s+/)&&!lines[i].trim().startsWith('|')){pl.push(lines[i]);i++;}
    if(pl.length>0)blocks.push({type:'paragraph',text:pl.join(' ')});
  }
  return blocks;
}

function mergeListsWithMath(text) {
  const lines=text.split('\n'), result=[];
  for(let i=0;i<lines.length;i++){
    const lm=lines[i].match(/^(\s*\d+[\.\)]\s*)$/);
    if(lm){
      let j=i+1; while(j<lines.length&&lines[j].trim()==='')j++;
      if(j<lines.length){
        const nl=lines[j].trim();
        if(nl.startsWith('$$')){
          let mc=nl.substring(2), ci=mc.indexOf('$$');
          if(ci>=0){result.push(lm[1]+'$'+mc.substring(0,ci).trim()+'$');i=j;continue;}
          else{const ml=[mc.trim()];j++;while(j<lines.length){if(lines[j].trim().includes('$$')){const lp=lines[j].trim().replace(/\$\$\s*$/,'').trim();if(lp)ml.push(lp);break;}ml.push(lines[j].trim());j++;}result.push(lm[1]+'$'+ml.join(' ')+'$');i=j;continue;}
        }
        if(nl.startsWith('\\(')||nl.startsWith('$')){result.push(lm[1]+nl);i=j;continue;}
      }
    }
    result.push(lines[i]);
  }
  return result.join('\n');
}

function consumeDisplayMath(lines,start){const f=lines[start].trim().substring(2),ci=f.indexOf('$$');if(ci>0)return{latex:f.substring(0,ci).trim(),next:start+1};const ml=[];if(f.trim())ml.push(f.trim());let j=start+1;while(j<lines.length){if(lines[j].trim().includes('$$')){const lp=lines[j].trim().replace(/\$\$\s*$/,'').trim();if(lp)ml.push(lp);j++;break;}ml.push(lines[j].trim());j++;}return{latex:ml.join(' '),next:j};}
function consumeBracketMath(lines,start){const f=lines[start].trim().replace(/^\\\[/,'').replace(/\\\]\s*$/,'').trim();if(f&&lines[start].trim().includes('\\]'))return{latex:f,next:start+1};const ml=[];if(f)ml.push(f);let j=start+1;while(j<lines.length){if(lines[j].includes('\\]')){const lp=lines[j].trim().replace(/\\\]\s*$/,'').trim();if(lp)ml.push(lp);j++;break;}ml.push(lines[j].trim());j++;}return{latex:ml.join(' '),next:j};}

// ── DOCX XML BUILDER (with per-list numbering) ──

function boldifyRuns(xml) {
  // Inject <w:b/> into existing run property blocks
  let r = xml.replace(/<w:rPr>/g, '<w:rPr><w:b/>');
  // Add run properties to bare <w:r> elements (skip drawing runs)
  r = r.replace(/<w:r>(?!<w:rPr>|<w:drawing>)/g, '<w:r><w:rPr><w:b/></w:rPr>');
  return r;
}

function buildTableXml(rows, imageMap, imageEntries) {
  if (!rows || rows.length === 0) return '';
  let tbl = '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>';
  rows.forEach((row, rowIdx) => {
    tbl += '<w:tr>';
    row.forEach(cell => {
      tbl += '<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr><w:p><w:pPr><w:spacing w:after="60"/></w:pPr>';
      if (rowIdx === 0) {
        const headerInner = buildMixedContent(cell, imageMap, imageEntries);
        tbl += headerInner ? boldifyRuns(headerInner) : `<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${escapeXml(cell)}</w:t></w:r>`;
      } else {
        const inner = buildMixedContent(cell, imageMap, imageEntries);
        tbl += inner || `<w:r><w:t xml:space="preserve">${escapeXml(cell)}</w:t></w:r>`;
      }
      tbl += '</w:p></w:tc>';
    });
    tbl += '</w:tr>';
  });
  tbl += '</w:tbl><w:p/>';
  return tbl;
}

function buildDocumentXml(blocks, imageMap) {
  let body='';
  let orderedListCount = 0;
  const imageEntries = [];

  for(const block of blocks){
    switch(block.type){
      case'heading':{const lv=Math.min(block.level,6);body+=`<w:p><w:pPr><w:pStyle w:val="Heading${lv}"/></w:pPr>${buildMixedContent(block.text,imageMap,imageEntries)}</w:p>`;break;}
      case'displayMath':{body+=`<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="120"/></w:pPr><m:oMathPara><m:oMath>${latexToOmml(block.latex)}</m:oMath></m:oMathPara></w:p>`;break;}
      case'code':{for(const cl of block.text.split('\n'))body+=`<w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">${escapeXml(cl)}</w:t></w:r></w:p>`;break;}
      case'orderedList':{
        orderedListCount++;
        const numId = orderedListCount;
        for(const item of block.items)
          body+=`<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr><w:spacing w:after="80"/></w:pPr>${buildMixedContent(item,imageMap,imageEntries)}</w:p>`;
        break;
      }
      case'unorderedList':{for(const item of block.items)body+=`<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="100"/></w:numPr><w:spacing w:after="80"/></w:pPr>${buildMixedContent(item,imageMap,imageEntries)}</w:p>`;break;}
      case'paragraph':{body+=`<w:p><w:pPr><w:spacing w:after="120"/></w:pPr>${buildMixedContent(block.text,imageMap,imageEntries)}</w:p>`;break;}
      case'table':{body+=buildTableXml(block.rows,imageMap,imageEntries);break;}
    }
  }
  return { xml: body, orderedListCount, imageEntries };
}

function cleanOrphanMarkers(text) {
  // Protect math segments
  const mathSegments = [];
  let cleaned = text.replace(/\$[^$]+\$/g, (match) => {
    mathSegments.push(match);
    return `\x00MATH${mathSegments.length - 1}\x00`;
  });

  // Protect valid italic: *text* (single * pairs with content)
  const italicSegments = [];
  cleaned = cleaned.replace(/\*([^*]+)\*/g, (match) => {
    italicSegments.push(match);
    return `\x00ITAL${italicSegments.length - 1}\x00`;
  });

  // Fix Gemini-style broken bold: odd number of ** means one is orphaned
  const dblParts = cleaned.split('**');
  if (dblParts.length > 1 && dblParts.length % 2 === 0) {
    // Odd number of ** markers = one is orphaned, remove the last one
    const lastIdx = cleaned.lastIndexOf('**');
    cleaned = cleaned.substring(0, lastIdx) + cleaned.substring(lastIdx + 2);
  }

  // Restore italic segments
  cleaned = cleaned.replace(/\x00ITAL(\d+)\x00/g, (_, idx) => italicSegments[parseInt(idx)]);

  // Restore math segments
  cleaned = cleaned.replace(/\x00MATH(\d+)\x00/g, (_, idx) => mathSegments[parseInt(idx)]);

  return cleaned;
}

function buildMixedContent(text, imageMap, imageEntries) {
  if(!text) return '';
  if (imageMap && imageEntries && text.includes('[[IMG:')) {
    return text.split(/(\[\[IMG:\d+\]\])/).map(part => {
      const m = part.match(/^\[\[IMG:(\d+)\]\]$/);
      if (m) {
        const img = imageMap[parseInt(m[1])];
        if (img && img.w && img.h) {
          const rIdNum = imageEntries.length + 4;
          const rId = `rId${rIdNum}`;
          const maxEmu = 5943600;
          let cx = img.w * 9525, cy = img.h * 9525;
          if (cx > maxEmu) { cy = Math.round(cy * maxEmu / cx); cx = maxEmu; }
          imageEntries.push({ rId, filename: `image${imageEntries.length + 1}.png`, data: img.data });
          return buildImageRunXml(rId, cx, cy);
        }
        return `<w:r><w:t>[Image]</w:t></w:r>`;
      }
      return _buildMixedContentText(part);
    }).join('');
  }
  return _buildMixedContentText(text);
}

function _buildMixedContentText(text) {
  if(!text)return'';
  text = cleanOrphanMarkers(text);
  
  // Parse into formatted segments (bold/italic spans that may contain math)
  const segments = parseFormattedSegments(text);
  
  let result = '';
  for (const seg of segments) {
    const mathRuns = extractMathRuns(seg.text);
    for (const run of mathRuns) {
      if (run.isMath) {
        result += `<m:oMath>${latexToOmml(run.text)}</m:oMath>`;
      } else {
        // Clean orphan markers from text
        let clean = run.text
          .replace(/\*\*/g, '')
          .replace(/^\*\s*/g, '').replace(/\s*\*$/g, '')
          .replace(/\*(?=\s*[\(\)])/g, '').replace(/(?<=[\(\)])\s*\*/g, '')
          .replace(/\*(?=\s)/g, '').replace(/(?<=\s)\*/g, '');
        if (!clean) continue;
        
        if (seg.bold && seg.italic) {
          result += `<w:r><w:rPr><w:b/><w:i/></w:rPr><w:t xml:space="preserve">${escapeXml(clean)}</w:t></w:r>`;
        } else if (seg.bold) {
          result += `<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${escapeXml(clean)}</w:t></w:r>`;
        } else if (seg.italic) {
          result += `<w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">${escapeXml(clean)}</w:t></w:r>`;
        } else {
          // Plain segment — check for nested bold/italic without math
          result += buildFormattedTextRuns(clean);
        }
      }
    }
  }
  return result;
}

// Split text into segments preserving bold/italic formatting
function parseFormattedSegments(text) {
  const segments = [];
  let remaining = text;
  
  // Match ***...***, **...**, *...* (allowing content with $ inside)
  const pattern = /(\*\*\*[\s\S]+?\*\*\*|\*\*[\s\S]+?\*\*|\*(?!\*|\s)[\s\S]+?(?<!\s|\*)\*(?!\*))/;
  
  while (remaining.length > 0) {
    const match = remaining.match(pattern);
    if (!match || match.index === undefined) {
      if (remaining) segments.push({ text: remaining, bold: false, italic: false });
      break;
    }
    
    if (match.index > 0) {
      segments.push({ text: remaining.substring(0, match.index), bold: false, italic: false });
    }
    
    const m = match[0];
    if (m.startsWith('***') && m.endsWith('***') && m.length > 6) {
      segments.push({ text: m.slice(3, -3), bold: true, italic: true });
    } else if (m.startsWith('**') && m.endsWith('**') && m.length > 4) {
      segments.push({ text: m.slice(2, -2), bold: true, italic: false });
    } else if (m.startsWith('*') && m.endsWith('*') && m.length > 2) {
      segments.push({ text: m.slice(1, -1), bold: false, italic: true });
    } else {
      segments.push({ text: m, bold: false, italic: false });
    }
    
    remaining = remaining.substring(match.index + m.length);
  }
  
  return segments;
}

// Extract math runs from text
function extractMathRuns(text) {
  const runs = [];
  let remaining = text;
  
  while (remaining.length > 0) {
    let mathStart = -1, latex = '', matchLen = 0;
    
    let ddIdx = remaining.indexOf('$$');
    let sdIdx = -1;
    for (let i = 0; i < remaining.length; i++) {
      if (remaining[i] === '$') {
        if (i + 1 < remaining.length && remaining[i + 1] === '$') { i++; continue; }
        sdIdx = i; break;
      }
    }
    let piIdx = remaining.indexOf('\\(');
    
    let bestIdx = Infinity, bestType = '';
    if (ddIdx >= 0 && ddIdx < bestIdx) { bestIdx = ddIdx; bestType = 'dd'; }
    if (sdIdx >= 0 && sdIdx < bestIdx) { bestIdx = sdIdx; bestType = 'sd'; }
    if (piIdx >= 0 && piIdx < bestIdx) { bestIdx = piIdx; bestType = 'pi'; }
    
    if (bestType === 'dd') {
      const endIdx = remaining.indexOf('$$', ddIdx + 2);
      if (endIdx > ddIdx) { mathStart = ddIdx; latex = remaining.substring(ddIdx + 2, endIdx); matchLen = endIdx + 2 - ddIdx; }
    } else if (bestType === 'sd') {
      let j = sdIdx + 1;
      while (j < remaining.length) {
        if (remaining[j] === '$' && (j + 1 >= remaining.length || remaining[j + 1] !== '$')) {
          mathStart = sdIdx; latex = remaining.substring(sdIdx + 1, j); matchLen = j + 1 - sdIdx; break;
        }
        j++;
      }
    } else if (bestType === 'pi') {
      const ci = remaining.indexOf('\\)', piIdx + 2);
      if (ci > piIdx) { mathStart = piIdx; latex = remaining.substring(piIdx + 2, ci); matchLen = ci + 2 - piIdx; }
    }
    
    if (mathStart >= 0 && latex) {
      if (mathStart > 0) runs.push({ text: remaining.substring(0, mathStart), isMath: false });
      runs.push({ text: latex, isMath: true });
      remaining = remaining.substring(mathStart + matchLen);
    } else {
      runs.push({ text: remaining, isMath: false });
      remaining = '';
    }
  }
  return runs;
}

function buildFormattedTextRuns(text) {
  if(!text)return'';
  let result='';
  const parts=text.split(/(\*\*\*[^*]+\*\*\*|\*\*[^*]+\*\*|\*[^*]+\*|~[^~]+~|\^[^^]+\^)/g);
  for(const part of parts){
    if(!part)continue;
    if(part.startsWith('***')&&part.endsWith('***')&&part.length>6)
      result+=`<w:r><w:rPr><w:b/><w:i/></w:rPr><w:t xml:space="preserve">${escapeXml(part.slice(3,-3))}</w:t></w:r>`;
    else if(part.startsWith('**')&&part.endsWith('**')&&part.length>4)
      result+=`<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${escapeXml(part.slice(2,-2))}</w:t></w:r>`;
    else if(part.startsWith('*')&&part.endsWith('*')&&part.length>2&&!part.startsWith('**'))
      result+=`<w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">${escapeXml(part.slice(1,-1))}</w:t></w:r>`;
    else if(part.startsWith('~')&&part.endsWith('~')&&part.length>2)
      result+=`<w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t xml:space="preserve">${escapeXml(part.slice(1,-1))}</w:t></w:r>`;
    else if(part.startsWith('^')&&part.endsWith('^')&&part.length>2)
      result+=`<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t xml:space="preserve">${escapeXml(part.slice(1,-1))}</w:t></w:r>`;
    else {
      let clean = part.replace(/\*\*/g, '').replace(/^\*\s*/g, '').replace(/\s*\*$/g, '');
      if(clean) result+=`<w:r><w:t xml:space="preserve">${escapeXml(clean)}</w:t></w:r>`;
    }
  }
  return result;
}

// ── DOCX FILE TEMPLATES ──

const CONTENT_TYPES=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>`;
const ROOT_RELS=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;
const DOC_RELS=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>`;
const STYLES=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:sz w:val="36"/><w:color w:val="2F5496"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="200" w:after="100"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:sz w:val="30"/><w:color w:val="2F5496"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="160" w:after="80"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:sz w:val="26"/><w:color w:val="2F5496"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/></w:pPr></w:style><w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style></w:styles>`;
const SETTINGS=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:mathPr><m:mathFont m:val="Cambria Math"/></m:mathPr><w:defaultTabStop w:val="720"/></w:settings>`;

// Generate numbering.xml dynamically based on how many ordered lists exist
function buildNumberingXml(orderedListCount) {
  let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">';

  // One abstract num for ordered lists (decimal)
  xml += '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>';

  // One abstract num for bullets
  xml += '<w:abstractNum w:abstractNumId="99"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="\u2022"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl></w:abstractNum>';

  // Create a w:num for EACH ordered list, each referencing abstractNumId 0 but with lvlOverride to restart at 1
  for (let i = 1; i <= Math.max(orderedListCount, 1); i++) {
    xml += `<w:num w:numId="${i}"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>`;
  }

  // Bullet num
  xml += '<w:num w:numId="100"><w:abstractNumId w:val="99"/></w:num>';

  xml += '</w:numbering>';
  return xml;
}

function makeDocXml(body){return`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" mc:Ignorable="w14"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`;}

// ── ZIP CREATOR ──

function createZip(files){
  const enc=new TextEncoder(),parts=[],cd=[];let offset=0;
  for(const file of files){
    const nb=enc.encode(file.name),cb=file.binary?base64ToBytes(file.content):enc.encode(file.content),crc=crc32(cb);
    const lh=new Uint8Array(30+nb.length),lv=new DataView(lh.buffer);
    lv.setUint32(0,0x04034b50,true);lv.setUint16(4,20,true);lv.setUint32(14,crc,true);
    lv.setUint32(18,cb.length,true);lv.setUint32(22,cb.length,true);lv.setUint16(26,nb.length,true);lh.set(nb,30);
    const ce=new Uint8Array(46+nb.length),cv=new DataView(ce.buffer);
    cv.setUint32(0,0x02014b50,true);cv.setUint16(4,20,true);cv.setUint16(6,20,true);cv.setUint32(16,crc,true);
    cv.setUint32(20,cb.length,true);cv.setUint32(24,cb.length,true);cv.setUint16(28,nb.length,true);cv.setUint32(42,offset,true);ce.set(nb,46);
    parts.push(lh,cb);cd.push(ce);offset+=lh.length+cb.length;
  }
  let cds=0;for(const c of cd)cds+=c.length;
  const eocd=new Uint8Array(22),ev=new DataView(eocd.buffer);
  ev.setUint32(0,0x06054b50,true);ev.setUint16(8,files.length,true);ev.setUint16(10,files.length,true);ev.setUint32(12,cds,true);ev.setUint32(16,offset,true);
  const all=[...parts,...cd,eocd];let tl=0;for(const a of all)tl+=a.length;
  const r=new Uint8Array(tl);let p=0;for(const a of all){r.set(a,p);p+=a.length;}return r;
}
const CRC32_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let j = 0; j < 8; j++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    t[i] = c;
  }
  return t;
})();
function crc32(bytes) {
  let crc = 0xFFFFFFFF;
  for (let i = 0; i < bytes.length; i++) crc = CRC32_TABLE[(crc ^ bytes[i]) & 0xFF] ^ (crc >>> 8);
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

// ── IMAGE HELPERS ──

function base64ToBytes(b64) {
  const bs = atob(b64), out = new Uint8Array(bs.length);
  for (let i = 0; i < bs.length; i++) out[i] = bs.charCodeAt(i);
  return out;
}

function buildImageRunXml(rId, cx, cy) {
  const id = parseInt(rId.replace('rId', '')) || 1;
  // Namespaces (wp, a, pic) are declared on the root w:document element
  return `<w:r><w:drawing>` +
    `<wp:inline distT="0" distB="114300" distL="0" distR="0">` +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    `<wp:docPr id="${id}" name="${rId}"/>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="${rId}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic>` +
    `</wp:inline></w:drawing></w:r>`;
}

// ── PUBLIC API ──

function convertChatGPTToDocx(markdownText, imageMap = {}) {
  let blocks, bodyXml, orderedListCount, imageEntries;
  try {
    blocks = parseChatGPTMarkdown(markdownText);
    ({ xml: bodyXml, orderedListCount, imageEntries } = buildDocumentXml(blocks, imageMap));
  } catch(e) {
    throw e;
  }

  // Build dynamic document relationships (add image entries after the 3 standard ones)
  let docRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>`;
  for (const e of imageEntries) {
    docRels += `<Relationship Id="${e.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${e.filename}"/>`;
  }
  docRels += `</Relationships>`;

  // Add PNG content type if images are present
  const contentTypes = imageEntries.length > 0
    ? CONTENT_TYPES.replace('</Types>', '<Default Extension="png" ContentType="image/png"/></Types>')
    : CONTENT_TYPES;

  const files = [
    {name:'[Content_Types].xml',content:contentTypes},
    {name:'_rels/.rels',content:ROOT_RELS},
    {name:'word/_rels/document.xml.rels',content:docRels},
    {name:'word/document.xml',content:makeDocXml(bodyXml)},
    {name:'word/styles.xml',content:STYLES},
    {name:'word/numbering.xml',content:buildNumberingXml(orderedListCount)},
    {name:'word/settings.xml',content:SETTINGS},
  ];
  for (const e of imageEntries) {
    files.push({ name: `word/media/${e.filename}`, content: e.data, binary: true });
  }
  return new Blob([createZip(files)],{type:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
}

if(typeof window!=='undefined'){window.convertChatGPTToDocx=convertChatGPTToDocx;}
