/* MHM MD — converter engine
 * Converts PDF / DOCX / TXT / MD / HTML / CSV into clean, structured Markdown
 * that Claude reads with as few tokens as possible.
 * Works in the browser (globals) and in Node (require) for testing.
 */
(function (root) {
  'use strict';

  let LIB = {};
  function init(libs) { LIB = Object.assign(LIB, libs || {}); }

  // ---------- helpers ----------
  const AR_RE = /[؀-ۿݐ-ݿﭐ-﷿ﹰ-﻿]/;
  const AR_G = /[؀-ۿݐ-ݿﭐ-﷿ﹰ-﻿]/g;
  const LAT_G = /[A-Za-z]/g;
  const PRES_FORMS = /[ﭐ-﷿ﹰ-﻿]/;

  function estTokens(s) {
    if (!s) return 0;
    const ar = (s.match(AR_G) || []).length;
    return Math.round(ar / 2.2 + (s.length - ar) / 4);
  }
  function isRTL(s) {
    return (s.match(AR_G) || []).length > (s.match(LAT_G) || []).length;
  }
  function detectLang(s) {
    const ar = (s.match(AR_G) || []).length, la = (s.match(LAT_G) || []).length;
    if (ar && la && Math.min(ar, la) / Math.max(ar, la) > 0.25) return 'Arabic + English';
    return ar > la ? 'Arabic' : 'English';
  }
  const AR_DIGITS = '٠١٢٣٤٥٦٧٨٩';
  function westernDigits(s) { return s.replace(/[٠-٩]/g, d => String(AR_DIGITS.indexOf(d))); }

  function cleanInline(s) {
    if (!s) return '';
    if (PRES_FORMS.test(s)) s = s.normalize('NFKC');
    return s
      .replace(/[​‎‏‪-‮⁦-⁩﻿­]/g, '')
      .replace(/ـ+/g, '')            // Arabic tatweel
      .replace(/ /g, ' ')
      .replace(/[ \t]+/g, ' ')
      .replace(/([.…·_\-=*~])\1{4,}/g, '$1$1$1') // long runs of dots/dashes
      .trim();
  }

  const BULLET_RE = /^\s*(?:[•●▪◦○■□➢➤►▶✓✔·\uF0B7\uF0A7\uF0D8\uF076\uF0FC]\s*|[-–—*+]\s+)/;
  const NUM_RE = /^\s*(?:(\d{1,3}|[٠-٩]{1,3})[.)\-]|\(\s*(\d{1,3}|[٠-٩]{1,3}|[a-z]|[أ-ي])\s*\)|([a-z]|[أ-ي])\))\s+/;

  function listInfo(text) {
    let m = text.match(BULLET_RE);
    if (m) return { ordered: false, text: text.slice(m[0].length).trim() };
    m = text.match(NUM_RE);
    if (m) {
      const num = m[1] || m[2] || m[3];
      if (/^\d+$|^[٠-٩]+$/.test(num)) return { ordered: true, num: +westernDigits(num), text: text.slice(m[0].length).trim() };
      return { ordered: false, label: num + ')', text: text.slice(m[0].length).trim() };
    }
    return null;
  }

  // Lines that are only page numbers / page labels
  const PAGE_NUM_RE = /^(?:[-–—\s]*(?:page|p\.|صفحة|الصفحة|ص)?\s*[\d٠-٩]{1,4}\s*(?:(?:of|من|\/)\s*[\d٠-٩]{1,4})?[-–—\s]*)$/i;
  // Table-of-contents leader lines: "Chapter 1 ........ 12"
  const TOC_LINE_RE = /(?:\.{3,}|…{2,}|·{3,}|_{4,})\s*[\d٠-٩ivxlc]{1,4}\s*$/i;

  // ---------- PDF ----------
  async function pdfToBlocks(data, opts, meta) {
    const pdfjsLib = LIB.pdfjsLib;
    if (!pdfjsLib) throw new Error('PDF library not loaded');
    const doc = await pdfjsLib.getDocument({
      data: data instanceof Uint8Array ? data : new Uint8Array(data),
      isEvalSupported: false, useSystemFonts: true, disableFontFace: true,
    }).promise;
    meta.pages = doc.numPages;
    try {
      const info = await doc.getMetadata();
      const t = info && info.info && info.info.Title ? String(info.info.Title).replace(/^Microsoft (Word|PowerPoint) - /i, '').trim() : '';
      if (t.length > 2 && !/^(untitled|document\d*|بدون عنوان)$/i.test(t) && !/\.(docx?|pdf|pptx?|html?|txt)$/i.test(t)) meta.title = t;
    } catch (e) { /* ignore */ }

    const pages = [];
    let totalChars = 0;
    for (let p = 1; p <= doc.numPages; p++) {
      const page = await doc.getPage(p);
      const vp = page.getViewport({ scale: 1 });
      const tc = await page.getTextContent();
      // resolve real font names for bold / monospace detection
      const fontInfo = {};
      try {
        await page.getOperatorList();
        for (const it of tc.items) {
          if (it.fontName && !(it.fontName in fontInfo)) {
            let name = '';
            try { const f = page.commonObjs.get(it.fontName); name = (f && (f.name || f.loadedName)) || ''; } catch (e) { /* not loaded */ }
            const fam = (tc.styles[it.fontName] && tc.styles[it.fontName].fontFamily) || '';
            fontInfo[it.fontName] = {
              bold: /bold|black|heavy|semibold|demi/i.test(name),
              mono: /mono|courier|consol|menlo|code/i.test(name) || fam === 'monospace',
            };
          }
        }
      } catch (e) { /* fonts unavailable */ }

      const items = [];
      for (const it of tc.items) {
        if (!it.str || !it.str.trim()) continue;
        const h = Math.abs(it.transform[3]) || it.height || 10;
        const fi = fontInfo[it.fontName] || {};
        items.push({ str: it.str, x: it.transform[4], y: it.transform[5], w: it.width || 0, h, bold: !!fi.bold, mono: !!fi.mono });
        totalChars += it.str.length;
      }
      pages.push({ items, width: vp.width, height: vp.height, lines: groupLines(items) });
      page.cleanup && page.cleanup();
    }
    if (totalChars < 20 * Math.max(1, doc.numPages) && totalChars < 200) {
      meta.scanned = true;
    }
    return linesToBlocks(pages, opts, meta);
  }

  function groupLines(items) {
    const allText = items.map(i => i.str).join('');
    const docRTL = (allText.match(AR_G) || []).length > (allText.match(LAT_G) || []).length;
    const sorted = items.slice().sort((a, b) => (b.y - a.y) || (a.x - b.x));
    const lines = [];
    for (const it of sorted) {
      let line = null;
      for (let i = lines.length - 1; i >= Math.max(0, lines.length - 3); i--) {
        const L = lines[i];
        if (Math.abs(L.y - it.y) < Math.min(L.h, it.h) * 0.55) { line = L; break; }
      }
      if (!line) { line = { y: it.y, h: it.h, items: [] }; lines.push(line); }
      line.items.push(it);
      line.h = Math.max(line.h, it.h);
    }
    return lines.map(L => buildLine(L, docRTL)).filter(l => l.text);
  }

  const charClass = s => AR_RE.test(s) ? 'R' : (/[A-Za-z0-9À-ɏ]/.test(s) ? 'L' : 'N');

  // Visual -> logical order for one line (simple bidi: runs of the minority
  // direction are reversed back after sorting by the main direction)
  function orderItems(items, rtl) {
    const its = items.slice().sort((a, b) => rtl ? (b.x + b.w) - (a.x + a.w) : a.x - b.x);
    const minor = rtl ? 'L' : 'R';
    const cls = its.map(i => charClass(i.str));
    const out = [];
    let i = 0;
    while (i < its.length) {
      if (cls[i] !== minor) { out.push(its[i]); i++; continue; }
      // extend run: minor items, allowing neutrals only when followed by another minor item
      let j = i;
      while (j + 1 < its.length) {
        if (cls[j + 1] === minor) { j++; continue; }
        if (cls[j + 1] === 'N') {
          let k = j + 1;
          while (k < its.length && cls[k] === 'N') k++;
          if (k < its.length && cls[k] === minor) { j = k; continue; }
        }
        break;
      }
      out.push(...its.slice(i, j + 1).reverse());
      i = j + 1;
    }
    return out;
  }

  function buildLine(L, docRTL) {
    const raw = L.items.map(i => i.str).join('');
    const ar = (raw.match(AR_G) || []).length, la = (raw.match(LAT_G) || []).length;
    const rtl = ar > 0 && (ar >= la * 0.4 || docRTL);
    const its = orderItems(L.items, rtl);
    let text = '';
    const cells = [];
    let cell = '';
    let prev = null;
    let boldChars = 0, monoChars = 0, chars = 0, sizeSum = 0;
    const edges = [];
    const parts = [];
    for (const it0 of its) {
      let it = it0;
      if (rtl && charClass(it.str) === 'L') {
        const m = it.str.match(/^([.,:;!?،؛]+)(.*[A-Za-z0-9].*)$/);
        if (m) it = Object.assign({}, it, { str: m[2] + m[1] });
      }
      if (!prev) edges.push(rtl ? it.x + it.w : it.x);
      if (prev) {
        const gap = rtl ? prev.x - (it.x + it.w) : it.x - (prev.x + prev.w);
        if (gap > Math.max(L.h * 0.6, 6)) {
          cells.push(cell.trim()); cell = '';
          edges.push(rtl ? it.x + it.w : it.x);
          text += '  ';
        } else if (gap > L.h * 0.12 && !/\s$/.test(text) && !/^\s/.test(it.str)) {
          text += ' '; cell += ' ';
        }
      }
      parts.push({ str: it.str, x: it.x, w: it.w });
      text += it.str; cell += it.str;
      const n = it.str.trim().length;
      chars += n; sizeSum += it.h * n;
      if (it.bold) boldChars += n;
      if (it.mono) monoChars += n;
      prev = it;
    }
    cells.push(cell.trim());
    const xs = L.items.map(i => i.x), xe = L.items.map(i => i.x + i.w);
    const fix = rtl ? fixLamAlef : (x => x);
    return {
      text: fix(cleanInline(text)),
      cells: cells.map(c => fix(cleanInline(c))).filter(Boolean),
      edges, parts,
      y: L.y, h: chars ? sizeSum / chars : L.h,
      bold: chars > 0 && boldChars / chars > 0.8,
      mono: chars > 0 && monoChars / chars > 0.8,
      rtl, x0: Math.min(...xs), x1: Math.max(...xe),
    };
  }

  function linesToBlocks(pages, opts, meta) {
    // 1) remove repeating headers / footers & page numbers
    // body font size (char-weighted mode)
    const sizeHist = {};
    for (const pg of pages) for (const l of pg.lines) {
      const k = Math.round(l.h * 2) / 2;
      sizeHist[k] = (sizeHist[k] || 0) + l.text.length;
    }
    let body = 10, best = -1;
    for (const k in sizeHist) if (sizeHist[k] > best) { best = sizeHist[k]; body = +k; }
    const nPages = pages.length;
    const norm = t => t.replace(/[\d٠-٩]+/g, '#').toLowerCase();
    // only lines in the top/bottom 11% of the page can be headers/footers
    const atEdge = (pg, l, i, n) => l.h <= body * 1.1 && (i < 3 || i >= n - 3) && (l.y > pg.height * 0.89 || l.y < pg.height * 0.11);
    if (opts.stripHeaders && nPages >= 3) {
      const count = {};
      for (const pg of pages) {
        const edge = new Set();
        const n = pg.lines.length;
        pg.lines.forEach((l, i) => { if (atEdge(pg, l, i, n)) edge.add(norm(l.text)); });
        edge.forEach(k => { count[k] = (count[k] || 0) + 1; });
      }
      const thr = Math.max(2, Math.ceil(nPages * 0.5));
      for (const pg of pages) {
        const n = pg.lines.length;
        pg.lines = pg.lines.filter((l, i) => !(atEdge(pg, l, i, n) && count[norm(l.text)] >= thr));
      }
    }
    if (opts.stripHeaders) {
      for (const pg of pages) {
        const n = pg.lines.length;
        pg.lines = pg.lines.filter((l, i) => !((i < 3 || i >= n - 3) && PAGE_NUM_RE.test(l.text)));
      }
    }
    // indentation relative to the page's text margin (for lists whose bullets are drawn as graphics)
    for (const pg of pages) {
      const lefts = pg.lines.filter(l => !l.rtl).map(l => l.x0).sort((a, b) => a - b);
      const rights = pg.lines.filter(l => l.rtl).map(l => l.x1).sort((a, b) => b - a);
      const mL = lefts.length ? lefts[Math.floor(lefts.length * 0.15)] : 0;
      const mR = rights.length ? rights[Math.floor(rights.length * 0.15)] : pg.width;
      for (const l of pg.lines) {
        l.indent = l.rtl ? mR - l.x1 : l.x0 - mL;
        const other = l.rtl ? l.x0 - (pg.width - mR) : (pg.width - mL) - l.x1;
        l.centered = Math.abs(l.indent - other) < l.h * 1.5 && l.indent > l.h * 3;
      }
    }
    if (opts.stripToc) {
      for (const pg of pages) pg.lines = pg.lines.filter(l => !TOC_LINE_RE.test(l.text));
    }

    const headSizes = Object.keys(sizeHist).map(Number)
      .filter(s => s >= body * 1.12)
      .sort((a, b) => b - a);
    // merge near-equal sizes into level buckets (max 3 levels)
    const buckets = [];
    for (const s of headSizes) {
      if (!buckets.length || buckets[buckets.length - 1] - s > body * 0.08) buckets.push(s);
    }
    const levelOf = h => {
      if (h < body * 1.12) return 0;
      for (let i = 0; i < buckets.length; i++) if (h >= buckets[i] - body * 0.04) return Math.min(i + 1, 3);
      return Math.min(buckets.length, 3);
    };
    const boldLevel = Math.min(Math.max(buckets.length, 1) + 1, 4);

    // 3) build blocks
    const blocks = [];
    let para = null; // {type:'p'|'li', text, lastLine}
    const flush = () => { if (para) { blocks.push(para); para = null; } };
    let typicalWidth = 0;
    {
      const ws = [];
      for (const pg of pages) for (const l of pg.lines) if (Math.abs(l.h - body) < 1) ws.push(l.x1 - l.x0);
      ws.sort((a, b) => a - b);
      typicalWidth = ws.length ? ws[Math.floor(ws.length * 0.8)] : 400;
    }

    for (let pi = 0; pi < pages.length; pi++) {
      const lines = pages[pi].lines;
      for (let i = 0; i < lines.length; i++) {
        const l = lines[i];
        // --- table detection: run of >=2 lines with same cell count >=2
        if (l.cells.length >= 2 && levelOf(l.h) === 0) {
          let j = i;
          const aligned = x => x.cells.length === l.cells.length && x.edges.length === l.edges.length &&
            x.edges.every((e, k) => Math.abs(e - l.edges[k]) < Math.max(l.h * 2.5, 20));
          while (j + 1 < lines.length && aligned(lines[j + 1])) j++;
          const rowsT = lines.slice(i, j + 1);
          const avgCell = rowsT.reduce((a, r) => a + r.cells.join('').length / r.cells.length, 0) / rowsT.length;
          // rows that the gap rule failed to split (tight columns): re-split by column edges
          while (j > i && j + 1 < lines.length && lines[j + 1].cells.length > 1 && lines[j + 1].cells.length < l.cells.length) {
            const rs = resplit(lines[j + 1], l.edges);
            if (!rs) break;
            lines[j + 1].cells = rs; j++;
          }
          if (j > i && avgCell < 40) {
            flush();
            blocks.push({ type: 'table', rows: lines.slice(i, j + 1).map(x => x.cells) });
            i = j; continue;
          }
        }
        // --- code (monospace)
        if (l.mono) {
          flush();
          const code = [l.text];
          while (i + 1 < lines.length && lines[i + 1].mono) code.push(lines[++i].text);
          blocks.push({ type: 'code', text: code.join('\n') });
          continue;
        }
        const lvl = levelOf(l.h);
        const shortish = l.text.length <= 120;
        if (lvl > 0 && shortish) {
          // merge multi-line headings of same size
          const prevB = blocks[blocks.length - 1];
          if (!para && prevB && prevB.type === 'h' && prevB.level === lvl && prevB.page === pi && Math.abs(prevB.y - l.y) < l.h * 1.8) {
            prevB.text += ' ' + l.text; prevB.y = l.y; continue;
          }
          flush();
          blocks.push({ type: 'h', level: lvl, text: l.text, page: pi, y: l.y });
          continue;
        }
        const li = listInfo(l.text);
        const prevLine = i > 0 ? lines[i - 1] : null;
        const gap = prevLine ? prevLine.y - l.y : 999;
        const nextLine = lines[i + 1];
        const nextGap = nextLine ? l.y - nextLine.y : 999;
        if (l.bold && !li && l.text.length <= 90 && !/[.،,;]$/.test(l.text) &&
            (!para || gap > l.h * 1.6) && (nextGap > l.h * 1.25 || !nextLine || !nextLine.bold)) {
          flush();
          blocks.push({ type: 'h', level: boldLevel, text: l.text.replace(/[:：]$/, ''), page: pi, y: l.y });
          continue;
        }
        if (li) {
          flush();
          para = { type: 'li', ordered: li.ordered, num: li.num, label: li.label, text: li.text, last: l, first0: l.x0 + 4, page: pi };
          continue;
        }
        // graphic bullets: consecutive indented lines → list items
        const ind = x => x && x.indent > x.h * 1.2 && !x.centered;
        if (ind(l) && levelOf(l.h) === 0) {
          const sameInd = x => ind(x) && Math.abs(x.indent - l.indent) < l.h;
          const endsSentence = t => /[.!?؟;:؛]$/.test(t);
          if (para && para.type === 'li' && para.auto && sameInd(para.last) && gap < l.h * 1.9) {
            if (!endsSentence(para.last.text)) { para.text = joinText(para.text, l.text); para.last = l; continue; }
            flush(); para = { type: 'li', ordered: false, auto: true, text: l.text, last: l, page: pi }; continue;
          }
          if ((nextLine && sameInd(nextLine) && endsSentence(l.text)) || (prevLine && sameInd(prevLine) && endsSentence(prevLine.text))) {
            flush(); para = { type: 'li', ordered: false, auto: true, text: l.text, last: l, page: pi }; continue;
          }
        }
        if (para && para.auto && !ind(l)) { flush(); }
        // continuation?
        if (para && pi === para.page && gap < l.h * 1.75 && Math.abs(l.h - para.last.h) < 1.2) {
          const endedShort = (para.last.x1 - para.last.x0) < typicalWidth * 0.7 && /[.!?؟:]$/.test(para.last.text);
          const indentOk = para.type !== 'li' ||
            (l.rtl ? l.x1 < para.last.x1 + 2 : l.x0 > para.first0 - 2);
          if (!endedShort && indentOk) { para.text = joinText(para.text, l.text); para.last = l; continue; }
        } else if (para && pi !== para.page && para.type === 'p' && !/[.!?؟:]$/.test(para.text) && i === 0) {
          para.text = joinText(para.text, l.text); para.last = l; para.page = pi; continue;
        }
        flush();
        para = { type: 'p', text: l.text, last: l };
        para.page = pi;
      }
      if (para) para.page = para.page == null ? pi : para.page;
    }
    flush();
    // fix: list items need page/first0 for continuation — set lazily
    return blocks;
  }

  function resplit(line, edges) {
    const cols = edges.map(() => '');
    let prev = null;
    for (const p of line.parts) {
      const pos = line.rtl ? p.x + p.w : p.x;
      let k = 0;
      for (let c = 0; c < edges.length; c++) {
        const tol = line.h * 0.6;
        if (line.rtl ? pos <= edges[c] + tol : pos >= edges[c] - tol) k = c;
      }
      const gap = prev ? (line.rtl ? prev.x - (p.x + p.w) : p.x - (prev.x + prev.w)) : 0;
      if (cols[k] && gap > line.h * 0.12) cols[k] += ' ';
      cols[k] += p.str;
      prev = p;
    }
    const out = cols.map(cleanInline);
    return out.every(Boolean) ? out : null;
  }

  // Some PDF writers emit the lam-alef ligature reversed ("ال" instead of "لا").
  // Only fix patterns that can never be valid Arabic.
  function fixLamAlef(t) {
    return t
      .replace(/(^|[\s(«"•●▪·])ال(?=[\s.,،؛:!?؟)»"]|$)/g, '$1لا')
      .replace(/ا([أإآ])ل/g, 'ال$1')
      .replace(/(^|\s)إال(?=\s|$)/g, '$1إلا')
      .replace(/(^|\s)أال(?=\s|$)/g, '$1ألا');
  }

  function joinText(a, b) {
    if (/[A-Za-z]-$/.test(a) && /^[a-z]/.test(b)) return a.slice(0, -1) + b;
    return a + ' ' + b;
  }

  // ---------- DOCX / HTML via Turndown ----------
  function makeTurndown(opts) {
    const TD = LIB.TurndownService;
    if (!TD) throw new Error('Turndown library not loaded');
    const td = new TD({ headingStyle: 'atx', bulletListMarker: '-', codeBlockStyle: 'fenced', emDelimiter: '*', strongDelimiter: '**' });
    td.escape = s => s; // no backslash noise — saves tokens
    td.remove(['script', 'style', 'noscript', 'head']);
    td.addRule('img', {
      filter: 'img',
      replacement: (c, node) => {
        if (!opts.imagePlaceholders) return '';
        const alt = (node.getAttribute('alt') || '').trim();
        return alt ? `[صورة: ${alt}]` : '[صورة]';
      },
    });
    td.addRule('emptyLink', {
      filter: node => node.nodeName === 'A' && !(node.getAttribute('href') || '').match(/^https?:|^mailto:/),
      replacement: c => c,
    });
    td.addRule('table', {
      filter: 'table',
      replacement: (c, node) => {
        const rows = [];
        const walk = n => {
          for (const ch of Array.from(n.childNodes || [])) {
            if (ch.nodeName === 'TR') rows.push(ch);
            else if (/^(THEAD|TBODY|TFOOT)$/.test(ch.nodeName)) walk(ch);
          }
        };
        walk(node);
        const data = rows.map(tr => Array.from(tr.childNodes)
          .filter(td2 => /^(TD|TH)$/.test(td2.nodeName))
          .map(cell => cleanInline(td.turndown(cell.innerHTML || '').replace(/\n+/g, ' ')).replace(/\|/g, '\\|')));
        return '\n\n' + tableMd(data) + '\n\n';
      },
    });
    td.addRule('cellOnly', { filter: ['td', 'th', 'tr', 'thead', 'tbody'], replacement: c => c });
    return td;
  }

  function tableMd(rows) {
    rows = rows.filter(r => r.some(c => c && c.trim()));
    if (!rows.length) return '';
    const n = Math.max(...rows.map(r => r.length));
    const pad = r => { const x = r.slice(); while (x.length < n) x.push(''); return x; };
    const out = ['| ' + pad(rows[0]).join(' | ') + ' |', '|' + ' --- |'.repeat(n)];
    for (const r of rows.slice(1)) out.push('| ' + pad(r).join(' | ') + ' |');
    return out.join('\n');
  }

  async function docxToMd(data, opts, meta) {
    const mammoth = LIB.mammoth;
    if (!mammoth) throw new Error('Word library not loaded');
    const input = (typeof Buffer !== 'undefined' && data instanceof Buffer) ? { buffer: data } : { arrayBuffer: data.buffer ? data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength) : data };
    const res = await mammoth.convertToHtml(input, {
      styleMap: [
        "p[style-name='Title'] => h1:fresh",
        "p[style-name='Subtitle'] => h2:fresh",
        "p[style-name='العنوان'] => h1:fresh",
        "p[style-name='Quote'] => blockquote:fresh",
        "p[style-name='Code'] => pre:separator('\\n')",
      ],
      convertImage: mammoth.images.imgElement(img => Promise.resolve({ src: '', alt: img.altText || '' })),
    });
    const td = makeTurndown(opts);
    return td.turndown(res.value || '');
  }

  function htmlToMd(html, opts) { return makeTurndown(opts).turndown(html); }

  // ---------- TXT ----------
  function txtToBlocks(text, opts) {
    const lines = text.replace(/\r\n?/g, '\n').split('\n');
    const chunks = [];
    let cur = [];
    for (const ln of lines) {
      if (!ln.trim()) { if (cur.length) { chunks.push(cur); cur = []; } }
      else cur.push(ln.replace(/\s+$/, ''));
    }
    if (cur.length) chunks.push(cur);

    const blocks = [];
    for (let ci = 0; ci < chunks.length; ci++) {
      const ch = chunks[ci];
      // code chunk
      const indented = ch.filter(l => /^( {4,}|\t)/.test(l)).length;
      const codey = ch.filter(l => /[{};]\s*$|^\s*(def |class |function |import |const |let |var |#include|<\/?\w+>)/.test(l)).length;
      if (ch.length >= 2 && (indented / ch.length > 0.7 || codey / ch.length > 0.6)) {
        blocks.push({ type: 'code', text: ch.join('\n') }); continue;
      }
      if (opts.stripToc && ch.every(l => TOC_LINE_RE.test(l))) continue;
      let start = 0;
      // single-line heading chunk
      if (ch.length === 1 && isTxtHeading(ch[0], chunks[ci + 1])) {
        const t = cleanInline(ch[0]);
        const caps = /[A-Z]{3}/.test(t) && t === t.toUpperCase();
        blocks.push({ type: 'h', level: caps ? 2 : 3, text: t.replace(/[:：]$/, '') });
        continue;
      }
      // "Rules:" followed by list
      if (ch.length > 1 && /[:：]$/.test(ch[0].trim()) && ch[0].trim().length <= 60 && listInfo(ch[1])) {
        blocks.push({ type: 'h', level: 3, text: cleanInline(ch[0]).replace(/[:：]$/, '') });
        start = 1;
      }
      let para = null;
      const avgLen = ch.reduce((a, l) => a + l.length, 0) / ch.length;
      const unwrap = avgLen > 55;
      for (let i = start; i < ch.length; i++) {
        const t = cleanInline(ch[i]);
        if (opts.stripToc && TOC_LINE_RE.test(t)) continue;
        const li = listInfo(t);
        if (li) {
          if (para) blocks.push(para);
          para = { type: 'li', ordered: li.ordered, num: li.num, label: li.label, text: li.text, indent: /^\s{2,}/.test(ch[i]) ? 1 : 0 };
          continue;
        }
        if (para && para.type === 'li' && /^\s{2,}/.test(ch[i])) { para.text += ' ' + t; continue; }
        if (para && para.type === 'p') { para.text = unwrap ? joinText(para.text, t) : para.text + '\n' + t; continue; }
        if (para) blocks.push(para);
        para = { type: 'p', text: t };
      }
      if (para) blocks.push(para);
    }
    return blocks;
  }

  function isTxtHeading(line, nextChunk) {
    const t = line.trim();
    if (t.length < 2 || t.length > 70 || !nextChunk) return false;
    if (/[.،,;!?؟…]$/.test(t) || listInfo(t)) return false;
    const caps = /[A-Z]{3}/.test(t) && t === t.toUpperCase();
    if (caps) return true;
    if (/^[^:：]+[:：]$/.test(t)) return true;
    if (/^#{1,6}\s/.test(t)) return true;
    return t.split(/\s+/).length <= 7 && !/[:：]/.test(t) && !/^[\d٠-٩]+$/.test(t);
  }

  // ---------- CSV ----------
  function csvToMd(text) {
    const delim = (text.split('\n')[0].match(/;/g) || []).length > (text.split('\n')[0].match(/,/g) || []).length ? ';' : ',';
    const rows = [];
    let row = [], cell = '', q = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (q) { if (c === '"' && text[i + 1] === '"') { cell += '"'; i++; } else if (c === '"') q = false; else cell += c; }
      else if (c === '"') q = true;
      else if (c === delim) { row.push(cell); cell = ''; }
      else if (c === '\n' || c === '\r') { if (c === '\r' && text[i + 1] === '\n') i++; row.push(cell); rows.push(row); row = []; cell = ''; }
      else cell += c;
    }
    if (cell || row.length) { row.push(cell); rows.push(row); }
    return tableMd(rows.map(r => r.map(c => cleanInline(c).replace(/\|/g, '\\|'))));
  }

  // ---------- blocks -> markdown ----------
  function blocksToMd(blocks) {
    const out = [];
    let prevType = null;
    for (const b of blocks) {
      let s = '';
      if (b.type === 'h') s = '#'.repeat(b.level) + ' ' + b.text;
      else if (b.type === 'p') s = b.text;
      else if (b.type === 'li') {
        const ind = b.indent ? '  ' : '';
        s = ind + (b.ordered ? `${b.num}. ` : '- ') + (b.label ? b.label + ' ' : '') + b.text;
      } else if (b.type === 'table') s = tableMd(b.rows.map(r => r.map(c => c.replace(/\|/g, '\\|'))));
      else if (b.type === 'code') s = '```\n' + b.text + '\n```';
      if (!s.trim()) continue;
      const tight = b.type === 'li' && prevType === 'li';
      out.push((out.length ? (tight ? '\n' : '\n\n') : '') + s);
      prevType = b.type;
    }
    return out.join('');
  }

  // ---------- post-processing ----------
  function finalize(md, opts) {
    let lines = md.replace(/\r\n?/g, '\n').split('\n');
    let inCode = false;
    const res = [];
    for (let l of lines) {
      if (/^```/.test(l.trim())) { inCode = !inCode; res.push(l.trim()); continue; }
      if (inCode) { res.push(l.replace(/\s+$/, '')); continue; }
      const lead = (l.match(/^(\s*)/) || ['', ''])[1];
      l = lead.slice(0, 4) + cleanInline(l);
      if (opts.stripHeaders && PAGE_NUM_RE.test(l.trim()) && l.trim()) continue;
      if (opts.dedupe && l.trim() && res.length && res[res.length - 1].trim() === l.trim() && !/^\|/.test(l.trim())) continue;
      // normalize bullets produced by other converters
      l = l.replace(/^(\s*)[*+]\s+/, '$1- ');
      // empty headings / emphasis
      if (/^#{1,6}\s*$/.test(l) || /^\*\*\s*\*\*$/.test(l.trim())) continue;
      res.push(l);
    }
    return res.join('\n')
      .replace(/\*\*\s*\*\*/g, '')
      .replace(/\*\*([^*\n]+)\*\*\*\*([^*\n]+)\*\*/g, '**$1$2**')
      .replace(/\n{3,}/g, '\n\n')
      .trim() + '\n';
  }

  // Extract instruction-like sentences so Claude sees them up-front
  const INSTR_RE = new RegExp([
    '\\b(must|must not|should|shall|never|always|do not|don\'t|required?|requirements?|ensure|make sure|mandatory|important|warning|note that|avoid|only use|forbidden|prohibited)\\b',
    '(يجب|ينبغي|لا ?بد|ممنوع|يمنع|يُمنع|مطلوب|ضروري|تأكد|احرص|التزم|اتبع|دائما|دائماً|أبدا|أبداً|ابدا|إلزامي|الزامي|تنبيه|مهم جدا|لا ت[\\u0621-\\u064A]{2,}|لا ي[\\u0621-\\u064A]{2,}|عدم )',
  ].join('|'), 'i');

  function extractInstructions(md, limit) {
    const found = [];
    const seen = new Set();
    let inCode = false, heading = '';
    for (const raw of md.split('\n')) {
      const l = raw.trim();
      if (/^```/.test(l)) { inCode = !inCode; continue; }
      if (inCode || !l || /^\|/.test(l)) continue;
      if (/^#{1,6}\s/.test(l)) { heading = l.replace(/^#+\s*/, ''); continue; }
      const text = l.replace(/^(-|\d+\.)\s+/, '');
      const sentences = text.split(/(?<=[.!؟?])\s+/);
      for (const s of sentences) {
        if (s.length < 6 || s.length > 300 || !INSTR_RE.test(s)) continue;
        const k = s.toLowerCase();
        if (seen.has(k)) continue;
        seen.add(k);
        found.push(s.replace(/\*\*/g, ''));
        if (found.length >= limit) return found;
      }
    }
    return found;
  }

  function buildToc(md) {
    const hs = md.split('\n').filter(l => /^#{1,3}\s/.test(l));
    if (hs.length < 5) return '';
    return hs.map(h => {
      const lvl = h.match(/^#+/)[0].length;
      return '  '.repeat(lvl - 1) + '- ' + h.replace(/^#+\s*/, '');
    }).join('\n');
  }

  // ---------- main ----------
  const DEFAULTS = {
    stripHeaders: true, stripToc: true, dedupe: true, imagePlaceholders: true,
    frontMatter: true, claudeNote: true, extractInstructions: true, toc: false,
    noteLang: 'en',
  };

  function extOf(name) { const m = /\.([a-z0-9]+)$/i.exec(name || ''); return m ? m[1].toLowerCase() : ''; }

  function decodeText(bytes) {
    const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    try { return new TextDecoder('utf-8', { fatal: true }).decode(u8).replace(/^﻿/, ''); }
    catch (e) {
      try { return new TextDecoder('windows-1256').decode(u8); } catch (e2) { return new TextDecoder('utf-8').decode(u8); }
    }
  }

  async function convert(file, options) {
    const opts = Object.assign({}, DEFAULTS, options || {});
    const ext = extOf(file.name);
    const meta = { source: file.name, format: ext.toUpperCase() };
    let body = '', rawText = '';

    if (ext === 'pdf') {
      const blocks = await pdfToBlocks(file.data, opts, meta);
      body = blocksToMd(blocks);
      rawText = blocks.map(b => b.type === 'table' ? b.rows.map(r => r.join(' ')).join('\n') : b.text).join('\n');
    } else if (ext === 'docx') {
      body = await docxToMd(file.data, opts, meta);
      rawText = body.replace(/[#*|`>-]/g, '');
    } else if (ext === 'html' || ext === 'htm') {
      const t = decodeText(file.data); rawText = t;
      body = htmlToMd(t, opts);
    } else if (ext === 'csv') {
      const t = decodeText(file.data); rawText = t; body = csvToMd(t);
    } else if (ext === 'md' || ext === 'markdown') {
      body = decodeText(file.data); rawText = body;
    } else if (ext === 'doc') {
      throw new Error('صيغة .doc القديمة غير مدعومة — احفظ الملف من Word بصيغة .docx ثم أعد المحاولة');
    } else {
      const t = decodeText(file.data); rawText = t;
      body = blocksToMd(txtToBlocks(t, opts));
    }

    body = finalize(body, opts);
    if (meta.scanned) {
      body = '> ⚠️ هذا الـPDF يبدو ممسوحاً ضوئياً (صور بدون نص) — لم يُستخرج نص كافٍ.\n\n' + body;
    }
    const lang = detectLang(body);

    // title
    let title = meta.title || '';
    const firstH = body.match(/^#\s+(.+)$/m);
    if (!title && firstH) title = firstH[1].trim();

    const parts = [];
    if (opts.frontMatter) {
      const fm = ['---', `source: ${file.name}`, `format: ${meta.format}`];
      if (meta.pages) fm.push(`pages: ${meta.pages}`);
      if (title) fm.push(`title: ${title.replace(/\n/g, ' ')}`);
      fm.push(`language: ${lang}`, '---');
      parts.push(fm.join('\n'));
    }
    if (opts.claudeNote) {
      parts.push(opts.noteLang === 'en'
        ? '> **For Claude:** This is the full content of the file above, converted to Markdown with its original structure (headings, lists, tables). Read it completely and follow every instruction and requirement it contains.'
        : '> **إلى Claude:** هذا هو المحتوى الكامل للملف المذكور أعلاه بعد تحويله إلى Markdown مع الحفاظ على بنيته (العناوين، القوائم، الجداول). اقرأه كاملاً والتزم بكل التعليمات والمتطلبات الواردة فيه.');
    }
    if (opts.extractInstructions && estTokens(body) >= 600) {
      const ins = extractInstructions(body, 40);
      if (ins.length) {
        const head = opts.noteLang === 'en' ? '## Key instructions (extracted from the document)' : '## أهم التعليمات والمتطلبات (مستخرجة من المستند)';
        parts.push(head + '\n' + ins.map(s => '- ' + s).join('\n'));
      }
    }
    if (opts.toc) {
      const toc = buildToc(body);
      if (toc) parts.push((opts.noteLang === 'en' ? '## Contents\n' : '## المحتويات\n') + toc);
    }
    if (parts.some(x => x.startsWith('## '))) parts.push('---');
    parts.push(body.trim());
    const md = parts.join('\n\n') + '\n';

    const rawTokens = estTokens(rawText);
    return {
      name: file.name,
      mdName: file.name.replace(/\.[^.]+$/, '') + '.md',
      markdown: md,
      meta: Object.assign(meta, { language: lang, title }),
      stats: {
        rawTokens,
        mdTokens: estTokens(md),
        chars: md.length,
        // Claude reads each PDF page as text + an image, roughly ≥1,500 extra tokens per page
        pdfDirectTokens: meta.pages ? rawTokens + meta.pages * 1500 : null,
      },
    };
  }

  // Combine several converted files into one Markdown
  function combine(results, opts) {
    opts = Object.assign({}, DEFAULTS, opts || {});
    const intro = opts.noteLang === 'en'
      ? `> **For Claude:** This file bundles ${results.length} documents converted to Markdown. Read all of them and follow every instruction they contain.`
      : `> **إلى Claude:** هذا الملف يجمع ${results.length} مستندات محوّلة إلى Markdown. اقرأها جميعاً والتزم بكل التعليمات الواردة فيها.`;
    const body = results.map((r, i) => {
      const md = r.markdown.replace(/^---[\s\S]*?\n---\n+/, '').replace(/^> \*\*(إلى Claude|For Claude):\*\*.*\n+/m, '');
      const demoted = md.replace(/^(#{1,5})\s/gm, '#$1 ');
      return `# ${i + 1}. ${r.name}\n\n${demoted.trim()}`;
    }).join('\n\n---\n\n');
    return (opts.claudeNote ? intro + '\n\n' : '') + body + '\n';
  }

  const api = { init, convert, combine, estTokens, DEFAULTS, _internal: { txtToBlocks, blocksToMd, finalize, extractInstructions, listInfo } };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.MHMConverter = api;
})(typeof window !== 'undefined' ? window : globalThis);
