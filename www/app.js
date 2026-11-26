(function () {
  'use strict';
  const $ = s => document.querySelector(s);
  const C = window.MHMConverter;

  // ---------- platform ----------
  const Cap = window.Capacitor;
  const isNative = !!(Cap && Cap.isNativePlatform && Cap.isNativePlatform());
  const plugin = name => {
    if (!isNative) return null;
    try { return (Cap.Plugins && Cap.Plugins[name]) || (Cap.registerPlugin ? Cap.registerPlugin(name) : null); } catch (e) { return null; }
  };
  const Filesystem = plugin('Filesystem');
  const Share = plugin('Share');
  const IncomingFile = plugin('IncomingFile');
  const Desktop = window.mhmDesktop || null; // Electron preload bridge

  // Claude artifact host: downloads go through its capability
  let artifactDl = null;
  try { if (window.claude && window.claude.use) window.claude.use('downloads').then(d => { artifactDl = d; }).catch(() => {}); } catch (e) { /* not an artifact */ }

  // ---------- OCR (lazy-loaded tesseract.js, fully offline) ----------
  const abs = p => new URL(p, location.href).href;
  function loadScript(src) {
    return new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = src; s.onload = res; s.onerror = () => rej(new Error('OCR غير متاح في هذه النسخة — استخدم تطبيق الهاتف أو الكمبيوتر'));
      document.head.appendChild(s);
    });
  }
  let ocrWorkerP = null;
  function getOcr() {
    if (ocrWorkerP) return ocrWorkerP;
    ocrWorkerP = (async () => {
      if (!window.Tesseract) await loadScript('lib/ocr/tesseract.min.js');
      setStatus('⏳ تجهيز محرك OCR (أول مرة فقط)…');
      return window.Tesseract.createWorker(['ara', 'eng'], 1, {
        workerPath: abs('lib/ocr/worker.min.js'),
        corePath: abs('lib/ocr/core'),
        langPath: abs('lib/ocr/lang'),
        // the Android build unpacks .gz assets, so the app ships plain .traineddata files there
        gzip: !isNative, cacheMethod: 'none', workerBlobURL: false,
      });
    })();
    ocrWorkerP.catch(() => { ocrWorkerP = null; });
    return ocrWorkerP;
  }
  async function ocr(input) {
    const w = await getOcr();
    let src = input instanceof Uint8Array ? new Blob([input]) : input;
    // small photos/screenshots: upscale ×2 — Tesseract reads Arabic much better at higher resolution
    if (src instanceof Blob && window.createImageBitmap) {
      try {
        const bmp = await createImageBitmap(src);
        if (bmp.width < 1600) {
          const k = Math.min(3, 2000 / bmp.width);
          const c = document.createElement('canvas');
          c.width = Math.round(bmp.width * k); c.height = Math.round(bmp.height * k);
          const g = c.getContext('2d'); g.imageSmoothingQuality = 'high'; g.drawImage(bmp, 0, 0, c.width, c.height);
          src = c;
        }
      } catch (e) { /* use original */ }
    }
    const { data } = await w.recognize(src, {}, { text: true, blocks: true });
    return ocrToText(data);
  }
  // Rebuild paragraphs and headings from Tesseract's line boxes (gaps → paragraphs, tall lines → headings)
  function ocrToText(data) {
    if (!data) return '';
    const lines = [];
    (data.blocks || []).forEach(b => (b.paragraphs || []).forEach(p => (p.lines || []).forEach(l => {
      const t = (l.text || '').replace(/[\u200e\u200f]/g, '').trim();
      if (t) lines.push({ t, y0: l.bbox.y0, y1: l.bbox.y1, h: l.bbox.y1 - l.bbox.y0 });
    })));
    if (!lines.length) return data.text || '';
    const hs = lines.map(l => l.h).sort((a, b) => a - b);
    const med = hs[Math.floor(hs.length / 2)] || 1;
    const out = [];
    let prev = null;
    for (const l of lines) {
      const big = l.h > med * 1.3 && l.t.length < 100;
      const gap = prev ? l.y0 - prev.y1 : 0;
      if (big) { out.push('', (l.h > med * 1.8 ? '# ' : '## ') + l.t, ''); prev = l; continue; }
      if (prev && gap > med * 0.75) out.push('');
      out.push(l.t);
      prev = l;
    }
    return out.join('\n').replace(/\n{3,}/g, '\n\n').trim();
  }

  C.init({
    pdfjsLib: window.pdfjsLib, mammoth: window.mammoth, TurndownService: window.TurndownService,
    JSZip: window.JSZip, XLSX: window.XLSX,
    ocr,
    createCanvas: (w, h) => { const c = document.createElement('canvas'); c.width = w; c.height = h; return c; },
    progress: msg => setStatus('⏳ ' + msg),
  });
  if (window.pdfjsLib) window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'lib/pdf.worker.min.js';

  // ---------- options (persisted on this device) ----------
  const store = {
    get(k, d) { try { const v = localStorage.getItem(k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* ignore */ } },
  };
  const opts = Object.assign({}, C.DEFAULTS, {
    model: 'claude-haiku-5-5',
    trAuto: false, trQuality: 'default', trReview: true, trBilingual: true, trAutoGlossary: true, trAppendGlossary: true,
  }, store.get('mhm-md-opts', {}));
  document.querySelectorAll('[data-opt]').forEach(el => {
    const k = el.dataset.opt;
    if (el.type === 'checkbox') el.checked = !!opts[k]; else el.value = opts[k];
    el.addEventListener('change', () => {
      opts[k] = el.type === 'checkbox' ? el.checked : el.value;
      store.set('mhm-md-opts', opts);
      if (k !== 'model' && !k.startsWith('tr') && files.length) run(); // re-convert with new settings
    });
  });
  // ---------- Claude backend: viewer's Claude account (web artifact) or API key (apps) ----------
  let sampleFn = null;
  try { if (window.claude && window.claude.use) window.claude.use('sample').then(f => { sampleFn = f; syncKey(); }).catch(() => {}); } catch (e) { /* not an artifact */ }
  const MODEL_BY_TIER = { quick: 'claude-haiku-5-5', default: 'claude-sonnet-5-5', complex: 'claude-opus-5-5' };
  const hasAI = () => !!sampleFn || !!apiKeyEl.value.trim();

  async function llm({ prompt, json, tier = 'default', maxTokens = 8000, signal, model }) {
    if (sampleFn) {
      try {
        const o = { modelTier: tier, cache: false };
        if (signal) o.signal = signal;
        if (json) return await sampleFn.json(prompt, o);
        const r = await sampleFn(prompt, o);
        return r.text;
      } catch (e) {
        const err = new Error(sampleMsg(e)); err.code = e && e.code; throw err;
      }
    }
    const key = apiKeyEl.value.trim();
    if (!key) { const e = new Error('أضف مفتاح API في قسم «🤖 Claude» من الإعدادات'); e.code = 'no_key'; throw e; }
    let res;
    try {
      res = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST', signal,
        headers: {
          'content-type': 'application/json', 'x-api-key': key, 'anthropic-version': '2023-06-01',
          'anthropic-dangerous-direct-browser-access': 'true',
        },
        body: JSON.stringify({ model: model || MODEL_BY_TIER[tier] || MODEL_BY_TIER.default, max_tokens: maxTokens, messages: [{ role: 'user', content: prompt }] }),
      });
    } catch (e) {
      const err = new Error(e && e.name === 'AbortError' ? 'أُوقف' : 'لا يمكن الاتصال بـ Claude API — تحقق من الإنترنت');
      err.code = e && e.name === 'AbortError' ? 'cancelled' : 'network'; throw err;
    }
    const j = await res.json().catch(() => ({}));
    if (!res.ok) {
      const m = (j.error && j.error.message) || ('HTTP ' + res.status);
      const err = new Error(res.status === 401 ? 'مفتاح API غير صحيح' : res.status === 429 ? 'تجاوزت حد الاستخدام — انتظر قليلاً' : 'Claude API: ' + m);
      err.code = 'api_' + res.status; throw err;
    }
    llm.usage += ((j.usage && (j.usage.input_tokens + j.usage.output_tokens)) || 0);
    return (j.content || []).filter(c => c.type === 'text').map(c => c.text).join('');
  }
  llm.usage = 0;
  function sampleMsg(e) {
    const c = e && e.code;
    if (c === 'cancelled') return 'أُوقف';
    if (c === 'not_granted' || c === 'sampling_disabled') return 'لم يُسمح باستخدام Claude في هذه الصفحة';
    if (c === 'rate_limited') return 'تجاوزت حد الاستخدام في Claude — حاول لاحقاً';
    if (c === 'prompt_too_large') return 'الجزء كبير جداً';
    return 'Claude: ' + ((e && e.message) || c || 'خطأ');
  }

  const apiKeyEl = $('#apiKey');
  apiKeyEl.value = store.get('mhm-md-key', '');
  const syncKey = () => {
    store.set('mhm-md-key', apiKeyEl.value.trim());
    document.body.classList.toggle('has-key', hasAI()); // shows the ✨/🌐 buttons without re-rendering cards
  };
  apiKeyEl.addEventListener('input', syncKey);
  syncKey();
  const glossaryEl = $('#glossary');
  glossaryEl.value = store.get('mhm-md-glossary', '');
  glossaryEl.addEventListener('input', () => store.set('mhm-md-glossary', glossaryEl.value));

  // ---------- state ----------
  let files = [];    // {name, data:Uint8Array}
  let results = [];  // converter results

  // ---------- input ----------
  const drop = $('#drop');
  $('#file').addEventListener('change', e => { addFiles(Array.from(e.target.files)); e.target.value = ''; });
  if (isNative) $('#pickFolder').hidden = true;
  $('#pickFolder').addEventListener('click', () => $('#folder').click());
  $('#folder').addEventListener('change', e => {
    const list = Array.from(e.target.files).filter(f => C.SUPPORTED.includes(C.extOf(f.name)) && !/(^|\/)(node_modules|\.git)\//.test(f.webkitRelativePath || ''));
    if (!list.length) toast('لا توجد ملفات مدعومة في هذا المجلد');
    addFiles(list); e.target.value = '';
  });
  ['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', e => addFiles(Array.from(e.dataTransfer.files || [])));
  document.addEventListener('paste', e => {
    if (document.activeElement && /TEXTAREA|INPUT/.test(document.activeElement.tagName)) return;
    const t = e.clipboardData && e.clipboardData.getData('text');
    if (t && t.trim().length > 20) addFiles([new File([t], 'نص-ملصق.txt', { type: 'text/plain' })]);
  });

  async function addFiles(list) {
    if (!list.length) return;
    for (const f of list) files.push({ name: f.name, data: new Uint8Array(await f.arrayBuffer()) });
    run();
  }
  function addRaw(list) { // [{name, data:Uint8Array}] from Android share / Windows right-click
    if (!list || !list.length) return;
    files = files.concat(list);
    run();
  }

  // Android: files shared to the app / "Open with"
  if (IncomingFile) {
    const take = async () => {
      try {
        const r = await IncomingFile.takePending();
        const got = ((r && r.files) || []).map(f => ({ name: f.name, data: b64ToBytes(f.data) }));
        if (got.length) addRaw(got);
      } catch (e) { /* none */ }
    };
    take();
    try { IncomingFile.addListener('filesReceived', take); } catch (e) { /* older bridge */ }
  }
  // Windows: files opened from the right-click menu
  if (Desktop && Desktop.onFiles) Desktop.onFiles(list => addRaw(list.map(f => ({ name: f.name, data: new Uint8Array(f.data) }))));

  // ---------- run ----------
  let runId = 0;
  async function run() {
    const my = ++runId;
    results = [];
    $('#results').innerHTML = '';
    let failed = 0;
    for (let i = 0; i < files.length; i++) {
      if (my !== runId) return;
      setStatus(`⏳ جاري تحليل ${files[i].name} (${i + 1}/${files.length})…`);
      try {
        const r = await C.convert({ name: files[i].name, data: files[i].data.slice() }, opts);
        if (my !== runId) return;
        results.push(r);
        renderCard(results.length - 1);
      } catch (err) {
        if (my !== runId) return;
        failed++;
        renderError(files[i].name, err);
      }
    }
    setStatus(`✅ تم تحويل ${results.length} من ${files.length} ملف` + (failed ? ` — تعذّر ${failed}` : ''));
    renderBulk();
    if (opts.trAuto && hasAI()) {
      for (let i = 0; i < results.length; i++) {
        if (my !== runId) return;
        if (isArabicDoc(results[i]) && !results[i].translated) await translate(i);
      }
    }
    console.log('MHM_DONE ' + JSON.stringify(results.map(r => ({ name: r.name, tokens: r.stats.mdTokens, ocr: (r.meta.ocrPages || []).length, head: r.body.slice(0, 80) }))));
  }

  function setStatus(t) { const st = $('#status'); st.hidden = !t; st.textContent = t || ''; }
  const fmt = n => Math.round(n).toLocaleString('en-US');

  // ---------- cards ----------
  function renderAll() {
    $('#results').querySelectorAll('.card:not(.err)').forEach(c => c.remove());
    results.forEach((r, i) => renderCard(i));
    renderBulk();
  }

  function renderCard(idx) {
    const r = results[idx];
    const s = r.stats;
    const card = document.createElement('div');
    card.className = 'card';
    card.dataset.idx = idx;
    let save = '';
    if (s.pdfDirectTokens) {
      const pct = Math.max(0, Math.round((1 - s.mdTokens / s.pdfDirectTokens) * 100));
      save = `<span class="save">توفير تقريبي ${pct}% مقارنة برفع الـPDF (~${fmt(s.pdfDirectTokens)})</span>`;
    }
    const quality = r.quality.concat(r.trQuality || []);
    const warns = quality.filter(q => q.type === 'warn').length;
    const qItems = quality.map(q => `<li class="q-${q.type}">${q.type === 'warn' ? '⚠️' : q.type === 'info' ? 'ℹ️' : '✅'} ${esc(q.msg)}</li>`).join('');
    card.innerHTML = `
      <div class="name"><span class="badge">${esc(r.meta.format)}</span><bdi dir="ltr">${esc(r.mdName)}</bdi>${r.improved ? '<span class="badge ai">✨ مُحسّن</span>' : ''}${r.translated ? '<span class="badge tr">🌐 EN</span>' : ''}</div>
      <div class="meta">
        ${r.meta.pages ? `<span>${r.meta.pages} ${r.meta.format === 'PPTX' ? 'شريحة' : 'صفحة'}</span>` : ''}
        ${r.meta.sheets ? `<span>${r.meta.sheets} ورقة</span>` : ''}
        <span>${esc(r.meta.language)}</span>
        <span>≈ ${fmt(s.mdTokens)} توكن</span>
        ${save}
      </div>
      ${r.trState ? `<div class="tr-progress"><span class="tr-msg">⏳ ${esc(r.trState.msg)}</span><button data-a="stop">⏹ إيقاف</button></div>` : ''}
      <details class="quality"${warns ? ' open' : ''}><summary>تقرير الجودة${warns ? ` — ${warns} تنبيه` : ' — سليم'}</summary><ul>${qItems}</ul></details>
      <div class="btns">
        <button data-a="view">👁 معاينة وتعديل</button>
        <button data-a="copy">📋 نسخ</button>
        <button data-a="save" class="primary">⬇ حفظ</button>
        ${canShare() ? '<button data-a="share">↗ مشاركة</button>' : ''}
        ${isArabicDoc(r) || r.translated ? `<button data-a="tr" class="ai-btn"${r.trState ? ' disabled' : ''}>${r.translated ? '↩ الأصل العربي' : '🌐 ترجمة للإنجليزية'}</button>` : ''}
        ${r.translated && r.trPairs ? '<button data-a="cmp">↔ مقارنة</button>' : ''}
        <button data-a="ai" class="ai-btn"${r.trState ? ' disabled' : ''}>${r.improved ? '↩ إلغاء التحسين' : '✨ تحسين بـ Claude'}</button>
      </div>`;
    card.addEventListener('click', e => {
      const b = e.target.closest('button');
      if (!b) return;
      const cur = results[idx];
      const a = b.dataset.a;
      if (a === 'view') openPreview({ title: cur.mdName, get: () => cur.markdown, set: v => { cur.markdown = v; cur.stats.mdTokens = C.estTokens(v); } });
      if (a === 'copy') copy(cur.markdown);
      if (a === 'save') saveFile(cur.mdName, cur.markdown);
      if (a === 'share') shareFile(cur.mdName, cur.markdown);
      if (a === 'ai') { if (cur.improved) undoImprove(idx); else improve(idx, b); }
      if (a === 'tr') { if (cur.translated) undoTranslate(idx); else translate(idx); }
      if (a === 'stop' && cur.trState && cur.trState.ctl) cur.trState.ctl.abort();
      if (a === 'cmp') openCompare(cur);
    });
    const old = $('#results').querySelector(`.card[data-idx="${idx}"]`);
    if (old) old.replaceWith(card); else $('#results').appendChild(card);
  }

  function renderError(name, err) {
    const card = document.createElement('div');
    card.className = 'card err';
    card.innerHTML = `<div class="name">${esc(name)}</div><div class="meta">⚠️ ${esc(err && err.message ? err.message : String(err))}</div>`;
    $('#results').appendChild(card);
  }

  function renderBulk() {
    const b = $('#bulk');
    b.hidden = results.length < 1;
    const total = results.reduce((a, r) => a + r.stats.mdTokens, 0);
    $('#bulkInfo').textContent = (results.length > 1 ? `${results.length} ملفات · ` : '') + `≈ ${fmt(total)} توكن`;
    const multi = results.length > 1;
    b.querySelector('[data-bulk="save"]').hidden = !multi;
    b.querySelector('[data-bulk="zip"]').hidden = !multi;
    b.querySelector('[data-bulk="copy"]').hidden = !multi;
  }

  $('#bulk').addEventListener('click', async e => {
    const btn = e.target.closest('button');
    const a = btn && btn.dataset.bulk;
    if (!a) return;
    if (a === 'clear') { files = []; results = []; $('#results').innerHTML = ''; setStatus(''); renderBulk(); return; }
    if (a === 'project') {
      const p = C.buildProject(results, opts);
      const zip = new window.JSZip();
      p.files.forEach(f => zip.file(f.path, f.content));
      const name = (results.length === 1 ? results[0].name.replace(/\.[^.]+$/, '') : 'project') + '-claude-code.zip';
      await saveBinary(name, await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
      toast(`الحزمة: CLAUDE.md + ${p.files.length - 1} ملف داخل docs/ (~${fmt(p.totalTokens)} توكن)`);
      return;
    }
    if (a === 'translate') {
      if (!hasAI()) { toast('أضف مفتاح API في قسم «🤖 Claude» من الإعدادات'); return; }
      const todo = results.map((r, i) => i).filter(i => isArabicDoc(results[i]) && !results[i].translated && !results[i].trState);
      if (!todo.length) { toast('لا توجد ملفات عربية تحتاج ترجمة'); return; }
      for (const i of todo) await translate(i);
      return;
    }
    if (a === 'claudemd') {
      let md = C.buildProject(results, opts).claudeMd;
      openPreview({ title: 'CLAUDE.md', get: () => md, set: v => { md = v; } });
      return;
    }
    if (a === 'zip') {
      const zip = new window.JSZip();
      const used = {};
      results.forEach(r => { let n = r.mdName; if (used[n]) n = n.replace(/\.md$/, `-${++used[n]}.md`); else used[n] = 1; zip.file(n, r.markdown); });
      await saveBinary('converted-md.zip', await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
      return;
    }
    const merged = C.combine(results, opts);
    if (merged.removed) toast(`حُذفت ${merged.removed} فقرة مكررة بين الملفات`);
    if (a === 'copy') copy(merged.markdown);
    if (a === 'save') saveFile('combined-for-claude.md', merged.markdown);
  });

  // ---------- preview ----------
  const dlg = $('#preview');
  let pv = null;
  function openPreview(target) {
    pv = target;
    $('#pvTitle').textContent = target.title;
    $('#pvText').value = target.get();
    updatePvStats();
    if (dlg.showModal) dlg.showModal(); else dlg.setAttribute('open', '');
  }
  function updatePvStats() { $('#pvStats').textContent = `≈ ${fmt(C.estTokens($('#pvText').value))} توكن`; }
  $('#pvText').addEventListener('input', () => { if (pv) pv.set($('#pvText').value); updatePvStats(); });
  $('#pvClose').addEventListener('click', () => { if (dlg.close) dlg.close(); else dlg.removeAttribute('open'); renderAll(); });
  $('#pvCopy').addEventListener('click', () => copy($('#pvText').value));
  $('#pvSave').addEventListener('click', () => saveFile(pv.title.endsWith('.md') ? pv.title : pv.title + '.md', $('#pvText').value));

  // ---------- ✨ improve with Claude API ----------
  const AI_SYSTEM = [
    'You restructure a document (already converted to Markdown) so that Claude and Claude Code understand it perfectly with as few tokens as possible.',
    'Rules:',
    '- Keep EVERY fact, number, name, requirement, constraint, URL and table value. Never invent or assume anything.',
    '- Keep the document\'s original language(s).',
    '- Structure: "# Title", a 2-4 sentence overview, then logical "##" sections (e.g. goals, scope, requirements, data, flows, rules, phases).',
    '- Turn requirements and rules into short bullet points; turn acceptance criteria and tasks into "- [ ]" checklists; keep tables (CSV in a ```csv block when wide).',
    '- Remove filler, repetition, marketing language and boilerplate.',
    '- Mark anything ambiguous or contradictory with "> Unclear:" instead of guessing.',
    '- Output ONLY the Markdown, with no preface.',
  ].join('\n');

  async function improve(idx, btn) {
    const r = results[idx];
    if (!hasAI()) { toast('أضف مفتاح API في قسم «🤖 Claude» من الإعدادات'); return; }
    if (r.stats.mdTokens > 60000) { toast('المستند كبير جداً للتحسين دفعة واحدة — استخدم حزمة Claude Code'); return; }
    btn.disabled = true; btn.textContent = '⏳ Claude يعمل…';
    try {
      const before = llm.usage;
      const text = await llm({
        prompt: AI_SYSTEM + `\n\nFile: ${r.name}\n\n<<<DOCUMENT\n${r.markdown}\nDOCUMENT>>>`,
        tier: 'default', model: opts.model,
        maxTokens: Math.min(32000, Math.max(4000, Math.round(r.stats.mdTokens * 1.3))),
      });
      const out = String(text || '').trim();
      if (!out) throw new Error('رد فارغ');
      r.original = { markdown: r.markdown, body: r.body, tokens: r.stats.mdTokens };
      r.markdown = out.replace(/^```(?:markdown|md)?\n([\s\S]*)\n```$/, '$1') + '\n';
      r.body = r.markdown;
      r.stats.mdTokens = C.estTokens(r.markdown);
      r.improved = true;
      const used = llm.usage - before;
      toast('تم التحسين ✓' + (used ? ` (استُهلك ${fmt(used)} توكن API)` : ''));
    } catch (err) {
      toast('فشل التحسين: ' + (err.message || err));
    }
    renderCard(idx); renderBulk();
  }

  // ---------- 🌐 translate Arabic → English ----------
  function isArabicDoc(r) { return r && /Arabic/.test((r.translated ? r.trOriginal.meta.language : r.meta.language) || ''); }

  async function translate(idx) {
    const r = results[idx];
    if (!hasAI()) { toast('الترجمة تحتاج مفتاح API — أضفه في قسم «🤖 Claude» من الإعدادات'); $('#aiOpts').open = true; return; }
    if (r.improved) undoImprove(idx);
    const ctl = new AbortController();
    r.trState = { msg: 'بدء الترجمة…', ctl };
    renderCard(idx);
    const setMsg = m => {
      r.trState.msg = m;
      const el = $('#results').querySelector(`.card[data-idx="${idx}"] .tr-msg`);
      if (el) el.textContent = '⏳ ' + m; else renderCard(idx);
    };
    const before = llm.usage;
    try {
      const tr = await C.translate.translateResult(r, llm, {
        glossary: C.translate.parseGlossary(glossaryEl.value),
        autoGlossary: opts.trAutoGlossary, bilingualRules: opts.trBilingual, review: opts.trReview,
        appendGlossary: opts.trAppendGlossary, tier: opts.trQuality || 'default',
        signal: ctl.signal, onProgress: setMsg,
      }, opts);
      r.trOriginal = tr.original;
      r.markdown = tr.markdown; r.body = tr.body; r.meta = tr.meta; r.mdName = tr.mdName;
      r.stats.mdTokens = tr.tokens;
      r.trPairs = tr.pairs;
      r.translated = true;
      const rep = tr.report;
      const q = [];
      const realIssues = rep.issues.filter(i => i.kind !== 'review');
      q.push({ type: 'ok', msg: `الترجمة: ${rep.chunks} جزء · قاموس ${rep.glossary.length} مصطلح${rep.retried ? ` · أُعيدت ترجمة ${rep.retried} جزء تلقائياً` : ''}${rep.reviewed ? ` · المراجعة صحّحت ${rep.fixes} موضع` : ''}` });
      if (!realIssues.length) q.push({ type: 'ok', msg: 'فحص الترجمة: كل الأرقام والكود والروابط والعناوين والقوائم والجداول مطابقة للأصل' });
      const AR_KIND = { numbers: 'أرقام', code: 'كود/روابط', structure: 'البنية', negation: 'النفي', untranslated: 'نص غير مترجم', omission: 'نقص محتمل', glossary: 'القاموس' };
      realIssues.forEach(i => q.push({ type: 'warn', msg: `ترجمة — الجزء ${i.part}، ${AR_KIND[i.kind] || i.kind}: ${i.msg}` }));
      rep.issues.filter(i => i.kind === 'review').forEach(i => q.push({ type: 'info', msg: `ملاحظة المراجعة — الجزء ${i.part}: ${i.msg}` }));
      r.trQuality = q;
      const used = llm.usage - before;
      toast(realIssues.length ? `تمت الترجمة مع ${realIssues.length} تنبيه — راجع «مقارنة»` : 'تمت الترجمة ✓' + (used ? ` (استُهلك ${fmt(used)} توكن API)` : ''));
    } catch (err) {
      toast(err && err.code === 'cancelled' ? 'أُوقفت الترجمة' : 'فشلت الترجمة: ' + ((err && err.message) || err));
    }
    r.trState = null;
    renderCard(idx); renderBulk();
  }

  function undoTranslate(idx) {
    const r = results[idx];
    if (!r.trOriginal) return;
    const o = r.trOriginal;
    r.markdown = o.markdown; r.body = o.body; r.meta = o.meta; r.mdName = o.mdName; r.stats.mdTokens = o.tokens;
    r.translated = false; r.trOriginal = null; r.trPairs = null; r.trQuality = null;
    renderCard(idx); renderBulk();
  }

  // side-by-side review of every translated part
  const cmp = $('#compare');
  function openCompare(r) {
    const pairs = r.trPairs || [];
    const probs = pairs.reduce((a, p) => a + p.problems.filter(x => x.kind !== 'review').length, 0);
    $('#cmpTitle').textContent = 'مقارنة: ' + r.trOriginal.mdName + ' ↔ ' + r.mdName;
    $('#cmpStats').textContent = `${pairs.length} جزء · ${probs ? probs + ' تنبيه' : 'بدون تنبيهات'}`;
    $('#cmpBody').innerHTML = pairs.map((p, i) => `
      <div class="cmp-pair">
        <div class="head"><span>الجزء ${i + 1}</span><span>${p.problems.length ? '⚠️ ' + p.problems.length : '✅'}</span></div>
        ${p.problems.length ? `<ul class="probs">${p.problems.map(x => `<li>${esc(x.msg)}</li>`).join('')}</ul>` : ''}
        <div class="cols"><pre class="ar" dir="rtl">${esc(p.src)}</pre><pre dir="ltr">${esc(p.out)}</pre></div>
      </div>`).join('');
    if (cmp.showModal) cmp.showModal(); else cmp.setAttribute('open', '');
  }
  $('#cmpClose').addEventListener('click', () => { if (cmp.close) cmp.close(); else cmp.removeAttribute('open'); });

  function undoImprove(idx) {
    const r = results[idx];
    if (!r.original) return;
    r.markdown = r.original.markdown; r.body = r.original.body; r.stats.mdTokens = r.original.tokens;
    r.improved = false; r.original = null;
    renderCard(idx); renderBulk();
  }

  // ---------- output ----------
  async function copy(text) {
    try { await navigator.clipboard.writeText(text); toast('تم النسخ ✓ — الصقه في Claude'); return; } catch (e) { /* fallback */ }
    const ta = document.createElement('textarea');
    ta.value = text; ta.style.position = 'fixed'; ta.style.opacity = '0';
    document.body.appendChild(ta); ta.select();
    try { document.execCommand('copy'); toast('تم النسخ ✓'); } catch (e) { toast('تعذّر النسخ'); }
    ta.remove();
  }

  function canShare() {
    if (isNative) return !!(Share && Filesystem);
    try { return !!(navigator.canShare && navigator.canShare({ files: [new File(['x'], 'x.md', { type: 'text/markdown' })] })); } catch (e) { return false; }
  }

  const saveFile = (name, text) => saveAny(name, text, false);
  const saveBinary = (name, bytes) => saveAny(name, bytes, true);

  async function saveAny(name, data, binary) {
    if (isNative && Filesystem) {
      try {
        await Filesystem.writeFile(Object.assign({ path: 'MHM MD/' + name, directory: 'DOCUMENTS', recursive: true },
          binary ? { data: bytesToB64(data) } : { data, encoding: 'utf8' }));
        toast('تم الحفظ في: Documents/MHM MD/' + name);
        return;
      } catch (e) {
        return shareAny(name, data, binary); // fall back to the share sheet (Files / Drive / Claude)
      }
    }
    if (artifactDl) {
      try { await artifactDl.save({ filename: name, data: binary ? new Blob([data]) : data }); toast('تم حفظ ' + name); }
      catch (e) { toast(e && e.code === 'declined' ? 'أُلغي الحفظ' : 'لم يتم الحفظ — استخدم زر النسخ'); }
      return;
    }
    const blob = new Blob([data], { type: binary ? 'application/zip' : 'text/markdown;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
    toast('تم تنزيل ' + name);
  }

  const shareFile = (name, text) => shareAny(name, text, false);
  async function shareAny(name, data, binary) {
    try {
      if (isNative && Filesystem && Share) {
        const w = await Filesystem.writeFile(Object.assign({ path: name, directory: 'CACHE' },
          binary ? { data: bytesToB64(data) } : { data, encoding: 'utf8' }));
        await Share.share({ title: name, files: [w.uri], dialogTitle: 'أرسل الملف إلى Claude' });
        return;
      }
      const file = new File([data], name, { type: binary ? 'application/zip' : 'text/markdown' });
      await navigator.share({ files: [file], title: name });
    } catch (e) {
      if (e && e.name !== 'AbortError' && !/cancel/i.test(String(e.message || e))) toast('تعذّرت المشاركة');
    }
  }

  function bytesToB64(u8) {
    let s = '';
    for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000));
    return btoa(s);
  }
  function b64ToBytes(b64) {
    const bin = atob(b64);
    const u8 = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
    return u8;
  }

  let tt;
  function toast(msg) {
    const t = $('#toast'); t.textContent = msg; t.hidden = false;
    clearTimeout(tt); tt = setTimeout(() => { t.hidden = true; }, 3200);
  }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }

  // hook for automated browser tests
  window.__mhm = { get results() { return results; }, buildProject: () => C.buildProject(results, opts), opts };
})();
