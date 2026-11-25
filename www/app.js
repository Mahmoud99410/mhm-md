(function () {
  'use strict';
  const $ = s => document.querySelector(s);

  // ---- libraries ----
  if (window.pdfjsLib) window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'lib/pdf.worker.min.js';
  MHMConverter.init({ pdfjsLib: window.pdfjsLib, mammoth: window.mammoth, TurndownService: window.TurndownService });

  const Cap = window.Capacitor;
  const isNative = !!(Cap && Cap.isNativePlatform && Cap.isNativePlatform());
  const Plugins = (Cap && Cap.Plugins) || {};

  // when hosted as a Claude artifact, downloads go through the page's downloads capability
  let artifactDl = null;
  try { if (window.claude && window.claude.use) window.claude.use('downloads').then(d => { artifactDl = d; }).catch(() => {}); } catch (e) { /* not an artifact */ }

  // ---- options (persisted) ----
  const store = {
    get() { try { return JSON.parse(localStorage.getItem('mhm-md-opts') || '{}'); } catch (e) { return {}; } },
    set(v) { try { localStorage.setItem('mhm-md-opts', JSON.stringify(v)); } catch (e) { /* ignore */ } },
  };
  const opts = Object.assign({}, MHMConverter.DEFAULTS, store.get());
  document.querySelectorAll('[data-opt]').forEach(el => {
    const k = el.dataset.opt;
    if (el.type === 'checkbox') el.checked = !!opts[k]; else el.value = opts[k];
    el.addEventListener('change', () => {
      opts[k] = el.type === 'checkbox' ? el.checked : el.value;
      store.set(opts);
      if (lastFiles.length) run(lastFiles); // re-convert with new settings
    });
  });

  // ---- state ----
  let results = [];
  let lastFiles = [];

  // ---- input ----
  const drop = $('#drop');
  $('#file').addEventListener('change', e => { if (e.target.files.length) handle(Array.from(e.target.files)); e.target.value = ''; });
  ['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', e => { const f = Array.from(e.dataTransfer.files || []); if (f.length) handle(f); });
  // paste text directly
  document.addEventListener('paste', e => {
    if (document.activeElement && /TEXTAREA|INPUT/.test(document.activeElement.tagName)) return;
    const t = e.clipboardData && e.clipboardData.getData('text');
    if (t && t.trim().length > 20) handle([new File([t], 'نص-ملصق.txt', { type: 'text/plain' })]);
  });

  async function handle(files) {
    const loaded = [];
    for (const f of files) loaded.push({ name: f.name, data: new Uint8Array(await f.arrayBuffer()) });
    lastFiles = lastFiles.concat(loaded);
    run(lastFiles);
  }

  let runId = 0;
  async function run(files) {
    const my = ++runId;
    results = [];
    $('#results').innerHTML = '';
    const st = $('#status');
    st.hidden = false;
    for (let i = 0; i < files.length; i++) {
      if (my !== runId) return;
      st.textContent = `⏳ جاري تحليل ${files[i].name} (${i + 1}/${files.length})…`;
      try {
        // copy so pdf.js can transfer the buffer without detaching our original
        const r = await MHMConverter.convert({ name: files[i].name, data: files[i].data.slice() }, opts);
        if (my !== runId) return;
        results.push(r);
        renderCard(r, results.length - 1);
      } catch (err) {
        if (my !== runId) return;
        renderError(files[i].name, err);
      }
    }
    st.textContent = `✅ تم تحويل ${results.length} من ${files.length} ملف`;
    renderBulk();
  }

  const fmt = n => n.toLocaleString('en-US');

  function renderCard(r, idx) {
    const s = r.stats;
    const card = document.createElement('div');
    card.className = 'card';
    let save = '';
    if (s.pdfDirectTokens) {
      const pct = Math.max(0, Math.round((1 - s.mdTokens / s.pdfDirectTokens) * 100));
      save = `<span class="save">توفير تقريبي ${pct}% مقارنة برفع الـPDF (~${fmt(s.pdfDirectTokens)})</span>`;
    }
    card.innerHTML = `
      <div class="name"><span class="badge">${esc(r.meta.format)}</span><bdi dir="ltr">${esc(r.mdName)}</bdi></div>
      <div class="meta">
        ${r.meta.pages ? `<span>${r.meta.pages} صفحة</span>` : ''}
        <span>${esc(r.meta.language)}</span>
        <span>≈ ${fmt(s.mdTokens)} توكن</span>
        ${save}
      </div>
      <div class="btns">
        <button data-a="view">👁 معاينة وتعديل</button>
        <button data-a="copy">📋 نسخ</button>
        <button data-a="save" class="primary">⬇ حفظ</button>
        ${canShare() ? '<button data-a="share">↗ مشاركة</button>' : ''}
      </div>`;
    card.addEventListener('click', e => {
      const a = e.target.closest('button') && e.target.closest('button').dataset.a;
      if (!a) return;
      const cur = results[idx];
      if (a === 'view') openPreview(idx);
      if (a === 'copy') copy(cur.markdown);
      if (a === 'save') saveFile(cur.mdName, cur.markdown);
      if (a === 'share') shareFile(cur.mdName, cur.markdown);
    });
    $('#results').appendChild(card);
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
    $('#bulkInfo').textContent = results.length > 1 ? `${results.length} ملفات · ≈ ${fmt(total)} توكن` : `≈ ${fmt(total)} توكن`;
    b.querySelector('[data-bulk="copy"]').hidden = results.length < 2;
    b.querySelector('[data-bulk="save"]').hidden = results.length < 2;
  }
  $('#bulk').addEventListener('click', e => {
    const a = e.target.closest('button') && e.target.closest('button').dataset.bulk;
    if (!a) return;
    if (a === 'clear') { lastFiles = []; results = []; $('#results').innerHTML = ''; $('#status').hidden = true; renderBulk(); return; }
    const md = MHMConverter.combine(results, opts);
    if (a === 'copy') copy(md);
    if (a === 'save') saveFile('combined-for-claude.md', md);
  });

  // ---- preview ----
  const dlg = $('#preview');
  let pvIdx = -1;
  function openPreview(idx) {
    pvIdx = idx;
    const r = results[idx];
    $('#pvTitle').textContent = r.mdName;
    $('#pvText').value = r.markdown;
    updatePvStats();
    if (dlg.showModal) dlg.showModal(); else dlg.setAttribute('open', '');
  }
  function updatePvStats() { $('#pvStats').textContent = `≈ ${fmt(MHMConverter.estTokens($('#pvText').value))} توكن`; }
  $('#pvText').addEventListener('input', () => {
    if (pvIdx >= 0) { results[pvIdx].markdown = $('#pvText').value; results[pvIdx].stats.mdTokens = MHMConverter.estTokens(results[pvIdx].markdown); }
    updatePvStats();
  });
  $('#pvClose').addEventListener('click', () => dlg.close ? dlg.close() : dlg.removeAttribute('open'));
  $('#pvCopy').addEventListener('click', () => copy($('#pvText').value));
  $('#pvSave').addEventListener('click', () => saveFile(results[pvIdx].mdName, $('#pvText').value));

  // ---- output helpers ----
  async function copy(text) {
    try { await navigator.clipboard.writeText(text); toast('تم النسخ ✓ — الصقه في Claude'); return; } catch (e) { /* fallback */ }
    const ta = document.createElement('textarea');
    ta.value = text; ta.style.position = 'fixed'; ta.style.opacity = '0';
    document.body.appendChild(ta); ta.select();
    try { document.execCommand('copy'); toast('تم النسخ ✓'); } catch (e) { toast('تعذّر النسخ'); }
    ta.remove();
  }

  function canShare() {
    if (isNative) return !!Plugins.Share;
    try { return !!(navigator.canShare && navigator.canShare({ files: [new File(['x'], 'x.md', { type: 'text/markdown' })] })); } catch (e) { return false; }
  }

  async function saveFile(name, text) {
    if (isNative && Plugins.Filesystem) {
      try {
        await Plugins.Filesystem.writeFile({ path: 'MHM MD/' + name, data: text, directory: 'DOCUMENTS', encoding: 'utf8', recursive: true });
        toast('تم الحفظ في: Documents/MHM MD/' + name);
        return;
      } catch (e) {
        return shareFile(name, text); // fall back to the share sheet (save to Files / Drive / Claude)
      }
    }
    if (artifactDl) {
      try { await artifactDl.save({ filename: name, data: text }); toast('تم حفظ ' + name); }
      catch (e) { toast('لم يتم الحفظ — استخدم زر النسخ'); }
      return;
    }
    const blob = new Blob([text], { type: 'text/markdown;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
    toast('تم تنزيل ' + name);
  }

  async function shareFile(name, text) {
    try {
      if (isNative && Plugins.Filesystem && Plugins.Share) {
        const w = await Plugins.Filesystem.writeFile({ path: name, data: text, directory: 'CACHE', encoding: 'utf8' });
        await Plugins.Share.share({ title: name, files: [w.uri], dialogTitle: 'أرسل الملف إلى Claude' });
        return;
      }
      const file = new File([text], name, { type: 'text/markdown' });
      await navigator.share({ files: [file], title: name });
    } catch (e) {
      if (e && e.name !== 'AbortError' && !/cancel/i.test(String(e.message || e))) toast('تعذّرت المشاركة');
    }
  }

  let tt;
  function toast(msg) {
    const t = $('#toast'); t.textContent = msg; t.hidden = false;
    clearTimeout(tt); tt = setTimeout(() => { t.hidden = true; }, 2600);
  }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
})();
