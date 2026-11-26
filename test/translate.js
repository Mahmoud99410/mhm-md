// Translation engine checks with a mock LLM (no network)
const fs = require('fs');
const path = require('path');
const C = require('../www/converter.js');
C.init({ mammoth: require('mammoth'), TurndownService: require('turndown'), JSZip: require('jszip'), XLSX: require('xlsx'), pdfjsLib: require('pdfjs-dist/legacy/build/pdf.js') });
const T = C.translate;
let fail = 0;
const check = (cond, label) => { if (cond) console.log('  ✓ ' + label); else { fail++; console.error('  ✗ ' + label); } };

const NEG = new Set(['لا', 'ليس', 'لن', 'لم', 'ممنوع', 'يمنع', 'يُمنع', 'عدم', 'بدون', 'دون']);
// fake translator: Arabic words → "w"/"not", everything else (markers, numbers, markdown) kept
const fakeTranslate = s => s.replace(/[ء-يً-ْ]+/g, w => NEG.has(w) ? 'not' : 'w');
const between = (p, a, b) => p.slice(p.indexOf(a) + a.length + 1, p.lastIndexOf(b) - 1);

function mockLLM({ dropOnce = false } = {}) {
  let dropped = false;
  const log = [];
  const fn = async ({ prompt, json }) => {
    log.push(prompt.slice(0, 40));
    if (prompt.includes('<<<DOC')) return JSON.stringify([{ ar: 'الكابتن', en: 'captain' }, { ar: 'الطلب', en: 'order' }]);
    if (prompt.includes('<<<ORIGINAL')) {
      const tr = between(prompt, '<<<TRANSLATION', 'TRANSLATION>>>');
      const first = tr.split('\n').find(l => / w /.test(l));
      return { issues: first ? [{ wrong: first, fix: first + ' [fixed]', why: 'test' }] : [] };
    }
    const src = between(prompt, '<<<ARABIC_MARKDOWN', 'ARABIC_MARKDOWN>>>');
    let out = fakeTranslate(src);
    if (dropOnce && !dropped && /\d/.test(out)) { dropped = true; out = out.replace(/\d+/, '').replace(/⟦\d+⟧/, ''); }
    return out;
  };
  fn.log = log;
  return fn;
}

(async () => {
  const load = async (name) => C.convert({ name, data: new Uint8Array(fs.readFileSync(path.join(__dirname, name))) }, {});
  const r = await load('sample.docx');

  console.log('Glossary & protection');
  const g = T.parseGlossary('طلب = order\nالكابتن: captain\nnot a line\nالزبون -> customer');
  check(g.length === 3 && g[1].en === 'captain', 'glossary parser reads =, : and ->');
  check(T.mergeGlossary([{ ar: 'الطلب', en: 'booking' }], [{ ar: 'الطلب', en: 'order' }])[0].en === 'booking', "user's glossary wins over automatic");
  const p = T.protect('شغّل `npm install` ثم افتح https://x.io/a?b=1\n- يجب استخدام PostgreSQL\n```js\nconst a=1\n```', true);
  check(!p.text.includes('npm install') && !p.text.includes('https://') && !p.text.includes('const a'), 'code, URLs and code blocks are protected');
  check(p.rules.length === 1 && p.text.includes('⟦R0⟧'), 'rule line marked');
  const back = T.restore(fakeTranslate(p.text), p.vault, p.rules, true);
  check(back.includes('`npm install`') && back.includes('https://x.io/a?b=1') && back.includes('const a=1'), 'protected items restored exactly');
  check(/_\(AR: يجب استخدام PostgreSQL\)_/.test(back), 'rule keeps original Arabic beside the English');

  console.log('Verification');
  const v = T.verifyChunk('## عنوان\n- عنصر 30 دقيقة\n- ممنوع ⟦0⟧', '## title\n- item minutes');
  const kinds = v.map(x => x.kind);
  check(kinds.includes('numbers') && kinds.includes('structure') && kinds.includes('code') && kinds.includes('negation'), 'detects missing number, list item, protected item and negation');

  console.log('Full translation (mock)');
  const llm = mockLLM();
  const tr = await C.translate.translateResult(r, llm, { glossary: g, review: false }, {});
  check(!/[ء-ي]{4,}/.test(tr.body.replace(/_\(AR:[^)]*\)_/g, '').replace(/## Glossary[\s\S]*$/, '')), 'no Arabic left outside rule notes and glossary');
  check(tr.body.includes('| --- |') && /2 JD/.test(tr.body) && /30/.test(tr.body), 'tables and numbers preserved');
  check(tr.report.issues.filter(i => i.kind !== 'review').length === 0, 'no verification problems');
  check(/## Glossary \(Arabic → English\)/.test(tr.body), 'glossary appended');
  check(/translated_from: Arabic/.test(tr.markdown) && tr.mdName.endsWith('.en.md'), 'front matter + file name marked as translation');
  check(tr.original.markdown === r.markdown, 'original kept for undo');

  console.log('Retry on a bad chunk');
  const bad = mockLLM({ dropOnce: true });
  const tr2 = await C.translate.translateResult(r, bad, { glossary: g, autoGlossary: false }, {});
  check(tr2.report.retried >= 1, `chunk with a missing number/code was re-translated (${tr2.report.retried})`);
  check(tr2.report.issues.filter(i => i.kind === 'numbers' || i.kind === 'code').length === 0, 'retry result has no missing numbers or code');

  console.log('Review pass');
  const tr3 = await C.translate.translateResult(r, mockLLM(), { glossary: g, review: true }, {});
  check(tr3.report.reviewed >= 1 && tr3.report.fixes >= 1 && tr3.body.includes('[fixed]'), `review fixes applied (${tr3.report.fixes})`);

  console.log('Long document chunking');
  const spec = await load('spec.docx');
  const ar = spec.body.replace(/[A-Za-z]+/g, 'كلمة'); // make it "Arabic"
  const tr4 = await C.translate.translateBody(ar, mockLLM(), { chunkTokens: 800, autoGlossary: false });
  check(tr4.report.chunks > 3, `split into ${tr4.report.chunks} parts`);
  check((tr4.body.match(/^#{1,6}\s/gm) || []).length === (ar.match(/^#{1,6}\s/gm) || []).length, 'all headings kept across parts');

  if (fail) { console.error(`\n${fail} translation check(s) failed`); process.exit(1); }
  console.log('\nall translation checks passed');
})().catch(e => { console.error(e); process.exit(1); });
