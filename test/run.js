// Automated checks for the converter engine (run by CI before building the apps)
const fs = require('fs');
const path = require('path');
const C = require('../www/converter.js');
C.init({
  pdfjsLib: require('pdfjs-dist/legacy/build/pdf.js'),
  mammoth: require('mammoth'),
  TurndownService: require('turndown'),
  JSZip: require('jszip'),
  XLSX: require('xlsx'),
});

let fail = 0;
const check = (cond, label) => { if (cond) console.log('  ✓ ' + label); else { fail++; console.error('  ✗ ' + label); } };
const load = async (name, opts) => C.convert({ name, data: new Uint8Array(fs.readFileSync(path.join(__dirname, name))) }, opts);

(async () => {
  console.log('Structure');
  const expect = {
    'rtl-chrome.pdf': ['# دليل تشغيل نظام التوصيل', '- لا تستخدم مكتبات', '2. شغ', '| تاكسي | 1.5 دينار | 15 دقيقة |'],
    'english.pdf': ['# Chapter 1: System Rules', '## Overview', '- Never commit secrets', '```\nnpm install'],
    'sample.docx': ['| توصيل | 2 JD | 30 دقيقة |', '1. حمّل المستودع.'],
    'sample.txt': ['## INSTRUCTIONS FOR THE AI', '### Rules', '- Always reply in Arabic.'],
    'deck.pptx': ['## 2. المميزات', '  - تتبع الكابتن', '> Notes: Mention the 60 second rule.', '| Taxi | 1.5 JD |'],
    'data.xlsx': ['## Captains', '```csv\nName,Phone,Car', '## Prices', '| Taxi | 1.5 |'],
  };
  const all = {};
  for (const [name, needles] of Object.entries(expect)) {
    const r = await load(name);
    all[name] = r;
    for (const n of needles) check(r.markdown.includes(n), `${name} contains ${JSON.stringify(n)}`);
  }
  check(!all['english.pdf'].markdown.includes('ACME Corp'), 'english.pdf: repeated header removed');

  console.log('Compression');
  const light = await load('sample.docx', { level: 'light' });
  const strong = await load('sample.docx', { level: 'strong' });
  check(light.markdown.includes('| --- |'), 'light keeps Markdown tables');
  check(strong.markdown.includes('```csv'), 'strong turns tables into CSV');
  check(strong.stats.mdTokens < light.stats.mdTokens, `strong uses fewer tokens (${strong.stats.mdTokens} < ${light.stats.mdTokens})`);
  const spec = await load('spec.docx');
  check(!/all rights reserved/i.test(spec.markdown), 'boilerplate copyright line removed');

  console.log('Quality report');
  const scan = await load('scan.pdf');
  check(scan.quality.some(q => q.type === 'warn' && /ممسوح/.test(q.msg)), 'scanned PDF flagged when OCR is unavailable');
  check(all['rtl-chrome.pdf'].quality[0].type === 'ok', 'structure summary present');

  console.log('Claude Code package');
  const p = C.buildProject([spec, all['sample.docx']], {});
  const paths = p.files.map(f => f.path);
  check(paths[0] === 'CLAUDE.md' && paths[1] === 'docs/INDEX.md', 'CLAUDE.md and docs/INDEX.md created');
  check(paths.filter(x => x.startsWith('docs/01-spec/')).length >= 6, 'long spec split into section files');
  check(/Node\.js/.test(p.claudeMd) && /PostgreSQL/.test(p.claudeMd) && /Flutter/.test(p.claudeMd), 'tech stack detected');
  check(/## Rules — must not[\s\S]*Never store card numbers/.test(p.claudeMd), 'must-not rules extracted');
  check(/## Phases[\s\S]*phase-1-mvp\.md/.test(p.claudeMd), 'phases linked');
  check(p.files.every(f => f.content.trim().length > 0), 'no empty files');

  console.log('Combine');
  const dup = C.combine([all['sample.docx'], all['sample.docx']], {});
  check(dup.removed > 0, `duplicate paragraphs across files removed (${dup.removed})`);

  if (fail) { console.error(`\n${fail} check(s) failed`); process.exit(1); }
  console.log('\nall checks passed');
})().catch(e => { console.error(e); process.exit(1); });
