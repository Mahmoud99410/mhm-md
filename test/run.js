// Converts the sample files and checks the structure Claude relies on.
const fs = require('fs');
const path = require('path');
const C = require('../www/converter.js');
C.init({
  pdfjsLib: require('pdfjs-dist/legacy/build/pdf.js'),
  mammoth: require('mammoth'),
  TurndownService: require('turndown'),
});
const expect = {
  'rtl-chrome.pdf': ['# دليل تشغيل نظام التوصيل', '- لا تستخدم مكتبات', '2. شغ', '| تاكسي | 1.5 دينار | 15 دقيقة |'],
  'english.pdf': ['# Chapter 1: System Rules', '## Overview', '- Never commit secrets', '```\nnpm install'],
  'sample.docx': ['## أهم', '| توصيل | 2 JD | 30 دقيقة |', '1. حمّل المستودع.'].slice(1),
  'sample.txt': ['## INSTRUCTIONS FOR THE AI', '### Rules', '- Always reply in Arabic.'],
};
(async () => {
  let fail = 0;
  for (const [name, needles] of Object.entries(expect)) {
    const data = fs.readFileSync(path.join(__dirname, name));
    const r = await C.convert({ name, data: name.endsWith('.docx') ? data : new Uint8Array(data) });
    for (const n of needles) {
      if (!r.markdown.includes(n)) { fail++; console.error(`✗ ${name}: missing "${n}"`); }
    }
    console.log(`✓ ${name} → ${r.stats.mdTokens} tokens`);
  }
  if (fail) { console.error(`${fail} check(s) failed`); process.exit(1); }
  console.log('all checks passed');
})().catch(e => { console.error(e); process.exit(1); });
