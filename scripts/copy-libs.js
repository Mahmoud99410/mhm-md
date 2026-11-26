// Copies browser builds of the libraries into www/lib (runs after npm install)
const fs = require('fs');
const path = require('path');
const nm = p => path.join(__dirname, '..', 'node_modules', p);
const out = path.join(__dirname, '..', 'www', 'lib');
const files = {
  'pdf.min.js': 'pdfjs-dist/legacy/build/pdf.min.js',
  'pdf.worker.min.js': 'pdfjs-dist/legacy/build/pdf.worker.min.js',
  'mammoth.browser.min.js': 'mammoth/mammoth.browser.min.js',
  'turndown.js': 'turndown/dist/turndown.js',
  'jszip.min.js': 'jszip/dist/jszip.min.js',
  'xlsx.full.min.js': 'xlsx/dist/xlsx.full.min.js',
  'ocr/tesseract.min.js': 'tesseract.js/dist/tesseract.min.js',
  'ocr/worker.min.js': 'tesseract.js/dist/worker.min.js',
  'ocr/core/tesseract-core-lstm.wasm.js': 'tesseract.js-core/tesseract-core-lstm.wasm.js',
  'ocr/core/tesseract-core-simd-lstm.wasm.js': 'tesseract.js-core/tesseract-core-simd-lstm.wasm.js',
  'ocr/lang/ara.traineddata.gz': '@tesseract.js-data/ara/4.0.0_best_int/ara.traineddata.gz',
  'ocr/lang/eng.traineddata.gz': '@tesseract.js-data/eng/4.0.0_best_int/eng.traineddata.gz',
};
for (const [dst, src] of Object.entries(files)) {
  const to = path.join(out, dst);
  fs.mkdirSync(path.dirname(to), { recursive: true });
  fs.copyFileSync(nm(src), to);
}
// SheetJS ships raw control characters inside string literals; escape them so the file is plain UTF-8 text
{
  const f = path.join(out, 'xlsx.full.min.js');
  const src = fs.readFileSync(f, 'utf8');
  fs.writeFileSync(f, src.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, c => '\\x' + c.charCodeAt(0).toString(16).padStart(2, '0')));
}
console.log(`copied ${Object.keys(files).length} library files to www/lib`);
