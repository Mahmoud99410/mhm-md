# MHM MD

Turn **PDF, Word, PowerPoint, Excel, TXT, images and more** into clean, structured Markdown that Claude and Claude Code understand well, using as few tokens as possible. Instructions inside the documents stay clear enough for Claude to follow.

## Features
- **Structure detection:** headings (font size / bold), bullet & numbered lists, tables, code blocks, slide titles & speaker notes, Excel sheets
- **Arabic RTL:** word order, presentation forms, lam-alef fixes, Windows-1256 text files
- **OCR (offline):** scanned PDFs and photos in Arabic + English (tesseract.js)
- **Compression levels:**
  - *light:* keep everything
  - *medium:* wide or long tables become CSV, boilerplate and tracking URLs are removed
  - *strong:* every table becomes CSV, emphasis is removed, repeated paragraphs are dropped
- **Quality report** per file: pages without text, OCR pages, possibly reversed Arabic, flat structure
- **Claude Code package (ZIP):**
  - `CLAUDE.md`: overview, detected tech stack, must / must-not rules, phases, key sections
  - `docs/` split by section, with `INDEX.md` (tokens + summary per file), so Claude Code reads only the parts it needs
- **Merge files** into one MD; duplicated paragraphs across files are removed
- **Optional ✨ improve with Claude API** (your key, stored on the device) restructures a document into a project spec
- **Android:** appears in the *Share* / *Open with* menu. **Windows:** right-click a file or folder and choose *Convert to Markdown (MHM MD)* (installer version)
- 100% local. Files never leave the device, except when you use the optional Claude API button.

## Download
Latest builds are on the [**latest** release](../../releases/tag/latest):
- `MHM-MD.apk` (Android)
- `MHM-MD-Setup-*.exe` (Windows installer, adds the right-click menu)
- `MHM-MD-Portable-*.exe` (Windows, no installation)

## Develop
```bash
npm install       # also copies browser libraries into www/lib
npm test          # converter checks on test/ samples
npm start         # desktop app (Electron)
npx cap sync android && cd android && ./gradlew assembleDebug
```
Core engine: `www/converter.js` (works in the browser and in Node). UI: `www/app.js`.
