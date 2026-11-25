# MHM MD

Convert **PDF / Word (DOCX) / TXT / MD / CSV / HTML** into clean, structured Markdown that Claude (chat or Claude Code) reads with far fewer tokens — and follows the instructions inside.

## What it does
- Detects headings (by font size / bold), bullet & numbered lists, tables, code blocks
- Arabic RTL support: fixes word order, presentation forms, lam-alef, Windows-1256 text files
- Removes repeated headers/footers, page numbers, the document's own dotted TOC, duplicate lines
- Adds a short front-matter + a note telling Claude to read everything and follow the instructions
- For long documents, lists the key instructions/requirements at the top
- Merges several files into one Markdown for a single upload
- 100% offline — files never leave the device

## Apps
| Platform | How |
| --- | --- |
| Android | `MHM-MD.apk` from the **latest** release (built by GitHub Actions with Capacitor) |
| Windows | `MHM-MD-Portable-*.exe` (no install) or `MHM-MD-Setup-*.exe` (built with Electron) |
| Any browser | open `www/index.html` |

## Develop
```bash
npm install
npm test          # converter checks on test/ samples
npm start         # desktop app (Electron)
npx cap sync android && cd android && ./gradlew assembleDebug
```
Core engine: `www/converter.js` (works in browser and Node).
