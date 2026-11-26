// MHM MD — Windows / desktop shell (Electron)
const { app, BrowserWindow, shell, Menu, protocol, net, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const { pathToFileURL } = require('url');

const WWW = path.join(__dirname, 'www');
const EXTS = /\.(pdf|docx|pptx|xlsx|xls|ods|txt|md|markdown|csv|html?|json|xml|log|png|jpe?g|webp|bmp|gif)$/i;

// Serve the app from app://local/ so fetch(), workers and WASM (OCR) work like on the web
protocol.registerSchemesAsPrivileged([{ scheme: 'app', privileges: { standard: true, secure: true, supportFetchAPI: true, corsEnabled: true, stream: true } }]);

let win = null;
let rendererReady = false;
let queue = [];

function filesFromArgv(argv) {
  const out = [];
  // packaged: [exe, ...files]; development: [electron, ., ...files]
  for (const a of argv.slice(app.isPackaged ? 1 : 2)) {
    if (!a || a.startsWith('-')) continue;
    try {
      const st = fs.statSync(a);
      if (st.isFile() && EXTS.test(a)) out.push(a);
      else if (st.isDirectory()) {
        for (const f of fs.readdirSync(a)) {
          const p = path.join(a, f);
          if (EXTS.test(f) && fs.statSync(p).isFile()) out.push(p);
        }
      }
    } catch (e) { /* not a path */ }
  }
  return out;
}

function sendFiles(paths) {
  if (!paths.length) return;
  if (!win || !rendererReady) { queue = queue.concat(paths); return; }
  const files = [];
  for (const p of paths) {
    try { files.push({ name: path.basename(p), data: fs.readFileSync(p) }); } catch (e) { /* skip */ }
  }
  if (files.length) win.webContents.send('open-files', files);
  if (win.isMinimized()) win.restore();
  win.focus();
}

function createWindow() {
  win = new BrowserWindow({
    width: 1000, height: 800, minWidth: 420, minHeight: 560,
    title: 'MHM MD', backgroundColor: '#0f766e', autoHideMenuBar: true,
    icon: path.join(WWW, 'icon.png'),
    webPreferences: { contextIsolation: true, nodeIntegration: false, sandbox: true, preload: path.join(__dirname, 'preload.js') },
  });
  win.loadURL('app://local/index.html');
  if (process.env.MHM_DEBUG) win.webContents.on('console-message', (e, lvl, msg) => console.log('[renderer]', msg));
  win.webContents.setWindowOpenHandler(({ url }) => { shell.openExternal(url); return { action: 'deny' }; });
  win.on('closed', () => { win = null; rendererReady = false; });
}

const gotLock = app.requestSingleInstanceLock();
if (!gotLock) {
  app.quit();
} else {
  // a second launch (e.g. right-click → "Convert to Markdown" on more files) forwards its files here
  app.on('second-instance', (e, argv) => sendFiles(filesFromArgv(argv)));

  app.whenReady().then(() => {
    protocol.handle('app', req => {
      const rel = decodeURIComponent(new URL(req.url).pathname).replace(/^\/+/, '');
      const file = path.normalize(path.join(WWW, rel || 'index.html'));
      if (!file.startsWith(WWW)) return new Response('forbidden', { status: 403 });
      return net.fetch(pathToFileURL(file).toString());
    });
    ipcMain.on('renderer-ready', () => {
      rendererReady = true;
      const q = queue; queue = [];
      sendFiles(q);
    });
    Menu.setApplicationMenu(null);
    createWindow();
    sendFiles(filesFromArgv(process.argv));
  });
  app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });
  app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow(); });
}
