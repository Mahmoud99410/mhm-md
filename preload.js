// Bridge between the Electron shell and the web app (files from the right-click menu)
const { contextBridge, ipcRenderer } = require('electron');
contextBridge.exposeInMainWorld('mhmDesktop', {
  onFiles(cb) {
    ipcRenderer.on('open-files', (e, files) => cb(files.map(f => ({ name: f.name, data: f.data }))));
    ipcRenderer.send('renderer-ready');
  },
});
