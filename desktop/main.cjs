// Finestra desktop di nikoffice (Electron): carica l'app già compilata in app/
// senza server né Node.js installato sul PC.
const { app, BrowserWindow, Menu, shell } = require('electron');
const path = require('path');

if (!app.requestSingleInstanceLock()) {
  app.quit();
}

let finestra = null;

function creaFinestra() {
  finestra = new BrowserWindow({
    width: 1280,
    height: 860,
    show: false,
    title: 'nikoffice',
    icon: path.join(__dirname, 'build', 'icon.png'),
    webPreferences: { contextIsolation: true, nodeIntegration: false, sandbox: true, spellcheck: false },
  });
  finestra.maximize();
  finestra.once('ready-to-show', () => finestra.show());
  // I link esterni si aprono nel browser, non dentro l'app.
  finestra.webContents.setWindowOpenHandler(({ url }) => {
    if (/^https?:/.test(url)) shell.openExternal(url);
    return { action: 'deny' };
  });
  finestra.webContents.on('will-navigate', (evento, url) => {
    if (!url.startsWith('file:')) {
      evento.preventDefault();
      if (/^https?:/.test(url)) shell.openExternal(url);
    }
  });
  finestra.loadFile(path.join(__dirname, 'app', 'index.html'));
}

app.on('second-instance', () => {
  if (!finestra) return;
  if (finestra.isMinimized()) finestra.restore();
  finestra.focus();
});

app.whenReady().then(() => {
  // L'app ha già la sua barra multifunzione: niente menu File/Modifica di Electron.
  Menu.setApplicationMenu(null);
  creaFinestra();
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) creaFinestra();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
