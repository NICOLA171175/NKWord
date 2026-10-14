// Finestra desktop di NKWord (Electron): carica l'app già compilata in app/
// senza server né Node.js installato sul PC.
const { app, BrowserWindow, Menu, ipcMain, session, shell } = require('electron');
const { execFile } = require('child_process');
const fs = require('fs');
const path = require('path');

// Il programma si chiamava NKoffice: i dati (documenti recenti, dizionario
// personale, impostazioni, salvataggi automatici) restano nella cartella col
// vecchio nome, così chi aggiorna non li perde.
app.setPath('userData', path.join(app.getPath('appData'), 'NKoffice'));

// Documenti che l'app sa aprire: solo questi si possono rileggere dai recenti.
const ESTENSIONI_DOCUMENTI = new Set(['.docx', '.html', '.htm', '.txt', '.md', '.pdf']);

if (!app.requestSingleInstanceLock()) {
  app.quit();
}

let finestra = null;

function creaFinestra() {
  finestra = new BrowserWindow({
    width: 1280,
    height: 860,
    show: false,
    title: 'NKWord',
    icon: path.join(__dirname, 'build', 'icon.png'),
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      spellcheck: false,
      preload: path.join(__dirname, 'preload.cjs'),
    },
  });
  finestra.maximize();
  finestra.once('ready-to-show', () => finestra.show());
  // I link esterni si aprono nel browser (o nel programma di posta), non dentro l'app.
  finestra.webContents.setWindowOpenHandler(({ url }) => {
    if (/^(https?|mailto|ftp):/i.test(url)) shell.openExternal(url);
    return { action: 'deny' };
  });
  finestra.webContents.on('will-navigate', (evento, url) => {
    if (!url.startsWith('file:')) {
      evento.preventDefault();
      if (/^(https?|mailto|ftp):/i.test(url)) shell.openExternal(url);
    }
  });
  finestra.loadFile(path.join(__dirname, 'app', 'index.html'));
}

app.on('second-instance', () => {
  if (!finestra) return;
  if (finestra.isMinimized()) finestra.restore();
  finestra.focus();
});

ipcMain.handle('nkoffice:leggi-file', async (_evento, percorso) => {
  if (typeof percorso !== 'string' || !path.isAbsolute(percorso) || !ESTENSIONI_DOCUMENTI.has(path.extname(percorso).toLowerCase())) {
    throw new Error('File non valido');
  }
  return new Uint8Array(await fs.promises.readFile(percorso));
});

// Dettatura: su Windows apre la dettatura del sistema premendo Win+H, che
// scrive nel campo con il focus (il documento). Il riconoscimento vocale
// del browser non funziona in Electron (vedi src/dictation/engine.ts).
const PREMI_WIN_H = [
  '$k = Add-Type -Name Tasti -Namespace NKWord -PassThru -MemberDefinition',
  "'[DllImport(\"user32.dll\")] public static extern void keybd_event(byte v, byte s, uint f, System.UIntPtr e);';",
  '$k::keybd_event(0x5B, 0, 0, [UIntPtr]::Zero);',
  '$k::keybd_event(0x48, 0, 0, [UIntPtr]::Zero);',
  '$k::keybd_event(0x48, 0, 2, [UIntPtr]::Zero);',
  '$k::keybd_event(0x5B, 0, 2, [UIntPtr]::Zero);',
].join(' ');
// -EncodedCommand evita i problemi con le virgolette nella riga di comando.
const COMANDO_WIN_H = Buffer.from(PREMI_WIN_H, 'utf16le').toString('base64');

ipcMain.handle('nkoffice:dettatura-sistema', () => {
  if (process.platform !== 'win32') throw new Error('Solo su Windows');
  return new Promise((risolvi, rifiuta) => {
    execFile(
      'powershell.exe',
      ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-WindowStyle', 'Hidden', '-EncodedCommand', COMANDO_WIN_H],
      { windowsHide: true, timeout: 15000 },
      (errore) => (errore ? rifiuta(errore) : risolvi(true)),
    );
  });
});

app.whenReady().then(() => {
  // L'app ha già la sua barra multifunzione: niente menu File/Modifica di Electron.
  Menu.setApplicationMenu(null);
  // Salvando, la finestra "Salva con nome" sceglie il percorso: lo si
  // comunica all'app per l'elenco dei documenti recenti.
  session.defaultSession.on('will-download', (_evento, elemento, contenuto) => {
    elemento.once('done', (_e, stato) => {
      const percorso = elemento.getSavePath();
      if (stato === 'completed' && percorso && !contenuto.isDestroyed()) contenuto.send('nkoffice:file-salvato', percorso);
    });
  });
  creaFinestra();
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) creaFinestra();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
