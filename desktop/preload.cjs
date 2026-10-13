// Ponte fra la finestra (sandbox) e il programma desktop: pubblica
// window.nkofficeDesktop per i documenti recenti (vedi src/recent/desktop.ts).
const { contextBridge, ipcRenderer, webUtils } = require('electron');

contextBridge.exposeInMainWorld('nkofficeDesktop', {
  pathForFile: (file) => {
    try {
      return webUtils.getPathForFile(file);
    } catch {
      return '';
    }
  },
  readFile: (percorso) => ipcRenderer.invoke('nkoffice:leggi-file', percorso),
  onFileSaved: (ascolta) => {
    const gestore = (_evento, percorso) => ascolta(percorso);
    ipcRenderer.on('nkoffice:file-salvato', gestore);
    return () => ipcRenderer.removeListener('nkoffice:file-salvato', gestore);
  },
});
