/**
 * Ponte con il programma desktop (Electron): `desktop/preload.cjs` pubblica
 * `window.nkofficeDesktop`. Nel browser non c'è.
 */

export interface DesktopBridge {
  /** Percorso sul disco di un file scelto con la finestra "Apri" ('' se non c'è). */
  pathForFile(file: File): string;
  /** Legge un file dal disco. */
  readFile(path: string): Promise<Uint8Array>;
  /** Avvisa quando un documento è stato salvato, con il percorso scelto. */
  onFileSaved(listener: (path: string) => void): () => void;
}

export function desktopBridge(): DesktopBridge | undefined {
  return (globalThis as { nkofficeDesktop?: DesktopBridge }).nkofficeDesktop;
}

/** Nome del file in un percorso Windows o Mac. */
export function baseName(path: string): string {
  return path.split(/[\\/]/).pop() ?? path;
}

/** Cartella di un percorso, per mostrarla sotto il nome. */
export function dirName(path: string): string {
  const i = Math.max(path.lastIndexOf('/'), path.lastIndexOf('\\'));
  return i > 0 ? path.slice(0, i) : '';
}
