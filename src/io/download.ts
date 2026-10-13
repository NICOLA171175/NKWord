import type { SaveHandler } from './commands';

/** `SaveHandler` per il browser: scarica il file con un link temporaneo. */
export const browserDownload: SaveHandler = ({ data, fileName, mimeType }) => {
  const blob = new Blob([data as BlobPart], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  a.style.display = 'none';
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 0);
};

export function hasDom(): boolean {
  return typeof document !== 'undefined' && typeof URL !== 'undefined' && typeof URL.createObjectURL === 'function';
}
