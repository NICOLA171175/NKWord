/**
 * Stampa vera e propria: le pagine scelte vengono disegnate a grandezza
 * reale in un contenitore nascosto e si apre la finestra di stampa del
 * sistema (`window.print`, anche nel programma desktop). Mentre si stampa,
 * il resto della pagina è nascosto, così su carta finisce solo il documento.
 */

import { ptToPx, type DocumentLayout } from '../core';
import { TabLeaders } from '../ui/tab-leaders';
import { renderPage } from './render';

export interface PrintJob {
  readonly layout: DocumentLayout;
  /** Indici delle pagine, nell'ordine di stampa. */
  readonly pages: readonly number[];
  readonly copies: number;
}

export type Printer = (job: PrintJob) => Promise<void>;

export const PRINT_ACTIVE_CLASS = 'print-job-active';

/** Prepara il contenitore con le pagine (copie fascicolate: 1,2,3,1,2,3). */
export function buildPrintOutput(job: PrintJob, d: Document = document): { root: HTMLElement; style: HTMLStyleElement } {
  const root = d.createElement('div');
  root.className = 'print-output';
  const leaders = new TabLeaders(job.layout.doc);
  for (let c = 0; c < job.copies; c++) {
    for (const i of job.pages) {
      const page = job.layout.pages[i];
      if (page) root.append(renderPage(page, { zoom: 1, doc: job.layout.doc, leaders, document: d }));
    }
  }
  // Il foglio prende le misure della prima pagina stampata.
  const first = job.layout.pages[job.pages[0] ?? 0];
  const style = d.createElement('style');
  style.textContent = first
    ? `@page { size: ${ptToPx(first.width)}px ${ptToPx(first.height)}px; margin: 0; }`
    : '@page { margin: 0; }';
  return { root, style };
}

/** Stampa con la finestra del sistema. `print` è iniettabile per i test. */
export function createDomPrinter(win: Window = window, print: () => void = () => win.print()): Printer {
  let pending: (() => void) | undefined;
  return async (job) => {
    pending?.();
    const d = win.document;
    const { root, style } = buildPrintOutput(job, d);
    d.head.append(style);
    d.body.append(root);
    d.body.classList.add(PRINT_ACTIVE_CLASS);
    const cleanup = () => {
      if (pending !== cleanup) return;
      pending = undefined;
      win.removeEventListener('afterprint', cleanup);
      root.remove();
      style.remove();
      d.body.classList.remove(PRINT_ACTIVE_CLASS);
    };
    pending = cleanup;
    // Le immagini devono essere caricate prima di aprire la finestra di stampa.
    await Promise.all(Array.from(root.querySelectorAll('img'), (img) => img.decode?.().catch(() => undefined)));
    // Si smonta a stampa finita; se l'evento non arriva, alla stampa successiva.
    // A schermo il contenitore è comunque invisibile.
    win.addEventListener('afterprint', cleanup);
    print();
  };
}
