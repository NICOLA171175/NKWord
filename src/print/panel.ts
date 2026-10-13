/**
 * Schermata Stampa (come File > Stampa in Word): copre la finestra, con le
 * impostazioni a sinistra e le pagine da stampare affiancate a destra.
 * Esc o la freccia in alto tornano al documento.
 */

import type { Editor } from '../core';
import { TabLeaders } from '../ui/tab-leaders';
import type { PrintModule } from './module';
import { MAX_COPIES, type Orientation, type PageChoice } from './pages';
import { renderPage } from './render';
import './styles.css';

export interface PrintHost {
  readonly editor: Editor;
  /** Riporta il focus al documento. */
  readonly input?: { focus(): void };
}

export const DEFAULT_PREVIEW_ZOOM = 0.5;

function el<K extends keyof HTMLElementTagNameMap>(d: Document, tag: K, className: string, text = ''): HTMLElementTagNameMap[K] {
  const e = d.createElement(tag);
  e.className = className;
  e.textContent = text;
  return e;
}

function select(d: Document, className: string, options: readonly [string, string][]): HTMLSelectElement {
  const s = el(d, 'select', className);
  for (const [value, label] of options) {
    const o = el(d, 'option', '', label);
    o.value = value;
    s.append(o);
  }
  return s;
}

function field(d: Document, label: string, control: HTMLElement): HTMLLabelElement {
  const l = el(d, 'label', 'print-field');
  l.append(el(d, 'span', 'print-field-label', label), control);
  return l;
}

const plural = (n: number, one: string, many: string) => `${n} ${n === 1 ? one : many}`;

export function mountPrint(host: PrintHost, module: PrintModule, d: Document = document): () => void {
  const { editor } = host;
  let zoom = DEFAULT_PREVIEW_ZOOM;

  const root = el(d, 'div', 'print-backstage');
  root.setAttribute('role', 'dialog');
  root.setAttribute('aria-label', 'Stampa');
  root.hidden = true;

  // Colonna delle impostazioni ------------------------------------------
  const side = el(d, 'div', 'print-side');
  const back = el(d, 'button', 'print-back', '←');
  back.type = 'button';
  back.title = 'Torna al documento (Esc)';
  back.setAttribute('aria-label', 'Torna al documento');
  const head = el(d, 'div', 'print-head');
  head.append(back, el(d, 'h1', 'print-title', 'Stampa'));

  const printButton = el(d, 'button', 'print-button');
  printButton.type = 'button';
  printButton.append(el(d, 'span', 'print-button-icon', '🖶'), el(d, 'span', '', 'Stampa'));
  const copies = el(d, 'input', 'print-copies');
  copies.type = 'number';
  copies.min = '1';
  copies.max = String(MAX_COPIES);
  const action = el(d, 'div', 'print-action');
  action.append(printButton, field(d, 'Copie', copies));

  const pages = select(d, 'print-pages', [
    ['all', 'Stampa tutte le pagine'],
    ['current', 'Stampa pagina corrente'],
    ['custom', 'Intervallo personalizzato'],
  ]);
  const range = el(d, 'input', 'print-range');
  range.type = 'text';
  range.placeholder = 'es. 1-3, 5';
  range.title = 'Scrivi i numeri di pagina e gli intervalli separati da virgole, per esempio 1-3, 5';
  const error = el(d, 'p', 'print-error');
  error.setAttribute('role', 'alert');
  const orientation = select(d, 'print-orientation', [
    ['', 'Come nel documento'],
    ['portrait', 'Verticale'],
    ['landscape', 'Orizzontale'],
  ]);
  const summary = el(d, 'p', 'print-summary');

  const settings = el(d, 'div', 'print-settings');
  settings.append(
    el(d, 'h2', 'print-subtitle', 'Impostazioni'),
    field(d, 'Pagine', pages),
    field(d, 'Intervallo', range),
    error,
    field(d, 'Orientamento', orientation),
    summary,
  );
  side.append(head, action, settings);

  // Anteprima -----------------------------------------------------------
  const main = el(d, 'div', 'print-main');
  const sheets = el(d, 'div', 'print-sheets');
  const bar = el(d, 'div', 'print-bar');
  const zoomInput = el(d, 'input', 'print-zoom');
  zoomInput.type = 'range';
  zoomInput.min = '20';
  zoomInput.max = '120';
  zoomInput.step = '5';
  zoomInput.setAttribute('aria-label', 'Zoom anteprima');
  const zoomLabel = el(d, 'span', 'print-zoom-label');
  bar.append(zoomInput, zoomLabel);
  main.append(sheets, bar);
  root.append(side, main);

  const render = () => {
    const s = module.settings;
    copies.value = String(s.copies);
    pages.value = s.pages;
    if (d.activeElement !== range) range.value = s.range;
    orientation.value = s.orientation ?? '';
    zoomInput.value = String(Math.round(zoom * 100));
    zoomLabel.textContent = `${Math.round(zoom * 100)}%`;

    const p = module.preview();
    sheets.replaceChildren();
    error.textContent = 'error' in p ? p.error : '';
    printButton.disabled = 'error' in p;
    if ('error' in p) {
      summary.textContent = '';
      if (!p.layout) sheets.append(el(d, 'p', 'print-empty', p.error));
      return;
    }
    const total = p.layout.pages.length;
    summary.textContent =
      s.copies > 1
        ? `${plural(p.pages.length, 'pagina', 'pagine')} di ${total}, ${s.copies} copie: ${plural(p.sheets, 'pagina stampata', 'pagine stampate')}`
        : `${plural(p.pages.length, 'pagina', 'pagine')} di ${total}`;
    const leaders = new TabLeaders(p.layout.doc);
    for (const i of p.pages) {
      const page = p.layout.pages[i]!;
      const sheet = el(d, 'figure', 'print-sheet');
      sheet.append(renderPage(page, { zoom, doc: p.layout.doc, leaders, document: d }), el(d, 'figcaption', 'print-caption', `Pagina ${i + 1}`));
      sheets.append(sheet);
    }
  };

  const show = () => {
    root.hidden = !module.isOpen;
    if (module.isOpen) render();
  };

  back.addEventListener('click', () => void editor.executeAsync('print.close'));
  printButton.addEventListener('click', () => void editor.executeAsync('print.print'));
  copies.addEventListener('change', () => module.update({ copies: Number(copies.value) }));
  pages.addEventListener('change', () => module.update({ pages: pages.value as PageChoice }));
  // Scrivere un intervallo sceglie "Intervallo personalizzato", come in Word.
  range.addEventListener('input', () => module.update({ range: range.value, pages: 'custom' }));
  orientation.addEventListener('change', () => module.update({ orientation: (orientation.value || null) as Orientation | null }));
  zoomInput.addEventListener('input', () => {
    zoom = Number(zoomInput.value) / 100;
    render();
  });

  let wasOpen = false;
  const offChange = module.onChange(() => {
    show();
    if (wasOpen && !module.isOpen) host.input?.focus();
    else if (!wasOpen && module.isOpen) printButton.focus();
    wasOpen = module.isOpen;
  });
  const onKey = (e: KeyboardEvent) => {
    if (!module.isOpen) return;
    if (e.key === 'Escape') {
      e.preventDefault();
      e.stopPropagation();
      void editor.executeAsync('print.close');
    } else if (e.key.toLowerCase() === 'p' && (e.ctrlKey || e.metaKey)) {
      // Ctrl+P con l'anteprima aperta: stampa, non riapre l'anteprima.
      e.preventDefault();
      e.stopPropagation();
      void editor.executeAsync('print.print');
    }
  };
  d.addEventListener('keydown', onKey, true);
  d.body.append(root);
  show();

  return () => {
    offChange();
    d.removeEventListener('keydown', onKey, true);
    root.remove();
  };
}
