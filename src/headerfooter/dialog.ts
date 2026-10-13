/** Finestra "Intestazione" / "Piè di pagina": testo, allineamento e numero di pagina. */

import type { HeaderFooterDialog } from './commands';
import type { HeaderFooterAlign } from './model';
import './styles.css';

function el<K extends keyof HTMLElementTagNameMap>(tag: K, className = '', text = ''): HTMLElementTagNameMap[K] {
  const e = document.createElement(tag);
  if (className) e.className = className;
  if (text) e.textContent = text;
  return e;
}

function button(text: string, className = ''): HTMLButtonElement {
  const b = el('button', className, text);
  b.type = 'button';
  return b;
}

const ALIGNS: readonly [HeaderFooterAlign, string][] = [
  ['left', 'A sinistra'],
  ['center', 'Al centro'],
  ['right', 'A destra'],
];

export const openHeaderFooterDialog: HeaderFooterDialog = ({ part, current }) =>
  new Promise((resolve) => {
    const title = part === 'header' ? 'Intestazione' : 'Piè di pagina';
    const backdrop = el('div', 'hf-backdrop');
    const box = el('form', 'hf-dialog');
    box.setAttribute('role', 'dialog');
    box.setAttribute('aria-label', title);

    const text = el('textarea', 'hf-text');
    text.name = 'text';
    text.rows = 3;
    text.value = current.text;
    text.placeholder = part === 'header' ? 'Es. il titolo del documento' : 'Es. il nome dell\'autore';
    const align = el('select', 'hf-align');
    align.name = 'align';
    for (const [value, label] of ALIGNS) {
      const o = el('option', '', label);
      o.value = value;
      align.append(o);
    }
    align.value = current.align;
    const number = el('input');
    number.type = 'checkbox';
    number.name = 'pageNumber';
    number.checked = current.pageNumber;
    const numberLabel = el('label', 'hf-check');
    numberLabel.append(number, ' Numero di pagina');
    const alignLabel = el('label', 'hf-row');
    alignLabel.append(el('span', '', 'Allineamento'), align);

    const intro = el('p', 'hf-intro', `Compare in tutte le pagine del documento. Più righe: vai a capo con Invio.`);
    box.append(el('h2', '', title), intro, text, alignLabel, numberLabel);
    if (current.complex) box.append(el('p', 'hf-warning', 'Questa parte contiene anche immagini o formattazione: se la modifichi qui resterà solo il testo.'));

    const buttons = el('div', 'hf-buttons');
    const remove = button('Rimuovi', 'hf-remove');
    const cancel = button('Annulla');
    const ok = button('OK', 'hf-primary');
    remove.disabled = !current.text && !current.pageNumber && !current.complex;
    buttons.append(remove, cancel, ok);
    box.append(buttons);
    backdrop.append(box);

    const previous = document.activeElement as HTMLElement | null;
    const done = (value: Parameters<typeof resolve>[0]) => {
      backdrop.remove();
      previous?.focus?.();
      resolve(value);
    };
    ok.addEventListener('click', () => done({ text: text.value, align: align.value as HeaderFooterAlign, pageNumber: number.checked }));
    remove.addEventListener('click', () => done({ text: '', align: 'left', pageNumber: false }));
    cancel.addEventListener('click', () => done(null));
    box.addEventListener('submit', (e) => e.preventDefault());
    backdrop.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        e.stopPropagation();
        done(null);
      } else if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) {
        e.preventDefault();
        ok.click();
      }
    });
    backdrop.addEventListener('mousedown', (e) => {
      if (e.target === backdrop) done(null);
    });
    document.body.append(backdrop);
    text.focus();
  });
