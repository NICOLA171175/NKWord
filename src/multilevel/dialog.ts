/**
 * Finestre degli elenchi a più livelli: "Definisci nuovo elenco a più
 * livelli" (formato di ogni livello, con anteprima) e "Imposta valore di
 * numerazione".
 */

import type { ListLevel, NumberFormat } from '../core';
import { buildPattern, LEVELS, markerText, parsePattern } from './model';

export interface DefineDialogInput {
  readonly levels: readonly ListLevel[];
  /** Livello mostrato all'apertura (quello del cursore). */
  readonly level: number;
}

export interface MultilevelDialogs {
  define(input: DefineDialogInput): Promise<ListLevel[] | null>;
  setValue(input: { value: number }): Promise<number | null>;
}

export const FORMAT_CHOICES: readonly { format: NumberFormat; label: string }[] = [
  { format: 'decimal', label: '1, 2, 3, …' },
  { format: 'upperRoman', label: 'I, II, III, …' },
  { format: 'lowerRoman', label: 'i, ii, iii, …' },
  { format: 'upperLetter', label: 'A, B, C, …' },
  { format: 'lowerLetter', label: 'a, b, c, …' },
  { format: 'bullet', label: 'Punto elenco' },
  { format: 'none', label: '(nessuno)' },
];

const CM = 28.3465;
const toCm = (pt: number) => Math.round((pt / CM) * 100) / 100;
const fromCm = (cm: number) => Math.round(cm * CM * 10) / 10;

function el<K extends keyof HTMLElementTagNameMap>(tag: K, className?: string, text?: string): HTMLElementTagNameMap[K] {
  const e = document.createElement(tag);
  if (className) e.className = className;
  if (text !== undefined) e.textContent = text;
  return e;
}

function field(label: string, control: HTMLElement): HTMLLabelElement {
  const row = el('label', 'multilevel-field');
  row.append(el('span', undefined, label), control);
  return row;
}

function numberInput(name: string, value: number, step: string, min = 0): HTMLInputElement {
  const i = el('input');
  i.type = 'number';
  i.name = name;
  i.step = step;
  i.min = String(min);
  i.value = String(value);
  return i;
}

/** Struttura comune: sfondo, finestra, titolo, pulsanti OK/Annulla. */
function frame<T>(title: string, resolve: (v: T | null) => void, collect: () => T | null) {
  const backdrop = el('div', 'multilevel-backdrop');
  const dialog = el('form', 'multilevel-dialog');
  dialog.setAttribute('role', 'dialog');
  dialog.setAttribute('aria-label', title);
  const body = el('div', 'multilevel-body');
  const buttons = el('div', 'multilevel-buttons');
  const ok = el('button', 'multilevel-ok', 'OK');
  ok.type = 'submit';
  const cancel = el('button', 'multilevel-cancel', 'Annulla');
  cancel.type = 'button';
  buttons.append(ok, cancel);
  dialog.append(el('h2', undefined, title), body, buttons);
  backdrop.append(dialog);

  let done = false;
  const close = (value: T | null) => {
    if (done) return;
    done = true;
    backdrop.remove();
    document.removeEventListener('keydown', onKey, true);
    resolve(value);
  };
  const onKey = (e: KeyboardEvent) => {
    if (e.key === 'Escape') {
      e.preventDefault();
      close(null);
    }
  };
  dialog.addEventListener('submit', (e) => {
    e.preventDefault();
    const v = collect();
    if (v !== null) close(v);
  });
  cancel.addEventListener('click', () => close(null));
  backdrop.addEventListener('mousedown', (e) => {
    if (e.target === backdrop) close(null);
  });
  document.addEventListener('keydown', onKey, true);
  return { backdrop, dialog, body, show: () => document.body.append(backdrop) };
}

/** Finestra "Definisci nuovo elenco a più livelli". */
export function openDefineDialog(input: DefineDialogInput): Promise<ListLevel[] | null> {
  return new Promise((resolve) => {
    const levels: ListLevel[] = Array.from({ length: LEVELS }, (_, i) => ({ ...input.levels[i]! }));
    let current = Math.max(0, Math.min(LEVELS - 1, input.level));
    const { body, show, dialog } = frame('Definisci nuovo elenco a più livelli', resolve, () => levels);

    const levelList = el('div', 'multilevel-levels');
    levelList.setAttribute('role', 'listbox');
    levelList.setAttribute('aria-label', 'Livello da modificare');
    const levelButtons = levels.map((_, i) => {
      const b = el('button', 'multilevel-level', String(i + 1));
      b.type = 'button';
      b.setAttribute('role', 'option');
      b.addEventListener('click', () => select(i));
      levelList.append(b);
      return b;
    });

    const format = el('select');
    format.name = 'format';
    for (const c of FORMAT_CHOICES) {
      const o = el('option', undefined, c.label);
      o.value = c.format;
      format.append(o);
    }
    const before = el('input');
    before.name = 'before';
    const after = el('input');
    after.name = 'after';
    const parents = el('input');
    parents.type = 'checkbox';
    parents.name = 'parents';
    const parentsRow = el('label', 'multilevel-check');
    parentsRow.append(parents, el('span', undefined, 'Includi i numeri dei livelli superiori (es. 1.1.1)'));
    const symbol = el('input');
    symbol.name = 'symbol';
    symbol.maxLength = 2;
    const start = numberInput('start', 1, '1', 0);
    const left = numberInput('indentLeft', 0, '0.01');
    const hanging = numberInput('hanging', 0, '0.01');

    const numberRows = el('div', 'multilevel-rows');
    numberRows.append(field('Testo prima del numero', before), field('Testo dopo il numero', after), parentsRow, field('Inizia da', start));
    const symbolRow = field('Simbolo', symbol);
    const indentRows = el('div', 'multilevel-indents');
    indentRows.append(field('Rientro testo (cm)', left), field('Spazio per il numero (cm)', hanging));

    const editor = el('div', 'multilevel-editor');
    const heading = el('div', 'multilevel-heading');
    editor.append(heading, field('Stile numerazione', format), numberRows, symbolRow, indentRows);

    const preview = el('div', 'multilevel-preview');
    preview.setAttribute('aria-label', 'Anteprima');

    const top = el('div', 'multilevel-top');
    top.append(levelList, editor);
    body.append(top, el('div', 'multilevel-preview-title', 'Anteprima'), preview);

    const render = () => {
      preview.replaceChildren();
      const counters: number[] = [];
      // Un esempio realistico: 1, 1.1, 1.1.1 … e una seconda voce al livello scelto.
      const rows: number[] = [...Array.from({ length: LEVELS }, (_, i) => i), current];
      for (const i of rows) {
        const lv = levels[i]!;
        counters[i] = (counters[i] ?? lv.start - 1) + 1;
        counters.length = i + 1;
        const row = el('div', 'multilevel-preview-row' + (i === current ? ' current' : ''));
        row.style.paddingLeft = `${Math.max(0, lv.indentLeft - lv.hanging) * 0.6}px`;
        const mark = el('span', 'multilevel-preview-mark', markerText(levels, i, counters) || ' ');
        mark.style.minWidth = `${lv.hanging * 0.6}px`;
        row.append(mark, el('span', 'multilevel-preview-line'));
        preview.append(row);
      }
      levelButtons.forEach((b, i) => {
        b.classList.toggle('active', i === current);
        b.setAttribute('aria-selected', String(i === current));
      });
    };

    const load = () => {
      const lv = levels[current]!;
      heading.textContent = `Livello ${current + 1}`;
      format.value = lv.format;
      const bullet = lv.format === 'bullet';
      numberRows.hidden = bullet;
      symbolRow.hidden = !bullet;
      parentsRow.hidden = current === 0;
      if (bullet) symbol.value = lv.text;
      else {
        const p = parsePattern(lv.text, current);
        before.value = p.before;
        after.value = p.after;
        parents.checked = p.parents;
      }
      start.value = String(lv.start);
      left.value = String(toCm(lv.indentLeft));
      hanging.value = String(toCm(lv.hanging));
      render();
    };

    const save = () => {
      const lv = levels[current]!;
      const f = format.value as NumberFormat;
      let text: string;
      if (f === 'bullet') text = symbol.value || (lv.format === 'bullet' ? lv.text : '•');
      else text = buildPattern({ before: before.value, after: after.value, parents: parents.checked }, current);
      const num = (i: HTMLInputElement, fallback: number) => (Number.isFinite(i.valueAsNumber) ? i.valueAsNumber : fallback);
      levels[current] = {
        ...lv,
        format: f,
        text,
        start: Math.max(0, Math.round(num(start, lv.start))),
        indentLeft: Math.max(0, fromCm(num(left, toCm(lv.indentLeft)))),
        hanging: Math.max(0, fromCm(num(hanging, toCm(lv.hanging)))),
      };
    };

    function select(i: number) {
      save();
      current = i;
      load();
    }

    format.addEventListener('change', () => {
      // Passando a/da punto elenco si cambia il campo mostrato.
      const wasBullet = levels[current]!.format === 'bullet';
      save();
      if (wasBullet !== (format.value === 'bullet')) {
        const lv = levels[current]!;
        levels[current] = { ...lv, text: format.value === 'bullet' ? '•' : `%${current + 1}.` };
      }
      load();
    });
    for (const i of [before, after, parents, symbol, start, left, hanging]) {
      i.addEventListener('input', () => {
        save();
        render();
      });
    }
    levelList.addEventListener('keydown', (e) => {
      if (e.key !== 'ArrowDown' && e.key !== 'ArrowUp') return;
      e.preventDefault();
      const next = Math.max(0, Math.min(LEVELS - 1, current + (e.key === 'ArrowDown' ? 1 : -1)));
      select(next);
      levelButtons[next]!.focus();
    });
    dialog.addEventListener('submit', save, true);

    load();
    show();
    levelButtons[current]!.focus();
  });
}

/** Finestra "Imposta valore di numerazione". */
export function openSetValueDialog(input: { value: number }): Promise<number | null> {
  return new Promise((resolve) => {
    const value = numberInput('value', input.value, '1', 0);
    const { body, show } = frame('Imposta valore di numerazione', resolve, () =>
      Number.isFinite(value.valueAsNumber) && value.valueAsNumber >= 0 ? Math.round(value.valueAsNumber) : null,
    );
    body.append(field('Imposta valore su', value));
    show();
    value.focus();
    value.select();
  });
}

export const domDialogs: MultilevelDialogs = { define: openDefineDialog, setValue: openSetValueDialog };
