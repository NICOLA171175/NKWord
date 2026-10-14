/**
 * Finestra "Correzione automatica", come quella di Word: in alto le opzioni
 * da attivare o spegnere, sotto l'elenco "Sostituisci / Con" da modificare.
 */

import { activeReplacements, DEFAULT_REPLACEMENTS, DEFAULT_SETTINGS, type AutoCorrectSettings, type Replacement } from './rules';

type Flag = 'twoInitialCaps' | 'capitalizeSentences' | 'capsLock' | 'smartQuotes' | 'replaceText';

const FLAGS: readonly { key: Flag; label: string }[] = [
  { key: 'twoInitialCaps', label: 'Correggi DUe INiziali MAiuscole' },
  { key: 'capitalizeSentences', label: "Maiuscola all'inizio delle frasi" },
  { key: 'capsLock', label: "Correggi l'uso accidentale del tasto bLOC mAIUSC" },
  { key: 'smartQuotes', label: 'Sostituisci le virgolette diritte con virgolette tipografiche' },
  { key: 'replaceText', label: 'Sostituisci il testo durante la digitazione' },
];

function el<K extends keyof HTMLElementTagNameMap>(tag: K, className = '', text = ''): HTMLElementTagNameMap[K] {
  const e = document.createElement(tag);
  if (className) e.className = className;
  if (text) e.textContent = text;
  return e;
}

/** Dall'elenco modificato nella finestra ricava le voci aggiunte e quelle tolte rispetto al predefinito. */
export function diffReplacements(list: readonly Replacement[]): Pick<AutoCorrectSettings, 'custom' | 'removed'> {
  const byFrom = new Map(list.map((r) => [r.from.toLowerCase(), r]));
  const defaults = new Map(DEFAULT_REPLACEMENTS.map((r) => [r.from.toLowerCase(), r]));
  const removed = DEFAULT_REPLACEMENTS.filter((r) => !byFrom.has(r.from.toLowerCase())).map((r) => r.from);
  const custom = list.filter((r) => defaults.get(r.from.toLowerCase())?.to !== r.to);
  return { custom, removed };
}

/** Mostra la finestra; la promessa si risolve con le nuove opzioni (o `null` se annullata). */
export function openAutoCorrectDialog(settings: AutoCorrectSettings): Promise<AutoCorrectSettings | null> {
  return new Promise((resolve) => {
    let list: Replacement[] = activeReplacements(settings).sort((a, b) => a.from.localeCompare(b.from, 'it'));
    let selected: string | null = null;

    const backdrop = el('div', 'autocorrect-backdrop');
    const dialog = el('form', 'autocorrect-dialog');
    dialog.setAttribute('role', 'dialog');
    dialog.setAttribute('aria-label', 'Correzione automatica');
    backdrop.append(dialog);

    const title = el('h2', '', 'Correzione automatica');
    const intro = el('p', 'autocorrect-intro', 'Mentre scrivi, NKWord corregge da solo questi errori. Per annullare una correzione premi Ctrl+Z subito dopo.');

    const checks = new Map<Flag, HTMLInputElement>();
    const flagsBox = el('div', 'autocorrect-flags');
    for (const f of FLAGS) {
      const row = el('label', 'autocorrect-check');
      const box = el('input');
      box.type = 'checkbox';
      box.name = f.key;
      box.checked = settings[f.key];
      checks.set(f.key, box);
      row.append(box, document.createTextNode(` ${f.label}`));
      flagsBox.append(row);
      if (f.key === 'smartQuotes') {
        const styleRow = el('label', 'autocorrect-quote-style', 'Virgolette doppie: ');
        const select = el('select');
        select.name = 'quoteStyle';
        for (const [value, label] of [['curly', '“alte”'], ['angle', '«basse»']] as const) {
          const o = el('option', '', label);
          o.value = value;
          o.selected = settings.quoteStyle === value;
          select.append(o);
        }
        styleRow.append(select);
        flagsBox.append(styleRow);
      }
    }

    // Elenco "Sostituisci / Con".
    const fromInput = el('input');
    fromInput.name = 'from';
    fromInput.spellcheck = false;
    const toInput = el('input');
    toInput.name = 'to';
    toInput.spellcheck = false;
    const head = el('div', 'autocorrect-entry');
    const fromLabel = el('label', '', 'Sostituisci:');
    fromLabel.append(fromInput);
    const toLabel = el('label', '', 'Con:');
    toLabel.append(toInput);
    head.append(fromLabel, toLabel);

    const table = el('div', 'autocorrect-list');
    table.setAttribute('role', 'listbox');
    table.setAttribute('aria-label', 'Voci della correzione automatica');

    const addBtn = el('button', 'autocorrect-btn', 'Aggiungi');
    addBtn.type = 'button';
    addBtn.name = 'add';
    const delBtn = el('button', 'autocorrect-btn', 'Elimina');
    delBtn.type = 'button';
    delBtn.name = 'delete';
    const resetBtn = el('button', 'autocorrect-btn', 'Ripristina predefiniti');
    resetBtn.type = 'button';
    resetBtn.name = 'reset';
    const listButtons = el('div', 'autocorrect-list-buttons');
    listButtons.append(addBtn, delBtn, resetBtn);

    const renderList = () => {
      table.replaceChildren(
        ...list.map((r) => {
          const row = el('div', 'autocorrect-row');
          row.setAttribute('role', 'option');
          row.dataset.from = r.from;
          row.classList.toggle('selected', r.from === selected);
          row.setAttribute('aria-selected', String(r.from === selected));
          row.append(el('span', 'autocorrect-from', r.from), el('span', 'autocorrect-to', r.to));
          row.addEventListener('click', () => {
            selected = r.from;
            fromInput.value = r.from;
            toInput.value = r.to;
            refresh();
          });
          return row;
        }),
      );
    };
    const refresh = () => {
      const from = fromInput.value.trim();
      const existing = list.find((r) => r.from.toLowerCase() === from.toLowerCase());
      addBtn.textContent = existing ? 'Sostituisci' : 'Aggiungi';
      addBtn.disabled = !from || !toInput.value || existing?.to === toInput.value;
      delBtn.disabled = !existing;
      renderList();
    };

    addBtn.addEventListener('click', () => {
      const from = fromInput.value.trim();
      if (!from || !toInput.value) return;
      list = [...list.filter((r) => r.from.toLowerCase() !== from.toLowerCase()), { from, to: toInput.value }].sort((a, b) =>
        a.from.localeCompare(b.from, 'it'),
      );
      selected = from;
      refresh();
    });
    delBtn.addEventListener('click', () => {
      const from = fromInput.value.trim().toLowerCase();
      list = list.filter((r) => r.from.toLowerCase() !== from);
      selected = null;
      fromInput.value = '';
      toInput.value = '';
      refresh();
    });
    resetBtn.addEventListener('click', () => {
      list = activeReplacements(DEFAULT_SETTINGS).sort((a, b) => a.from.localeCompare(b.from, 'it'));
      for (const f of FLAGS) checks.get(f.key)!.checked = DEFAULT_SETTINGS[f.key];
      (dialog.elements.namedItem('quoteStyle') as HTMLSelectElement).value = DEFAULT_SETTINGS.quoteStyle;
      selected = null;
      refresh();
    });
    for (const input of [fromInput, toInput]) {
      input.addEventListener('input', () => {
        const match = list.find((r) => r.from.toLowerCase() === fromInput.value.trim().toLowerCase());
        selected = match?.from ?? null;
        refresh();
        if (match) table.querySelector('.selected')?.scrollIntoView?.({ block: 'nearest' });
      });
    }

    const footer = el('div', 'autocorrect-footer');
    const ok = el('button', 'autocorrect-btn primary', 'OK');
    ok.type = 'submit';
    const cancel = el('button', 'autocorrect-btn', 'Annulla');
    cancel.type = 'button';
    footer.append(ok, cancel);

    dialog.append(title, intro, flagsBox, el('h3', '', 'Sostituisci il testo durante la digitazione'), head, table, listButtons, footer);

    const close = (result: AutoCorrectSettings | null) => {
      backdrop.remove();
      document.removeEventListener('keydown', onKey, true);
      resolve(result);
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        close(null);
      }
    };
    dialog.addEventListener('submit', (e) => {
      e.preventDefault();
      const flags = Object.fromEntries(FLAGS.map((f) => [f.key, checks.get(f.key)!.checked])) as Record<Flag, boolean>;
      const quoteStyle = (dialog.elements.namedItem('quoteStyle') as HTMLSelectElement).value === 'angle' ? 'angle' : 'curly';
      close({ ...settings, ...flags, quoteStyle, ...diffReplacements(list) });
    });
    cancel.addEventListener('click', () => close(null));
    backdrop.addEventListener('mousedown', (e) => {
      if (e.target === backdrop) close(null);
    });
    document.addEventListener('keydown', onKey, true);

    refresh();
    document.body.append(backdrop);
    checks.get('twoInitialCaps')!.focus();
  });
}
