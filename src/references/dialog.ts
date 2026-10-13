/**
 * Finestre di Riferimenti: "Didascalia", "Riferimento incrociato" e
 * "Segna voce di indice", con le stesse scelte di Word.
 */

import type { CaptionArgs, CaptionPosition } from './captions';
import { REF_KINDS, type CrossRef, type RefKind, type RefTarget, type RefTargetType } from './crossrefs';
import type { MarkArgs } from './indexing';
import { joinEntry, splitEntry } from './indexing';
import type { ReferencesDialogs } from './module';
import './styles.css';

function el<K extends keyof HTMLElementTagNameMap>(tag: K, className = '', text = ''): HTMLElementTagNameMap[K] {
  const e = document.createElement(tag);
  if (className) e.className = className;
  if (text) e.textContent = text;
  return e;
}

function option(text: string, value: string): HTMLOptionElement {
  const o = el('option', '', text);
  o.value = value;
  return o;
}

function field(label: string, control: HTMLElement): HTMLLabelElement {
  const row = el('label', 'refs-field');
  row.append(el('span', '', label), control);
  return row;
}

function button(text: string, primary = false): HTMLButtonElement {
  const b = el('button', primary ? 'refs-primary' : '', text);
  b.type = 'button';
  return b;
}

/**
 * Finestra modale. `build` riceve la funzione per chiuderla con un
 * risultato; Esc o "Annulla" la chiudono con `null`.
 */
function modal<T>(title: string, build: (body: HTMLElement, buttons: HTMLElement, done: (value: T | null) => void) => HTMLElement | void): Promise<T | null> {
  return new Promise((resolve) => {
    const backdrop = el('div', 'refs-backdrop');
    const box = el('form', 'refs-dialog');
    box.setAttribute('role', 'dialog');
    box.setAttribute('aria-label', title);
    const body = el('div', 'refs-body');
    const buttons = el('div', 'refs-buttons');
    box.append(el('h2', '', title), body, buttons);
    backdrop.append(box);
    const previous = document.activeElement as HTMLElement | null;
    const done = (value: T | null) => {
      backdrop.remove();
      previous?.focus?.();
      resolve(value);
    };
    const focus = build(body, buttons, done);
    const cancel = button('Annulla');
    cancel.addEventListener('click', () => done(null));
    buttons.append(cancel);
    box.addEventListener('submit', (e) => e.preventDefault());
    backdrop.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        e.stopPropagation();
        done(null);
      }
    });
    backdrop.addEventListener('mousedown', (e) => {
      if (e.target === backdrop) done(null);
    });
    document.body.append(backdrop);
    (focus ?? box.querySelector<HTMLElement>('input, select, button'))?.focus();
  });
}

// ---------------------------------------------------------------------------
// Didascalia

export function openCaptionDialog(data: Parameters<ReferencesDialogs['caption']>[0]): Promise<CaptionArgs | null> {
  return modal<CaptionArgs>('Didascalia', (body, buttons, done) => {
    const preview = el('div', 'refs-preview');
    const label = el('select');
    label.name = 'label';
    for (const l of data.labels) label.append(option(l, l));
    const custom = el('input');
    custom.name = 'newLabel';
    custom.placeholder = 'Nuova etichetta (es. Grafico)';
    custom.spellcheck = false;
    const text = el('input');
    text.name = 'text';
    text.placeholder = 'Testo dopo il numero (facoltativo)';
    const position = el('select');
    position.name = 'position';
    position.append(option("Sopra l'elemento selezionato", 'above'), option("Sotto l'elemento selezionato", 'below'));
    const chosen = () => custom.value.trim() || label.value;
    const valid = () => /^[\p{L}][\p{L}\p{M}'’]*$/u.test(chosen());
    const ok = button('OK', true);
    const refresh = () => {
      const l = chosen();
      preview.textContent = valid() ? `${l} ${data.next(l)}${text.value.trim() ? `: ${text.value.trim()}` : ''}` : "L'etichetta deve essere una sola parola";
      ok.disabled = !valid();
    };
    let positionTouched = false;
    const syncPosition = () => {
      if (!positionTouched) position.value = data.defaultPosition(chosen());
      refresh();
    };
    label.addEventListener('change', syncPosition);
    custom.addEventListener('input', syncPosition);
    text.addEventListener('input', refresh);
    position.addEventListener('change', () => (positionTouched = true));
    body.append(
      el('p', 'refs-intro', 'La didascalia si numera da sola: se ne aggiungi una prima delle altre, i numeri si aggiornano.'),
      field('Didascalia', preview),
      field('Etichetta', label),
      field('Oppure', custom),
      field('Testo', text),
      field('Posizione', position),
    );
    syncPosition();
    const submit = () => {
      if (!valid()) return;
      done({ label: chosen(), text: text.value.trim(), position: position.value as CaptionPosition });
    };
    ok.addEventListener('click', submit);
    for (const input of [custom, text]) {
      input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
          e.preventDefault();
          submit();
        }
      });
    }
    buttons.append(ok);
    return text;
  });
}

// ---------------------------------------------------------------------------
// Riferimento incrociato

/** Tipi di destinazione proposti: titoli, una voce per etichetta di didascalia, note. */
export function crossRefCategories(targets: readonly RefTarget[]): { key: string; label: string; type: RefTargetType }[] {
  const out: { key: string; label: string; type: RefTargetType }[] = [{ key: 'heading', label: 'Titolo', type: 'heading' }];
  const labels = new Set(['Figura', 'Tabella', 'Equazione']);
  for (const t of targets) if (t.type === 'caption' && t.label) labels.add(t.label);
  for (const l of labels) out.push({ key: `caption:${l}`, label: l, type: 'caption' });
  out.push({ key: 'note', label: 'Nota a piè di pagina', type: 'note' });
  return out;
}

export function openCrossRefDialog(data: Parameters<ReferencesDialogs['crossRef']>[0]): Promise<CrossRef | null> {
  return modal<CrossRef>('Riferimento incrociato', (body, buttons, done) => {
    const categories = crossRefCategories(data.targets);
    const type = el('select');
    type.name = 'type';
    for (const c of categories) type.append(option(c.label, c.key));
    const kind = el('select');
    kind.name = 'kind';
    const list = el('select', 'refs-list');
    list.name = 'target';
    list.size = 8;
    const empty = el('p', 'refs-empty');
    const ok = button('Inserisci', true);

    const current = () => categories.find((c) => c.key === type.value)!;
    const items = () => {
      const c = current();
      return data.targets.filter((t) => t.type === c.type && (c.type !== 'caption' || `caption:${t.label}` === c.key));
    };
    const refresh = () => {
      const c = current();
      kind.replaceChildren(...REF_KINDS[c.type].map((k) => option(k.label, k.kind)));
      const found = items();
      list.replaceChildren(
        ...found.map((t) => option(`${'\u00a0'.repeat(((t.level ?? 1) - 1) * 3)}${t.display}`, t.id)),
      );
      if (found.length) list.selectedIndex = 0;
      list.hidden = !found.length;
      empty.hidden = !!found.length;
      empty.textContent =
        c.type === 'heading'
          ? 'Nel documento non ci sono titoli: usa gli stili "Titolo 1", "Titolo 2"… nella scheda Home.'
          : c.type === 'note'
            ? 'Nel documento non ci sono note a piè di pagina.'
            : `Nel documento non ci sono didascalie "${c.label}": aggiungine una con "Inserisci didascalia".`;
      ok.disabled = !found.length;
    };
    type.addEventListener('change', refresh);
    const submit = () => {
      const id = list.value;
      if (!id) return;
      done({ type: current().type, kind: kind.value as RefKind, id });
    };
    ok.addEventListener('click', submit);
    list.addEventListener('dblclick', submit);
    list.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        submit();
      }
    });
    body.append(field('Tipo di riferimento', type), field('Inserisci riferimento a', kind), el('div', 'refs-label', 'A quale elemento'), list, empty);
    refresh();
    buttons.append(ok);
    return type;
  });
}

// ---------------------------------------------------------------------------
// Segna voce di indice

export function openMarkEntryDialog(data: Parameters<ReferencesDialogs['markEntry']>[0]): Promise<MarkArgs | null> {
  return modal<MarkArgs>('Segna voce di indice', (body, buttons, done) => {
    const [main0, sub0] = splitEntry(data.entry);
    const main = el('input');
    main.name = 'main';
    main.value = main0;
    const sub = el('input');
    sub.name = 'sub';
    sub.value = sub0 ?? '';
    sub.placeholder = 'Facoltativa';
    const one = button('Segna', true);
    const all = button('Segna tutto');
    const refresh = () => {
      one.disabled = all.disabled = !main.value.trim();
    };
    main.addEventListener('input', refresh);
    const submit = (everywhere: boolean) => {
      if (!main.value.trim()) return;
      done({ entry: joinEntry(main.value, sub.value), all: everywhere });
    };
    one.addEventListener('click', () => submit(false));
    all.addEventListener('click', () => submit(true));
    for (const input of [main, sub]) {
      input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
          e.preventDefault();
          submit(false);
        }
      });
    }
    body.append(
      el('p', 'refs-intro', 'Il segno non si vede nel testo: la voce comparirà nell\'indice con il numero di pagina quando lo inserisci o lo aggiorni. "Segna tutto" segna ogni punto in cui compare la stessa parola.'),
      field('Voce principale', main),
      field('Sottovoce', sub),
    );
    refresh();
    buttons.append(one, all);
    return main;
  });
}

export const domDialogs: ReferencesDialogs = {
  caption: openCaptionDialog,
  crossRef: openCrossRefDialog,
  markEntry: openMarkEntryDialog,
};
