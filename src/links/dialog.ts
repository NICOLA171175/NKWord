/**
 * Finestra "Inserisci collegamento", come quella di Word: a sinistra il
 * tipo (pagina web, indirizzo email, punto di questo documento), in alto il
 * testo da visualizzare, sotto l'indirizzo o l'elenco dei titoli.
 */

import type { LinkDialogInput, LinkDialogResult } from './commands';
import { documentPlaces, linkKind, normalizeAddress, type LinkKind } from './model';

const KINDS: readonly { kind: LinkKind; label: string }[] = [
  { kind: 'web', label: 'Pagina web' },
  { kind: 'email', label: 'Indirizzo email' },
  { kind: 'place', label: 'Questo documento' },
];

/** Indirizzo email e oggetto da un collegamento `mailto:`. */
export function parseMailto(url: string): { address: string; subject: string } {
  const body = url.replace(/^mailto:/i, '');
  const [address = '', query = ''] = body.split('?');
  const subject = new URLSearchParams(query).get('subject') ?? '';
  return { address: decodeURIComponentSafe(address), subject };
}

export function buildMailto(address: string, subject: string): string | null {
  const base = normalizeAddress(address, 'email');
  if (!base) return null;
  return subject.trim() ? `${base}?subject=${encodeURIComponent(subject.trim())}` : base;
}

function decodeURIComponentSafe(s: string): string {
  try {
    return decodeURIComponent(s);
  } catch {
    return s;
  }
}

function field(labelText: string, control: HTMLElement): HTMLElement {
  const row = document.createElement('label');
  row.className = 'link-dialog-field';
  const span = document.createElement('span');
  span.textContent = labelText;
  row.append(span, control);
  return row;
}

function input(name: string, value = ''): HTMLInputElement {
  const el = document.createElement('input');
  el.type = 'text';
  el.name = name;
  el.value = value;
  el.spellcheck = false;
  return el;
}

/** Mostra la finestra; la promessa si risolve con la scelta (o `null` se annullata). */
export function openLinkDialog(init: LinkDialogInput): Promise<LinkDialogResult> {
  return new Promise((resolve) => {
    const editing = init.url !== undefined;
    let kind: LinkKind = init.url ? linkKind(init.url) : 'web';
    const mail = init.url && kind === 'email' ? parseMailto(init.url) : { address: '', subject: '' };

    const backdrop = document.createElement('div');
    backdrop.className = 'link-dialog-backdrop';
    const dialog = document.createElement('form');
    dialog.className = 'link-dialog';
    dialog.setAttribute('role', 'dialog');
    dialog.setAttribute('aria-label', editing ? 'Modifica collegamento' : 'Inserisci collegamento');
    backdrop.append(dialog);

    const title = document.createElement('h2');
    title.textContent = editing ? 'Modifica collegamento ipertestuale' : 'Inserisci collegamento ipertestuale';

    const textInput = input('text', init.text ?? '');
    if (init.text === undefined) {
      textInput.disabled = true;
      textInput.value = '<<Selezione nel documento>>';
    }

    const side = document.createElement('div');
    side.className = 'link-dialog-kinds';
    side.setAttribute('role', 'tablist');
    const kindButtons = new Map<LinkKind, HTMLButtonElement>();
    for (const k of KINDS) {
      const b = document.createElement('button');
      b.type = 'button';
      b.className = 'link-dialog-kind';
      b.dataset.kind = k.kind;
      b.setAttribute('role', 'tab');
      b.textContent = k.label;
      b.addEventListener('click', () => select(k.kind));
      kindButtons.set(k.kind, b);
      side.append(b);
    }

    // Pagina web
    const webInput = input('address', init.url && kind === 'web' ? init.url : '');
    webInput.placeholder = 'https://www.esempio.it';
    const webPane = field('Indirizzo:', webInput);

    // Indirizzo email
    const mailInput = input('email', mail.address);
    mailInput.placeholder = 'nome@esempio.it';
    const subjectInput = input('subject', mail.subject);
    const mailPane = document.createElement('div');
    mailPane.append(field('Indirizzo email:', mailInput), field('Oggetto:', subjectInput));

    // Questo documento
    const places = documentPlaces(init.doc);
    const placeList = document.createElement('select');
    placeList.name = 'place';
    placeList.size = Math.min(10, Math.max(4, places.length));
    placeList.className = 'link-dialog-places';
    for (const place of places) {
      const opt = document.createElement('option');
      opt.value = place.url;
      opt.textContent = `${' '.repeat(Math.max(0, place.level - 1))}${place.label}`;
      if (place.url === init.url) opt.selected = true;
      placeList.append(opt);
    }
    const placePane = document.createElement('div');
    const placeLabel = document.createElement('div');
    placeLabel.className = 'link-dialog-hint';
    placeLabel.textContent = places.length > 1 ? 'Scegli un punto del documento:' : 'Scegli un punto del documento (i titoli compaiono qui):';
    placePane.append(placeLabel, placeList);

    const panes = new Map<LinkKind, HTMLElement>([['web', webPane], ['email', mailPane], ['place', placePane]]);
    const content = document.createElement('div');
    content.className = 'link-dialog-content';
    content.append(field('Testo da visualizzare:', textInput), webPane, mailPane, placePane);

    const body = document.createElement('div');
    body.className = 'link-dialog-body';
    body.append(side, content);

    const error = document.createElement('div');
    error.className = 'link-dialog-error';
    error.setAttribute('role', 'alert');

    const footer = document.createElement('div');
    footer.className = 'link-dialog-footer';
    const removeButton = document.createElement('button');
    removeButton.type = 'button';
    removeButton.className = 'link-dialog-remove';
    removeButton.textContent = 'Rimuovi collegamento';
    removeButton.hidden = !editing;
    const cancel = document.createElement('button');
    cancel.type = 'button';
    cancel.textContent = 'Annulla';
    const ok = document.createElement('button');
    ok.type = 'submit';
    ok.className = 'link-dialog-ok';
    ok.textContent = 'OK';
    footer.append(removeButton, cancel, ok);

    dialog.append(title, body, error, footer);

    function select(next: LinkKind): void {
      kind = next;
      for (const [k, b] of kindButtons) {
        b.classList.toggle('active', k === kind);
        b.setAttribute('aria-selected', String(k === kind));
      }
      for (const [k, pane] of panes) pane.hidden = k !== kind;
      error.textContent = '';
      const first = kind === 'web' ? webInput : kind === 'email' ? mailInput : placeList;
      first.focus();
    }

    const finish = (result: LinkDialogResult) => {
      document.removeEventListener('keydown', onKey, true);
      backdrop.remove();
      resolve(result);
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key !== 'Escape') return;
      e.preventDefault();
      e.stopPropagation();
      finish(null);
    };

    const address = (): string | null => {
      if (kind === 'web') return normalizeAddress(webInput.value, 'web');
      if (kind === 'email') return buildMailto(mailInput.value, subjectInput.value);
      return placeList.value || null;
    };

    dialog.addEventListener('submit', (e) => {
      e.preventDefault();
      const url = address();
      if (!url) {
        error.textContent =
          kind === 'web' ? 'Scrivi l\'indirizzo della pagina.' : kind === 'email' ? 'Scrivi l\'indirizzo email.' : 'Scegli un punto del documento.';
        return;
      }
      finish(textInput.disabled ? { url } : { url, text: textInput.value });
    });
    placeList.addEventListener('dblclick', () => dialog.requestSubmit?.() ?? dialog.dispatchEvent(new Event('submit', { cancelable: true })));
    removeButton.addEventListener('click', () => finish({ remove: true }));
    cancel.addEventListener('click', () => finish(null));
    backdrop.addEventListener('mousedown', (e) => {
      if (e.target === backdrop) finish(null);
    });
    document.addEventListener('keydown', onKey, true);

    document.body.append(backdrop);
    select(kind);
  });
}
