/**
 * Comandi dei collegamenti ipertestuali.
 *
 * - `link.dialog`: finestra "Inserisci collegamento" (Ctrl+K), per inserire,
 *   modificare o rimuovere il collegamento alla selezione o al cursore.
 * - `link.insert` / `link.remove`: le stesse operazioni senza finestra.
 * - `link.open`: segue un collegamento (Ctrl+clic): i punti del documento
 *   spostano il cursore, gli indirizzi web ed email si aprono fuori dall'app.
 * - `link.auto`: collegamento creato riconoscendo un indirizzo appena
 *   scritto; è un passo a sé, così Ctrl+Z lo annulla lasciando il testo.
 */

import {
  caret,
  findParagraph,
  isCollapsed,
  orderedRange,
  paragraphLength,
  paragraphText,
  replaceBlock,
  type AnyCommand,
  type CharFormat,
  type DocumentModel,
  type Editor,
  type EditorState,
  type Id,
  type Inline,
  type Position,
  type Transaction,
} from '../core';
import { deleteInParagraph, formatAt, insertInParagraph, mapFormatInRange, textToInlines } from '../editing/inlines';
import { paragraphsBetween } from '../editing/range';
import { displayText, ensureHyperlinkStyle, hyperlinkStyleId, isExternalSafe, linkAt, linkInside, placeTarget } from './model';

export interface LinkSpan {
  readonly paragraphId: Id;
  readonly start: number;
  readonly end: number;
}

export interface LinkInsertArgs {
  readonly url: string;
  /** Testo da mostrare; se diverso da quello presente lo sostituisce. */
  readonly text?: string;
  /** Testo da collegare; di default la selezione o il collegamento al cursore. */
  readonly range?: LinkSpan;
}

export interface LinkCommandArgs {
  'link.dialog': void;
  'link.insert': LinkInsertArgs;
  'link.remove': { readonly range?: LinkSpan } | void;
  'link.open': { readonly url: string };
  'link.auto': LinkSpan & { readonly url: string };
}

export type LinkCommandId = keyof LinkCommandArgs;

/** Valori iniziali e risultato della finestra. */
export interface LinkDialogInput {
  readonly doc: DocumentModel;
  /** Indirizzo del collegamento che si modifica, se c'è. */
  readonly url?: string;
  /** Testo da mostrare; `undefined` quando la selezione non è un testo semplice. */
  readonly text?: string;
}

export type LinkDialogResult = { readonly url: string; readonly text?: string } | { readonly remove: true } | null;

export interface LinkCommandOptions {
  /** Mostra la finestra; di default quella del modulo (`openLinkDialog`). */
  readonly openDialog?: (input: LinkDialogInput) => Promise<LinkDialogResult>;
  /** Apre un indirizzo esterno; di default una nuova finestra del browser (in Electron il browser di sistema). */
  readonly openExternal?: (url: string) => void;
  /** Avvisa dopo aver portato il cursore a un punto del documento (per scorrere fin lì). */
  readonly onNavigate?: () => void;
}

const span = (r: LinkSpan): [Position, Position] => [
  { paragraphId: r.paragraphId, offset: r.start },
  { paragraphId: r.paragraphId, offset: r.end },
];

/** Intervallo su cui agire: quello dato, la selezione o il collegamento al cursore. */
function targetRange(state: EditorState, range?: LinkSpan): [Position, Position] | null {
  if (range) return span(range);
  if (!isCollapsed(state.selection)) return orderedRange(state.doc, state.selection);
  const link = linkInside(state.doc, state.selection.focus);
  return link ? span(link) : null;
}

function withoutLink(format: CharFormat): CharFormat {
  const { link: _l, ...rest } = format;
  return rest;
}

function textBetween(doc: DocumentModel, [a, b]: [Position, Position]): string | undefined {
  if (a.paragraphId !== b.paragraphId) return undefined;
  const p = findParagraph(doc, a.paragraphId)?.paragraph;
  return p ? paragraphText(p).slice(a.offset, b.offset) : undefined;
}

/** Testo da mostrare nella finestra per la selezione: solo se è testo semplice in un paragrafo. */
export function selectionText(state: EditorState): string | undefined {
  const range = targetRange(state);
  if (!range) return '';
  const text = textBetween(state.doc, range);
  return text !== undefined && !/[\t\n￼]/.test(text) ? text : undefined;
}

export function insertLink(state: EditorState, args: LinkInsertArgs, label = 'Collegamento'): Transaction | null {
  const url = args.url.trim();
  if (!url) return null;
  const { sheet, styleId } = ensureHyperlinkStyle(state.doc.styles);
  let doc: DocumentModel = sheet === state.doc.styles ? state.doc : { ...state.doc, styles: sheet };
  const range = targetRange(state, args.range);
  const text = args.text?.replace(/[\t\n]/g, ' ');

  if (range && range[0].paragraphId === range[1].paragraphId && text && text !== textBetween(doc, range)) {
    // Testo nuovo al posto del vecchio, con il formato del primo carattere.
    const [a, b] = range;
    const p = findParagraph(doc, a.paragraphId)!.paragraph;
    const base = formatAt(p, Math.min(a.offset + 1, paragraphLength(p)));
    const inlines = textToInlines(text, { ...withoutLink(base.format), link: url }, styleId);
    doc = replaceBlock(doc, p.id, [insertInParagraph(deleteInParagraph(p, a.offset, b.offset), a.offset, inlines)]);
    return { state: { doc, selection: caret(p.id, a.offset + text.length) }, meta: { label } };
  }
  if (range) {
    const [a, b] = range;
    for (const { paragraph: p } of paragraphsBetween(doc, a, b)) {
      const from = p.id === a.paragraphId ? a.offset : 0;
      const to = p.id === b.paragraphId ? b.offset : paragraphLength(p);
      doc = replaceBlock(doc, p.id, [mapFormatInRange(p, from, to, (format) => ({ format: { ...format, link: url }, styleId }))]);
    }
    return { state: { ...state, doc }, meta: { label } };
  }
  // Nessun testo selezionato: si inserisce il testo (o l'indirizzo stesso) al cursore.
  const pos = state.selection.focus;
  const p = findParagraph(doc, pos.paragraphId)?.paragraph;
  if (!p) return null;
  const shown = text || displayText(url);
  const base = formatAt(p, pos.offset);
  const inlines = textToInlines(shown, { ...withoutLink({ ...base.format, ...state.pendingFormat }), link: url }, styleId);
  doc = replaceBlock(doc, p.id, [insertInParagraph(p, pos.offset, inlines)]);
  return { state: { doc, selection: caret(p.id, pos.offset + shown.length) }, meta: { label } };
}

export function removeLink(state: EditorState, range?: LinkSpan): Transaction | null {
  const target = targetRange(state, range);
  if (!target) return null;
  const hyperlink = hyperlinkStyleId(state.doc.styles);
  const [a, b] = target;
  let doc = state.doc;
  let changed = false;
  for (const { paragraph: p } of paragraphsBetween(doc, a, b)) {
    const from = p.id === a.paragraphId ? a.offset : 0;
    const to = p.id === b.paragraphId ? b.offset : paragraphLength(p);
    const next = mapFormatInRange(p, from, to, (format, inline: Inline) => {
      const styleId = inline.type === 'text' && inline.styleId !== hyperlink ? inline.styleId : undefined;
      if (format.link || styleId !== (inline.type === 'text' ? inline.styleId : undefined)) changed = true;
      return { format: withoutLink(format), styleId };
    });
    doc = replaceBlock(doc, p.id, [next]);
  }
  return changed ? { state: { ...state, doc }, meta: { label: 'Rimuovi collegamento' } } : null;
}

/** Apertura predefinita: nuova scheda del browser; nel programma desktop Electron passa al browser di sistema. */
export function defaultOpenExternal(url: string): void {
  globalThis.open?.(url, '_blank', 'noopener');
}

export function createLinkCommands(editor: Editor, options: LinkCommandOptions = {}): AnyCommand[] {
  const openExternal = options.openExternal ?? defaultOpenExternal;
  return [
    {
      id: 'link.dialog',
      label: 'Collegamento',
      async: true,
      isEnabled: () => !!options.openDialog,
      run: async ({ state }) => {
        const link = isCollapsed(state.selection) ? linkInside(state.doc, state.selection.focus) : null;
        const fromSelection = !isCollapsed(state.selection) ? targetRange(state) : null;
        const existing = link ?? (fromSelection ? linkAt(state.doc, fromSelection[0]) : null);
        const sameLink = existing && fromSelection
          ? fromSelection[0].paragraphId === existing.paragraphId && fromSelection[0].offset >= existing.start && fromSelection[1].offset <= existing.end
          : !!existing;
        const input: LinkDialogInput = sameLink && existing
          ? { doc: state.doc, url: existing.url, text: existing.text }
          : { doc: state.doc, text: selectionText(state) };
        const result = await options.openDialog!(input);
        if (!result) return null;
        // La finestra è modale: lo stato è quello di prima, ma si rilegge per sicurezza.
        const now = editor.state;
        const range = sameLink && existing ? { paragraphId: existing.paragraphId, start: existing.start, end: existing.end } : undefined;
        if ('remove' in result) return removeLink(now, range);
        return insertLink(now, { url: result.url, ...(result.text !== undefined ? { text: result.text } : {}), ...(range ? { range } : {}) });
      },
    },
    {
      id: 'link.insert',
      label: 'Inserisci collegamento',
      isEnabled: (_c, a: LinkInsertArgs) => !!a?.url?.trim(),
      run: ({ state }, a: LinkInsertArgs) => insertLink(state, a),
    },
    {
      id: 'link.remove',
      label: 'Rimuovi collegamento',
      isEnabled: ({ state }, a?: { range?: LinkSpan }) => !!targetRange(state, a?.range),
      run: ({ state }, a?: { range?: LinkSpan }) => removeLink(state, a?.range),
    },
    {
      id: 'link.auto',
      label: 'Collegamento automatico',
      run: ({ state }, a: LinkSpan & { url: string }) => insertLink(state, { url: a.url, range: a }, 'Formattazione automatica'),
    },
    {
      id: 'link.open',
      label: 'Apri collegamento',
      async: true,
      isEnabled: ({ state }, a: { url: string }) => !!a?.url && (a.url.startsWith('#') ? !!placeTarget(state.doc, a.url) : isExternalSafe(a.url)),
      run: async ({ state }, a: { url: string }) => {
        if (!a.url.startsWith('#')) {
          openExternal(a.url.trim());
          return null;
        }
        const target = placeTarget(state.doc, a.url)!;
        queueMicrotask(() => options.onNavigate?.());
        return { state: { ...state, selection: caret(target, 0) }, meta: { label: 'Vai al collegamento', addToHistory: false } };
      },
    },
  ];
}
