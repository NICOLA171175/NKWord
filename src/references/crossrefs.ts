/**
 * Riferimenti incrociati (Riferimenti > Riferimento incrociato), come in
 * Word: un pezzo di testo che mostra il titolo, la didascalia, il numero di
 * nota o la pagina di un altro punto del documento e si aggiorna.
 *
 * Nel modello è testo normale con due proprietà del carattere:
 * - `crossRef` = "<cosa mostra>:<tipo>:<id>" dice come rigenerarlo;
 * - `link` = `#_Toc<id del paragrafo>` lo rende cliccabile verso la
 *   destinazione (stessa convenzione del sommario: in DOCX è un segnalibro).
 */

import {
  caret,
  findParagraph,
  footnoteNumbers,
  inlineLength,
  isCollapsed,
  paragraphLength,
  paragraphs,
  replaceBlock,
  textRun,
  type CharFormat,
  type DocumentModel,
  type EditorState,
  type Id,
  type Paragraph,
  type Transaction,
  type TransactionFilter,
} from '../core';
import { footnoteRefs, footnoteText } from '../footnotes/model';
import { formatAt, insertInParagraph, mapFormatInRange } from '../editing/inlines';
import { deleteSelection } from '../editing/text';
import { collectHeadings, tocAnchor } from '../toc/toc';
import { collectCaptions } from './captions';
import { readingOrder, replaceText } from './common';

export type RefTargetType = 'heading' | 'caption' | 'note';
export type RefKind = 'text' | 'label' | 'caption' | 'number' | 'page' | 'position';

export interface CrossRef {
  readonly kind: RefKind;
  readonly type: RefTargetType;
  /** Id del paragrafo (titolo, didascalia) o della nota. */
  readonly id: Id;
}

/** Testo mostrato quando la destinazione non esiste più, come in Word. */
export const MISSING_TARGET = "Errore. L'origine riferimento non è stata trovata.";

/** Cosa si può mostrare per ogni tipo di destinazione, con le parole di Word. */
export const REF_KINDS: Readonly<Record<RefTargetType, readonly { kind: RefKind; label: string }[]>> = {
  heading: [
    { kind: 'text', label: 'Testo del titolo' },
    { kind: 'page', label: 'Numero di pagina' },
    { kind: 'position', label: 'Sopra/sotto' },
  ],
  caption: [
    { kind: 'label', label: 'Solo etichetta e numero' },
    { kind: 'text', label: 'Intera didascalia' },
    { kind: 'caption', label: 'Solo testo didascalia' },
    { kind: 'page', label: 'Numero di pagina' },
    { kind: 'position', label: 'Sopra/sotto' },
  ],
  note: [
    { kind: 'number', label: 'Numero nota' },
    { kind: 'page', label: 'Numero di pagina' },
    { kind: 'position', label: 'Sopra/sotto' },
  ],
};

const TYPES: readonly RefTargetType[] = ['heading', 'caption', 'note'];

export function encodeCrossRef(ref: CrossRef): string {
  return `${ref.kind}:${ref.type}:${ref.id}`;
}

export function parseCrossRef(value: string | undefined): CrossRef | null {
  if (!value) return null;
  const [kind, type, ...rest] = value.split(':');
  const id = rest.join(':');
  if (!id || !TYPES.includes(type as RefTargetType)) return null;
  if (!REF_KINDS[type as RefTargetType].some((k) => k.kind === kind)) return null;
  return { kind: kind as RefKind, type: type as RefTargetType, id };
}

// ---------------------------------------------------------------------------
// Destinazioni

export interface RefTarget {
  readonly type: RefTargetType;
  readonly id: Id;
  /** Paragrafo del corpo a cui porta il collegamento. */
  readonly paragraphId: Id;
  /** Etichetta della didascalia ("Figura"); solo per le didascalie. */
  readonly label?: string;
  /** Riga mostrata nell'elenco della finestra. */
  readonly display: string;
  /** Livello del titolo (rientro nell'elenco). */
  readonly level?: number;
}

/** Tutto ciò a cui si può fare riferimento, in ordine di lettura. */
export function listTargets(doc: DocumentModel): RefTarget[] {
  const out: RefTarget[] = [];
  for (const h of collectHeadings(doc, 9)) out.push({ type: 'heading', id: h.paragraphId, paragraphId: h.paragraphId, display: h.text, level: h.level });
  for (const c of collectCaptions(doc)) out.push({ type: 'caption', id: c.paragraphId, paragraphId: c.paragraphId, label: c.label, display: c.text });
  const numbers = footnoteNumbers(doc);
  for (const { noteId, position } of footnoteRefs(doc)) {
    const n = numbers.get(noteId);
    if (n === undefined || out.some((t) => t.type === 'note' && t.id === noteId)) continue;
    const text = (footnoteText(doc, noteId) ?? '').replace(/\s+/g, ' ').trim();
    out.push({ type: 'note', id: noteId, paragraphId: position.paragraphId, display: text ? `${n} ${text}` : String(n) });
  }
  return out;
}

/** Paragrafo della destinazione (per le note: quello che contiene il richiamo). */
function targetParagraph(doc: DocumentModel, ref: CrossRef): Id | undefined {
  if (ref.type !== 'note') return findParagraph(doc, ref.id) ? ref.id : undefined;
  return footnoteRefs(doc).find((r) => r.noteId === ref.id)?.position.paragraphId;
}

export interface RefContext {
  readonly pages: ReadonlyMap<Id, number>;
  /** Ordine di lettura dei paragrafi, per "sopra"/"sotto". */
  readonly order?: ReadonlyMap<Id, number>;
}

/** Testo del riferimento posto nel paragrafo `at`. */
export function crossRefText(doc: DocumentModel, ref: CrossRef, at: Id, ctx: RefContext): string {
  const target = targetParagraph(doc, ref);
  if (!target) return MISSING_TARGET;
  switch (ref.kind) {
    case 'page': {
      const page = ctx.pages.get(target);
      return page === undefined ? MISSING_TARGET : String(page);
    }
    case 'position': {
      const order = ctx.order ?? readingOrder(doc);
      return (order.get(target) ?? 0) <= (order.get(at) ?? 0) ? 'sopra' : 'sotto';
    }
    case 'number': {
      const n = footnoteNumbers(doc).get(ref.id);
      return n === undefined ? MISSING_TARGET : String(n);
    }
  }
  if (ref.type === 'heading') return collectHeadings(doc, 9).find((h) => h.paragraphId === ref.id)?.text ?? MISSING_TARGET;
  if (ref.type === 'caption') {
    const c = collectCaptions(doc).find((x) => x.paragraphId === ref.id);
    if (!c) return MISSING_TARGET;
    if (ref.kind === 'label') return `${c.label} ${c.number}`;
    if (ref.kind === 'caption') return c.caption || `${c.label} ${c.number}`;
    return c.text;
  }
  return MISSING_TARGET;
}

// ---------------------------------------------------------------------------
// Inserimento e aggiornamento

export function canInsertCrossRef(state: EditorState): boolean {
  return !!findParagraph(state.doc, state.selection.focus.paragraphId);
}

/** Inserisce il riferimento al cursore (al posto della selezione), con il formato del testo intorno. */
export function insertCrossRef(state: EditorState, ref: CrossRef, pages: ReadonlyMap<Id, number>): Transaction | null {
  const target = targetParagraph(state.doc, ref);
  if (!target) return null;
  const cleared = isCollapsed(state.selection) ? null : deleteSelection(state);
  const doc = cleared?.state.doc ?? state.doc;
  const at = cleared?.state.selection.focus ?? state.selection.focus;
  const p = findParagraph(doc, at.paragraphId)!.paragraph;
  const base = formatAt(p, at.offset);
  const { link: _l, crossRef: _c, indexEntry: _i, ...plain } = base.format;
  const text = crossRefText(doc, ref, p.id, { pages });
  const format: CharFormat = { ...plain, ...state.pendingFormat, link: tocAnchor(target), crossRef: encodeCrossRef(ref) };
  const run = base.styleId ? { ...textRun(text, format), styleId: base.styleId } : textRun(text, format);
  const next = insertInParagraph(p, at.offset, [run]);
  return {
    state: { doc: replaceBlock(doc, p.id, [next]), selection: caret(p.id, at.offset + text.length) },
    meta: { label: 'Riferimento incrociato' },
  };
}

interface Span {
  readonly from: number;
  readonly to: number;
  readonly value: string;
  readonly format: CharFormat;
  readonly styleId?: Id;
}

/** Tratti consecutivi di testo con lo stesso `crossRef`. */
export function crossRefSpans(p: Paragraph): Span[] {
  const out: Span[] = [];
  let pos = 0;
  for (const c of p.children) {
    const len = inlineLength(c);
    const value = c.type === 'text' ? c.format.crossRef : undefined;
    const last = out[out.length - 1];
    if (value && last && last.value === value && last.to === pos) out[out.length - 1] = { ...last, to: pos + len };
    else if (value && c.type === 'text') out.push({ from: pos, to: pos + len, value, format: c.format, ...(c.styleId ? { styleId: c.styleId } : {}) });
    pos += len;
  }
  return out;
}

/** Rigenera il testo di tutti i riferimenti incrociati del documento. */
export function updateCrossRefs(doc: DocumentModel, pages: ReadonlyMap<Id, number>): DocumentModel {
  const order = readingOrder(doc);
  let out = doc;
  for (const { paragraph: p } of paragraphs(doc)) {
    const spans = crossRefSpans(p);
    if (!spans.length) continue;
    let next = p;
    // Dalla fine, così gli offset dei tratti precedenti restano validi.
    for (const span of [...spans].reverse()) {
      const ref = parseCrossRef(span.value);
      if (!ref) continue;
      const text = crossRefText(doc, ref, p.id, { pages, order });
      const current = spanText(next, span);
      const target = targetParagraph(doc, ref);
      const format: CharFormat = target ? { ...span.format, link: tocAnchor(target) } : span.format;
      if (current === text && format.link === span.format.link) continue;
      next = replaceText(next, span.from, span.to, text, format, span.styleId);
    }
    if (next !== p) out = replaceBlock(out, p.id, [next]);
  }
  return out;
}

function spanText(p: Paragraph, span: Span): string {
  let pos = 0;
  let text = '';
  for (const c of p.children) {
    const len = inlineLength(c);
    if (pos >= span.from && pos + len <= span.to) text += c.type === 'text' ? c.text : '\u0000';
    pos += len;
  }
  return text;
}

/**
 * Il testo scritto appena prima o appena dopo un riferimento non ne fa
 * parte (altrimenti l'aggiornamento lo cancellerebbe): gli si tolgono
 * `crossRef` e collegamento. Scrivere dentro il riferimento invece lo
 * modifica, come in Word, fino al prossimo aggiornamento.
 */
export const crossRefTypingFilter: TransactionFilter = (tr, previous) => {
  if (tr.meta?.mergeKey !== 'typing' || !isCollapsed(previous.selection) || !isCollapsed(tr.state.selection)) return tr;
  const at = previous.selection.focus;
  const end = tr.state.selection.focus;
  if (end.paragraphId !== at.paragraphId || end.offset <= at.offset) return tr;
  const before = findParagraph(previous.doc, at.paragraphId)?.paragraph;
  const after = findParagraph(tr.state.doc, at.paragraphId)?.paragraph;
  if (!before || !after) return tr;
  const left = at.offset > 0 ? formatAt(before, at.offset).format.crossRef : undefined;
  const right = at.offset < paragraphLength(before) ? formatAt(before, at.offset + 1).format.crossRef : undefined;
  const typed = formatAt(after, end.offset).format.crossRef;
  if (!typed || (left === typed && right === typed)) return tr;
  const link = formatAt(after, end.offset).format.link;
  const cleaned = mapFormatInRange(after, at.offset, end.offset, (format, inline) => {
    const { crossRef: _c, link: l, ...rest } = format;
    const keepLink = l && l !== link ? { link: l } : {};
    const styleId = inline.type === 'text' ? inline.styleId : undefined;
    return { format: { ...rest, ...keepLink }, styleId };
  });
  return { ...tr, state: { ...tr.state, doc: replaceBlock(tr.state.doc, after.id, [cleaned]) } };
};
