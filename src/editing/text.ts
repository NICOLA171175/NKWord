/**
 * Digitazione e cancellazione: funzioni pure stato → transazione.
 */

import {
  caret,
  isCollapsed,
  newId,
  paragraphLength,
  paragraphText,
  paragraphs,
  replaceBlock,
  resolveParaFormat,
  type Block,
  type CharFormat,
  type DocumentModel,
  type EditorState,
  type Inline,
  type ListRef,
  type Paragraph,
  type Transaction,
} from '../core';
import { typingFormat } from './format';
import { formatAt, inlinesLength, insertInParagraph, splitInlines, textToInlines } from './inlines';
import { deleteRange, insertBlocksAt, locate, sameContainer, withFreshIds } from './range';
import { selectionRange, stateOf } from './selection';
import { nextChar, prevChar, wordLeft, wordRight } from './words';

/** Cancella la selezione (se c'è) e restituisce documento e cursore risultanti. */
function removeSelection(state: EditorState): { doc: DocumentModel; at: { paragraphId: string; offset: number } } {
  const [start, end] = selectionRange(state);
  if (isCollapsed(state.selection)) return { doc: state.doc, at: start };
  return { doc: deleteRange(state.doc, start, end), at: start };
}

export function deleteSelection(state: EditorState): Transaction | null {
  if (isCollapsed(state.selection)) return null;
  const { doc, at } = removeSelection(state);
  return { state: stateOf(doc, caret(at.paragraphId, at.offset)), meta: { label: 'Elimina' } };
}

/**
 * Come in Word, il testo scritto all'inizio o alla fine di un collegamento
 * non ne fa parte: si toglie il collegamento (e lo stile "Hyperlink").
 */
function outsideLink(
  state: EditorState,
  p: Paragraph,
  offset: number,
  base: { format: CharFormat; styleId?: string },
): { format: CharFormat; styleId?: string } {
  const link = base.format.link;
  if (!link || !isCollapsed(state.selection)) return base;
  const before = offset > 0 ? formatAt(p, offset).format.link : undefined;
  const after = offset < paragraphLength(p) ? formatAt(p, offset + 1).format.link : undefined;
  if (before === link && after === link) return base;
  const { link: _l, ...format } = base.format;
  const style = base.styleId ? state.doc.styles.styles[base.styleId] : undefined;
  const isHyperlink = style?.type === 'character' && /^hyperlink$/i.test(style.name);
  return base.styleId && !isHyperlink ? { format, styleId: base.styleId } : { format };
}

/**
 * Inserisce testo al cursore, sostituendo la selezione. '\n' crea nuovi
 * paragrafi, '\t' tabulazioni. Il testo prende la formattazione del
 * carattere prima del cursore (o del primo carattere selezionato), più la
 * formattazione in attesa.
 */
export function insertText(state: EditorState, text: string): Transaction | null {
  if (!text && isCollapsed(state.selection)) return null;
  const [start] = selectionRange(state);
  const original = locate(state.doc, start.paragraphId).paragraph;
  const refOffset =
    !isCollapsed(state.selection) && start.offset < paragraphLength(original) ? start.offset + 1 : start.offset;
  const base = outsideLink(state, original, start.offset, formatAt(original, refOffset));
  const format = typingFormat(state.doc.styles, original, base.format, state.pendingFormat, base.styleId);

  const { doc, at } = removeSelection(state);
  const p = locate(doc, at.paragraphId).paragraph;
  const lines = text.replace(/\r\n?/g, '\n').split('\n');
  let result: { doc: DocumentModel; caret: { paragraphId: string; offset: number } };
  if (lines.length === 1) {
    const inlines = textToInlines(lines[0]!, format, base.styleId);
    result = {
      doc: replaceBlock(doc, p.id, [insertInParagraph(p, at.offset, inlines)]),
      caret: { paragraphId: p.id, offset: at.offset + inlinesLength(inlines) },
    };
  } else {
    const { id: _id, children: _c, ...attrs } = p;
    const blocks: Block[] = lines.map((line) => ({
      ...attrs,
      id: newId(),
      children: textToInlines(line, format, base.styleId),
    }));
    result = insertBlocksAt(doc, at, blocks);
  }
  return {
    state: stateOf(result.doc, caret(result.caret.paragraphId, result.caret.offset)),
    meta: { label: 'Digitazione', mergeKey: 'typing' },
  };
}

/** Inserisce un singolo inline non testuale (tab, interruzione di riga). */
export function insertInline(state: EditorState, inline: Inline, label: string): Transaction {
  const { doc, at } = removeSelection(state);
  const p = locate(doc, at.paragraphId).paragraph;
  return {
    state: stateOf(replaceBlock(doc, p.id, [insertInParagraph(p, at.offset, [inline])]), caret(p.id, at.offset + 1)),
    meta: { label, mergeKey: 'typing' },
  };
}

/**
 * Invio: divide il paragrafo al cursore. Il primo pezzo conserva l'id; il
 * nuovo paragrafo copia la formattazione e, se il cursore era alla fine,
 * prende lo stile "successivo" (es. Titolo 1 → Normale).
 */
export function splitParagraph(state: EditorState): Transaction {
  // Invio su una voce di elenco vuota: come Word risale di un livello o chiude l'elenco.
  if (isCollapsed(state.selection)) {
    const p = locate(state.doc, state.selection.focus.paragraphId).paragraph;
    const list = resolveParaFormat(state.doc.styles, p).list;
    if (list && paragraphLength(p) === 0) {
      const next = list.level > 0 ? withList(state, p, { ...list, level: list.level - 1 }) : withList(state, p, null);
      return { state: { ...state, doc: replaceBlock(state.doc, p.id, [next]) }, meta: { label: 'Elenco' } };
    }
  }
  const { doc, at } = removeSelection(state);
  const p = locate(doc, at.paragraphId).paragraph;
  const [before, after] = splitInlines(p.children, at.offset);
  const atEnd = at.offset === paragraphLength(p);
  const ref = formatAt(p, at.offset);
  const markFormat = typingFormat(doc.styles, p, ref.format, state.pendingFormat, ref.styleId);
  let second: Paragraph = { ...p, id: newId(), children: after, markFormat };
  if (atEnd && p.styleId) {
    const next = doc.styles.styles[p.styleId]?.next;
    if (next && next !== p.styleId) {
      const { styleId: _s, ...rest } = second;
      const { outlineLevel: _o, ...format } = second.format;
      second = next === doc.styles.defaultParagraphStyle ? { ...rest, format } : { ...rest, format, styleId: next };
    }
  }
  // Il paragrafo successivo non ripete l'interruzione di pagina prima.
  if (second.format.pageBreakBefore) {
    const { pageBreakBefore: _b, ...format } = second.format;
    second = { ...second, format };
  }
  const first: Paragraph = { ...p, children: before };
  const out = replaceBlock(doc, p.id, [first, second]);
  return { state: stateOf(out, caret(second.id, 0)), meta: { label: 'Nuovo paragrafo' } };
}

/** Il paragrafo con un altro riferimento di elenco (`null` = fuori dall'elenco). */
function withList(state: EditorState, p: Paragraph, list: ListRef | null): Paragraph {
  const { list: _l, ...format } = p.format;
  const fromStyle = resolveParaFormat(state.doc.styles, { ...p, format }).list;
  return { ...p, format: list || fromStyle ? { ...format, list } : format };
}

function neighbour(doc: DocumentModel, id: string, dir: -1 | 1) {
  const all = [...paragraphs(doc)];
  const i = all.findIndex((l) => l.paragraph.id === id);
  const self = all[i]!;
  const other = all[i + dir];
  return other && sameContainer(self, other) ? other : undefined;
}

/** Backspace. Con `unit: 'word'` (Ctrl+Backspace) cancella fino all'inizio della parola. */
export function deleteBackward(state: EditorState, unit: 'char' | 'word' = 'char'): Transaction | null {
  if (!isCollapsed(state.selection)) return deleteSelection(state);
  const pos = state.selection.focus;
  const p = locate(state.doc, pos.paragraphId).paragraph;
  // Backspace all'inizio di una voce di elenco toglie il punto elenco o il numero, come in Word.
  if (pos.offset === 0 && resolveParaFormat(state.doc.styles, p).list) {
    const doc = replaceBlock(state.doc, p.id, [withList(state, p, null)]);
    return { state: { ...state, doc }, meta: { label: 'Elenco' } };
  }
  if (pos.offset > 0) {
    const text = paragraphText(p);
    const from = unit === 'word' ? wordLeft(text, pos.offset) : prevChar(text, pos.offset);
    const doc = deleteRange(state.doc, { paragraphId: p.id, offset: from }, pos);
    return { state: stateOf(doc, caret(p.id, from)), meta: { label: 'Elimina', mergeKey: 'deleteBackward' } };
  }
  const prev = neighbour(state.doc, p.id, -1);
  if (!prev) return null;
  const at = { paragraphId: prev.paragraph.id, offset: paragraphLength(prev.paragraph) };
  const doc = deleteRange(state.doc, at, pos);
  return { state: stateOf(doc, caret(at.paragraphId, at.offset)), meta: { label: 'Elimina', mergeKey: 'deleteBackward' } };
}

/** Canc. Con `unit: 'word'` (Ctrl+Canc) cancella fino all'inizio della parola successiva. */
export function deleteForward(state: EditorState, unit: 'char' | 'word' = 'char'): Transaction | null {
  if (!isCollapsed(state.selection)) return deleteSelection(state);
  const pos = state.selection.focus;
  const p = locate(state.doc, pos.paragraphId).paragraph;
  const len = paragraphLength(p);
  if (pos.offset < len) {
    const text = paragraphText(p);
    const to = unit === 'word' ? wordRight(text, pos.offset) : nextChar(text, pos.offset);
    const doc = deleteRange(state.doc, pos, { paragraphId: p.id, offset: to });
    return { state: stateOf(doc, caret(p.id, pos.offset)), meta: { label: 'Elimina', mergeKey: 'deleteForward' } };
  }
  const next = neighbour(state.doc, p.id, 1);
  if (!next) return null;
  const doc = deleteRange(state.doc, pos, { paragraphId: next.paragraph.id, offset: 0 });
  return { state: stateOf(doc, caret(p.id, pos.offset)), meta: { label: 'Elimina', mergeKey: 'deleteForward' } };
}

/**
 * Incolla un frammento (già convertito dagli appunti). I blocchi ricevono
 * id nuovi; stili, elenchi e risorse mancanti vengono copiati nel documento.
 */
export function insertFragment(state: EditorState, fragment: DocumentModel): Transaction | null {
  const blocks = withFreshIds(fragment.sections.flatMap((s) => s.blocks));
  if (!blocks.length) return null;
  const { doc, at } = removeSelection(state);
  const merged: DocumentModel = {
    ...doc,
    styles: { ...doc.styles, styles: { ...fragment.styles.styles, ...doc.styles.styles } },
    lists: { ...fragment.lists, ...doc.lists },
    media: { ...fragment.media, ...doc.media },
  };
  const result = insertBlocksAt(merged, at, blocks);
  return {
    state: stateOf(result.doc, caret(result.caret.paragraphId, result.caret.offset)),
    meta: { label: 'Incolla' },
  };
}
