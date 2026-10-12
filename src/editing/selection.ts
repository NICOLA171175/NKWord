/** Helper di selezione e di stato condivisi dai comandi di editing. */

import {
  caret,
  findParagraph,
  isCollapsed,
  orderedRange,
  paragraphLength,
  paragraphs,
  type DocumentModel,
  type EditorState,
  type Position,
  type Selection,
} from '../core';

export function selectionRange(state: EditorState): [Position, Position] {
  return orderedRange(state.doc, state.selection);
}

/** Nuovo stato con selezione data; la formattazione in attesa si azzera. */
export function stateOf(doc: DocumentModel, selection: Selection): EditorState {
  return { doc, selection };
}

export function sameSelection(a: Selection, b: Selection): boolean {
  return (
    a.anchor.paragraphId === b.anchor.paragraphId &&
    a.anchor.offset === b.anchor.offset &&
    a.focus.paragraphId === b.focus.paragraphId &&
    a.focus.offset === b.focus.offset
  );
}

/** Riporta una posizione dentro il documento; null se il paragrafo non esiste. */
export function clampPosition(doc: DocumentModel, pos: Position): Position | null {
  const loc = findParagraph(doc, pos.paragraphId);
  if (!loc) return null;
  const len = paragraphLength(loc.paragraph);
  const offset = Math.max(0, Math.min(len, Math.floor(pos.offset)));
  return offset === pos.offset ? pos : { paragraphId: pos.paragraphId, offset };
}

export function docStart(doc: DocumentModel): Position {
  const first = paragraphs(doc).next();
  if (first.done) throw new Error('Documento senza paragrafi');
  return { paragraphId: first.value.paragraph.id, offset: 0 };
}

export function docEnd(doc: DocumentModel): Position {
  let last: Position | undefined;
  for (const loc of paragraphs(doc)) last = { paragraphId: loc.paragraph.id, offset: paragraphLength(loc.paragraph) };
  if (!last) throw new Error('Documento senza paragrafi');
  return last;
}

export { caret, isCollapsed };
