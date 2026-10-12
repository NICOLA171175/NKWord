/** Spostamento del cursore e gestione della selezione. */

import {
  caret,
  isCollapsed,
  paragraphLength,
  paragraphText,
  paragraphs,
  type CommandArgs,
  type DocumentModel,
  type EditorState,
  type Position,
  type Selection,
  type Transaction,
} from '../core';
import { clampPosition, docEnd, docStart, sameSelection, selectionRange, stateOf } from './selection';
import { nextChar, prevChar, wordLeft, wordRight } from './words';

export type CaretDirection = CommandArgs['edit.moveCaret']['direction'];

/** Inizio e fine della riga visiva che contiene la posizione (dal layout, se c'è). */
export type LineBounds = (doc: DocumentModel, pos: Position) => [number, number] | null;

function selectionTr(state: EditorState, selection: Selection): Transaction | null {
  if (sameSelection(selection, state.selection) && !state.pendingFormat) return null;
  return { state: stateOf(state.doc, selection), meta: { addToHistory: false } };
}

function adjacentParagraph(doc: DocumentModel, id: string, dir: -1 | 1): Position | null {
  const all = [...paragraphs(doc)];
  const i = all.findIndex((l) => l.paragraph.id === id);
  const other = all[i + dir]?.paragraph;
  if (!other) return null;
  return { paragraphId: other.id, offset: dir < 0 ? paragraphLength(other) : 0 };
}

function step(doc: DocumentModel, pos: Position, direction: CaretDirection, lineBounds?: LineBounds): Position {
  const p = [...paragraphs(doc)].find((l) => l.paragraph.id === pos.paragraphId)!.paragraph;
  const text = paragraphText(p);
  switch (direction) {
    case 'left':
      return pos.offset > 0
        ? { paragraphId: p.id, offset: prevChar(text, pos.offset) }
        : (adjacentParagraph(doc, p.id, -1) ?? pos);
    case 'right':
      return pos.offset < text.length
        ? { paragraphId: p.id, offset: nextChar(text, pos.offset) }
        : (adjacentParagraph(doc, p.id, 1) ?? pos);
    case 'wordLeft':
      return pos.offset > 0
        ? { paragraphId: p.id, offset: wordLeft(text, pos.offset) }
        : (adjacentParagraph(doc, p.id, -1) ?? pos);
    case 'wordRight':
      return pos.offset < text.length
        ? { paragraphId: p.id, offset: wordRight(text, pos.offset) }
        : (adjacentParagraph(doc, p.id, 1) ?? pos);
    case 'lineStart':
      return { paragraphId: p.id, offset: lineBounds?.(doc, pos)?.[0] ?? 0 };
    case 'lineEnd':
      return { paragraphId: p.id, offset: lineBounds?.(doc, pos)?.[1] ?? text.length };
    case 'docStart':
      return docStart(doc);
    case 'docEnd':
      return docEnd(doc);
  }
}

export function moveCaret(
  state: EditorState,
  direction: CaretDirection,
  extend = false,
  lineBounds?: LineBounds,
): Transaction | null {
  const { doc, selection } = state;
  if (!extend && !isCollapsed(selection) && (direction === 'left' || direction === 'right')) {
    // Con una selezione, freccia sinistra/destra la chiude sul bordo.
    const [start, end] = selectionRange(state);
    const at = direction === 'left' ? start : end;
    return selectionTr(state, caret(at.paragraphId, at.offset));
  }
  const focus = step(doc, selection.focus, direction, lineBounds);
  return selectionTr(state, extend ? { anchor: selection.anchor, focus } : { anchor: focus, focus });
}

export function selectAll(state: EditorState): Transaction | null {
  return selectionTr(state, { anchor: docStart(state.doc), focus: docEnd(state.doc) });
}

export function setSelection(state: EditorState, selection: Selection): Transaction | null {
  const anchor = clampPosition(state.doc, selection.anchor);
  const focus = clampPosition(state.doc, selection.focus);
  if (!anchor || !focus) return null;
  return selectionTr(state, { anchor, focus });
}
