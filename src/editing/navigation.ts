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

/**
 * Inizio e fine della riga visiva che contiene la posizione (dal layout, se
 * c'è). Il terzo valore, se presente, è dove si ferma il cursore con Fine:
 * prima dello spazio o dell'a capo che chiude la riga, così resta su di essa.
 */
export type LineBounds = (doc: DocumentModel, pos: Position) => readonly [number, number, number?] | null;

/** Riga sopra o sotto (`distance` assente) o spostamento di una schermata, dal layout. */
export type VerticalMove = (doc: DocumentModel, pos: Position, dir: -1 | 1, distance?: number) => Position | null;

/** Spostamenti che dipendono dall'impaginazione. */
export interface LayoutNavigation {
  readonly lineBounds?: LineBounds;
  readonly vertical?: VerticalMove;
}

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

/** Senza impaginazione la "riga" sopra o sotto è il paragrafo vicino, alla stessa distanza dall'inizio. */
function paragraphLine(doc: DocumentModel, pos: Position, dir: -1 | 1): Position | null {
  const all = [...paragraphs(doc)];
  const other = all[all.findIndex((l) => l.paragraph.id === pos.paragraphId) + dir]?.paragraph;
  return other ? { paragraphId: other.id, offset: Math.min(pos.offset, paragraphLength(other)) } : null;
}

function step(
  doc: DocumentModel,
  pos: Position,
  direction: CaretDirection,
  nav: LayoutNavigation,
  extend: boolean,
  distance?: number,
): Position {
  const { lineBounds, vertical } = nav;
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
    case 'lineEnd': {
      // Maiusc+Fine seleziona anche lo spazio finale; Fine lascia il cursore sulla riga.
      const b = lineBounds?.(doc, pos);
      return { paragraphId: p.id, offset: b ? (extend ? b[1] : (b[2] ?? b[1])) : text.length };
    }
    case 'lineUp':
    case 'lineDown': {
      const dir = direction === 'lineUp' ? -1 : 1;
      if (vertical) {
        // Sulla prima (ultima) riga Word porta il cursore all'inizio (alla fine) del documento.
        return vertical(doc, pos, dir) ?? (dir < 0 ? docStart(doc) : docEnd(doc));
      }
      return paragraphLine(doc, pos, dir) ?? (dir < 0 ? docStart(doc) : docEnd(doc));
    }
    case 'pageUp':
    case 'pageDown': {
      const dir = direction === 'pageUp' ? -1 : 1;
      const moved = vertical?.(doc, pos, dir, distance ?? 720);
      return moved ?? (dir < 0 ? docStart(doc) : docEnd(doc));
    }
    case 'paragraphUp':
      // Ctrl+Su: inizio del paragrafo, o del precedente se si è già all'inizio.
      return pos.offset > 0 ? { paragraphId: p.id, offset: 0 } : { ...(adjacentParagraph(doc, p.id, -1) ?? pos), offset: 0 };
    case 'paragraphDown':
      // Ctrl+Giù: inizio del paragrafo seguente; nell'ultimo, la sua fine.
      return adjacentParagraph(doc, p.id, 1) ?? { paragraphId: p.id, offset: text.length };
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
  nav: LayoutNavigation = {},
  distance?: number,
): Transaction | null {
  const { doc, selection } = state;
  let from = selection.focus;
  if (!extend && !isCollapsed(selection)) {
    const [start, end] = selectionRange(state);
    // Con una selezione, freccia sinistra/destra la chiude sul bordo.
    if (direction === 'left' || direction === 'right') {
      const at = direction === 'left' ? start : end;
      return selectionTr(state, caret(at.paragraphId, at.offset));
    }
    // Su e Giù partono dal bordo della selezione nel verso dello spostamento.
    if (direction === 'lineUp' || direction === 'pageUp') from = start;
    if (direction === 'lineDown' || direction === 'pageDown') from = end;
  }
  const focus = step(doc, from, direction, nav, extend, distance);
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
