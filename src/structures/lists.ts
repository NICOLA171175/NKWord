/**
 * Elenchi puntati e numerati: `list.toggleBullet`, `list.toggleNumbered`,
 * `list.increaseLevel`, `list.decreaseLevel`.
 */

import {
  newId,
  replaceBlock,
  type DocumentModel,
  type EditorState,
  type ListDefinition,
  type ListLevel,
  type NumberFormat,
  type Paragraph,
  type ParaFormat,
} from '../core';
import { defineCommand, requireParagraph, selectedParagraphs, containerOf } from './util';

export const MAX_LIST_LEVEL = 8;
/** Rientro per livello (0,63 cm come Word). */
export const LIST_INDENT_STEP = 18;
export const LIST_HANGING = 18;

const BULLETS = ['•', '◦', '▪'];
const NUMBER_CYCLE: NumberFormat[] = ['decimal', 'lowerLetter', 'lowerRoman'];

function level(format: NumberFormat, text: string, i: number): ListLevel {
  return { format, text, start: 1, indentLeft: LIST_INDENT_STEP * (i + 1), hanging: LIST_HANGING };
}

export function bulletList(): ListDefinition {
  return {
    id: newId(),
    levels: Array.from({ length: MAX_LIST_LEVEL + 1 }, (_, i) => level('bullet', BULLETS[i % BULLETS.length]!, i)),
  };
}

/** Elenco numerato: il primo livello usa `format`, i successivi alternano 1, a, i. */
export function numberedList(format: NumberFormat = 'decimal'): ListDefinition {
  return {
    id: newId(),
    levels: Array.from({ length: MAX_LIST_LEVEL + 1 }, (_, i) => {
      const f = i === 0 ? format : NUMBER_CYCLE[i % NUMBER_CYCLE.length]!;
      return level(f, `%${i + 1}.`, i);
    }),
  };
}

/** Tipo di elenco: 'bullet' oppure il formato del primo livello. */
export function listKind(def: ListDefinition | undefined): NumberFormat | undefined {
  return def?.levels[0]?.format;
}

function paragraphKind(doc: DocumentModel, p: Paragraph): NumberFormat | undefined {
  const ref = p.format.list;
  return ref ? listKind(doc.lists[ref.listId]) : undefined;
}

/** Toglie i rientri diretti, che altrimenti coprirebbero quelli del livello. */
function withoutIndent(format: ParaFormat): ParaFormat {
  const { indentLeft: _l, indentFirstLine: _f, ...rest } = format;
  return rest;
}

/**
 * Attiva o disattiva un tipo di elenco sui paragrafi selezionati. Se sono già
 * tutti di quel tipo l'elenco viene tolto; altrimenti entrano nell'elenco del
 * paragrafo precedente, se è dello stesso tipo, o in un elenco nuovo.
 */
function toggleList(state: EditorState, kind: NumberFormat): EditorState | null {
  const locs = selectedParagraphs(state);
  if (!locs.length) return null;
  let doc = state.doc;
  if (locs.every((l) => paragraphKind(doc, l.paragraph) === kind)) {
    for (const { paragraph: p } of locs) doc = replaceBlock(doc, p.id, [{ ...p, format: { ...p.format, list: null } }]);
    return { ...state, doc };
  }
  const first = locs[0]!;
  const container = containerOf(doc, first);
  const prev = container[first.path[first.path.length - 1]! - 1];
  let listId: string;
  if (prev?.type === 'paragraph' && prev.format.list && paragraphKind(doc, prev) === kind) {
    listId = prev.format.list.listId;
  } else {
    const def = kind === 'bullet' ? bulletList() : numberedList(kind);
    listId = def.id;
    doc = { ...doc, lists: { ...doc.lists, [def.id]: def } };
  }
  for (const { paragraph: p } of locs) {
    const level = p.format.list?.level ?? 0;
    doc = replaceBlock(doc, p.id, [{ ...p, format: { ...withoutIndent(p.format), list: { listId, level } } }]);
  }
  return { ...state, doc };
}

function changeLevel(state: EditorState, delta: number): EditorState | null {
  let doc = state.doc;
  let changed = false;
  for (const { paragraph: p } of selectedParagraphs(state)) {
    const ref = p.format.list;
    if (!ref) continue;
    const level = Math.min(MAX_LIST_LEVEL, Math.max(0, ref.level + delta));
    if (level === ref.level) continue;
    changed = true;
    doc = replaceBlock(doc, p.id, [{ ...p, format: { ...p.format, list: { ...ref, level } } }]);
  }
  return changed ? { ...state, doc } : null;
}

function focusKind(state: EditorState): NumberFormat | undefined {
  return paragraphKind(state.doc, requireParagraph(state.doc, state.selection.focus.paragraphId).paragraph);
}

function inList(state: EditorState): boolean {
  return selectedParagraphs(state).some((l) => !!l.paragraph.format.list);
}

/**
 * Tab cambia livello solo all'inizio della voce o con più paragrafi
 * selezionati, come in Word; in mezzo al testo inserisce una tabulazione.
 */
function atItemStart(state: EditorState): boolean {
  const { anchor, focus } = state.selection;
  return anchor.paragraphId !== focus.paragraphId || Math.min(anchor.offset, focus.offset) === 0;
}

export const listCommands = [
  defineCommand('list.toggleBullet', {
    label: 'Elenco puntato',
    isActive: ({ state }) => focusKind(state) === 'bullet',
    run: ({ state }) => {
      const next = toggleList(state, 'bullet');
      return next && { state: next, meta: { label: 'Elenco puntato' } };
    },
  }),
  defineCommand('list.toggleNumbered', {
    label: 'Elenco numerato',
    isActive: ({ state }) => {
      const k = focusKind(state);
      return k !== undefined && k !== 'bullet';
    },
    run: ({ state }, args) => {
      const next = toggleList(state, args?.format ?? 'decimal');
      return next && { state: next, meta: { label: 'Elenco numerato' } };
    },
  }),
  defineCommand('list.increaseLevel', {
    label: 'Aumenta livello elenco',
    isEnabled: ({ state }) => inList(state) && atItemStart(state),
    run: ({ state }) => {
      const next = changeLevel(state, 1);
      return next && { state: next, meta: { label: 'Aumenta livello elenco' } };
    },
  }),
  defineCommand('list.decreaseLevel', {
    label: 'Riduci livello elenco',
    isEnabled: ({ state }) => inList(state),
    run: ({ state }) => {
      const next = changeLevel(state, -1);
      return next && { state: next, meta: { label: 'Riduci livello elenco' } };
    },
  }),
];
