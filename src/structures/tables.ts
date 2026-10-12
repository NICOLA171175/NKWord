/**
 * Comandi `table.*`. La selezione vive sui paragrafi: la "cella corrente" è
 * quella che contiene il cursore; una selezione che va da una cella a
 * un'altra della stessa tabella individua un rettangolo di righe × colonne.
 */

import {
  caret,
  findParagraph,
  isCollapsed,
  paragraph,
  paragraphLength,
  replaceBlock,
  table as newTable,
  type Block,
  type DocumentModel,
  type EditorState,
  type Id,
  type Table,
  type Transaction,
} from '../core';
import {
  deleteColumns,
  deleteRows,
  gridStarts,
  insertColumns,
  insertRows,
  mergeCells,
  paragraphAt,
  span,
  splitCell,
  canMerge,
} from './table-ops';
import { defineCommand, firstParagraphIn, insertBlocks, lastParagraphIn, requireParagraph } from './util';

/** Livello di annidamento: una tabella e la cella in cui si trova il paragrafo. */
export interface TableLevel {
  readonly table: Table;
  /** Blocchi che contengono la tabella e sua posizione fra essi. */
  readonly container: readonly Block[];
  readonly index: number;
  readonly row: number;
  readonly cell: number;
}

/** Tabelle che contengono il paragrafo, dalla più esterna alla più interna. */
export function tableChain(doc: DocumentModel, paragraphId: Id): TableLevel[] {
  const loc = findParagraph(doc, paragraphId);
  if (!loc) return [];
  const out: TableLevel[] = [];
  let blocks = doc.sections[loc.sectionIndex]!.blocks;
  for (let i = 0; i + 3 < loc.path.length; i += 3) {
    const index = loc.path[i]!;
    const table = blocks[index] as Table;
    const row = loc.path[i + 1]!;
    const cell = loc.path[i + 2]!;
    out.push({ table, container: blocks, index, row, cell });
    blocks = table.rows[row]!.cells[cell]!.blocks;
  }
  return out;
}

/** Tabella corrente e rettangolo selezionato (righe r0..r1, colonne di griglia g0..g1). */
export interface TableSelection {
  readonly level: TableLevel;
  readonly r0: number;
  readonly r1: number;
  readonly g0: number;
  readonly g1: number;
  /** Colonna di griglia del cursore. */
  readonly focusGrid: number;
}

export function tableSelection(state: EditorState): TableSelection | undefined {
  const focusChain = tableChain(state.doc, state.selection.focus.paragraphId);
  if (!focusChain.length) return undefined;
  const anchorChain = isCollapsed(state.selection) ? focusChain : tableChain(state.doc, state.selection.anchor.paragraphId);
  // Tabella comune più interna; se l'ancora è fuori conta solo la cella del cursore.
  let depth = 0;
  while (
    depth < focusChain.length &&
    depth < anchorChain.length &&
    focusChain[depth]!.table.id === anchorChain[depth]!.table.id
  ) depth++;
  const focus = depth > 0 ? focusChain[depth - 1]! : focusChain[focusChain.length - 1]!;
  const anchor = depth > 0 ? anchorChain[depth - 1]! : focus;
  const t = focus.table;
  const bounds = (l: TableLevel) => {
    const s = gridStarts(t.rows[l.row]!)[l.cell]!;
    return [s, s + span(t.rows[l.row]!.cells[l.cell]!) - 1] as const;
  };
  const [fa, fb] = bounds(focus);
  const [aa, ab] = bounds(anchor);
  return {
    level: focus,
    r0: Math.min(focus.row, anchor.row),
    r1: Math.max(focus.row, anchor.row),
    g0: Math.min(fa, aa),
    g1: Math.max(fb, ab),
    focusGrid: fa,
  };
}

function inTable(state: EditorState): boolean {
  return tableChain(state.doc, state.selection.focus.paragraphId).length > 0;
}

/** Sostituisce la tabella corrente e mette il cursore nella cella (row, g). */
function commit(state: EditorState, old: Table, next: Table, row: number, g: number, label: string): Transaction {
  const doc = replaceBlock(state.doc, old.id, [next]);
  const p = paragraphAt(next, row, Math.min(g, next.columnWidths.length - 1)) ?? firstParagraphIn([next])!;
  return { state: { ...state, doc, selection: caret(p.id, 0), pendingFormat: undefined }, meta: { label } };
}

/** Elimina una tabella e porta il cursore al paragrafo successivo (o precedente). */
function removeTable(state: EditorState, level: TableLevel, label: string): Transaction {
  const nextBlock = level.container[level.index + 1];
  const prevBlock = level.container[level.index - 1];
  let doc: DocumentModel;
  let target;
  if (nextBlock) {
    doc = replaceBlock(state.doc, level.table.id, []);
    target = caret(firstParagraphIn([nextBlock])!.id, 0);
  } else if (prevBlock) {
    doc = replaceBlock(state.doc, level.table.id, []);
    const p = lastParagraphIn([prevBlock])!;
    target = caret(p.id, paragraphLength(p));
  } else {
    const p = paragraph();
    doc = replaceBlock(state.doc, level.table.id, [p]);
    target = caret(p.id, 0);
  }
  return { state: { ...state, doc, selection: target, pendingFormat: undefined }, meta: { label } };
}

/** Larghezza disponibile per una nuova tabella al cursore. */
function availableWidth(state: EditorState): number {
  const chain = tableChain(state.doc, state.selection.focus.paragraphId);
  const inner = chain[chain.length - 1];
  if (inner) {
    const s = gridStarts(inner.table.rows[inner.row]!)[inner.cell]!;
    const n = span(inner.table.rows[inner.row]!.cells[inner.cell]!);
    return inner.table.columnWidths.slice(s, s + n).reduce((a, b) => a + b, 0);
  }
  const loc = requireParagraph(state.doc, state.selection.focus.paragraphId);
  const page = state.doc.sections[loc.sectionIndex]!.page;
  const content = page.width - page.margins.left - page.margins.right;
  const cols = page.columns;
  return cols && cols.count > 1 ? (content - cols.gap * (cols.count - 1)) / cols.count : content;
}

export const MAX_TABLE_COLUMNS = 63;
export const MAX_TABLE_ROWS = 32767;

const validSize = (rows: number, cols: number) =>
  Number.isInteger(rows) && Number.isInteger(cols) && rows >= 1 && cols >= 1 && rows <= MAX_TABLE_ROWS && cols <= MAX_TABLE_COLUMNS;

export const tableCommands = [
  defineCommand('table.insert', {
    label: 'Inserisci tabella',
    isEnabled: (_ctx, args) => !!args && validSize(args.rows, args.cols),
    run: ({ state }, { rows, cols }) => {
      if (!validSize(rows, cols)) return null;
      const t = newTable(rows, cols, availableWidth(state));
      const { doc } = insertBlocks(state, [t]);
      const first = firstParagraphIn([t])!;
      return { state: { ...state, doc, selection: caret(first.id, 0), pendingFormat: undefined }, meta: { label: 'Inserisci tabella' } };
    },
  }),
  defineCommand('table.insertRowAbove', {
    label: 'Inserisci riga sopra',
    isEnabled: ({ state }) => inTable(state),
    run: ({ state }) => {
      const sel = tableSelection(state);
      if (!sel) return null;
      const next = insertRows(sel.level.table, sel.r0, sel.r1, false);
      return commit(state, sel.level.table, next, sel.r0, sel.focusGrid, 'Inserisci riga sopra');
    },
  }),
  defineCommand('table.insertRowBelow', {
    label: 'Inserisci riga sotto',
    isEnabled: ({ state }) => inTable(state),
    run: ({ state }) => {
      const sel = tableSelection(state);
      if (!sel) return null;
      const next = insertRows(sel.level.table, sel.r0, sel.r1, true);
      return commit(state, sel.level.table, next, sel.r1 + 1, sel.focusGrid, 'Inserisci riga sotto');
    },
  }),
  defineCommand('table.insertColumnLeft', {
    label: 'Inserisci colonna a sinistra',
    isEnabled: ({ state }) => inTable(state),
    run: ({ state }) => {
      const sel = tableSelection(state);
      if (!sel) return null;
      const next = insertColumns(sel.level.table, sel.g0, sel.g1, false);
      return commit(state, sel.level.table, next, sel.level.row, sel.g0, 'Inserisci colonna a sinistra');
    },
  }),
  defineCommand('table.insertColumnRight', {
    label: 'Inserisci colonna a destra',
    isEnabled: ({ state }) => inTable(state),
    run: ({ state }) => {
      const sel = tableSelection(state);
      if (!sel) return null;
      const next = insertColumns(sel.level.table, sel.g0, sel.g1, true);
      return commit(state, sel.level.table, next, sel.level.row, sel.g1 + 1, 'Inserisci colonna a destra');
    },
  }),
  defineCommand('table.deleteRow', {
    label: 'Elimina righe',
    isEnabled: ({ state }) => inTable(state),
    run: ({ state }) => {
      const sel = tableSelection(state);
      if (!sel) return null;
      const next = deleteRows(sel.level.table, sel.r0, sel.r1);
      if (!next) return removeTable(state, sel.level, 'Elimina righe');
      return commit(state, sel.level.table, next, Math.min(sel.r0, next.rows.length - 1), sel.focusGrid, 'Elimina righe');
    },
  }),
  defineCommand('table.deleteColumn', {
    label: 'Elimina colonne',
    isEnabled: ({ state }) => inTable(state),
    run: ({ state }) => {
      const sel = tableSelection(state);
      if (!sel) return null;
      const next = deleteColumns(sel.level.table, sel.g0, sel.g1);
      if (!next) return removeTable(state, sel.level, 'Elimina colonne');
      return commit(state, sel.level.table, next, sel.level.row, Math.max(0, sel.g0 - (sel.g0 >= next.columnWidths.length ? 1 : 0)), 'Elimina colonne');
    },
  }),
  defineCommand('table.delete', {
    label: 'Elimina tabella',
    isEnabled: ({ state }) => inTable(state),
    run: ({ state }) => {
      const sel = tableSelection(state);
      return sel ? removeTable(state, sel.level, 'Elimina tabella') : null;
    },
  }),
  defineCommand('table.mergeCells', {
    label: 'Unisci celle',
    isEnabled: ({ state }) => {
      const sel = tableSelection(state);
      return !!sel && canMerge(sel.level.table, sel.r0, sel.r1, sel.g0, sel.g1);
    },
    run: ({ state }) => {
      const sel = tableSelection(state);
      if (!sel) return null;
      const next = mergeCells(sel.level.table, sel.r0, sel.r1, sel.g0, sel.g1);
      return next && commit(state, sel.level.table, next, sel.r0, sel.g0, 'Unisci celle');
    },
  }),
  defineCommand('table.splitCell', {
    label: 'Dividi celle',
    isEnabled: ({ state }, args) => {
      const sel = tableSelection(state);
      return !!sel && !!args && splitCell(sel.level.table, sel.level.row, sel.focusGrid, args.rows, args.cols) !== null;
    },
    run: ({ state }, { rows, cols }) => {
      const sel = tableSelection(state);
      if (!sel) return null;
      const next = splitCell(sel.level.table, sel.level.row, sel.focusGrid, rows, cols);
      if (!next) return null;
      // La prima parte della cella conserva il contenuto: il cursore resta dov'era.
      const doc = replaceBlock(state.doc, sel.level.table.id, [next]);
      return { state: { ...state, doc, selection: caret(state.selection.focus.paragraphId, state.selection.focus.offset) }, meta: { label: 'Dividi celle' } };
    },
  }),
];
