/**
 * Operazioni pure su una `Table`. La tabella è vista come una griglia di
 * `columnWidths.length` colonne: ogni cella parte da una colonna della
 * griglia e ne occupa `colSpan`; le unioni verticali usano `vMerge`
 * ('restart' sulla prima cella, 'continue' sulle successive, alla stessa
 * colonna di partenza e con la stessa ampiezza).
 */

import {
  newId,
  paragraph,
  type Block,
  type Paragraph,
  type Table,
  type TableCell,
  type TableCellFormat,
  type TableRow,
} from '../core';
import { firstParagraphIn } from './util';

export const span = (c: TableCell): number => c.colSpan ?? 1;

/** Colonna di griglia in cui inizia ciascuna cella della riga. */
export function gridStarts(row: TableRow): number[] {
  const out: number[] = [];
  let g = 0;
  for (const c of row.cells) {
    out.push(g);
    g += span(c);
  }
  return out;
}

/** Indice della cella che copre la colonna di griglia `g`, oppure -1. */
export function cellAtGrid(row: TableRow, g: number): number {
  let start = 0;
  for (let i = 0; i < row.cells.length; i++) {
    const s = span(row.cells[i]!);
    if (g >= start && g < start + s) return i;
    start += s;
  }
  return -1;
}

export function makeCell(
  blocks: readonly Block[],
  format: TableCellFormat,
  colSpan = 1,
  vMerge?: 'restart' | 'continue',
  id: string = newId(),
): TableCell {
  return {
    id,
    blocks,
    format,
    ...(colSpan > 1 ? { colSpan } : {}),
    ...(vMerge ? { vMerge } : {}),
  };
}

/** Paragrafo vuoto con lo stesso stile del primo paragrafo della cella. */
export function emptyLike(cell: TableCell): Paragraph {
  const p = firstParagraphIn(cell.blocks);
  return paragraph([], p?.format ?? {}, p?.styleId);
}

function isEmptyCell(cell: TableCell): boolean {
  return cell.blocks.length === 1 && cell.blocks[0]!.type === 'paragraph' && cell.blocks[0]!.children.length === 0;
}

function withCells(row: TableRow, cells: readonly TableCell[]): TableRow {
  return { ...row, cells };
}

/** Sistema le unioni verticali rimaste incoerenti dopo un'operazione. */
export function normalizeVMerge(table: Table): Table {
  const rows = table.rows.map((r) => ({ ...r, cells: [...r.cells] }));
  for (let r = 0; r < rows.length; r++) {
    const starts = gridStarts(rows[r]!);
    rows[r]!.cells.forEach((c, i) => {
      if (c.vMerge !== 'continue') return;
      const above = r > 0 ? rows[r - 1]! : undefined;
      const j = above ? cellAtGrid(above, starts[i]!) : -1;
      const ok = above && j >= 0 && gridStarts(above)[j] === starts[i] && span(above.cells[j]!) === span(c) && above.cells[j]!.vMerge;
      if (!ok) rows[r]!.cells[i] = { ...c, vMerge: 'restart' };
    });
  }
  for (let r = 0; r < rows.length; r++) {
    const starts = gridStarts(rows[r]!);
    rows[r]!.cells.forEach((c, i) => {
      if (c.vMerge !== 'restart') return;
      const below = rows[r + 1];
      const j = below ? cellAtGrid(below, starts[i]!) : -1;
      const ok = below && j >= 0 && gridStarts(below)[j] === starts[i] && below.cells[j]!.vMerge === 'continue';
      if (!ok) {
        const { vMerge: _v, ...rest } = c;
        rows[r]!.cells[i] = rest;
      }
    });
  }
  const changed = rows.some((row, r) => row.cells.some((c, i) => c !== table.rows[r]!.cells[i]));
  return changed ? { ...table, rows: rows.map((row, r) => (row.cells.some((c, i) => c !== table.rows[r]!.cells[i]) ? row : table.rows[r]!)) } : table;
}

/**
 * Riga effettiva della cella in (row, g): se la cella prosegue un'unione
 * verticale risale alla cella che la apre.
 */
export function mergeOrigin(table: Table, row: number, g: number): number {
  let r = row;
  while (r > 0) {
    const i = cellAtGrid(table.rows[r]!, g);
    if (i < 0 || table.rows[r]!.cells[i]!.vMerge !== 'continue') break;
    r--;
  }
  return r;
}

/** Primo paragrafo della cella che copre (row, g), risalendo le unioni verticali. */
export function paragraphAt(table: Table, row: number, g: number): Paragraph | undefined {
  const r = mergeOrigin(table, Math.min(row, table.rows.length - 1), g);
  const cells = table.rows[r]!.cells;
  let i = cellAtGrid(table.rows[r]!, g);
  if (i < 0) i = cells.length - 1;
  return firstParagraphIn(cells[i]!.blocks);
}

// ---------------------------------------------------------------------------
// Righe

export function insertRows(table: Table, r0: number, r1: number, below: boolean): Table {
  const count = r1 - r0 + 1;
  const ref = below ? r1 : r0;
  const refRow = table.rows[ref]!;
  const next = table.rows[ref + 1];
  const nextStarts = next ? gridStarts(next) : [];
  const starts = gridStarts(refRow);
  const newRows: TableRow[] = Array.from({ length: count }, () => ({
    ...refRow,
    id: newId(),
    cells: refRow.cells.map((c, i) => {
      let vMerge: 'continue' | undefined;
      if (below && c.vMerge && next) {
        const j = cellAtGrid(next, starts[i]!);
        if (j >= 0 && nextStarts[j] === starts[i] && next.cells[j]!.vMerge === 'continue') vMerge = 'continue';
      } else if (!below && c.vMerge === 'continue') vMerge = 'continue';
      return makeCell([emptyLike(c)], c.format, span(c), vMerge);
    }),
  }));
  const at = below ? r1 + 1 : r0;
  return normalizeVMerge({ ...table, rows: [...table.rows.slice(0, at), ...newRows, ...table.rows.slice(at)] });
}

/** Elimina le righe r0..r1; `null` se non resta nessuna riga. */
export function deleteRows(table: Table, r0: number, r1: number): Table | null {
  const rows = table.rows.filter((_, r) => r < r0 || r > r1);
  return rows.length ? normalizeVMerge({ ...table, rows }) : null;
}

// ---------------------------------------------------------------------------
// Colonne

/** Inserisce tante colonne quante sono g0..g1, a sinistra o a destra; la larghezza totale resta uguale. */
export function insertColumns(table: Table, g0: number, g1: number, right: boolean): Table {
  const count = g1 - g0 + 1;
  const at = right ? g1 + 1 : g0;
  const rows = table.rows.map((row) => {
    const starts = gridStarts(row);
    const crossing = row.cells.findIndex((c, i) => starts[i]! < at && at < starts[i]! + span(c));
    if (crossing >= 0) {
      const c = row.cells[crossing]!;
      return withCells(row, row.cells.map((x, i) => (i === crossing ? { ...c, colSpan: span(c) + count } : x)));
    }
    let idx = starts.findIndex((s) => s >= at);
    if (idx < 0) idx = row.cells.length;
    const ref = row.cells[right ? Math.max(0, idx - 1) : Math.min(idx, row.cells.length - 1)];
    const fresh = Array.from({ length: count }, () =>
      ref ? makeCell([emptyLike(ref)], ref.format) : makeCell([paragraph()], {}),
    );
    return withCells(row, [...row.cells.slice(0, idx), ...fresh, ...row.cells.slice(idx)]);
  });
  const total = table.columnWidths.reduce((a, b) => a + b, 0);
  const added = table.columnWidths.slice(g0, g1 + 1);
  const widths = [...table.columnWidths.slice(0, at), ...added, ...table.columnWidths.slice(at)];
  const scale = total / widths.reduce((a, b) => a + b, 0);
  return { ...table, rows, columnWidths: widths.map((w) => w * scale) };
}

/** Elimina le colonne di griglia g0..g1; `null` se non resta nulla. */
export function deleteColumns(table: Table, g0: number, g1: number): Table | null {
  const columnWidths = table.columnWidths.filter((_, g) => g < g0 || g > g1);
  if (!columnWidths.length) return null;
  const rows: TableRow[] = [];
  for (const row of table.rows) {
    const starts = gridStarts(row);
    const cells: TableCell[] = [];
    row.cells.forEach((c, i) => {
      const s = starts[i]!;
      const e = s + span(c) - 1;
      const overlap = Math.max(0, Math.min(e, g1) - Math.max(s, g0) + 1);
      if (overlap === span(c)) return;
      const n = span(c) - overlap;
      if (overlap === 0) cells.push(c);
      else {
        const { colSpan: _c, ...rest } = c;
        cells.push(n > 1 ? { ...rest, colSpan: n } : rest);
      }
    });
    if (cells.length) rows.push(withCells(row, cells));
  }
  return rows.length ? normalizeVMerge({ ...table, rows, columnWidths }) : null;
}

// ---------------------------------------------------------------------------
// Unione e divisione

/** Celle della riga comprese fra le colonne g0..g1, oppure null se una cella sporge. */
function cellsInRange(row: TableRow, g0: number, g1: number): { first: number; last: number } | null {
  const starts = gridStarts(row);
  let first = -1;
  let last = -1;
  let covered = 0;
  for (let i = 0; i < row.cells.length; i++) {
    const s = starts[i]!;
    const e = s + span(row.cells[i]!) - 1;
    if (e < g0 || s > g1) continue;
    if (s < g0 || e > g1) return null;
    if (first < 0) first = i;
    last = i;
    covered += e - s + 1;
  }
  return first >= 0 && covered === g1 - g0 + 1 ? { first, last } : null;
}

/** Il rettangolo r0..r1 × g0..g1 si può unire? */
export function canMerge(table: Table, r0: number, r1: number, g0: number, g1: number): boolean {
  if (r0 === r1 && g0 === g1) return false;
  for (let r = r0; r <= r1; r++) {
    const range = cellsInRange(table.rows[r]!, g0, g1);
    if (!range) return false;
    if (r === r0 && table.rows[r]!.cells.slice(range.first, range.last + 1).some((c) => c.vMerge === 'continue')) return false;
  }
  const after = table.rows[r1 + 1];
  if (after) {
    const starts = gridStarts(after);
    if (after.cells.some((c, i) => c.vMerge === 'continue' && starts[i]! + span(c) - 1 >= g0 && starts[i]! <= g1)) return false;
  }
  return true;
}

/** Unisce il rettangolo in una cella; il contenuto non vuoto viene accodato in ordine di lettura. */
export function mergeCells(table: Table, r0: number, r1: number, g0: number, g1: number): Table | null {
  if (!canMerge(table, r0, r1, g0, g1)) return null;
  const blocks: Block[] = [];
  const ranges = table.rows.slice(r0, r1 + 1).map((row) => cellsInRange(row, g0, g1)!);
  ranges.forEach((range, k) => {
    for (const c of table.rows[r0 + k]!.cells.slice(range.first, range.last + 1)) if (!isEmptyCell(c)) blocks.push(...c.blocks);
  });
  const first = table.rows[r0]!.cells[ranges[0]!.first]!;
  const width = g1 - g0 + 1;
  const rows = table.rows.map((row, r) => {
    if (r < r0 || r > r1) return row;
    const range = ranges[r - r0]!;
    const merged =
      r === r0
        ? makeCell(blocks.length ? blocks : first.blocks, first.format, width, r1 > r0 ? 'restart' : undefined, first.id)
        : makeCell([emptyLike(row.cells[range.first]!)], row.cells[range.first]!.format, width, 'continue');
    return withCells(row, [...row.cells.slice(0, range.first), merged, ...row.cells.slice(range.last + 1)]);
  });
  return normalizeVMerge({ ...table, rows });
}

function distribute(total: number, parts: number): number[] {
  const base = Math.floor(total / parts);
  return Array.from({ length: parts }, (_, i) => base + (i < total % parts ? 1 : 0));
}

/**
 * Divide la cella che copre (row, g) in `rows` × `cols` celle. Una cella
 * unita verticalmente su k righe si può dividere solo in 1 o k righe.
 * Restituisce null se la divisione non è possibile o non cambia nulla.
 */
export function splitCell(table: Table, row: number, g: number, rows: number, cols: number): Table | null {
  if (!Number.isInteger(rows) || !Number.isInteger(cols) || rows < 1 || cols < 1) return null;
  if (rows === 1 && cols === 1) return null;
  const r = mergeOrigin(table, row, g);
  const ti = cellAtGrid(table.rows[r]!, g);
  if (ti < 0) return null;
  const s = gridStarts(table.rows[r]!)[ti]!;
  let sp = span(table.rows[r]!.cells[ti]!);
  let k = 1;
  if (table.rows[r]!.cells[ti]!.vMerge === 'restart') {
    while (r + k < table.rows.length) {
      const nr = table.rows[r + k]!;
      const j = cellAtGrid(nr, s);
      if (j < 0 || gridStarts(nr)[j] !== s || nr.cells[j]!.vMerge !== 'continue') break;
      k++;
    }
  }
  if (k > 1 && rows !== 1 && rows !== k) return null;

  let t = table;
  if (cols > sp) {
    // Servono nuove colonne di griglia: si divide l'ultima colonna della cella.
    const extra = cols - sp;
    const gl = s + sp - 1;
    const w = t.columnWidths[gl]! / (extra + 1);
    const columnWidths = [...t.columnWidths.slice(0, gl), ...Array<number>(extra + 1).fill(w), ...t.columnWidths.slice(gl + 1)];
    const newRows = t.rows.map((rw) => {
      const i = cellAtGrid(rw, gl);
      if (i < 0) return rw;
      const c = rw.cells[i]!;
      return withCells(rw, rw.cells.map((x, j) => (j === i ? { ...c, colSpan: span(c) + extra } : x)));
    });
    t = { ...t, columnWidths, rows: newRows };
    sp = cols;
  }
  const spans = distribute(sp, cols);
  const pieces = (cell: TableCell, vMerge: 'restart' | 'continue' | undefined, keep: boolean): TableCell[] =>
    spans.map((w, j) =>
      j === 0 && keep
        ? makeCell(cell.blocks, cell.format, w, vMerge, cell.id)
        : makeCell([emptyLike(cell)], cell.format, w, vMerge),
    );
  const replaceAt = (rw: TableRow, cells: (c: TableCell) => TableCell[]): TableRow => {
    const i = cellAtGrid(rw, s);
    return withCells(rw, [...rw.cells.slice(0, i), ...cells(rw.cells[i]!), ...rw.cells.slice(i + 1)]);
  };

  let outRows: TableRow[];
  if (k > 1) {
    const unmerge = rows === k;
    outRows = t.rows.map((rw, q) => {
      if (q < r || q >= r + k) return rw;
      const vm = unmerge ? undefined : q === r ? 'restart' : 'continue';
      return replaceAt(rw, (c) => pieces(c, vm, true));
    });
  } else {
    const base = t.rows[r]!;
    const target = base.cells[cellAtGrid(base, s)]!;
    const first = replaceAt(
      withCells(base, base.cells.map((c) => (c === target || rows === 1 ? c : { ...c, vMerge: c.vMerge ?? 'restart' }))),
      (c) => pieces(c, undefined, true),
    );
    const added: TableRow[] = Array.from({ length: rows - 1 }, () =>
      withCells(
        { ...base, id: newId() },
        base.cells.flatMap((c) =>
          c === target
            ? pieces(c, undefined, false)
            : [makeCell([emptyLike(c)], c.format, span(c), 'continue')],
        ),
      ),
    );
    outRows = [...t.rows.slice(0, r), first, ...added, ...t.rows.slice(r + 1)];
  }
  return normalizeVMerge({ ...t, rows: outRows });
}
