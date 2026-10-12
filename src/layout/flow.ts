/**
 * Impaginazione di sequenze di blocchi senza interruzioni di pagina:
 * intestazioni, piè di pagina e contenuto delle celle. Qui c'è anche il
 * calcolo delle righe di tabella, usato anche dall'impaginatore del corpo.
 */

import type {
  Block,
  BorderLine,
  Borders,
  Decoration,
  DocumentModel,
  Id,
  LineBox,
  ParaFormat,
  Paragraph,
  Points,
  Table,
  TextMeasurer,
} from '../core';
import type { FieldContext, MarkerInfo, ParaBox, ParaLine } from './paragraph';

export interface LayoutContext {
  readonly doc: DocumentModel;
  readonly measurer: TextMeasurer;
  readonly markers: ReadonlyMap<Id, MarkerInfo>;
  readonly fields: FieldContext;
  /** Divide un paragrafo in righe (con cache). */
  paragraph(p: Paragraph, width: Points, fields: FieldContext): ParaBox;
}

/** Contenuto impaginato con coordinate relative a un'origine. */
export interface Flow {
  readonly height: Points;
  readonly lines: LineBox[];
  readonly decorations: Decoration[];
}

/** Spaziatura interna predefinita delle celle (0,19 cm a destra e a sinistra). */
export const DEFAULT_CELL_PADDING = { top: 0, right: 5.4, bottom: 0, left: 5.4 } as const;

export function toLineBox(paragraphId: Id, l: ParaLine, ox: Points, oy: Points): LineBox {
  return {
    paragraphId,
    start: l.start,
    end: l.end,
    x: ox + l.x,
    y: oy + l.y,
    width: l.width,
    height: l.height,
    baseline: l.baseline,
    runs: l.runs,
  };
}

export function translateLine(l: LineBox, dx: Points, dy: Points): LineBox {
  return { ...l, x: l.x + dx, y: l.y + dy };
}

export function translateDecoration(d: Decoration, dx: Points, dy: Points): Decoration {
  return { ...d, rect: { ...d.rect, x: d.rect.x + dx, y: d.rect.y + dy } };
}

function visible(b: BorderLine | undefined): b is BorderLine {
  return !!b && b.style !== 'none' && b.width > 0;
}

/** Bordi di un rettangolo come rettangoli pieni centrati sui lati. */
export function borderDecorations(
  x: Points,
  y: Points,
  width: Points,
  height: Points,
  borders: Borders,
): Decoration[] {
  const out: Decoration[] = [];
  const { top, right, bottom, left } = borders;
  if (visible(top)) out.push({ kind: 'border', color: top.color, width: top.width, rect: { x, y: y - top.width / 2, width, height: top.width } });
  if (visible(bottom))
    out.push({ kind: 'border', color: bottom.color, width: bottom.width, rect: { x, y: y + height - bottom.width / 2, width, height: bottom.width } });
  if (visible(left)) out.push({ kind: 'border', color: left.color, width: left.width, rect: { x: x - left.width / 2, y, width: left.width, height } });
  if (visible(right))
    out.push({ kind: 'border', color: right.color, width: right.width, rect: { x: x + width - right.width / 2, y, width: right.width, height } });
  return out;
}

/** Sfondo e bordi di un paragrafo (o di un suo pezzo su una pagina). */
export function paragraphDecorations(
  format: ParaFormat,
  ox: Points,
  width: Points,
  top: Points,
  height: Points,
): Decoration[] {
  const left = ox + (format.indentLeft ?? 0);
  const w = width - (format.indentLeft ?? 0) - (format.indentRight ?? 0);
  const out: Decoration[] = [];
  if (format.shading) out.push({ kind: 'fill', color: format.shading, rect: { x: left, y: top, width: w, height } });
  if (format.borders) out.push(...borderDecorations(left, top, w, height, format.borders));
  return out;
}

/** Impagina una sequenza di blocchi in una colonna larga `width`, senza interruzioni. */
export function layoutFlow(ctx: LayoutContext, blocks: readonly Block[], width: Points): Flow {
  const lines: LineBox[] = [];
  const decorations: Decoration[] = [];
  let y = 0;
  blocks.forEach((b, i) => {
    if (b.type === 'paragraph') {
      const box = ctx.paragraph(b, width, ctx.fields);
      if (i > 0) y += box.spaceBefore;
      for (const l of box.lines) lines.push(toLineBox(b.id, l, 0, y));
      decorations.push(...paragraphDecorations(box.format, 0, width, y, box.height));
      y += box.height + box.spaceAfter;
    } else {
      const t = layoutTable(ctx, b, width);
      for (const row of t.rows) {
        lines.push(...row.lines.map((l) => translateLine(l, 0, y)));
        decorations.push(...row.decorations.map((d) => translateDecoration(d, 0, y)));
        y += row.height;
      }
    }
  });
  return { height: y, lines, decorations };
}

// ---------------------------------------------------------------------------
// Tabelle

export interface RowBox {
  readonly height: Points;
  readonly isHeader: boolean;
  /** Coordinate relative all'area disponibile (x) e all'inizio della riga (y). */
  readonly lines: LineBox[];
  readonly decorations: Decoration[];
}

export interface TableBox {
  readonly width: Points;
  readonly rows: RowBox[];
}

interface CellLayout {
  readonly row: number;
  readonly gridCol: number;
  readonly span: number;
  readonly x: Points;
  readonly width: Points;
  readonly flow: Flow;
  readonly contentHeight: Points;
  /** Righe coperte dall'unione verticale (1 se non unita). */
  rowSpan: number;
  readonly continues: boolean;
}

export function layoutTable(ctx: LayoutContext, t: Table, available: Points): TableBox {
  const cols = t.columnWidths;
  const colX: Points[] = [0];
  for (const w of cols) colX.push(colX[colX.length - 1]! + w);
  const tableWidth = colX[colX.length - 1]!;
  const offset =
    t.format.align === 'center'
      ? (available - tableWidth) / 2
      : t.format.align === 'right'
        ? available - tableWidth
        : t.format.indentLeft ?? 0;
  const tb = t.format.borders ?? {};
  const nCols = cols.length;

  // 1. Contenuto di ogni cella.
  const grid: CellLayout[][] = t.rows.map((row, r) => {
    let gc = 0;
    return row.cells.map((cell) => {
      const span = Math.max(1, cell.colSpan ?? 1);
      const x = colX[Math.min(gc, nCols)] ?? tableWidth;
      const width = (colX[Math.min(gc + span, nCols)] ?? tableWidth) - x;
      const pad = cell.format.padding ?? DEFAULT_CELL_PADDING;
      const continues = cell.vMerge === 'continue';
      const flow = continues
        ? { height: 0, lines: [], decorations: [] }
        : layoutFlow(ctx, cell.blocks, Math.max(1, width - pad.left - pad.right));
      const layout: CellLayout = {
        row: r,
        gridCol: gc,
        span,
        x,
        width,
        flow,
        contentHeight: flow.height + pad.top + pad.bottom,
        rowSpan: 1,
        continues,
      };
      gc += span;
      return layout;
    });
  });

  // 2. Unioni verticali: una cella 'restart' copre le 'continue' sotto di lei.
  grid.forEach((cells, r) => {
    cells.forEach((c, ci) => {
      if (t.rows[r]!.cells[ci]!.vMerge !== 'restart') return;
      let rr = r + 1;
      while (rr < grid.length && grid[rr]!.some((o) => o.gridCol === c.gridCol && o.continues)) rr++;
      c.rowSpan = rr - r;
    });
  });

  // 3. Altezze delle righe.
  const heights = t.rows.map((row, r) => {
    let h = 0;
    for (const c of grid[r]!) if (c.rowSpan === 1 && !c.continues) h = Math.max(h, c.contentHeight);
    if (h === 0) h = ctx.measurer.measure('', ctx.doc.styles.defaults.char).ascent * 1.25;
    const rule = row.height;
    if (rule?.rule === 'exact') return rule.value;
    if (rule?.rule === 'atLeast') return Math.max(h, rule.value);
    return h;
  });
  for (const cells of grid) {
    for (const c of cells) {
      if (c.rowSpan <= 1) continue;
      const last = c.row + c.rowSpan - 1;
      let sum = 0;
      for (let r = c.row; r <= last; r++) sum += heights[r]!;
      if (c.contentHeight > sum && t.rows[last]!.height?.rule !== 'exact') heights[last]! += c.contentHeight - sum;
    }
  }

  // 4. Posizionamento di contenuto, sfondi e bordi.
  const rows: RowBox[] = t.rows.map((row, r) => {
    const lines: LineBox[] = [];
    const decorations: Decoration[] = [];
    for (const [ci, c] of grid[r]!.entries()) {
      if (c.continues) continue;
      const cell = row.cells[ci]!;
      const pad = cell.format.padding ?? DEFAULT_CELL_PADDING;
      let h = 0;
      for (let rr = r; rr < r + c.rowSpan; rr++) h += heights[rr]!;
      const x = offset + c.x;
      if (cell.format.shading) decorations.push({ kind: 'fill', color: cell.format.shading, rect: { x, y: 0, width: c.width, height: h } });
      const lastRow = r + c.rowSpan - 1 === t.rows.length - 1;
      const own = cell.format.borders ?? {};
      const borders: Borders = {
        top: own.top ?? (r === 0 ? tb.top : tb.insideH),
        bottom: own.bottom ?? (lastRow ? tb.bottom : tb.insideH),
        left: own.left ?? (c.gridCol === 0 ? tb.left : tb.insideV),
        right: own.right ?? (c.gridCol + c.span >= nCols ? tb.right : tb.insideV),
      };
      decorations.push(...borderDecorations(x, 0, c.width, h, borders));
      const free = Math.max(0, h - c.contentHeight);
      const va = cell.format.verticalAlign;
      const dy = pad.top + (va === 'center' ? free / 2 : va === 'bottom' ? free : 0);
      const dx = x + pad.left;
      lines.push(...c.flow.lines.map((l) => translateLine(l, dx, dy)));
      decorations.push(...c.flow.decorations.map((d) => translateDecoration(d, dx, dy)));
    }
    return { height: heights[r]!, isHeader: !!row.isHeader, lines, decorations };
  });

  return { width: tableWidth, rows };
}
