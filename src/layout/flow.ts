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

/** Spazio fra il testo e i bordi del paragrafo (valori predefiniti di Word). */
export const BORDER_SPACE = { top: 1, right: 4, bottom: 1, left: 4 } as const;

const hasBorder = (b: Borders | undefined) => !!b && (visible(b.top) || visible(b.right) || visible(b.bottom) || visible(b.left));

function sameLine(a: BorderLine | undefined, b: BorderLine | undefined): boolean {
  if (!visible(a) || !visible(b)) return visible(a) === visible(b);
  return a.style === b.style && a.width === b.width && a.color === b.color;
}

/**
 * Due paragrafi consecutivi con gli stessi bordi, sfondo e rientri formano un
 * unico riquadro, come in Word: niente bordo fra l'uno e l'altro e uno sfondo
 * continuo anche sullo spazio che li separa.
 */
export function sameBox(a: ParaFormat, b: ParaFormat): boolean {
  if (!hasBorder(a.borders) && !a.shading) return false;
  const sides = ['top', 'right', 'bottom', 'left'] as const;
  return (
    sides.every((s) => sameLine(a.borders?.[s], b.borders?.[s])) &&
    (a.shading ?? null) === (b.shading ?? null) &&
    (a.indentLeft ?? 0) === (b.indentLeft ?? 0) &&
    (a.indentRight ?? 0) === (b.indentRight ?? 0)
  );
}

export interface ParagraphBoxDecorations {
  readonly decorations: Decoration[];
  /** Bordo inferiore, da togliere se il paragrafo successivo continua il riquadro. */
  readonly bottomBorder?: Decoration;
}

/**
 * Sfondo e bordi di un paragrafo (o di un suo pezzo su una pagina). Con
 * `joinTop` il riquadro continua quello del paragrafo precedente: niente bordo
 * superiore, e `top` è già il fondo del precedente.
 */
export function paragraphBox(
  format: ParaFormat,
  ox: Points,
  width: Points,
  top: Points,
  height: Points,
  joinTop = false,
): ParagraphBoxDecorations {
  const borders = format.borders;
  const pad = hasBorder(borders) ? BORDER_SPACE : { top: 0, right: 0, bottom: 0, left: 0 };
  const left = ox + (format.indentLeft ?? 0) - pad.left;
  const w = width - (format.indentLeft ?? 0) - (format.indentRight ?? 0) + pad.left + pad.right;
  const y = joinTop ? top : top - pad.top;
  const h = height + (joinTop ? 0 : pad.top) + pad.bottom;
  const decorations: Decoration[] = [];
  if (format.shading) decorations.push({ kind: 'fill', color: format.shading, rect: { x: left, y, width: w, height: h } });
  let bottomBorder: Decoration | undefined;
  if (borders) {
    const { top: t, bottom, left: l, right } = borders;
    decorations.push(...borderDecorations(left, y, w, h, { ...(joinTop ? {} : { top: t }), left: l, right }));
    bottomBorder = borderDecorations(left, y, w, h, { bottom })[0];
    if (bottomBorder) decorations.push(bottomBorder);
  }
  return { decorations, ...(bottomBorder ? { bottomBorder } : {}) };
}

/** Sfondo e bordi di un paragrafo isolato (o di un suo pezzo su una pagina). */
export function paragraphDecorations(format: ParaFormat, ox: Points, width: Points, top: Points, height: Points): Decoration[] {
  return paragraphBox(format, ox, width, top, height).decorations;
}

/** Impagina una sequenza di blocchi in una colonna larga `width`, senza interruzioni. */
export function layoutFlow(ctx: LayoutContext, blocks: readonly Block[], width: Points): Flow {
  const lines: LineBox[] = [];
  const decorations: Decoration[] = [];
  let y = 0;
  let prev: { format: ParaFormat; bottom: Points; bottomBorder?: Decoration } | null = null;
  blocks.forEach((b, i) => {
    if (b.type === 'paragraph') {
      const box = ctx.paragraph(b, width, ctx.fields);
      if (i > 0) y += box.spaceBefore;
      for (const l of box.lines) lines.push(toLineBox(b.id, l, 0, y));
      const join = !!prev && sameBox(prev.format, box.format);
      if (join && prev!.bottomBorder) decorations.splice(decorations.indexOf(prev!.bottomBorder), 1);
      const top = join ? prev!.bottom : y;
      const deco = paragraphBox(box.format, 0, width, top, y + box.height - top, join);
      decorations.push(...deco.decorations);
      prev = { format: box.format, bottom: y + box.height, ...(deco.bottomBorder ? { bottomBorder: deco.bottomBorder } : {}) };
      y += box.height + box.spaceAfter;
    } else {
      prev = null;
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
