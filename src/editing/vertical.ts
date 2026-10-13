/**
 * Spostamenti verticali del cursore (frecce Su/Giù, PagSu/PagGiù) sulle
 * righe impaginate, come in Word: il cursore va alla riga sopra o sotto
 * restando nella stessa colonna (anche nelle celle di tabella e fra due
 * pagine) e, premendo più volte, ricorda la colonna di partenza anche se
 * attraversa righe più corte.
 */

import type { Block, DocumentLayout, DocumentModel, Id, LayoutEngine, LineBox, Points, Position } from '../core';

/** Una riga con la pagina che la contiene. */
interface PlacedLine {
  readonly page: number;
  readonly line: LineBox;
}

const EPS = 0.5;

/** Righe fra cui muoversi: il corpo di tutte le pagine o l'intestazione/piè di pagina della pagina del cursore. */
function lineGroup(layout: DocumentLayout, pos: Position, page: number, y: Points): PlacedLine[] | null {
  const body: PlacedLine[] = [];
  for (const p of layout.pages) for (const line of p.lines) body.push({ page: p.index, line });
  if (body.some((l) => l.page === page && l.line.paragraphId === pos.paragraphId && Math.abs(l.line.y - y) < EPS)) return body;
  const p = layout.pages[page];
  for (const part of [p?.header, p?.footer]) {
    if (part?.some((l) => l.paragraphId === pos.paragraphId && Math.abs(l.y - y) < EPS)) return part.map((line) => ({ page, line }));
  }
  return null;
}

const containsX = (l: LineBox, x: Points) => x >= l.x - EPS && x <= l.x + l.width + EPS;
const dxOf = (l: LineBox, x: Points) => (x < l.x ? l.x - x : x > l.x + l.width ? x - (l.x + l.width) : 0);

/** Riga di tabella di ogni paragrafo che sta in una cella (la più interna). */
const rowsCache = new WeakMap<DocumentModel, Map<Id, object>>();
function tableRows(doc: DocumentModel): Map<Id, object> {
  let rows = rowsCache.get(doc);
  if (rows) return rows;
  const out = new Map<Id, object>();
  const walk = (blocks: readonly Block[], row: object | null) => {
    for (const b of blocks) {
      if (b.type === 'paragraph') {
        if (row) out.set(b.id, row);
      } else for (const r of b.rows) for (const cell of r.cells) walk(cell.blocks, r);
    }
  };
  for (const section of doc.sections) walk(section.blocks, null);
  rowsCache.set(doc, (rows = out));
  return rows;
}

/**
 * Riga più vicina in verticale (`distance` più piccola) fra quelle date; a
 * parità di altezza quella che contiene la colonna `x` o le è più vicina.
 * In una tabella, se la riga trovata è in un'altra colonna, si preferisce la
 * cella della stessa riga di tabella che sta sotto (o sopra) la colonna `x`.
 */
function nearest(
  layout: DocumentLayout,
  lines: readonly PlacedLine[],
  x: Points,
  distance: (l: LineBox) => number,
): PlacedLine | null {
  let best: PlacedLine | null = null;
  for (const l of lines) {
    if (!best) best = l;
    else {
      const d = distance(l.line) - distance(best.line);
      if (d < -0.01 || (Math.abs(d) <= 0.01 && dxOf(l.line, x) < dxOf(best.line, x))) best = l;
    }
  }
  if (!best || containsX(best.line, x)) return best;
  const rows = tableRows(layout.doc);
  const row = rows.get(best.line.paragraphId);
  if (!row) return best;
  const sameRow = lines.filter((l) => l.page === best!.page && rows.get(l.line.paragraphId) === row && containsX(l.line, x));
  let alt: PlacedLine | null = null;
  for (const l of sameRow) if (!alt || distance(l.line) < distance(alt.line)) alt = l;
  return alt ?? best;
}

/** Le due righe si sovrappongono in orizzontale (stessa colonna di testo o di tabella). */
const overlaps = (a: LineBox, b: LineBox) => a.x < b.x + b.width && b.x < a.x + a.width;

/** Riga immediatamente sopra (dir -1) o sotto (dir 1), anche sulle pagine vicine. */
function adjacentLine(layout: DocumentLayout, group: readonly PlacedLine[], cur: PlacedLine, x: Points, dir: -1 | 1): PlacedLine | null {
  const { line } = cur;
  const distance = (l: LineBox) => (dir > 0 ? l.y : -(l.y + l.height));
  // Le righe delle celle vicine (altre colonne della tabella) non sono "sotto" il cursore.
  const samePage = group.filter(
    (l) =>
      l.page === cur.page &&
      l.line !== line &&
      overlaps(l.line, line) &&
      (dir > 0 ? l.line.y >= line.y + line.height - EPS : l.line.y + l.line.height <= line.y + EPS),
  );
  const target = nearest(layout, samePage, x, distance);
  if (target) return target;
  // Nessuna riga sulla pagina: la prima (o l'ultima) riga della pagina seguente (o precedente).
  const pages = [...new Set(group.map((l) => l.page))].filter((p) => (dir > 0 ? p > cur.page : p < cur.page));
  if (pages.length === 0) return null;
  const page = dir > 0 ? Math.min(...pages) : Math.max(...pages);
  return nearest(
    layout,
    group.filter((l) => l.page === page),
    x,
    distance,
  );
}

/** Ultimo offset in cui il cursore resta sulla riga (non dopo lo spazio o l'a capo che la chiude). */
export function lastCaretOffset(lines: readonly LineBox[], line: LineBox): number {
  const wrapped = line.end > line.start && lines.some((l) => l !== line && l.paragraphId === line.paragraphId && l.start === line.end);
  return wrapped ? line.end - 1 : line.end;
}

/** Offset della riga il cui cursore è più vicino alla coordinata `x` (ricerca binaria). */
function offsetAtX(engine: LayoutEngine, layout: DocumentLayout, group: readonly PlacedLine[], line: LineBox, x: Points): number {
  const xOf = (offset: number) => engine.caretRect(layout, { paragraphId: line.paragraphId, offset })?.rect.x ?? line.x;
  const last = Math.max(line.start, lastCaretOffset(group.map((l) => l.line), line));
  let lo = line.start;
  let hi = last;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (xOf(mid) <= x) lo = mid;
    else hi = mid - 1;
  }
  if (lo < last && x - xOf(lo) > xOf(lo + 1) - x) lo++;
  return lo;
}

export interface VerticalTarget {
  readonly focus: Position;
  /** Colonna (x sulla pagina) da ricordare per lo spostamento successivo. */
  readonly x: Points;
}

/**
 * Posizione dopo una freccia Su/Giù (`dir` = -1/1) o, con `distance`, dopo
 * PagSu/PagGiù (spostamento di `distance` punti). `goalX` è la colonna da
 * mantenere; se manca vale la x attuale del cursore. Null se non si può
 * calcolare (niente impaginazione) o se non c'è dove andare.
 */
export function verticalTarget(
  engine: LayoutEngine,
  layout: DocumentLayout,
  pos: Position,
  dir: -1 | 1,
  goalX?: Points,
  distance?: Points,
): VerticalTarget | null {
  const caret = engine.caretRect(layout, pos);
  if (!caret) return null;
  const group = lineGroup(layout, pos, caret.page, caret.rect.y);
  if (!group) return null;
  const cur = group.find((l) => l.page === caret.page && l.line.paragraphId === pos.paragraphId && Math.abs(l.line.y - caret.rect.y) < EPS);
  if (!cur) return null;
  const x = goalX ?? caret.rect.x;
  let target: PlacedLine | null;
  if (distance === undefined) target = adjacentLine(layout, group, cur, x, dir);
  else target = lineAtDistance(layout, group, cur, x, dir * distance);
  if (!target) return null;
  return { focus: { paragraphId: target.line.paragraphId, offset: offsetAtX(engine, layout, group, target.line, x) }, x };
}

/** Riga a `delta` punti in verticale (positivi verso il basso) contando le pagine una sotto l'altra. */
function lineAtDistance(layout: DocumentLayout, group: readonly PlacedLine[], cur: PlacedLine, x: Points, delta: Points): PlacedLine | null {
  const top: number[] = [];
  let acc = 0;
  for (const p of layout.pages) {
    top.push(acc);
    acc += p.height;
  }
  const absY = (l: PlacedLine) => (top[l.page] ?? 0) + l.line.y + l.line.height / 2;
  const goal = absY(cur) + delta;
  let best: PlacedLine | null = null;
  let bestD = Infinity;
  for (const l of group) {
    const d = Math.abs(absY(l) - goal);
    if (d < bestD - 0.01 || (Math.abs(d - bestD) <= 0.01 && best && containsX(l.line, x) && !containsX(best.line, x))) {
      best = l;
      bestD = d;
    }
  }
  if (!best) return null;
  // Sulla riga trovata, preferisce la cella della stessa colonna alla stessa altezza.
  const row = group.filter((l) => l.page === best!.page && Math.abs(l.line.y - best!.line.y) < best!.line.height / 2);
  const same = row.find((l) => containsX(l.line, x));
  const target = same ?? best;
  return target === cur ? null : target;
}
