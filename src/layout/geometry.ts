/**
 * Geometria del cursore su un'impaginazione già calcolata: dalla posizione
 * nel modello al rettangolo (caretRect), dal punto alla posizione (hitTest)
 * e i rettangoli della selezione.
 */

import type { DocumentLayout, GlyphRun, Id, LayoutEngine, LineBox, Points, Position, Rect, TextMeasurer, Block } from '../core';
import { paragraphs } from '../core';

interface PlacedLine {
  readonly page: number;
  readonly line: LineBox;
  readonly body: boolean;
}

function* allLines(layout: DocumentLayout): Generator<PlacedLine> {
  for (const p of layout.pages) for (const line of p.lines) yield { page: p.index, line, body: true };
  for (const p of layout.pages) {
    for (const line of p.header) yield { page: p.index, line, body: false };
    for (const line of p.footer) yield { page: p.index, line, body: false };
  }
}

/** Un run "a capo" (interruzione di riga/pagina) occupa un offset ma non ha testo. */
function isBreakRun(r: GlyphRun): boolean {
  return r.kind === 'text' && r.text === '' && r.end > r.start;
}

function collectBlockParagraphIds(blocks: readonly Block[], out: Id[]): void {
  for (const b of blocks) {
    if (b.type === 'paragraph') out.push(b.id);
    else for (const row of b.rows) for (const cell of row.cells) collectBlockParagraphIds(cell.blocks, out);
  }
}

const startsCache = new WeakMap<DocumentLayout, Set<string>>();

/** Inizi di riga del corpo come "paragrafo:offset", per riconoscere gli a capo automatici. */
function lineStarts(layout: DocumentLayout): Set<string> {
  let set = startsCache.get(layout);
  if (!set) {
    set = new Set();
    for (const p of layout.pages) for (const l of p.lines) if (l.end > l.start || l.runs.length === 0) set.add(`${l.paragraphId}:${l.start}`);
    startsCache.set(layout, set);
  }
  return set;
}

const orderCache = new WeakMap<DocumentLayout, Map<Id, number>>();

/** Ordine di lettura dei paragrafi: corpo, poi intestazioni e piè di pagina. */
function paragraphOrder(layout: DocumentLayout): Map<Id, number> {
  let order = orderCache.get(layout);
  if (order) return order;
  order = new Map();
  for (const { paragraph } of paragraphs(layout.doc)) order.set(paragraph.id, order.size);
  const extra: Id[] = [];
  for (const s of layout.doc.sections) {
    for (const hf of [s.header, s.footer]) {
      if (!hf) continue;
      for (const blocks of [hf.first, hf.default, hf.even]) if (blocks) collectBlockParagraphIds(blocks, extra);
    }
  }
  for (const id of extra) if (!order.has(id)) order.set(id, order.size);
  orderCache.set(layout, order);
  return order;
}

export function createGeometry(measurer: TextMeasurer): Pick<LayoutEngine, 'caretRect' | 'hitTest' | 'selectionRects'> {
  /** Distanza orizzontale dall'inizio della riga per un offset del paragrafo. */
  const xInLine = (line: LineBox, offset: number): Points => {
    const runs = line.runs.filter((r) => r.kind !== 'listMarker');
    let candidate: GlyphRun | undefined;
    for (const r of runs) {
      if (r.start <= offset && offset < r.end) {
        candidate = r;
        break;
      }
      if (r.end === offset) candidate = r;
    }
    if (!candidate) {
      const first = runs[0];
      if (first && offset <= first.start) return first.x;
      const last = runs[runs.length - 1];
      return last ? last.x + last.width : 0;
    }
    const r = candidate;
    const k = offset - r.start;
    if (k <= 0) return r.x;
    if (r.kind === 'text') {
      if (isBreakRun(r)) return r.x;
      return r.x + Math.min(r.width, measurer.measure(r.text.slice(0, k), r.format).width);
    }
    return r.x + r.width;
  };

  /** Offset del paragrafo più vicino alla coordinata `rx` (relativa alla riga). */
  const offsetInRun = (r: GlyphRun, rx: Points): number => {
    if (r.kind !== 'text') return rx < r.x + r.width / 2 ? r.start : r.end;
    if (isBreakRun(r)) return r.start;
    // Ricerca binaria sulla larghezza dei prefissi, poi il più vicino dei due.
    let lo = 0;
    let hi = r.text.length;
    const at = (k: number) => r.x + Math.min(r.width, measurer.measure(r.text.slice(0, k), r.format).width);
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (at(mid) <= rx) lo = mid;
      else hi = mid - 1;
    }
    if (lo < r.text.length && rx - at(lo) > at(lo + 1) - rx) lo++;
    return Math.min(r.end, r.start + lo);
  };

  const isWrapped = (layout: DocumentLayout, line: LineBox): boolean =>
    line.end > line.start && lineStarts(layout).has(`${line.paragraphId}:${line.end}`);

  const lineEndOffset = (layout: DocumentLayout, line: LineBox): number => {
    const last = line.runs[line.runs.length - 1];
    if (last && isBreakRun(last)) return last.start;
    if (last && last.kind === 'text' && /\s$/.test(last.text) && isWrapped(layout, line)) return line.end - 1;
    return line.end;
  };

  const caretRect: LayoutEngine['caretRect'] = (layout, pos) => {
    let fallback: PlacedLine | undefined;
    let inBody = false;
    for (const pl of allLines(layout)) {
      const { line } = pl;
      if (line.paragraphId !== pos.paragraphId) continue;
      // Le intestazioni si ripetono su ogni pagina: basta la prima.
      if (!pl.body && inBody) break;
      if (!pl.body && fallback && pl.page !== fallback.page) break;
      if (pl.body) inBody = true;
      if (line.start <= pos.offset && pos.offset < line.end) {
        fallback = pl;
        break;
      }
      if (line.start <= pos.offset && pos.offset <= line.end) fallback = pl;
      else if (line.start > pos.offset) break;
    }
    if (!fallback) return null;
    const { line, page } = fallback;
    return { page, rect: { x: line.x + xInLine(line, pos.offset), y: line.y, width: 0, height: line.height } };
  };

  const hitTest: LayoutEngine['hitTest'] = (layout, pageIndex, x, y) => {
    const page = layout.pages[pageIndex];
    if (!page || page.lines.length === 0) return null;
    let best: LineBox | undefined;
    let bestDy = Infinity;
    let bestDx = Infinity;
    for (const l of page.lines) {
      const dy = y < l.y ? l.y - y : y > l.y + l.height ? y - (l.y + l.height) : 0;
      const dx = x < l.x ? l.x - x : x > l.x + l.width ? x - (l.x + l.width) : 0;
      if (dy < bestDy - 0.01 || (Math.abs(dy - bestDy) <= 0.01 && dx < bestDx)) {
        best = l;
        bestDy = dy;
        bestDx = dx;
      }
    }
    const line = best!;
    const rx = x - line.x;
    const runs = line.runs.filter((r) => r.kind !== 'listMarker');
    let offset = lineEndOffset(layout, line);
    if (runs.length === 0 || rx <= runs[0]!.x) offset = runs[0]?.start ?? line.start;
    else {
      for (const r of runs) {
        if (rx <= r.x + r.width) {
          offset = offsetInRun(r, rx);
          break;
        }
      }
    }
    return { paragraphId: line.paragraphId, offset: Math.min(offset, Math.max(line.end, line.start)) };
  };

  const selectionRects: LayoutEngine['selectionRects'] = (layout, from, to) => {
    const order = paragraphOrder(layout);
    const idx = (p: Position) => order.get(p.paragraphId) ?? -1;
    const cmp = (a: Position, b: Position) => idx(a) - idx(b) || a.offset - b.offset;
    if (cmp(from, to) > 0) [from, to] = [to, from];
    if (cmp(from, to) === 0) return [];
    const out: { page: number; rect: Rect }[] = [];
    for (const { page, line } of allLines(layout)) {
      const pid = line.paragraphId;
      const lineStart: Position = { paragraphId: pid, offset: line.start };
      const lineEnd: Position = { paragraphId: pid, offset: line.end };
      if (!order.has(pid)) continue;
      const a = cmp(from, lineStart) > 0 ? from : lineStart;
      const b = cmp(to, lineEnd) < 0 ? to : lineEnd;
      if (cmp(a, b) > 0) continue;
      const pastEnd = cmp(to, lineEnd) > 0;
      // La riga che termina dove inizia la selezione non è selezionata (a capo automatico).
      if (cmp(a, b) === 0 && !pastEnd) continue;
      if (cmp(a, b) === 0 && isWrapped(layout, line) && a.offset === line.end) continue;
      const xa = xInLine(line, a.offset);
      const xb = xInLine(line, b.offset);
      // Fine paragrafo selezionata: si evidenzia anche il segno di paragrafo.
      const markWidth = pastEnd && !isWrapped(layout, line) ? line.height * 0.3 : 0;
      out.push({ page, rect: { x: line.x + xa, y: line.y, width: Math.max(0, xb - xa) + markWidth, height: line.height } });
    }
    return out;
  };

  return { caretRect, hitTest, selectionRects };
}
