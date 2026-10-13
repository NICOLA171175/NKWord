/**
 * Estrazione di un frammento di documento fra due posizioni, usata dagli
 * appunti per copiare/tagliare. Le tabelle coperte per intero restano
 * tabelle; di quelle coperte in parte si tengono solo i paragrafi
 * selezionati, come testo normale.
 */

import {
  comparePositions,
  inlineLength,
  normalizeInlines,
  paragraph,
  type Block,
  type DocumentModel,
  type Inline,
  type Paragraph,
  type Position,
  type Section,
} from '../core';

/** Inline compresi fra gli offset `start` (incluso) e `end` (escluso). */
export function sliceInlines(inlines: readonly Inline[], start: number, end: number): Inline[] {
  const out: Inline[] = [];
  let pos = 0;
  for (const inline of inlines) {
    const len = inlineLength(inline);
    const from = Math.max(start, pos);
    const to = Math.min(end, pos + len);
    if (from < to) {
      if (inline.type === 'text') out.push({ ...inline, text: inline.text.slice(from - pos, to - pos) });
      else out.push(inline);
    }
    pos += len;
    if (pos >= end) break;
  }
  return normalizeInlines(out);
}

interface Walk {
  /** 'before': prima dell'inizio; 'inside': in selezione; 'after': finito. */
  phase: 'before' | 'inside' | 'after';
}

function sliceParagraph(p: Paragraph, from: Position, to: Position, w: Walk): Paragraph | null {
  if (w.phase === 'after') return null;
  const isStart = p.id === from.paragraphId;
  const isEnd = p.id === to.paragraphId;
  if (w.phase === 'before' && !isStart) return null;
  w.phase = 'inside';
  const start = isStart ? from.offset : 0;
  const end = isEnd ? to.offset : Number.POSITIVE_INFINITY;
  if (isEnd) w.phase = 'after';
  if (start === 0 && end >= Number.POSITIVE_INFINITY) return p;
  return { ...p, children: sliceInlines(p.children, start, end) };
}

function sliceBlocks(blocks: readonly Block[], from: Position, to: Position, w: Walk): Block[] {
  const out: Block[] = [];
  for (const b of blocks) {
    if (b.type === 'paragraph') {
      const sliced = sliceParagraph(b, from, to, w);
      if (sliced) out.push(sliced);
      continue;
    }
    const startPhase = w.phase;
    const cellBlocks: Block[] = [];
    for (const row of b.rows) for (const cell of row.cells) cellBlocks.push(...sliceBlocks(cell.blocks, from, to, w));
    // Una tabella resta tale solo se la selezione la attraversa per intero
    // (inizia prima e finisce dopo).
    if (startPhase === 'inside' && w.phase === 'inside') out.push(b);
    else out.push(...cellBlocks);
  }
  return out;
}

/** Frammento di documento compreso fra `from` e `to` (in qualsiasi ordine). */
export function sliceDocument(doc: DocumentModel, a: Position, b: Position): DocumentModel {
  const [from, to] = comparePositions(doc, a, b) <= 0 ? [a, b] : [b, a];
  const w: Walk = { phase: 'before' };
  const sections: Section[] = [];
  for (const s of doc.sections) {
    if (w.phase === 'after') break;
    const blocks = sliceBlocks(s.blocks, from, to, w);
    if (blocks.length) sections.push({ ...s, blocks, header: undefined, footer: undefined });
  }
  if (!sections.length) sections.push({ ...doc.sections[0]!, blocks: [paragraph()], header: undefined, footer: undefined });
  return { ...doc, sections };
}
