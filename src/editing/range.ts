/**
 * Operazioni su intervalli che possono attraversare più paragrafi:
 * cancellazione di una selezione e inserimento di blocchi al cursore.
 */

import {
  findParagraph,
  newId,
  normalizeInlines,
  paragraphLength,
  paragraphs,
  replaceBlock,
  type Block,
  type DocumentModel,
  type Paragraph,
  type ParagraphLocation,
  type Position,
} from '../core';
import { deleteInParagraph, inlinesLength, splitInlines } from './inlines';

export function locate(doc: DocumentModel, id: string): ParagraphLocation {
  const loc = findParagraph(doc, id);
  if (!loc) throw new Error(`Paragrafo non trovato: ${id}`);
  return loc;
}

/** Paragrafi (in ordine di lettura) dal paragrafo di `start` a quello di `end` inclusi. */
export function paragraphsBetween(doc: DocumentModel, start: Position, end: Position): ParagraphLocation[] {
  const out: ParagraphLocation[] = [];
  let inside = false;
  for (const loc of paragraphs(doc)) {
    if (loc.paragraph.id === start.paragraphId) inside = true;
    if (inside) out.push(loc);
    if (inside && loc.paragraph.id === end.paragraphId) break;
  }
  return out;
}

/** I due paragrafi sono fratelli nello stesso contenitore (sezione o cella)? */
export function sameContainer(a: ParagraphLocation, b: ParagraphLocation): boolean {
  if (a.sectionIndex !== b.sectionIndex || a.path.length !== b.path.length) return false;
  return a.path.slice(0, -1).every((v, i) => v === b.path[i]);
}

function containerBlocks(doc: DocumentModel, loc: ParagraphLocation): readonly Block[] {
  let blocks = doc.sections[loc.sectionIndex]!.blocks;
  const path = loc.path;
  for (let i = 0; i + 1 < path.length; i += 3) {
    const t = blocks[path[i]!] as Extract<Block, { type: 'table' }>;
    blocks = t.rows[path[i + 1]!]!.cells[path[i + 2]!]!.blocks;
  }
  return blocks;
}

/**
 * Cancella il testo fra `start` e `end` (già ordinati). Se i due estremi
 * sono nello stesso contenitore i paragrafi si uniscono e i blocchi in mezzo
 * (tabelle comprese) spariscono; altrimenti, come in Word quando la
 * selezione entra in una tabella, si svuota solo il testo selezionato.
 */
export function deleteRange(doc: DocumentModel, start: Position, end: Position): DocumentModel {
  if (start.paragraphId === end.paragraphId) {
    const p = locate(doc, start.paragraphId).paragraph;
    return replaceBlock(doc, p.id, [deleteInParagraph(p, start.offset, end.offset)]);
  }
  const first = locate(doc, start.paragraphId);
  const last = locate(doc, end.paragraphId);
  if (sameContainer(first, last)) {
    const blocks = containerBlocks(doc, first);
    const from = first.path[first.path.length - 1]!;
    const to = last.path[last.path.length - 1]!;
    const [head] = splitInlines(first.paragraph.children, start.offset);
    const [, tail] = splitInlines(last.paragraph.children, end.offset);
    // Se il primo paragrafo sparisce del tutto, il risultato prende gli
    // attributi dell'ultimo (stile, allineamento...) ma conserva l'id.
    const attrs = start.offset === 0 ? last.paragraph : first.paragraph;
    const merged: Paragraph = { ...attrs, id: first.paragraph.id, children: normalizeInlines([...head, ...tail]) };
    let out = replaceBlock(doc, merged.id, [merged]);
    for (const b of blocks.slice(from + 1, to + 1)) out = replaceBlock(out, b.id, []);
    return out;
  }
  let out = doc;
  for (const loc of paragraphsBetween(doc, start, end)) {
    const p = loc.paragraph;
    const from = p.id === start.paragraphId ? start.offset : 0;
    const to = p.id === end.paragraphId ? end.offset : paragraphLength(p);
    out = replaceBlock(out, p.id, [deleteInParagraph(p, from, to)]);
  }
  return out;
}

/** Copia profonda di blocchi con id nuovi (per incollare senza collisioni). */
export function withFreshIds(blocks: readonly Block[]): Block[] {
  return blocks.map((b): Block => {
    if (b.type === 'paragraph') return { ...b, id: newId() };
    return {
      ...b,
      id: newId(),
      rows: b.rows.map((row) => ({
        ...row,
        id: newId(),
        cells: row.cells.map((cell) => ({ ...cell, id: newId(), blocks: withFreshIds(cell.blocks) })),
      })),
    };
  });
}

/**
 * Inserisce blocchi alla posizione `pos`. Il primo paragrafo inserito si
 * fonde con il testo prima del cursore, l'ultimo con quello dopo; il
 * paragrafo di partenza conserva id e formattazione. Restituisce anche la
 * posizione alla fine del contenuto inserito.
 */
export function insertBlocksAt(
  doc: DocumentModel,
  pos: Position,
  blocks: readonly Block[],
): { doc: DocumentModel; caret: Position } {
  const p = locate(doc, pos.paragraphId).paragraph;
  if (!blocks.length) return { doc, caret: pos };
  const [before, after] = splitInlines(p.children, pos.offset);
  const first = blocks[0]!;
  const last = blocks[blocks.length - 1]!;

  if (blocks.length === 1 && first.type === 'paragraph') {
    const children = normalizeInlines([...before, ...first.children, ...after]);
    const caret = { paragraphId: p.id, offset: inlinesLength(before) + inlinesLength(first.children) };
    return { doc: replaceBlock(doc, p.id, [{ ...p, children }]), caret };
  }

  // Il primo paragrafo si fonde con la testa, l'ultimo con la coda; tutto
  // il resto (tabelle comprese) va in mezzo così com'è.
  const mergeFirst = first.type === 'paragraph';
  const mergeLast = last.type === 'paragraph';
  const middle = blocks.slice(mergeFirst ? 1 : 0, mergeLast ? -1 : blocks.length);
  const out: Block[] = [];
  // Una tabella incollata a inizio paragrafo va prima del paragrafo.
  const headOmitted = !mergeFirst && !before.length;
  if (mergeFirst) out.push({ ...p, children: normalizeInlines([...before, ...first.children]) });
  else if (!headOmitted) out.push({ ...p, children: normalizeInlines(before) });
  out.push(...middle);
  const tailId = headOmitted ? p.id : newId();
  let tail: Paragraph;
  let caretOffset: number;
  if (mergeLast) {
    tail = { ...last, id: tailId, children: normalizeInlines([...last.children, ...after]) };
    caretOffset = inlinesLength(last.children);
  } else {
    tail = { ...p, id: tailId, children: normalizeInlines(after) };
    caretOffset = 0;
  }
  out.push(tail);
  return { doc: replaceBlock(doc, p.id, out), caret: { paragraphId: tail.id, offset: caretOffset } };
}
