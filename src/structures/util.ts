/**
 * Funzioni di supporto del modulo strutture: selezione, divisione dei
 * paragrafi e inserimento al cursore. Sono volutamente locali al modulo:
 * le operazioni di modifica generali appartengono al modulo di editing.
 */

import {
  caret,
  findParagraph,
  isCollapsed,
  normalizeInlines,
  orderedRange,
  paragraphLength,
  paragraphs,
  replaceBlock,
  newId,
  type Block,
  type CharFormat,
  type Command,
  type CommandArgs,
  type CommandId,
  type DocumentModel,
  type EditorState,
  type Inline,
  type Paragraph,
  type ParagraphLocation,
  type Position,
  type Section,
  type Table,
  type TableCell,
} from '../core';

/** Comando tipizzato secondo `CommandArgs`. */
export function defineCommand<K extends CommandId>(
  id: K,
  def: Omit<Command<CommandArgs[K]>, 'id'>,
): Command<CommandArgs[K]> {
  return { id, ...def };
}

export function requireParagraph(doc: DocumentModel, id: string): ParagraphLocation {
  const loc = findParagraph(doc, id);
  if (!loc) throw new Error(`Paragrafo non trovato: ${id}`);
  return loc;
}

/** Paragrafi coperti dalla selezione, in ordine di lettura. */
export function selectedParagraphs(state: EditorState): ParagraphLocation[] {
  const [from, to] = orderedRange(state.doc, state.selection);
  const out: ParagraphLocation[] = [];
  let inside = false;
  for (const loc of paragraphs(state.doc)) {
    if (loc.paragraph.id === from.paragraphId) inside = true;
    if (inside) out.push(loc);
    if (loc.paragraph.id === to.paragraphId) break;
  }
  return out;
}

/** Primo paragrafo (anche annidato in tabelle) di una lista di blocchi. */
export function firstParagraphIn(blocks: readonly Block[]): Paragraph | undefined {
  for (const b of blocks) {
    if (b.type === 'paragraph') return b;
    for (const row of b.rows) for (const cell of row.cells) {
      const p = firstParagraphIn(cell.blocks);
      if (p) return p;
    }
  }
  return undefined;
}

/** Ultimo paragrafo (anche annidato in tabelle) di una lista di blocchi. */
export function lastParagraphIn(blocks: readonly Block[]): Paragraph | undefined {
  for (let i = blocks.length - 1; i >= 0; i--) {
    const b = blocks[i]!;
    if (b.type === 'paragraph') return b;
    for (let r = b.rows.length - 1; r >= 0; r--) {
      const cells = b.rows[r]!.cells;
      for (let c = cells.length - 1; c >= 0; c--) {
        const p = lastParagraphIn(cells[c]!.blocks);
        if (p) return p;
      }
    }
  }
  return undefined;
}

/** Blocchi che contengono il paragrafo indicato da `loc` (sezione o cella). */
export function containerOf(doc: DocumentModel, loc: ParagraphLocation): readonly Block[] {
  let blocks = doc.sections[loc.sectionIndex]!.blocks;
  for (let i = 0; i + 3 < loc.path.length; i += 3) {
    const t = blocks[loc.path[i]!] as Table;
    blocks = t.rows[loc.path[i + 1]!]!.cells[loc.path[i + 2]!]!.blocks;
  }
  return blocks;
}

/** Cella più interna che contiene il paragrafo, se esiste. */
export function innermostCell(doc: DocumentModel, loc: ParagraphLocation): { table: Table; cell: TableCell } | undefined {
  let blocks = doc.sections[loc.sectionIndex]!.blocks;
  let found: { table: Table; cell: TableCell } | undefined;
  for (let i = 0; i + 3 < loc.path.length; i += 3) {
    const table = blocks[loc.path[i]!] as Table;
    const cell = table.rows[loc.path[i + 1]!]!.cells[loc.path[i + 2]!]!;
    found = { table, cell };
    blocks = cell.blocks;
  }
  return found;
}

/** Divide gli inline all'offset indicato. */
export function splitInlines(children: readonly Inline[], offset: number): [Inline[], Inline[]] {
  const left: Inline[] = [];
  const right: Inline[] = [];
  let pos = 0;
  for (const c of children) {
    const len = c.type === 'text' ? c.text.length : 1;
    if (pos + len <= offset) left.push(c);
    else if (pos >= offset) right.push(c);
    else if (c.type === 'text') {
      left.push({ ...c, text: c.text.slice(0, offset - pos) });
      right.push({ ...c, text: c.text.slice(offset - pos) });
    }
    pos += len;
  }
  return [normalizeInlines(left), normalizeInlines(right)];
}

/**
 * Divide un paragrafo in due: la parte sinistra mantiene l'id, la destra ne
 * riceve uno nuovo e conserva stile e formattazione del paragrafo.
 */
export function splitParagraph(p: Paragraph, offset: number): [Paragraph, Paragraph] {
  const [l, r] = splitInlines(p.children, offset);
  return [
    { ...p, children: l },
    { ...p, id: newId(), children: r },
  ];
}

/** Formattazione di carattere al cursore: run precedente (o successivo), poi segno di paragrafo. */
export function charFormatAt(state: EditorState, pos: Position): CharFormat {
  const p = requireParagraph(state.doc, pos.paragraphId).paragraph;
  let at = 0;
  let before: CharFormat | undefined;
  let after: CharFormat | undefined;
  for (const c of p.children) {
    const len = c.type === 'text' ? c.text.length : 1;
    if (c.type === 'text' || c.type === 'field') {
      if (at < pos.offset) before = c.format;
      else if (!after) after = c.format;
    }
    at += len;
  }
  return { ...(before ?? after ?? p.markFormat ?? {}), ...state.pendingFormat };
}

/**
 * Elimina il contenuto selezionato e restituisce il documento e la posizione
 * in cui inserire. Gestisce selezioni in un solo paragrafo o fra paragrafi
 * dello stesso contenitore; negli altri casi (es. da una cella a fuori della
 * tabella) non elimina nulla e inserisce all'inizio della selezione.
 */
export function deleteSelection(state: EditorState): { doc: DocumentModel; pos: Position } {
  const { doc, selection } = state;
  const [from, to] = orderedRange(doc, selection);
  if (isCollapsed(selection)) return { doc, pos: from };
  const a = requireParagraph(doc, from.paragraphId);
  if (from.paragraphId === to.paragraphId) {
    const [left] = splitInlines(a.paragraph.children, from.offset);
    const [, right] = splitInlines(a.paragraph.children, to.offset);
    const p = { ...a.paragraph, children: normalizeInlines([...left, ...right]) };
    return { doc: replaceBlock(doc, p.id, [p]), pos: from };
  }
  const b = requireParagraph(doc, to.paragraphId);
  const sameContainer =
    a.sectionIndex === b.sectionIndex &&
    a.path.length === b.path.length &&
    a.path.slice(0, -1).every((v, i) => v === b.path[i]);
  if (!sameContainer) return { doc, pos: from };
  const container = containerOf(doc, a);
  const i0 = a.path[a.path.length - 1]!;
  const i1 = b.path[b.path.length - 1]!;
  const [left] = splitInlines(a.paragraph.children, from.offset);
  const [, right] = splitInlines(b.paragraph.children, to.offset);
  let next = replaceBlock(doc, a.paragraph.id, [{ ...a.paragraph, children: normalizeInlines([...left, ...right]) }]);
  for (let i = i0 + 1; i <= i1; i++) next = replaceBlock(next, container[i]!.id, []);
  return { doc: next, pos: from };
}

/** Inserisce un inline al cursore (sostituendo la selezione); il cursore va dopo l'inline. */
export function insertInline(state: EditorState, inline: Inline): EditorState {
  const { doc, pos } = deleteSelection(state);
  const p = requireParagraph(doc, pos.paragraphId).paragraph;
  const [l, r] = splitInlines(p.children, pos.offset);
  const updated = { ...p, children: normalizeInlines([...l, inline, ...r]) };
  return {
    doc: replaceBlock(doc, p.id, [updated]),
    selection: caret(p.id, pos.offset + 1),
  };
}

/**
 * Inserisce blocchi al cursore dividendo il paragrafo. Restituisce il nuovo
 * documento e il paragrafo che segue i blocchi inseriti.
 */
export function insertBlocks(
  state: EditorState,
  blocks: readonly Block[],
): { doc: DocumentModel; after: Paragraph } {
  const { doc, pos } = deleteSelection(state);
  const p = requireParagraph(doc, pos.paragraphId).paragraph;
  if (pos.offset === 0) {
    return { doc: replaceBlock(doc, p.id, [...blocks, p]), after: p };
  }
  const [left, right] = splitParagraph(p, pos.offset);
  return { doc: replaceBlock(doc, p.id, [left, ...blocks, right]), after: right };
}

/** Indici delle sezioni coperte dalla selezione. */
export function selectedSections(state: EditorState): number[] {
  const [from, to] = orderedRange(state.doc, state.selection);
  const s0 = requireParagraph(state.doc, from.paragraphId).sectionIndex;
  const s1 = requireParagraph(state.doc, to.paragraphId).sectionIndex;
  return Array.from({ length: s1 - s0 + 1 }, (_, i) => s0 + i);
}

export function replaceSection(doc: DocumentModel, index: number, section: Section): DocumentModel {
  return { ...doc, sections: doc.sections.map((s, i) => (i === index ? section : s)) };
}
