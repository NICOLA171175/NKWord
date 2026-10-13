/**
 * Traduzione nel modello del documento, senza rete: qui si decide quali
 * pezzi di testo tradurre e come rimetterli al loro posto.
 *
 * Un "pezzo" è una sequenza di run di testo consecutivi dello stesso
 * paragrafo, interrotta da tabulazioni, immagini, campi, interruzioni o
 * richiami di nota (che restano dove sono). Ogni pezzo si traduce intero,
 * così le frasi restano frasi; il testo tradotto prende la formattazione
 * del run più lungo del pezzo (una parola in grassetto in mezzo a una
 * frase non può essere ritrovata nella traduzione). Gli spazi all'inizio e
 * alla fine del pezzo restano quelli originali.
 *
 * Tutto passa da una funzione `mapText(testo) → testo`: con l'identità si
 * raccolgono i testi da mandare al servizio, con il dizionario delle
 * traduzioni si costruisce il documento tradotto.
 */

import {
  isCollapsed,
  normalizeInlines,
  orderedRange,
  paragraphLength,
  updateParagraph,
  type Block,
  type DocumentModel,
  type EditorState,
  type HeaderFooter,
  type Inline,
  type Paragraph,
  type Section,
  type TextRun,
  type Transaction,
} from '../core';
import { splitRange } from '../editing/inlines';
import { paragraphsBetween } from '../editing/range';

export type MapText = (text: string) => string;

/** Dizionario testo originale → testo tradotto. */
export type Translations = ReadonlyMap<string, string>;

const EDGES = /^(\s*)([\s\S]*?)(\s*)$/;

/** Ha qualcosa da tradurre (non solo spazi, numeri o punteggiatura)? */
export function isTranslatable(text: string): boolean {
  return /\p{L}/u.test(text);
}

/** Traduce un testo mantenendo gli spazi ai bordi; i testi senza lettere restano uguali. */
export function mapKeepingEdges(text: string, mapText: MapText): string {
  const [, lead = '', core = '', trail = ''] = EDGES.exec(text) ?? [];
  if (!isTranslatable(core)) return text;
  const translated = mapText(core).replace(/[\r\n\t]+/g, ' ').trim();
  return translated ? lead + translated + trail : text;
}

/** Riscrive i pezzi di testo di una lista di inline. */
export function rewriteInlines(inlines: readonly Inline[], mapText: MapText): Inline[] {
  const out: Inline[] = [];
  let group: TextRun[] = [];
  const flush = () => {
    if (!group.length) return;
    const text = group.map((r) => r.text).join('');
    const translated = mapKeepingEdges(text, mapText);
    if (translated === text) out.push(...group);
    else {
      const main = group.reduce((a, b) => (b.text.length > a.text.length ? b : a));
      out.push({ ...main, text: translated });
    }
    group = [];
  };
  for (const inline of inlines) {
    // Il testo eliminato in revisione resta com'è e interrompe il pezzo.
    if (inline.type === 'text' && inline.revision?.kind !== 'delete') group.push(inline);
    else {
      flush();
      out.push(inline);
    }
  }
  flush();
  return out;
}

/** Riscrive il paragrafo fra gli offset dati (tutto, se non indicati). */
export function rewriteParagraph(p: Paragraph, mapText: MapText, from = 0, to = paragraphLength(p)): Paragraph {
  const [head, middle, tail] = splitRange(p.children, from, to);
  const rewritten = rewriteInlines(middle, mapText);
  if (rewritten.length === middle.length && rewritten.every((x, i) => x === middle[i])) return p;
  return { ...p, children: normalizeInlines([...head, ...rewritten, ...tail]) };
}

export function rewriteBlocks(blocks: readonly Block[], mapText: MapText): readonly Block[] {
  let changed = false;
  const out = blocks.map((b): Block => {
    let next: Block;
    if (b.type === 'paragraph') next = rewriteParagraph(b, mapText);
    else {
      let rowsChanged = false;
      const rows = b.rows.map((row) => {
        let cellsChanged = false;
        const cells = row.cells.map((cell) => {
          const inner = rewriteBlocks(cell.blocks, mapText);
          if (inner === cell.blocks) return cell;
          cellsChanged = true;
          return { ...cell, blocks: inner };
        });
        if (!cellsChanged) return row;
        rowsChanged = true;
        return { ...row, cells };
      });
      next = rowsChanged ? { ...b, rows } : b;
    }
    if (next !== b) changed = true;
    return next;
  });
  return changed ? out : blocks;
}

function rewriteHeaderFooter(hf: HeaderFooter | undefined, mapText: MapText): HeaderFooter | undefined {
  if (!hf) return hf;
  const out: { -readonly [K in keyof HeaderFooter]: HeaderFooter[K] } = { ...hf };
  for (const key of ['default', 'first', 'even'] as const) {
    const blocks = hf[key];
    if (blocks) out[key] = rewriteBlocks(blocks, mapText);
  }
  return out;
}

/**
 * Riscrive tutto il testo del documento, come "Traduci documento" di
 * Word: corpo, tabelle, intestazioni e piè di pagina, note a piè di pagina
 * e testo dentro forme e caselle di testo. I commenti restano com'erano.
 */
export function rewriteDocument(doc: DocumentModel, mapText: MapText): DocumentModel {
  const sections = doc.sections.map(
    (s): Section => ({
      ...s,
      blocks: rewriteBlocks(s.blocks, mapText),
      ...(s.header ? { header: rewriteHeaderFooter(s.header, mapText) } : {}),
      ...(s.footer ? { footer: rewriteHeaderFooter(s.footer, mapText) } : {}),
    }),
  );
  const next: { -readonly [K in keyof DocumentModel]: DocumentModel[K] } = { ...doc, sections };
  if (doc.footnotes) {
    next.footnotes = Object.fromEntries(
      Object.entries(doc.footnotes).map(([id, note]) => [id, { ...note, blocks: rewriteBlocks(note.blocks, mapText) }]),
    );
  }
  if (doc.drawings) {
    next.drawings = Object.fromEntries(
      Object.entries(doc.drawings).map(([id, d]) => [
        id,
        d.text ? { ...d, text: { ...d.text, text: d.text.text.split('\n').map((line) => mapKeepingEdges(line, mapText)).join('\n') } } : d,
      ]),
    );
  }
  if (doc.metadata.title && isTranslatable(doc.metadata.title)) {
    next.metadata = { ...doc.metadata, title: mapKeepingEdges(doc.metadata.title, mapText) };
  }
  return next;
}

/** Raccoglie i testi distinti che `rewrite` chiederebbe di tradurre, nell'ordine. */
export function collectTexts(rewrite: (mapText: MapText) => unknown): string[] {
  const seen = new Set<string>();
  rewrite((text) => {
    seen.add(text);
    return text;
  });
  return [...seen];
}

/** `mapText` che legge il dizionario e lascia invariato ciò che non c'è. */
export function lookup(translations: Translations): MapText {
  return (text) => translations.get(text) ?? text;
}

// ---------------------------------------------------------------------------
// Selezione

interface SelectedPart {
  readonly paragraph: Paragraph;
  readonly from: number;
  readonly to: number;
}

function selectedParts(state: EditorState): SelectedPart[] {
  if (isCollapsed(state.selection)) return [];
  const [start, end] = orderedRange(state.doc, state.selection);
  return paragraphsBetween(state.doc, start, end).map(({ paragraph }) => ({
    paragraph,
    from: paragraph.id === start.paragraphId ? start.offset : 0,
    to: paragraph.id === end.paragraphId ? end.offset : paragraphLength(paragraph),
  }));
}

/** Testo visibile di una parte di paragrafo ('\t' per le tabulazioni, niente oggetti). */
function partText(inlines: readonly Inline[]): string {
  return inlines.map((c) => (c.type === 'text' ? c.text : c.type === 'tab' ? '\t' : c.type === 'break' ? '\n' : '')).join('');
}

/** C'è testo selezionato da tradurre? */
export function hasSelectedText(state: EditorState): boolean {
  return selectedParts(state).some((part) => isTranslatable(partText(splitRange(part.paragraph.children, part.from, part.to)[1])));
}

/** Testo della selezione, un paragrafo per riga (con `mapText`, già tradotto). */
export function selectionText(state: EditorState, mapText: MapText = (t) => t): string {
  return selectedParts(state)
    .map((part) => partText(rewriteInlines(splitRange(part.paragraph.children, part.from, part.to)[1], mapText)))
    .join('\n');
}

/** Sostituisce la selezione con la sua traduzione e la lascia selezionata. */
export function rewriteSelection(state: EditorState, mapText: MapText): Transaction | null {
  const parts = selectedParts(state);
  if (!parts.length) return null;
  let doc = state.doc;
  let endOffset = 0;
  const last = parts[parts.length - 1]!;
  for (const part of parts) {
    const next = rewriteParagraph(part.paragraph, mapText, part.from, part.to);
    if (part === last) endOffset = part.to + paragraphLength(next) - paragraphLength(part.paragraph);
    if (next !== part.paragraph) doc = updateParagraph(doc, next);
  }
  if (doc === state.doc) return null;
  const [start] = orderedRange(state.doc, state.selection);
  return {
    state: { doc, selection: { anchor: start, focus: { paragraphId: last.paragraph.id, offset: endOffset } } },
    meta: { label: 'Traduci selezione' },
  };
}
