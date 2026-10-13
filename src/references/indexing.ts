/**
 * Indice analitico (Riferimenti > Segna voce / Inserisci indice), come in
 * Word.
 *
 * - Segnare una voce mette `indexEntry` sul testo scelto ("voce" oppure
 *   "voce:sottovoce", come il campo XE di Word). Il segno non si vede.
 * - L'indice è una sequenza di paragrafi di primo livello con gli stili di
 *   Word "index heading" (le lettere A, B, C...) e "index 1".."index 2" (le
 *   voci). Ogni voce elenca le pagine in cui compare; ogni numero di pagina
 *   è un collegamento al testo segnato.
 */

import {
  findParagraph,
  isCollapsed,
  paragraph,
  paragraphText,
  paragraphs,
  replaceBlock,
  textRun,
  STYLE_NORMAL,
  type Block,
  type DocumentModel,
  type EditorState,
  type Id,
  type Inline,
  type LayoutEngine,
  type Paragraph,
  type Style,
  type StyleSheet,
  type Transaction,
} from '../core';
import { mapFormatInRange } from '../editing/inlines';
import { paragraphsBetween } from '../editing/range';
import { selectionRange } from '../editing/selection';
import { isWordChar } from '../editing/words';
import { tocAnchor } from '../toc/toc';
import { ensureStyles, findStyleId, pageMap, replaceBlocks } from './common';

export const INDEX_HEADING_STYLE = 'IndexHeading';
export const NO_INDEX_ENTRIES = 'Nessuna voce di indice trovata.';
export const MAX_INDEX_LEVEL = 2;

export function indexEntryStyleId(level: number): Id {
  return `Index${level}`;
}

const HEADING_NAME = /^index heading$/i;

/** Livello (1..9) se lo stile è una voce d'indice, altrimenti 0. */
export function indexEntryLevel(sheet: StyleSheet, styleId: Id | undefined): number {
  if (!styleId) return 0;
  const byId = /^Index([1-9])$/.exec(styleId);
  if (byId) return Number(byId[1]);
  const byName = /^index ([1-9])$/i.exec(sheet.styles[styleId]?.name ?? '');
  return byName ? Number(byName[1]) : 0;
}

export function isIndexStyle(sheet: StyleSheet, styleId: Id | undefined): boolean {
  if (!styleId) return false;
  return indexEntryLevel(sheet, styleId) > 0 || styleId === INDEX_HEADING_STYLE || HEADING_NAME.test(sheet.styles[styleId]?.name ?? '');
}

const entryStyle = (level: number): Style => ({
  id: indexEntryStyleId(level),
  name: `index ${level}`,
  type: 'paragraph',
  basedOn: STYLE_NORMAL,
  next: STYLE_NORMAL,
  // Rientro sporgente: le righe lunghe continuano più a destra, come in Word.
  paraFormat: { indentLeft: level * 11, indentFirstLine: -11, spaceAfter: 0 },
});

const headingStyle: Style = {
  id: INDEX_HEADING_STYLE,
  name: 'index heading',
  type: 'paragraph',
  basedOn: STYLE_NORMAL,
  next: STYLE_NORMAL,
  paraFormat: { spaceBefore: 10, spaceAfter: 4, keepWithNext: true },
  charFormat: { bold: true, fontSize: 13, color: '#2f5496' },
};

export function ensureIndexStyles(sheet: StyleSheet): StyleSheet {
  return ensureStyles(sheet, [headingStyle, entryStyle(1), entryStyle(2)]);
}

// ---------------------------------------------------------------------------
// Voci segnate

/** Parola che tocca il cursore (anche se il cursore è all'inizio o alla fine). */
function wordAt(text: string, offset: number): [number, number] | null {
  let a = offset;
  let b = offset;
  while (a > 0 && isWordChar(text[a - 1])) a--;
  while (b < text.length && isWordChar(text[b])) b++;
  return a < b ? [a, b] : null;
}

/** "voce:sottovoce" → ["voce", "sottovoce"]; spazi in eccesso tolti. */
export function splitEntry(entry: string): [string, string | undefined] {
  const i = entry.indexOf(':');
  const clean = (s: string) => s.replace(/\s+/g, ' ').trim();
  if (i < 0) return [clean(entry), undefined];
  const sub = clean(entry.slice(i + 1));
  return [clean(entry.slice(0, i)), sub || undefined];
}

export function joinEntry(main: string, sub?: string): string {
  const m = main.replace(/\s+/g, ' ').replace(/:/g, '').trim();
  const s = sub?.replace(/\s+/g, ' ').replace(/:/g, '').trim();
  return s ? `${m}:${s}` : m;
}

export interface IndexMark {
  readonly entry: string;
  readonly paragraphId: Id;
}

/** Voci segnate nel corpo, in ordine di lettura (una per tratto di testo segnato). */
export function collectMarks(doc: DocumentModel): IndexMark[] {
  const out: IndexMark[] = [];
  for (const { paragraph: p } of paragraphs(doc)) {
    let previous: string | undefined;
    for (const c of p.children) {
      const entry = c.type === 'text' ? c.format.indexEntry : undefined;
      if (entry && entry !== previous) out.push({ entry, paragraphId: p.id });
      previous = entry;
    }
  }
  return out;
}

/** Testo che la finestra "Segna voce" propone: la selezione (su una riga) o la parola al cursore. */
export function suggestedEntry(state: EditorState): string {
  const [start, end] = selectionRange(state);
  const p = findParagraph(state.doc, start.paragraphId)?.paragraph;
  if (!p) return '';
  const text = paragraphText(p);
  if (isCollapsed(state.selection)) {
    const word = wordAt(text, start.offset);
    return word ? text.slice(word[0], word[1]) : '';
  }
  const to = end.paragraphId === p.id ? end.offset : text.length;
  return text.slice(start.offset, to).replace(/[\t\n￼]/g, ' ').replace(/\s+/g, ' ').trim();
}

/** Intervalli [from, to) di `word` intera nel paragrafo (maiuscole comprese, come "Segna tutto" di Word). */
function occurrences(p: Paragraph, word: string): [number, number][] {
  const text = paragraphText(p);
  const out: [number, number][] = [];
  const isWord = (ch: string | undefined) => !!ch && /[\p{L}\p{N}_]/u.test(ch);
  for (let i = text.indexOf(word); i >= 0; i = text.indexOf(word, i + word.length)) {
    if (!isWord(text[i - 1]) && !isWord(text[i + word.length])) out.push([i, i + word.length]);
  }
  return out;
}

const mark = (p: Paragraph, from: number, to: number, entry: string) =>
  mapFormatInRange(p, from, to, (format, inline) => ({
    format: { ...format, indexEntry: entry },
    styleId: inline.type === 'text' ? inline.styleId : undefined,
  }));

export interface MarkArgs {
  /** "voce" oppure "voce:sottovoce". */
  readonly entry: string;
  /** Segna tutte le volte che il testo selezionato compare nel documento. */
  readonly all?: boolean;
}

/**
 * Segna una voce d'indice sulla selezione (o sulla parola al cursore). Con
 * `all` segna anche ogni altra occorrenza dello stesso testo.
 */
export function markEntry(state: EditorState, args: MarkArgs): Transaction | null {
  const [main, sub] = splitEntry(args.entry);
  if (!main) return null;
  const entry = joinEntry(main, sub);
  let [start, end] = selectionRange(state);
  if (isCollapsed(state.selection)) {
    const p = findParagraph(state.doc, start.paragraphId)?.paragraph;
    const word = p && wordAt(paragraphText(p), start.offset);
    if (word) {
      start = { paragraphId: p.id, offset: word[0] };
      end = { paragraphId: p.id, offset: word[1] };
    }
  }
  let doc = state.doc;
  if (start.paragraphId !== end.paragraphId || start.offset !== end.offset) {
    for (const { paragraph: p } of paragraphsBetween(doc, start, end)) {
      const from = p.id === start.paragraphId ? start.offset : 0;
      const to = p.id === end.paragraphId ? end.offset : paragraphText(p).length;
      doc = replaceBlock(doc, p.id, [mark(p, from, to, entry)]);
    }
  } else {
    // Nessun testo vicino: un segno invisibile non si può attaccare a niente.
    return null;
  }
  if (args.all) {
    const word = suggestedEntry(state);
    if (word) {
      for (const { paragraph: p } of paragraphs(doc)) {
        if (isIndexStyle(doc.styles, p.styleId)) continue;
        let next = p;
        for (const [from, to] of occurrences(p, word)) next = mark(next, from, to, entry);
        if (next !== p) doc = replaceBlock(doc, p.id, [next]);
      }
    }
  }
  if (doc === state.doc) return null;
  return { state: { doc, selection: state.selection }, meta: { label: 'Segna voce di indice' } };
}

// ---------------------------------------------------------------------------
// Costruzione dell'indice

export interface IndexEntry {
  readonly text: string;
  readonly level: number;
  /** Pagine in ordine, ognuna con il paragrafo a cui porta il collegamento. */
  readonly pages: readonly { page: number; paragraphId: Id }[];
}

const compare = (a: string, b: string) => a.localeCompare(b, 'it', { sensitivity: 'base' }) || a.localeCompare(b, 'it');

/** Lettera sotto cui va una voce: iniziale senza accenti, "#" per numeri e simboli. */
export function indexLetter(text: string): string {
  const first = text.normalize('NFD').replace(/\p{M}/gu, '').charAt(0).toLocaleUpperCase('it');
  return /\p{L}/u.test(first) ? first : '#';
}

/** Voci ordinate (principali e sottovoci) con le pagine, raggruppate per lettera. */
export function buildEntries(marks: readonly IndexMark[], pages: ReadonlyMap<Id, number>): { letter: string; entries: IndexEntry[] }[] {
  type Node = { text: string; pages: Map<number, Id>; subs: Map<string, { text: string; pages: Map<number, Id> }> };
  const mains = new Map<string, Node>();
  for (const m of marks) {
    const [main, sub] = splitEntry(m.entry);
    if (!main) continue;
    const key = main.toLocaleLowerCase('it');
    let node = mains.get(key);
    if (!node) mains.set(key, (node = { text: main, pages: new Map(), subs: new Map() }));
    const page = pages.get(m.paragraphId) ?? 1;
    let target = node.pages;
    if (sub) {
      const subKey = sub.toLocaleLowerCase('it');
      let s = node.subs.get(subKey);
      if (!s) node.subs.set(subKey, (s = { text: sub, pages: new Map() }));
      target = s.pages;
    }
    if (!target.has(page)) target.set(page, m.paragraphId);
  }
  const listPages = (map: Map<number, Id>) => [...map].sort((a, b) => a[0] - b[0]).map(([page, paragraphId]) => ({ page, paragraphId }));
  const groups: { letter: string; entries: IndexEntry[] }[] = [];
  for (const node of [...mains.values()].sort((a, b) => compare(a.text, b.text))) {
    const letter = indexLetter(node.text);
    let group = groups[groups.length - 1];
    if (group?.letter !== letter) groups.push((group = { letter, entries: [] }));
    group.entries.push({ text: node.text, level: 1, pages: listPages(node.pages) });
    for (const s of [...node.subs.values()].sort((a, b) => compare(a.text, b.text))) {
      group.entries.push({ text: s.text, level: 2, pages: listPages(s.pages) });
    }
  }
  // "#" (numeri e simboli) va in testa, come in Word.
  return groups.sort((a, b) => (a.letter === '#' ? -1 : b.letter === '#' ? 1 : 0));
}

/** Paragrafi dell'indice per le voci segnate. */
export function buildIndexBlocks(sheet: StyleSheet, marks: readonly IndexMark[], pages: ReadonlyMap<Id, number>): Paragraph[] {
  const headingId = findStyleId(sheet, INDEX_HEADING_STYLE, HEADING_NAME) ?? INDEX_HEADING_STYLE;
  const levelStyle = (level: number) => findStyleId(sheet, indexEntryStyleId(level), new RegExp(`^index ${level}$`, 'i')) ?? indexEntryStyleId(level);
  const groups = buildEntries(marks, pages);
  if (!groups.length) return [paragraph(NO_INDEX_ENTRIES, {}, levelStyle(1))];
  const blocks: Paragraph[] = [];
  for (const g of groups) {
    blocks.push(paragraph(g.letter, {}, headingId));
    for (const e of g.entries) {
      const children: Inline[] = [textRun(e.text)];
      for (const { page, paragraphId } of e.pages) children.push(textRun(', '), textRun(String(page), { link: tocAnchor(paragraphId) }));
      blocks.push(paragraph(children, {}, levelStyle(e.level)));
    }
  }
  return blocks;
}

// ---------------------------------------------------------------------------
// Posizione dell'indice

export interface IndexLocation {
  readonly sectionIndex: number;
  /** Indici dei blocchi della sezione: [start, end). */
  readonly start: number;
  readonly end: number;
}

const isIndexBlock = (sheet: StyleSheet, b: Block | undefined): b is Paragraph => b?.type === 'paragraph' && isIndexStyle(sheet, b.styleId);

/** Primo indice del documento (blocchi consecutivi con stili dell'indice). */
export function findIndex(doc: DocumentModel): IndexLocation | undefined {
  for (let s = 0; s < doc.sections.length; s++) {
    const blocks = doc.sections[s]!.blocks;
    const start = blocks.findIndex((b) => isIndexBlock(doc.styles, b));
    if (start < 0) continue;
    let end = start;
    while (isIndexBlock(doc.styles, blocks[end])) end++;
    return { sectionIndex: s, start, end };
  }
  return undefined;
}

/** Numero massimo di impaginazioni: l'indice stesso può spostare il testo segnato. */
const MAX_PASSES = 3;

/**
 * Scrive l'indice al posto dei blocchi [start, end) della sezione, con le
 * voci e le pagine aggiornate. I paragrafi riusano gli id di quelli che
 * sostituiscono, così il cursore dentro l'indice resta valido.
 */
export function writeIndex(doc: DocumentModel, at: IndexLocation, layout?: LayoutEngine): DocumentModel {
  const old = doc.sections[at.sectionIndex]!.blocks.slice(at.start, at.end);
  const base: DocumentModel = { ...doc, styles: ensureIndexStyles(doc.styles) };
  const marks = collectMarks(base);
  const write = (pages: ReadonlyMap<Id, number>) => {
    const blocks = buildIndexBlocks(base.styles, marks, pages).map((b, i) => (old[i] ? { ...b, id: old[i]!.id } : b));
    return replaceBlocks(base, at.sectionIndex, at.start, at.end, blocks, () => paragraph());
  };
  let pages: ReadonlyMap<Id, number> = pageMap(base, layout);
  for (let pass = 0; pass < MAX_PASSES; pass++) {
    const result = write(pages);
    const next = pageMap(result, layout);
    if (marks.every((m) => next.get(m.paragraphId) === pages.get(m.paragraphId))) return result;
    pages = next;
  }
  return write(pages);
}

