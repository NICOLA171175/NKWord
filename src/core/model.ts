/**
 * Costruttori e funzioni di sola lettura sul modello, condivisi da tutti i
 * moduli. Le operazioni di modifica vere e proprie (inserire testo, unire
 * paragrafi...) appartengono al modulo di editing, non al core.
 */

import type {
  Block,
  CharFormat,
  DocumentModel,
  Id,
  Inline,
  PageSetup,
  ParaFormat,
  Paragraph,
  Position,
  Revision,
  Section,
  Selection,
  StyleSheet,
  Table,
  TableCell,
  TextRun,
} from './types';
import { PAGE_SIZES, cmToPt } from './units';

// ---------------------------------------------------------------------------
// Id

let idCounter = 0;
const idPrefix = Math.random().toString(36).slice(2, 7);

/** Genera un id univoco nella sessione. */
export function newId(): Id {
  idCounter += 1;
  return `${idPrefix}${idCounter.toString(36)}`;
}

// ---------------------------------------------------------------------------
// Costruttori

export function textRun(text: string, format: CharFormat = {}): TextRun {
  return { type: 'text', text, format };
}

export function paragraph(
  children: readonly Inline[] | string = [],
  format: ParaFormat = {},
  styleId?: Id,
): Paragraph {
  const inlines = typeof children === 'string' ? (children ? [textRun(children)] : []) : children;
  const p: Paragraph = { type: 'paragraph', id: newId(), format, children: inlines };
  return styleId ? { ...p, styleId } : p;
}

export function tableCell(blocks: readonly Block[] = [paragraph()]): TableCell {
  return { id: newId(), blocks: blocks.length ? blocks : [paragraph()], format: {} };
}

export function table(rows: number, cols: number, totalWidth: number): Table {
  return {
    type: 'table',
    id: newId(),
    columnWidths: Array.from({ length: cols }, () => totalWidth / cols),
    rows: Array.from({ length: rows }, () => ({
      id: newId(),
      cells: Array.from({ length: cols }, () => tableCell()),
    })),
    format: {},
  };
}

export function defaultPageSetup(): PageSetup {
  return {
    ...PAGE_SIZES.A4,
    orientation: 'portrait',
    margins: {
      top: cmToPt(2.5),
      right: cmToPt(2),
      bottom: cmToPt(2),
      left: cmToPt(2),
      header: cmToPt(1.25),
      footer: cmToPt(1.25),
    },
  };
}

export function section(blocks: readonly Block[] = [paragraph()], page = defaultPageSetup()): Section {
  return { id: newId(), page, blocks: blocks.length ? blocks : [paragraph()] };
}

export const STYLE_NORMAL = 'Normal';

export function defaultStyleSheet(): StyleSheet {
  const heading = (level: number, size: number) => ({
    id: `Heading${level}`,
    name: `Titolo ${level}`,
    type: 'paragraph' as const,
    basedOn: STYLE_NORMAL,
    next: STYLE_NORMAL,
    quickStyle: true,
    paraFormat: { spaceBefore: 12, spaceAfter: 3, keepWithNext: true, outlineLevel: level },
    charFormat: { fontSize: size, bold: true, color: '#2f5496' },
  });
  return {
    defaults: {
      char: { fontFamily: 'Calibri', fontSize: 11, color: '#000000' },
      para: { align: 'left', spaceAfter: 8, lineSpacing: { rule: 'multiple', value: 1.08 } },
    },
    defaultParagraphStyle: STYLE_NORMAL,
    styles: {
      [STYLE_NORMAL]: { id: STYLE_NORMAL, name: 'Normale', type: 'paragraph', quickStyle: true },
      Heading1: heading(1, 16),
      Heading2: heading(2, 13),
      Heading3: heading(3, 12),
      Title: {
        id: 'Title',
        name: 'Titolo',
        type: 'paragraph',
        basedOn: STYLE_NORMAL,
        next: STYLE_NORMAL,
        quickStyle: true,
        paraFormat: { spaceAfter: 0 },
        charFormat: { fontSize: 28, fontFamily: 'Calibri Light' },
      },
    },
  };
}

/** Documento nuovo: una sezione A4 con un paragrafo vuoto. */
export function createDocument(blocks: readonly Block[] = [paragraph()]): DocumentModel {
  const now = new Date().toISOString();
  return {
    sections: [section(blocks)],
    styles: defaultStyleSheet(),
    lists: {},
    metadata: { created: now, modified: now },
    media: {},
  };
}

// ---------------------------------------------------------------------------
// Lunghezze e testo

export function inlineLength(inline: Inline): number {
  return inline.type === 'text' ? inline.text.length : 1;
}

export function paragraphLength(p: Paragraph): number {
  let n = 0;
  for (const c of p.children) n += inlineLength(c);
  return n;
}

/** Testo semplice del paragrafo; i tab diventano '\t', le interruzioni '\n'. */
export function paragraphText(p: Paragraph): string {
  return p.children
    .map((c) => {
      switch (c.type) {
        case 'text':
          return c.text;
        case 'tab':
          return '\t';
        case 'break':
          return '\n';
        default:
          return '￼';
      }
    })
    .join('');
}

function sameFormat(a: CharFormat, b: CharFormat): boolean {
  const ka = Object.keys(a) as (keyof CharFormat)[];
  const kb = Object.keys(b);
  return ka.length === kb.length && ka.every((k) => a[k] === b[k]);
}

/** Due revisioni sono uguali se hanno stesso tipo, autore e data (o mancano entrambe). */
export function sameRevision(a: Revision | undefined, b: Revision | undefined): boolean {
  if (a === b) return true;
  return !!a && !!b && a.kind === b.kind && a.author === b.author && a.date === b.date;
}

/** Unisce i run di testo adiacenti con stessa formattazione e toglie quelli vuoti. */
export function normalizeInlines(inlines: readonly Inline[]): Inline[] {
  const out: Inline[] = [];
  for (const inline of inlines) {
    if (inline.type === 'text') {
      if (inline.text === '') continue;
      const prev = out[out.length - 1];
      if (
        prev?.type === 'text' &&
        prev.styleId === inline.styleId &&
        sameFormat(prev.format, inline.format) &&
        sameRevision(prev.revision, inline.revision)
      ) {
        out[out.length - 1] = { ...prev, text: prev.text + inline.text };
        continue;
      }
    }
    out.push(inline);
  }
  return out;
}

// ---------------------------------------------------------------------------
// Attraversamento

export interface ParagraphLocation {
  readonly paragraph: Paragraph;
  readonly sectionIndex: number;
  /** Indici dei blocchi dalla sezione al paragrafo: [blocco, riga, cella, blocco, ...]. */
  readonly path: readonly number[];
}

function* walkBlocks(
  blocks: readonly Block[],
  sectionIndex: number,
  prefix: number[],
): Generator<ParagraphLocation> {
  for (let i = 0; i < blocks.length; i++) {
    const b = blocks[i]!;
    if (b.type === 'paragraph') {
      yield { paragraph: b, sectionIndex, path: [...prefix, i] };
    } else {
      for (let r = 0; r < b.rows.length; r++) {
        const row = b.rows[r]!;
        for (let c = 0; c < row.cells.length; c++) {
          yield* walkBlocks(row.cells[c]!.blocks, sectionIndex, [...prefix, i, r, c]);
        }
      }
    }
  }
}

/** Tutti i paragrafi del corpo del documento in ordine di lettura (tabelle incluse). */
export function* paragraphs(doc: DocumentModel): Generator<ParagraphLocation> {
  for (let s = 0; s < doc.sections.length; s++) {
    yield* walkBlocks(doc.sections[s]!.blocks, s, []);
  }
}

export function findParagraph(doc: DocumentModel, id: Id): ParagraphLocation | undefined {
  for (const loc of paragraphs(doc)) if (loc.paragraph.id === id) return loc;
  return undefined;
}

export function firstParagraph(doc: DocumentModel): Paragraph {
  const first = paragraphs(doc).next();
  if (first.done) throw new Error('Documento senza paragrafi');
  return first.value.paragraph;
}

/** Sostituisce un blocco qualsiasi (anche annidato in una tabella) per id. */
export function replaceBlock(doc: DocumentModel, id: Id, replacement: readonly Block[]): DocumentModel {
  let found = false;
  const visit = (blocks: readonly Block[]): readonly Block[] => {
    const out: Block[] = [];
    let changed = false;
    for (const b of blocks) {
      if (!found && b.id === id) {
        found = true;
        changed = true;
        out.push(...replacement);
      } else if (!found && b.type === 'table') {
        const rows = b.rows.map((row) => {
          const cells = row.cells.map((cell) => {
            const inner = visit(cell.blocks);
            return inner === cell.blocks ? cell : { ...cell, blocks: inner };
          });
          return cells.some((c, i) => c !== row.cells[i]) ? { ...row, cells } : row;
        });
        if (rows.some((r, i) => r !== b.rows[i])) {
          changed = true;
          out.push({ ...b, rows });
        } else out.push(b);
      } else out.push(b);
    }
    return changed ? out : blocks;
  };
  const sections = doc.sections.map((s) => {
    if (found) return s;
    const blocks = visit(s.blocks);
    return blocks === s.blocks ? s : { ...s, blocks };
  });
  if (!found) throw new Error(`Blocco non trovato: ${id}`);
  return { ...doc, sections };
}

/** Comodità: sostituisce un paragrafo con la sua versione aggiornata. */
export function updateParagraph(doc: DocumentModel, p: Paragraph): DocumentModel {
  return replaceBlock(doc, p.id, [p]);
}

/**
 * Numero di ogni nota a piè di pagina (1, 2, 3...) nell'ordine in cui i
 * richiami compaiono nel corpo. Le note senza richiamo non hanno numero.
 */
export function footnoteNumbers(doc: DocumentModel): Map<Id, number> {
  const numbers = new Map<Id, number>();
  if (!doc.footnotes) return numbers;
  for (const { paragraph: p } of paragraphs(doc)) {
    for (const c of p.children) {
      if (c.type === 'footnoteRef' && doc.footnotes[c.noteId] && !numbers.has(c.noteId)) numbers.set(c.noteId, numbers.size + 1);
    }
  }
  return numbers;
}

// ---------------------------------------------------------------------------
// Posizioni

export function comparePositions(doc: DocumentModel, a: Position, b: Position): number {
  if (a.paragraphId === b.paragraphId) return a.offset - b.offset;
  for (const loc of paragraphs(doc)) {
    if (loc.paragraph.id === a.paragraphId) return -1;
    if (loc.paragraph.id === b.paragraphId) return 1;
  }
  throw new Error('Posizioni fuori dal documento');
}

export function caret(paragraphId: Id, offset: number): Selection {
  const p = { paragraphId, offset };
  return { anchor: p, focus: p };
}

export function isCollapsed(sel: Selection): boolean {
  return sel.anchor.paragraphId === sel.focus.paragraphId && sel.anchor.offset === sel.focus.offset;
}

/** Restituisce [inizio, fine] in ordine di documento. */
export function orderedRange(doc: DocumentModel, sel: Selection): [Position, Position] {
  return comparePositions(doc, sel.anchor, sel.focus) <= 0
    ? [sel.anchor, sel.focus]
    : [sel.focus, sel.anchor];
}

// ---------------------------------------------------------------------------
// Risoluzione degli stili

function styleChain(sheet: StyleSheet, styleId: Id | undefined): Id[] {
  const chain: Id[] = [];
  let id: Id | undefined = styleId ?? sheet.defaultParagraphStyle;
  while (id && !chain.includes(id)) {
    chain.unshift(id);
    id = sheet.styles[id]?.basedOn;
  }
  return chain;
}

/** Formattazione di paragrafo effettiva: default ← stili ← formattazione diretta. */
export function resolveParaFormat(sheet: StyleSheet, p: Paragraph): ParaFormat {
  let f: ParaFormat = { ...sheet.defaults.para };
  for (const id of styleChain(sheet, p.styleId)) f = { ...f, ...sheet.styles[id]?.paraFormat };
  return { ...f, ...p.format };
}

/** Formattazione di carattere effettiva di un run dentro un paragrafo. */
export function resolveCharFormat(sheet: StyleSheet, p: Paragraph, run?: TextRun | CharFormat): CharFormat {
  let f: CharFormat = { ...sheet.defaults.char };
  for (const id of styleChain(sheet, p.styleId)) f = { ...f, ...sheet.styles[id]?.charFormat };
  if (!run) return { ...f, ...p.markFormat };
  if ('type' in run) {
    if (run.styleId) for (const id of styleChain(sheet, run.styleId)) f = { ...f, ...sheet.styles[id]?.charFormat };
    return { ...f, ...run.format };
  }
  return { ...f, ...run };
}
