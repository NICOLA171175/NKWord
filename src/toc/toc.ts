/**
 * Sommario automatico: raccolta dei titoli, costruzione delle voci e
 * ricerca del sommario nel documento.
 *
 * Nel modello il sommario è una sequenza di paragrafi di primo livello di
 * una sezione: un titolo facoltativo (stile "TOC Heading") seguito dalle
 * voci (stili "toc 1".."toc 9"). Ogni voce contiene il testo del titolo,
 * una tabulazione a destra con puntini e il numero di pagina; testo e numero
 * sono collegamenti interni `#_Toc<id del paragrafo>` al titolo.
 */

import {
  paragraph,
  paragraphText,
  paragraphs,
  resolveParaFormat,
  textRun,
  type Block,
  type DocumentModel,
  type Id,
  type Inline,
  type LayoutEngine,
  type Paragraph,
  type Section,
  type StyleSheet,
} from '../core';
import {
  TOC_HEADING_STYLE,
  ensureTocStyles,
  isTocHeadingStyle,
  isTocStyle,
  tocEntryLevel,
  tocEntryStyleId,
} from './styles';

export const DEFAULT_TOC_LEVELS = 3;
export const TOC_TITLE = 'Sommario';
export const NO_ENTRIES_TEXT = 'Nessuna voce di sommario trovata.';
/** Prefisso dei collegamenti alle voci; è anche il nome del segnalibro in DOCX. */
export const TOC_ANCHOR_PREFIX = '_Toc';

export function tocAnchor(paragraphId: Id): string {
  return `#${TOC_ANCHOR_PREFIX}${paragraphId}`;
}

/** Id del paragrafo a cui punta un collegamento di sommario, se lo è. */
export function anchorTarget(link: string | undefined): Id | undefined {
  return link?.startsWith(`#${TOC_ANCHOR_PREFIX}`) ? link.slice(TOC_ANCHOR_PREFIX.length + 1) : undefined;
}

// ---------------------------------------------------------------------------
// Titoli

export interface TocHeading {
  readonly paragraphId: Id;
  readonly level: number;
  readonly text: string;
}

/** Testo di un titolo su una riga: tab e interruzioni diventano spazi, gli oggetti spariscono. */
function headingText(p: Paragraph): string {
  return paragraphText(p)
    .replace(/[\t\n]/g, ' ')
    .replace(/￼/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

/** Titoli del documento (livello 1..`levels`) in ordine di lettura, tabelle incluse. */
export function collectHeadings(doc: DocumentModel, levels = DEFAULT_TOC_LEVELS): TocHeading[] {
  const out: TocHeading[] = [];
  for (const { paragraph: p } of paragraphs(doc)) {
    if (isTocStyle(doc.styles, p.styleId)) continue;
    const level = resolveParaFormat(doc.styles, p).outlineLevel;
    if (!level || level < 1 || level > levels) continue;
    const text = headingText(p);
    if (text) out.push({ paragraphId: p.id, level, text });
  }
  return out;
}

// ---------------------------------------------------------------------------
// Posizione del sommario

export interface TocLocation {
  readonly sectionIndex: number;
  /** Indici dei blocchi della sezione: [start, end). */
  readonly start: number;
  readonly end: number;
  /** Livelli mostrati, dedotti dalle voci presenti (minimo `DEFAULT_TOC_LEVELS`). */
  readonly levels: number;
}

function isTocBlock(sheet: StyleSheet, b: Block | undefined): b is Paragraph {
  return b?.type === 'paragraph' && isTocStyle(sheet, b.styleId);
}

/** Primo sommario del documento (blocchi consecutivi con stili del sommario). */
export function findToc(doc: DocumentModel): TocLocation | undefined {
  for (let s = 0; s < doc.sections.length; s++) {
    const blocks = doc.sections[s]!.blocks;
    const start = blocks.findIndex((b) => isTocBlock(doc.styles, b));
    if (start < 0) continue;
    let end = start;
    let levels = DEFAULT_TOC_LEVELS;
    while (isTocBlock(doc.styles, blocks[end])) {
      levels = Math.max(levels, tocEntryLevel(doc.styles, (blocks[end] as Paragraph).styleId));
      end++;
    }
    return { sectionIndex: s, start, end, levels };
  }
  return undefined;
}

// ---------------------------------------------------------------------------
// Numeri di pagina

export type PageMap = ReadonlyMap<Id, number>;

/** Numeri di pagina dei paragrafi secondo l'impaginazione. */
export function pagesFromLayout(engine: LayoutEngine, doc: DocumentModel): Map<Id, number> {
  const pages = new Map<Id, number>();
  for (const page of engine.layout(doc).pages) {
    for (const line of page.lines) if (!pages.has(line.paragraphId)) pages.set(line.paragraphId, page.number);
  }
  return pages;
}

/**
 * Stima senza impaginazione: conta interruzioni di pagina esplicite,
 * "pagina successiva" nei paragrafi e sezioni che iniziano su una nuova pagina.
 */
export function estimatePages(doc: DocumentModel): Map<Id, number> {
  const pages = new Map<Id, number>();
  let page = 1;
  let started = false;
  const visit = (blocks: readonly Block[]) => {
    for (const b of blocks) {
      if (b.type === 'table') {
        for (const row of b.rows) for (const cell of row.cells) visit(cell.blocks);
        continue;
      }
      if (started && resolveParaFormat(doc.styles, b).pageBreakBefore) page++;
      started = true;
      pages.set(b.id, page);
      for (const c of b.children) if (c.type === 'break' && c.kind === 'page') page++;
    }
  };
  doc.sections.forEach((s, i) => {
    if (i === 0) page = s.pageNumberStart ?? 1;
    else if (s.breakType !== 'continuous') page = s.pageNumberStart ?? page + 1;
    else if (s.pageNumberStart !== undefined) page = s.pageNumberStart;
    visit(s.blocks);
  });
  return pages;
}

// ---------------------------------------------------------------------------
// Costruzione

function contentWidth(section: Section): number {
  const { width, margins } = section.page;
  return width - margins.left - margins.right;
}

/** Id di stile da usare: preferisce uno stile esistente con il nome di Word (es. importato). */
function styleIdByName(sheet: StyleSheet, fallback: Id, matches: (id: Id) => boolean): Id {
  if (sheet.styles[fallback]) return fallback;
  return Object.keys(sheet.styles).find(matches) ?? fallback;
}

export interface TocBuildOptions {
  readonly levels: number;
  readonly pages: PageMap;
  /** Larghezza utile: posizione della tabulazione destra dei numeri di pagina. */
  readonly width: number;
  readonly title?: boolean;
}

/** Paragrafi del sommario (titolo + voci) per i titoli dati. */
export function buildTocBlocks(sheet: StyleSheet, headings: readonly TocHeading[], options: TocBuildOptions): Paragraph[] {
  const blocks: Paragraph[] = [];
  if (options.title !== false) {
    const id = styleIdByName(sheet, TOC_HEADING_STYLE, (s) => isTocHeadingStyle(sheet, s));
    blocks.push(paragraph(TOC_TITLE, {}, id));
  }
  const tabs = [{ position: options.width, align: 'right' as const, leader: 'dot' as const }];
  const entryStyle = (level: number) => styleIdByName(sheet, tocEntryStyleId(level), (s) => tocEntryLevel(sheet, s) === level);
  if (!headings.length) {
    blocks.push(paragraph(NO_ENTRIES_TEXT, {}, entryStyle(1)));
    return blocks;
  }
  for (const h of headings) {
    const link = tocAnchor(h.paragraphId);
    const page = options.pages.get(h.paragraphId);
    const children: Inline[] = [textRun(h.text, { link }), { type: 'tab' }];
    if (page !== undefined) children.push(textRun(String(page), { link }));
    blocks.push(paragraph(children, { tabs }, entryStyle(h.level)));
  }
  return blocks;
}

function replaceBlocks(doc: DocumentModel, sectionIndex: number, start: number, end: number, blocks: readonly Block[]): DocumentModel {
  const sections = doc.sections.map((s, i) => {
    if (i !== sectionIndex) return s;
    const next = [...s.blocks.slice(0, start), ...blocks, ...s.blocks.slice(end)];
    return { ...s, blocks: next.length ? next : [paragraph()] };
  });
  return { ...doc, sections };
}

export interface TocWriteOptions {
  readonly levels?: number;
  /** Impaginazione per i numeri di pagina; senza, i numeri sono stimati. */
  readonly layout?: LayoutEngine;
}

/** Numero massimo di impaginazioni: il sommario stesso può spostare i titoli. */
const MAX_PASSES = 3;

/**
 * Scrive il sommario nella posizione data (sostituendo i blocchi
 * [start, end) della sezione) con titoli e numeri di pagina aggiornati.
 * Mantiene il titolo se c'era, o lo crea se il sommario è nuovo.
 */
export function writeToc(
  doc: DocumentModel,
  at: { sectionIndex: number; start: number; end: number },
  options: TocWriteOptions = {},
): DocumentModel {
  const old = doc.sections[at.sectionIndex]!.blocks.slice(at.start, at.end);
  const levels = Math.min(9, Math.max(1, options.levels ?? DEFAULT_TOC_LEVELS));
  const title = at.end === at.start || old.some((b) => b.type === 'paragraph' && isTocHeadingStyle(doc.styles, b.styleId));
  const base: DocumentModel = { ...doc, styles: ensureTocStyles(doc.styles, levels) };
  const width = contentWidth(base.sections[at.sectionIndex]!);
  const headings = collectHeadings(base, levels);
  const pagesOf = (d: DocumentModel) => (options.layout ? pagesFromLayout(options.layout, d) : estimatePages(d));

  // Si impagina con il sommario già al suo posto e si ricalcola finché i
  // numeri non cambiano: le voci possono spingere i titoli alla pagina dopo.
  // Le voci riusano gli id dei paragrafi che sostituiscono: il cursore
  // dentro il sommario resta valido dopo un aggiornamento.
  const write = (pages: PageMap) => {
    const blocks = buildTocBlocks(base.styles, headings, { levels, pages, width, title }).map((b, i) =>
      old[i] ? { ...b, id: old[i]!.id } : b,
    );
    return replaceBlocks(base, at.sectionIndex, at.start, at.end, blocks);
  };
  let pages: PageMap = new Map();
  for (let pass = 0; pass < MAX_PASSES; pass++) {
    const result = write(pages);
    const next = pagesOf(result);
    if (headings.every((h) => next.get(h.paragraphId) === pages.get(h.paragraphId))) return result;
    pages = next;
  }
  return write(pages);
}

/** Toglie il sommario dal documento. */
export function removeToc(doc: DocumentModel, at: TocLocation): DocumentModel {
  return replaceBlocks(doc, at.sectionIndex, at.start, at.end, []);
}
