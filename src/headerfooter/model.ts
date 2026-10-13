/**
 * Intestazione e piè di pagina come testo semplice: righe, allineamento e
 * numero di pagina. È ciò che mostra e modifica la finestra "Intestazione"
 * (vedi `dialog.ts`); i contenuti più ricchi letti da un DOCX si conservano
 * finché non si modificano.
 */

import {
  findParagraph,
  newId,
  textRun,
  type Block,
  type DocumentModel,
  type HeaderFooter,
  type Inline,
  type Paragraph,
  type ParaFormat,
  type Section,
} from '../core';

export type HeaderFooterPart = 'header' | 'footer';
export type HeaderFooterAlign = 'left' | 'center' | 'right';

export interface HeaderFooterContent {
  /** Righe separate da '\n'. */
  readonly text: string;
  readonly align: HeaderFooterAlign;
  /** Numero di pagina alla fine dell'ultima riga. */
  readonly pageNumber: boolean;
}

export interface HeaderFooterInfo extends HeaderFooterContent {
  /** Contiene altro oltre a testo e numero di pagina (immagini, tabelle, formattazione): modificarla lo semplifica. */
  readonly complex: boolean;
}

export const EMPTY_CONTENT: HeaderFooterContent = { text: '', align: 'left', pageNumber: false };

/** Sezione del cursore (0 se il cursore non è in un paragrafo del corpo). */
export function sectionAt(doc: DocumentModel, paragraphId: string): number {
  return findParagraph(doc, paragraphId)?.sectionIndex ?? 0;
}

/** Contenuto della parte (pagine normali) di una sezione, come testo semplice. */
export function readHeaderFooter(section: Section, part: HeaderFooterPart): HeaderFooterInfo {
  const blocks = section[part]?.default ?? [];
  const lines: string[] = [];
  let pageNumber = false;
  let complex = false;
  let align: HeaderFooterAlign | undefined;
  for (const b of blocks) {
    if (b.type !== 'paragraph') {
      complex = true;
      continue;
    }
    const a = b.format.align;
    align ??= a === 'center' || a === 'right' ? a : 'left';
    let line = '';
    for (const c of b.children) {
      if (c.type === 'text') {
        line += c.text;
        if (Object.keys(c.format).length) complex = true;
      } else if (c.type === 'tab') line += ' ';
      else if (c.type === 'field' && c.kind === 'page') pageNumber = true;
      else complex = true;
    }
    lines.push(line);
  }
  while (lines.length && lines[lines.length - 1]!.trim() === '') lines.pop();
  return { text: lines.join('\n').trimEnd(), align: align ?? 'left', pageNumber, complex };
}

/** Paragrafi della parte; nessuno se è vuota. */
export function headerFooterBlocks(content: HeaderFooterContent): Paragraph[] {
  const text = content.text.replace(/\r\n|\r/g, '\n').replace(/\t/g, ' ').trimEnd();
  if (!text && !content.pageNumber) return [];
  const lines = text ? text.split('\n') : [''];
  const format: ParaFormat = content.align === 'left' ? {} : { align: content.align };
  return lines.map((line, i) => {
    const children: Inline[] = line ? [textRun(line)] : [];
    if (content.pageNumber && i === lines.length - 1) {
      if (line && !line.endsWith(' ')) children.push(textRun(' '));
      children.push({ type: 'field', kind: 'page', format: {} });
    }
    const p: Paragraph = { type: 'paragraph', id: newId(), format, children: children.length ? normalize(children) : [] };
    return p;
  });
}

function normalize(children: Inline[]): Inline[] {
  // Due testi consecutivi (riga + spazio) diventano uno solo.
  const out: Inline[] = [];
  for (const c of children) {
    const last = out[out.length - 1];
    if (c.type === 'text' && last?.type === 'text') out[out.length - 1] = textRun(last.text + c.text);
    else out.push(c);
  }
  return out;
}

/**
 * Scrive la parte in tutte le sezioni (in Word le sezioni sono "collegate
 * alla precedente" finché non si scollegano). Le varianti prima pagina e
 * pagine pari restano com'erano.
 */
export function writeHeaderFooter(doc: DocumentModel, part: HeaderFooterPart, content: HeaderFooterContent): DocumentModel {
  const blocks: readonly Block[] = headerFooterBlocks(content);
  const sections = doc.sections.map((s) => {
    const { default: _old, ...others } = s[part] ?? {};
    const next: HeaderFooter = blocks.length ? { ...others, default: blocks } : others;
    const { [part]: _drop, ...rest } = s;
    return Object.keys(next).length ? { ...rest, [part]: next } : rest;
  });
  return { ...doc, sections: sections as Section[] };
}

/** Aggiunge un paragrafo centrato con il numero di pagina, senza toccare il resto (contenuti ricchi). */
export function appendPageNumber(doc: DocumentModel, part: HeaderFooterPart): DocumentModel {
  const sections = doc.sections.map((s) => {
    const [p] = headerFooterBlocks({ text: '', align: 'center', pageNumber: true });
    const hf = s[part] ?? {};
    return { ...s, [part]: { ...hf, default: [...(hf.default ?? []), p!] } };
  });
  return { ...doc, sections };
}

export function sameContent(a: HeaderFooterContent, b: HeaderFooterContent): boolean {
  return a.text === b.text && a.align === b.align && a.pageNumber === b.pageNumber;
}
