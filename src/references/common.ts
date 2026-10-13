/**
 * Funzioni condivise da didascalie, riferimenti incrociati e indice:
 * numeri di pagina, stili cercati anche per nome, sostituzione di un pezzo
 * di testo che mantiene la formattazione.
 */

import {
  paragraphText,
  paragraphs,
  textRun,
  type Block,
  type CharFormat,
  type DocumentModel,
  type Id,
  type LayoutEngine,
  type Paragraph,
  type Style,
  type StyleSheet,
} from '../core';
import { deleteInParagraph, insertInParagraph } from '../editing/inlines';
import { estimatePages, pagesFromLayout } from '../toc/toc';

/** Numero di pagina di ogni paragrafo: dall'impaginazione se c'è, altrimenti stimato. */
export function pageMap(doc: DocumentModel, layout?: LayoutEngine): Map<Id, number> {
  return layout ? pagesFromLayout(layout, doc) : estimatePages(doc);
}

/** Ordine di lettura dei paragrafi del corpo (0, 1, 2...). */
export function readingOrder(doc: DocumentModel): Map<Id, number> {
  const order = new Map<Id, number>();
  for (const { paragraph: p } of paragraphs(doc)) order.set(p.id, order.size);
  return order;
}

/** Testo del paragrafo su una riga: tab e interruzioni diventano spazi, gli oggetti spariscono. */
export function oneLine(p: Paragraph): string {
  return paragraphText(p)
    .replace(/[\t\n]/g, ' ')
    .replace(/￼/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Id dello stile con quel ruolo: quello predefinito se c'è, altrimenti uno
 * con il nome di Word (es. importato da DOCX con un id localizzato).
 */
export function findStyleId(sheet: StyleSheet, id: Id, name: RegExp): Id | undefined {
  if (sheet.styles[id]) return id;
  return Object.values(sheet.styles).find((s) => s.type === 'paragraph' && name.test(s.name))?.id;
}

/** Aggiunge gli stili che mancano (per id o per nome); non tocca quelli esistenti. */
export function ensureStyles(sheet: StyleSheet, styles: readonly Style[]): StyleSheet {
  const missing = styles.filter((s) => !findStyleId(sheet, s.id, new RegExp(`^${s.name}$`, 'i')));
  if (!missing.length) return sheet;
  const all = { ...sheet.styles };
  for (const s of missing) all[s.id] = s;
  return { ...sheet, styles: all };
}

/** Sostituisce [from, to) del paragrafo con `text`, nel formato del primo carattere sostituito. */
export function replaceText(p: Paragraph, from: number, to: number, text: string, format: CharFormat, styleId?: Id): Paragraph {
  const run = styleId ? { ...textRun(text, format), styleId } : textRun(text, format);
  return insertInParagraph(deleteInParagraph(p, from, to), from, text ? [run] : []);
}

/** Sostituisce i blocchi [start, end) di una sezione (la sezione resta con almeno un paragrafo). */
export function replaceBlocks(doc: DocumentModel, sectionIndex: number, start: number, end: number, blocks: readonly Block[], fallback: () => Block): DocumentModel {
  const sections = doc.sections.map((s, i) => {
    if (i !== sectionIndex) return s;
    const next = [...s.blocks.slice(0, start), ...blocks, ...s.blocks.slice(end)];
    return { ...s, blocks: next.length ? next : [fallback()] };
  });
  return { ...doc, sections };
}
