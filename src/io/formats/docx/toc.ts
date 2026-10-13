/**
 * Sommario e collegamenti interni in DOCX.
 *
 * - Un collegamento `#_Toc<id>` punta al paragrafo con quell'id: il
 *   paragrafo riceve il segnalibro `_Toc<id>` (convenzione di Word per i
 *   titoli del sommario).
 * - I paragrafi consecutivi con stile "toc 1".."toc 9" sono avvolti in un
 *   campo TOC, così Word li riconosce come sommario e li sa aggiornare.
 */

import { paragraphs, updateParagraph, type Block, type DocumentModel, type Id, type Paragraph, type StyleSheet } from '../../../core';
import { escapeAttr } from '../../util/markup';

const BOOKMARK_PREFIX = '_Toc';

/** Nomi dei segnalibri richiesti dai collegamenti interni del documento. */
export function internalAnchors(doc: DocumentModel): Set<string> {
  const out = new Set<string>();
  for (const { paragraph: p } of paragraphs(doc)) {
    for (const c of p.children) if (c.type === 'text' && c.format.link?.startsWith(`#${BOOKMARK_PREFIX}`)) out.add(c.format.link.slice(1));
  }
  return out;
}

/**
 * In lettura: i collegamenti `#nome` a un segnalibro diventano `#_Toc<id>`
 * del paragrafo che lo contiene, così si possono seguire e risalvare anche
 * se i paragrafi hanno id nuovi.
 */
export function relinkBookmarks(doc: DocumentModel, bookmarks: ReadonlyMap<string, Id>): DocumentModel {
  if (!bookmarks.size) return doc;
  let out = doc;
  for (const { paragraph: p } of paragraphs(doc)) {
    let changed = false;
    const children = p.children.map((c) => {
      const target = c.type === 'text' && c.format.link?.startsWith('#') ? bookmarks.get(c.format.link.slice(1)) : undefined;
      if (!target || c.type !== 'text') return c;
      changed = true;
      return { ...c, format: { ...c.format, link: `#${BOOKMARK_PREFIX}${target}` } };
    });
    if (changed) out = updateParagraph(out, { ...p, children });
  }
  return out;
}

export function paragraphBookmark(p: Paragraph, anchors: ReadonlySet<string>): string | undefined {
  const name = `${BOOKMARK_PREFIX}${p.id}`;
  return p.id && anchors.has(name) ? name : undefined;
}

function tocLevel(sheet: StyleSheet, styleId: string | undefined): number {
  if (!styleId) return 0;
  const m = /^TOC([1-9])$/.exec(styleId) ?? /^toc ([1-9])$/i.exec(sheet.styles[styleId]?.name ?? '');
  return m ? Number(m[1]) : 0;
}

export interface FieldWrap {
  before?: string;
  after?: string;
}

/** Inizio e fine del campo TOC per ogni gruppo di voci di sommario consecutive. */
export function tocFieldWraps(blocks: readonly Block[], sheet: StyleSheet): Map<Paragraph, FieldWrap> {
  const out = new Map<Paragraph, FieldWrap>();
  let i = 0;
  while (i < blocks.length) {
    const first = blocks[i]!;
    if (first.type !== 'paragraph' || !tocLevel(sheet, first.styleId)) {
      i++;
      continue;
    }
    let levels = 3;
    let last = first;
    while (i < blocks.length) {
      const b = blocks[i]!;
      const level = b.type === 'paragraph' ? tocLevel(sheet, b.styleId) : 0;
      if (!level) break;
      levels = Math.max(levels, level);
      last = b as Paragraph;
      i++;
    }
    const instr = ` TOC \\o "1-${levels}" \\h \\z \\u `;
    const begin =
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      `<w:r><w:instrText xml:space="preserve">${escapeAttr(instr)}</w:instrText></w:r>` +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>';
    const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
    out.set(first, { ...out.get(first), before: begin });
    out.set(last, { ...out.get(last), after: end });
  }
  return out;
}
