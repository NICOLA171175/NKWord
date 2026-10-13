/** Testo semplice (.txt): un paragrafo per riga, tab come `TabInline`. */

import {
  createDocument,
  normalizeInlines,
  paragraph,
  textRun,
  type Block,
  type DocumentFormat,
  type DocumentModel,
  type Inline,
  type Paragraph,
} from '../../core';
import { dataToText } from '../util/encoding';

export function textToBlocks(text: string): Paragraph[] {
  const lines = text.replace(/\r\n?/g, '\n').split('\n');
  return lines.map((line) => {
    const inlines: Inline[] = [];
    line.split('\t').forEach((part, i) => {
      if (i > 0) inlines.push({ type: 'tab' });
      // I caratteri di controllo (salvo il tab) non sono ammessi nei run.
      const clean = part.replace(/[\u0000-\u0008\u000b-\u001f\u007f]/g, '');
      if (clean) inlines.push(textRun(clean));
    });
    return paragraph(normalizeInlines(inlines));
  });
}

export function textToDocument(text: string): DocumentModel {
  return createDocument(textToBlocks(text));
}

function paragraphPlainText(p: Paragraph): string {
  return p.children
    .map((c) => {
      switch (c.type) {
        case 'text':
          return c.text;
        case 'tab':
          return '\t';
        case 'break':
          return c.kind === 'line' ? '\n' : '\f';
        case 'image':
          return c.alt ? `[${c.alt}]` : '';
        case 'field':
          return '';
      }
    })
    .join('');
}

function blockLines(block: Block, out: string[]): void {
  if (block.type === 'paragraph') {
    out.push(paragraphPlainText(block));
    return;
  }
  for (const row of block.rows) {
    const cells = row.cells
      .filter((c) => c.vMerge !== 'continue')
      .map((c) => {
        const inner: string[] = [];
        for (const b of c.blocks) blockLines(b, inner);
        return inner.join(' ').replace(/[\t\n]/g, ' ');
      });
    out.push(cells.join('\t'));
  }
}

/** Testo semplice di un documento: un paragrafo per riga, celle separate da tab. */
export function documentToText(doc: DocumentModel, newline = '\n'): string {
  const lines: string[] = [];
  for (const s of doc.sections) for (const b of s.blocks) blockLines(b, lines);
  return lines.join(newline).replace(/\f/g, newline);
}

export const textFormat: DocumentFormat = {
  id: 'txt',
  label: 'Testo semplice',
  extensions: ['txt', 'text'],
  mimeType: 'text/plain',
  async import(data) {
    return textToDocument(dataToText(data));
  },
  async export(doc) {
    return documentToText(doc, '\r\n');
  },
};
