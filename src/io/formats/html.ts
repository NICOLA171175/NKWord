import type { DocumentFormat } from '../../core';
import { dataToText } from '../util/encoding';
import { documentToHtml } from './html-export';
import { htmlToDocument } from './html-import';

export { documentToHtml, type HtmlExportOptions } from './html-export';
export { htmlToDocument } from './html-import';

export const htmlFormat: DocumentFormat = {
  id: 'html',
  label: 'Pagina Web',
  extensions: ['html', 'htm'],
  mimeType: 'text/html',
  async import(data) {
    return htmlToDocument(dataToText(data));
  },
  async export(doc) {
    return documentToHtml(doc);
  },
};
