import type { DocumentFormat } from '../../../core';
import { dataToBytes } from '../../util/encoding';
import { DOCX_MIME } from './common';
import { exportDocx } from './export';
import { importDocx } from './import';

export { exportDocx } from './export';
export { importDocx } from './import';

export const docxFormat: DocumentFormat = {
  id: 'docx',
  label: 'Documento di Word',
  extensions: ['docx'],
  mimeType: DOCX_MIME,
  import(data) {
    return importDocx(dataToBytes(data));
  },
  export(doc) {
    return exportDocx(doc);
  },
};
