/**
 * Modulo PDF: aggiunge il formato 'pdf' al registro dei formati del modulo
 * io, così `file.export` con `{ formatId: 'pdf' }` salva il documento come
 * PDF impaginato come a schermo e `file.open` apre un PDF ricostruendo
 * paragrafi, formattazione e immagini. Non registra comandi propri.
 */

import type { EditorModule } from '../core';
import { createPdfFormat, type PdfFormatOptions } from './format';
import { browserRasterizer } from './images';

export const PDF_MODULE_ID = 'pdf';

export function createPdfModule(options: Omit<PdfFormatOptions, 'layout'> = {}): EditorModule {
  return {
    id: PDF_MODULE_ID,
    activate(editor) {
      const format = createPdfFormat({ rasterize: browserRasterizer, ...options, layout: () => editor.services.layout });
      let registered: NonNullable<typeof editor.services.formats> | undefined;
      let disposed = false;
      const register = () => {
        const formats = editor.services.formats;
        if (disposed || registered || !formats || formats.get(format.id)) return;
        formats.register(format);
        registered = formats;
      };
      register();
      // Il modulo io può essere attivato dopo questo: si riprova a fine attivazione.
      if (!registered) queueMicrotask(register);
      return () => {
        disposed = true;
        const r = registered as { unregister?: (id: string) => void } | undefined;
        if (r?.unregister && registered?.get(format.id) === format) r.unregister(format.id);
      };
    },
  };
}
