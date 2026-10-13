/**
 * Modulo io: formati di file, appunti e comandi file.*.
 *
 * Pubblica `editor.services.formats` (DOCX, HTML, testo) e
 * `editor.services.clipboard`; registra i comandi di cui è proprietario.
 */

import { COMMAND_OWNERS, type DocumentFormat, type Editor, type EditorModule } from '../core';
import { DefaultClipboardService } from './clipboard';
import { DOCUMENT_REPLACED_LABELS, fileExportCommand, fileNewCommand, fileOpenCommand, type SaveHandler } from './commands';
import { browserDownload, hasDom } from './download';
import { docxFormat } from './formats/docx';
import { htmlFormat } from './formats/html';
import { textFormat } from './formats/text';
import { DefaultFormatRegistry } from './registry';

export interface IoModuleOptions {
  /** Destinazione di `file.export`; per default il download del browser (se c'è un DOM). */
  readonly save?: SaveHandler;
  /** Formati aggiuntivi oltre a DOCX, HTML e testo. */
  readonly formats?: readonly DocumentFormat[];
}

export const BUILTIN_FORMATS: readonly DocumentFormat[] = [docxFormat, htmlFormat, textFormat];

export function createIoModule(options: IoModuleOptions = {}): EditorModule {
  return {
    id: 'io',
    activate(editor: Editor) {
      const formats = new DefaultFormatRegistry([...BUILTIN_FORMATS, ...(options.formats ?? [])]);
      const clipboard = new DefaultClipboardService();
      const save = options.save ?? (hasDom() ? browserDownload : undefined);
      editor.services.formats = formats;
      editor.services.clipboard = clipboard;

      const commands = [fileNewCommand(), fileOpenCommand(formats), fileExportCommand(formats, save)];
      for (const c of commands) {
        if (COMMAND_OWNERS[c.id as keyof typeof COMMAND_OWNERS] !== 'io') throw new Error(`Il modulo io non possiede ${c.id}`);
      }
      const disposers = commands.map((c) => editor.commands.register(c));

      // Un documento nuovo o aperto non si annulla: si azzera la cronologia.
      disposers.push(
        editor.events.on('change', ({ meta }) => {
          if (meta.label && DOCUMENT_REPLACED_LABELS.includes(meta.label)) editor.history.clear();
        }),
      );

      return () => {
        for (const d of disposers) d();
        if (editor.services.formats === formats) delete editor.services.formats;
        if (editor.services.clipboard === clipboard) delete editor.services.clipboard;
      };
    },
  };
}

export const ioModule: EditorModule = createIoModule();
