/**
 * Comandi file.* (proprietario: modulo io, vedi COMMAND_OWNERS).
 *
 * - `file.new` è puro: stato → transazione con un documento vuoto.
 * - `file.open` è asincrono perché la conversione dei formati lo è, ma non
 *   ha effetti esterni: riceve i byte già letti dalla UI.
 * - `file.export` converte il documento e lo consegna al `SaveHandler`
 *   (per default lo scaricamento nel browser); non cambia lo stato.
 *
 * Aprire o creare un documento non entra nella cronologia e la azzera
 * (vedi `ioModule`), come in Word.
 */

import {
  createState,
  type AsyncCommand,
  type Command,
  type CommandArgs,
  type DocumentModel,
  type FileData,
  type FormatRegistry,
} from '../core';

/** Etichette delle transazioni che sostituiscono il documento. */
export const DOCUMENT_REPLACED_LABELS: readonly string[] = ['file.new', 'file.open'];

export interface SaveRequest {
  readonly data: FileData;
  /** Nome proposto, es. 'Relazione.docx'. */
  readonly fileName: string;
  readonly mimeType: string;
  readonly formatId: string;
}

/** Riceve il file esportato: download nel browser, scrittura su disco in Electron, test... */
export type SaveHandler = (file: SaveRequest) => void | Promise<void>;

export function suggestedFileName(doc: DocumentModel, extension: string): string {
  const base = (doc.metadata.title ?? '').replace(/[\\/:*?"<>|\u0000-\u001f]+/g, ' ').trim() || 'Documento';
  return `${base}.${extension}`;
}

export function fileNewCommand(): Command<CommandArgs['file.new']> {
  return {
    id: 'file.new',
    label: 'Nuovo',
    run: () => ({ state: createState(), meta: { addToHistory: false, label: 'file.new' } }),
  };
}

export function fileOpenCommand(formats: FormatRegistry): AsyncCommand<CommandArgs['file.open']> {
  return {
    id: 'file.open',
    label: 'Apri',
    async: true,
    isEnabled: (_ctx, args) => !args || !!formats.get(args.formatId)?.import,
    async run(_ctx, { data, formatId }) {
      const format = formats.get(formatId);
      if (!format?.import) throw new Error(`Formato non apribile: ${formatId}`);
      const doc = await format.import(data);
      return { state: createState(doc), meta: { addToHistory: false, label: 'file.open' } };
    },
  };
}

export function fileExportCommand(formats: FormatRegistry, save: SaveHandler | undefined): AsyncCommand<CommandArgs['file.export']> {
  return {
    id: 'file.export',
    label: 'Salva con nome',
    async: true,
    isEnabled: (_ctx, args) => !!save && (!args || !!formats.get(args.formatId)?.export),
    async run({ state }, { formatId }) {
      const format = formats.get(formatId);
      if (!format?.export) throw new Error(`Formato non esportabile: ${formatId}`);
      if (!save) throw new Error('Nessuna destinazione per il salvataggio');
      const data = await format.export(state.doc);
      await save({
        data,
        fileName: suggestedFileName(state.doc, format.extensions[0] ?? format.id),
        mimeType: format.mimeType,
        formatId,
      });
      return null;
    },
  };
}
