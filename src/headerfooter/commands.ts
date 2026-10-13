/**
 * Comandi di intestazione e piè di pagina. Non fanno parte del catalogo di
 * `src/core`: la barra multifunzione li trova tramite `ribbonItems`, e
 * `page.headerFooter` (Inserisci > Intestazione) apre la finestra di modifica.
 */

import type { AsyncCommand, Command, Editor, EditorState } from '../core';
import {
  appendPageNumber,
  readHeaderFooter,
  sameContent,
  sectionAt,
  writeHeaderFooter,
  type HeaderFooterContent,
  type HeaderFooterInfo,
  type HeaderFooterPart,
} from './model';

export interface HeaderFooterCommandArgs {
  /** Apre la finestra per modificare la parte. */
  'headerFooter.edit': { part: HeaderFooterPart };
  /** Sostituisce il contenuto della parte in tutto il documento. */
  'headerFooter.set': { part: HeaderFooterPart } & HeaderFooterContent;
  /** Inserisci > Numero di pagina: in alto, in basso o al cursore. */
  'headerFooter.pageNumber': { position: 'header' | 'footer' | 'here' };
}

export type HeaderFooterCommandId = keyof HeaderFooterCommandArgs;

/** Finestra di modifica; `null` se l'utente annulla. */
export type HeaderFooterDialog = (data: { part: HeaderFooterPart; current: HeaderFooterInfo }) => Promise<HeaderFooterContent | null>;

const LABELS: Record<HeaderFooterPart, string> = { header: 'Intestazione', footer: 'Piè di pagina' };

export function currentHeaderFooter(state: EditorState, part: HeaderFooterPart): HeaderFooterInfo {
  const index = sectionAt(state.doc, state.selection.focus.paragraphId);
  return readHeaderFooter(state.doc.sections[index]!, part);
}

export function createHeaderFooterCommands(editor: Editor, dialog?: HeaderFooterDialog): (Command<any> | AsyncCommand<any>)[] {
  const isPart = (part: unknown): part is HeaderFooterPart => part === 'header' || part === 'footer';

  const set: Command<HeaderFooterCommandArgs['headerFooter.set']> = {
    id: 'headerFooter.set',
    label: 'Modifica intestazione o piè di pagina',
    isEnabled: (_ctx, args) => !args || isPart(args.part),
    run: ({ state }, { part, text, align, pageNumber }) => {
      const content = { text, align, pageNumber };
      const current = currentHeaderFooter(state, part);
      if (!current.complex && sameContent(current, content)) return null;
      const doc = writeHeaderFooter(state.doc, part, content);
      return { state: { ...state, doc }, meta: { label: LABELS[part] } };
    },
  };

  const edit: AsyncCommand<HeaderFooterCommandArgs['headerFooter.edit']> = {
    id: 'headerFooter.edit',
    label: 'Intestazione e piè di pagina',
    async: true,
    isEnabled: (_ctx, args) => !!dialog && (!args || isPart(args.part)),
    run: async ({ state }, { part }) => {
      const result = await dialog!({ part, current: currentHeaderFooter(state, part) });
      if (result) editor.execute('headerFooter.set', { part, ...result });
      return null;
    },
  };

  const pageNumber: Command<HeaderFooterCommandArgs['headerFooter.pageNumber']> = {
    id: 'headerFooter.pageNumber',
    label: 'Numero di pagina',
    isEnabled: (_ctx, args) => !args || args.position !== 'here' || editor.canExecute('insert.field', { kind: 'page' }),
    run: ({ state }, { position }) => {
      if (position === 'here') {
        editor.execute('insert.field', { kind: 'page' });
        return null;
      }
      // Come "Numero normale 2" di Word: al centro, se la parte è ancora vuota.
      const current = currentHeaderFooter(state, position);
      if (current.pageNumber) return null;
      const doc = current.complex
        ? appendPageNumber(state.doc, position)
        : writeHeaderFooter(state.doc, position, { text: current.text, align: current.text ? current.align : 'center', pageNumber: true });
      return { state: { ...state, doc }, meta: { label: 'Numero di pagina' } };
    },
  };

  return [set, edit, pageNumber];
}
