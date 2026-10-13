/**
 * Modulo note a piè di pagina: registra i comandi `footnote.*` assegnati a
 * 'footnotes' in `COMMAND_OWNERS`. Il testo in fondo alla pagina lo disegna
 * il motore di impaginazione (vedi `PageLayout.footnotes`).
 */

import type { Command, CommandArgs, CommandId, EditorModule } from '../core';
import { canInsertFootnote, deleteFootnote, editFootnote, footnoteAtSelection, insertFootnote } from './model';

function defineCommand<K extends CommandId>(id: K, def: Omit<Command<CommandArgs[K]>, 'id'>): Command<CommandArgs[K]> {
  return { id, ...def };
}

export const footnoteCommands = [
  defineCommand('footnote.insert', {
    label: 'Inserisci nota a piè di pagina',
    isEnabled: ({ state }) => canInsertFootnote(state),
    run: ({ state }, args) => {
      const result = insertFootnote(state, args?.text ?? '');
      return result && { state: result.state, meta: { label: 'Inserisci nota a piè di pagina' } };
    },
  }),
  defineCommand('footnote.edit', {
    label: 'Modifica nota a piè di pagina',
    isEnabled: ({ state }, args) => !!(args?.noteId ?? footnoteAtSelection(state)),
    run: ({ state }, { text, noteId }) => {
      const id = noteId ?? footnoteAtSelection(state);
      const next = id ? editFootnote(state, id, text) : null;
      return next && { state: next, meta: { label: 'Modifica nota a piè di pagina' } };
    },
  }),
  defineCommand('footnote.delete', {
    label: 'Elimina nota a piè di pagina',
    isEnabled: ({ state }, args) => !!(args?.noteId ?? footnoteAtSelection(state)),
    run: ({ state }, args) => {
      const id = args?.noteId ?? footnoteAtSelection(state);
      const next = id ? deleteFootnote(state, id) : null;
      return next && { state: next, meta: { label: 'Elimina nota a piè di pagina' } };
    },
  }),
];

export const footnotesModule: EditorModule = {
  id: 'footnotes',
  activate(editor) {
    const disposers = footnoteCommands.map((c) => editor.commands.register(c));
    return () => disposers.forEach((d) => d());
  },
};
