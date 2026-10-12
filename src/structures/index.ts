/**
 * Modulo strutture: elenchi, tabelle, oggetti inseriti, interruzioni e
 * impostazioni di pagina. Registra i comandi assegnati a 'structures' in
 * `COMMAND_OWNERS`.
 */

import type { AnyCommand, EditorModule } from '../core';
import { insertCommands } from './inserts';
import { listCommands } from './lists';
import { tableCommands } from './tables';

export const structuresCommands: readonly AnyCommand[] = [...listCommands, ...tableCommands, ...insertCommands];

export const structuresModule: EditorModule = {
  id: 'structures',
  activate(editor) {
    const disposers = structuresCommands.map((c) => editor.commands.register(c));
    return () => disposers.forEach((d) => d());
  },
};

export { bulletList, numberedList, listKind } from './lists';
export { tableChain, tableSelection } from './tables';
export { applyPageSetup, isValidPageSetup } from './inserts';
export * as tableOps from './table-ops';
