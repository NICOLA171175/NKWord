/**
 * Modulo sommario (Riferimenti > Sommario): inserisce un sommario generato
 * dai titoli del documento, lo aggiorna con i numeri di pagina
 * dell'impaginazione e lo rimuove.
 */

import type { EditorModule, RibbonItem } from '../core';
import { createTocCommands } from './commands';

export const TOC_MODULE_ID = 'toc';

export const tocModule: EditorModule = {
  id: TOC_MODULE_ID,
  activate(editor) {
    const disposers = createTocCommands(editor).map((c) => editor.commands.register(c));
    return () => disposers.forEach((d) => d());
  },
};

/** Voci della barra multifunzione, raccolte dalla UI (vedi `collectRibbonItems`). */
export const ribbonItems: readonly (RibbonItem & { readonly label: string })[] = [
  { tab: 'references', group: 'Sommario', commandId: 'toc.insert', label: 'Sommario', icon: '☰', kind: 'button' },
  { tab: 'references', group: 'Sommario', commandId: 'toc.update', label: 'Aggiorna sommario', icon: '⟳', kind: 'button' },
  { tab: 'references', group: 'Sommario', commandId: 'toc.remove', label: 'Rimuovi sommario', icon: '✕', kind: 'button' },
];

export { createTocCommands, type TocCommandArgs, type TocCommandId } from './commands';
export * from './styles';
export * from './toc';
