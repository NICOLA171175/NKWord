/**
 * Modulo Riferimenti: didascalie numerate (Figura 1, Tabella 1),
 * riferimenti incrociati che si aggiornano e indice analitico, come nella
 * scheda Riferimenti di Word. Il sommario resta nel suo modulo (src/toc),
 * da cui questo riprende numeri di pagina e collegamenti interni.
 */

import type { EditorModule, RibbonItem } from '../core';
import { domDialogs } from './dialog';
import { createReferencesModule as createModule, type ReferencesOptions } from './module';

export function createReferencesModule(options: ReferencesOptions = {}): EditorModule {
  return createModule({ ...options, dialogs: options.dialogs ?? (typeof document !== 'undefined' ? domDialogs : undefined) });
}

/** Voci della barra multifunzione, raccolte dalla UI (vedi `collectRibbonItems`). */
export const ribbonItems: readonly (RibbonItem & { readonly label: string })[] = [
  { tab: 'references', group: 'Didascalie', commandId: 'references.captionDialog', label: 'Inserisci didascalia', kind: 'button', icon: '▭' },
  { tab: 'references', group: 'Didascalie', commandId: 'references.crossRefDialog', label: 'Riferimento incrociato', kind: 'button', icon: '↪' },
  { tab: 'references', group: 'Didascalie', commandId: 'references.updateFields', label: 'Aggiorna tutto', kind: 'button', icon: '⟳' },
  { tab: 'references', group: 'Indice', commandId: 'references.markEntryDialog', label: 'Segna voce', kind: 'button', icon: '✎' },
  { tab: 'references', group: 'Indice', commandId: 'references.insertIndex', label: 'Inserisci indice', kind: 'button', icon: '☰' },
  { tab: 'references', group: 'Indice', commandId: 'references.updateIndex', label: 'Aggiorna indice', kind: 'button', icon: '⟳' },
];

export {
  REFERENCES_MODULE_ID,
  createReferencesCommands,
  type ReferencesCommandArgs,
  type ReferencesCommandId,
  type ReferencesDialogs,
  type ReferencesOptions,
} from './module';
export * from './captions';
export * from './crossrefs';
export * from './indexing';
export { crossRefCategories, domDialogs, openCaptionDialog, openCrossRefDialog, openMarkEntryDialog } from './dialog';
