/**
 * Modulo salvataggio automatico: copia periodicamente il documento non
 * salvato e, dopo una chiusura improvvisa, propone di recuperarlo.
 */

import type { RibbonItem } from '../core';
import { mountAutosave, type AutosaveHost } from './banner';
import { createAutosaveModule as createModule, type AutosaveModule, type AutosaveOptions } from './module';

export interface AutosaveAppModule extends AutosaveModule {
  /** Avviso di recupero e indicatore nella barra di stato (chiamato da src/ui/main.ts). */
  mount(host: AutosaveHost): () => void;
}

export function createAutosaveModule(options: AutosaveOptions = {}): AutosaveAppModule {
  const module = createModule(options) as AutosaveAppModule;
  module.mount = (host) => mountAutosave(host, module);
  return module;
}

/** Voci della barra multifunzione, raccolte dalla UI (vedi `collectRibbonItems`). */
export const ribbonItems: readonly (RibbonItem & { readonly label: string })[] = [
  { tab: 'file', group: 'Recupera', commandId: 'autosave.show', label: 'Documenti recuperati', icon: '⟲', kind: 'button' },
];

export { AUTOSAVE_MODULE_ID, RESTORE_LABEL, parseSnapshot, type AutosaveCommandArgs, type AutosaveModule, type AutosaveOptions } from './module';
export { createPresence, type Presence } from './presence';
export * from './store';
export { mountAutosave, type AutosaveHost } from './banner';
