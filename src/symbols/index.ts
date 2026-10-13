/**
 * Modulo Simbolo: la tabella di Inserisci > Simbolo, per scrivere con un
 * clic i caratteri che non sono sulla tastiera (€, ©, ½, frecce, È...).
 */

import type { RibbonItem } from '../core';
import { createSymbolsModule as createModule, type SymbolsModule, type SymbolsOptions } from './module';
import { mountSymbols, type SymbolsHost } from './panel';

export interface SymbolsAppModule extends SymbolsModule {
  /** Tabella sotto il pulsante (chiamato da src/ui/main.ts). */
  mount(host: SymbolsHost): () => void;
}

export function createSymbolsModule(options: SymbolsOptions = {}): SymbolsAppModule {
  const module = createModule(options) as SymbolsAppModule;
  module.mount = (host) => mountSymbols(host, module);
  return module;
}

/** Voci della barra multifunzione, raccolte dalla UI (vedi `collectRibbonItems`). */
export const ribbonItems: readonly (RibbonItem & { readonly label: string })[] = [
  { tab: 'insert', group: 'Simboli', commandId: 'symbol.show', label: 'Simbolo', kind: 'button', icon: 'Ω' },
];

export { SYMBOLS_MODULE_ID, type SymbolCommandArgs, type SymbolsModule, type SymbolsOptions } from './module';
export { charCode, charFromCode, SYMBOL_CATEGORIES, symbolName, type SymbolCategory } from './catalog';
export { DEFAULT_RECENT, MAX_RECENT_SYMBOLS, pushRecent, type SymbolStorage } from './recent';
export { mountSymbols, type SymbolsHost } from './panel';
