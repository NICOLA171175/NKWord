/**
 * Modulo documenti recenti: l'elenco degli ultimi file aperti o salvati,
 * nella scheda File, riapribili con un clic.
 */

import type { RibbonItem } from '../core';
import { createRecentModule as createModule, type RecentModule, type RecentOptions } from './module';
import { mountRecent, type RecentHost } from './panel';

export interface RecentAppModule extends RecentModule {
  /** Elenco sotto il pulsante e collegamento con la shell (chiamato da src/ui/main.ts). */
  mount(host: RecentHost): () => void;
}

export function createRecentModule(options: RecentOptions = {}): RecentAppModule {
  const module = createModule(options) as RecentAppModule;
  module.mount = (host) => mountRecent(host, module);
  return module;
}

/** Voci della barra multifunzione, raccolte dalla UI (vedi `collectRibbonItems`). */
export const ribbonItems: readonly (RibbonItem & { readonly label: string })[] = [
  { tab: 'file', group: 'File', commandId: 'recent.show', label: 'Recenti', kind: 'button' },
];

export { MAX_RECENT, RECENT_MODULE_ID, type RecentCommandArgs, type RecentModule, type RecentOpener, type RecentOpenResult, type RecentOptions } from './module';
export { baseName, desktopBridge, dirName, type DesktopBridge } from './desktop';
export { mountRecent, whenUsed, whereIs, type RecentHost } from './panel';
export * from './store';
