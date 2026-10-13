/**
 * Modulo Copia formato: il pennello della scheda Home, come in Word.
 */

import type { RibbonItem } from '../core';
import { createFormatPainterModule as createModule, type FormatPainterModule } from './module';
import { mountFormatPainter, type FormatPainterHost } from './panel';

export interface FormatPainterAppModule extends FormatPainterModule {
  /** Doppio clic, puntatore e applicazione con il mouse (chiamato da src/ui/main.ts). */
  mount(host: FormatPainterHost): () => void;
}

export function createFormatPainterModule(): FormatPainterAppModule {
  const module = createModule() as FormatPainterAppModule;
  module.mount = (host) => mountFormatPainter(host, module);
  return module;
}

/** Voci della barra multifunzione, raccolte dalla UI (vedi `collectRibbonItems`). */
export const ribbonItems: readonly (RibbonItem & { readonly label: string })[] = [
  { tab: 'home', group: 'Appunti', commandId: 'formatPainter.toggle', label: 'Copia formato', kind: 'toggle', icon: '🖌' },
];

export { FORMAT_PAINTER_MODULE_ID, type FormatPainterCommandArgs, type FormatPainterModule, type PainterMode } from './module';
export { copyFormat, pasteFormat, type CopiedFormat } from './painter';
export { mountFormatPainter, PAINTER_ACTIVE_CLASS, type FormatPainterHost } from './panel';
