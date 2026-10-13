/**
 * Modulo Modalità lettura: Visualizza > Modalità lettura, per leggere il
 * documento a schermo intero, senza barra multifunzione, come in Word.
 */

import type { RibbonItem } from '../core';
import { createReadModeModule as createModule, type ReadModeModule } from './module';
import { mountReadMode, type ReadModeHost } from './view';

export interface ReadModeAppModule extends ReadModeModule {
  /** Schermata della modalità lettura (chiamato da src/ui/main.ts). */
  mount(host: ReadModeHost): () => void;
}

export function createReadModeModule(): ReadModeAppModule {
  const module = createModule() as ReadModeAppModule;
  module.mount = (host) => mountReadMode(host, module);
  return module;
}

export {
  clampReadZoom,
  DEFAULT_READ_ZOOM,
  isPageColor,
  isReadLayout,
  MAX_READ_ZOOM,
  MIN_READ_ZOOM,
  PAGE_COLOR_LABELS,
  PAGE_COLORS,
  READ_LAYOUT_LABELS,
  READ_LAYOUTS,
  READ_MODE_MODULE_ID,
  READ_ZOOM_STEP,
  screenStart,
  type PageColor,
  type ReadLayout,
  type ReadModeCommandArgs,
  type ReadModeModule,
} from './module';
export { MAX_PAGES_PER_SCREEN, mountReadMode, pageLabel, pagesPerScreen, type ReadModeHost } from './view';

/** Voci della barra multifunzione, raccolte dalla UI (vedi `collectRibbonItems`). */
export const ribbonItems: readonly (RibbonItem & { readonly label: string })[] = [
  { tab: 'view', group: 'Visualizzazioni', commandId: 'readmode.open', label: 'Modalità lettura', icon: '📖' },
];
