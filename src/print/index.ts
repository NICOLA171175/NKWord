/**
 * Modulo stampa: anteprima di stampa con scelta di pagine, copie e
 * orientamento (File > Stampa, Ctrl+P), e stampa con la finestra del sistema.
 */

import { createPrintModule as createModule, type PrintModule, type PrintOptions } from './module';
import { mountPrint, type PrintHost } from './panel';

export interface PrintAppModule extends PrintModule {
  /** Schermata Stampa (chiamato da src/ui/main.ts). */
  mount(host: PrintHost): () => void;
}

export function createPrintModule(options: PrintOptions = {}): PrintAppModule {
  const module = createModule(options) as PrintAppModule;
  module.mount = (host) => mountPrint(host, module);
  return module;
}

export { PRINT_MODULE_ID, type PrintCommandArgs, type PrintModule, type PrintOptions, type PrintPreview } from './module';
export { buildPrintOutput, createDomPrinter, PRINT_ACTIVE_CLASS, type Printer, type PrintJob } from './output';
export { clampCopies, DEFAULT_SETTINGS, MAX_COPIES, parsePageRanges, selectPages, withOrientation, type Orientation, type PageChoice, type PageSelection, type PrintSettings } from './pages';
export { DEFAULT_PREVIEW_ZOOM, mountPrint, type PrintHost } from './panel';
export { renderPage, type PageRenderOptions } from './render';
