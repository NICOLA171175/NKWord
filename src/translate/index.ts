/**
 * Modulo Traduttore: Revisione > Traduci (Traduci selezione, Traduci
 * documento), come in Word.
 */

import type { RibbonItem } from '../core';
import { createTranslateModule as createModule, type TranslateModule } from './module';
import { mountTranslatePanel, type TranslatePanelHost } from './panel';

export interface TranslateAppModule extends TranslateModule {
  /** Riquadro Traduttore (chiamato da src/ui/main.ts). */
  mount(host: TranslatePanelHost): () => void;
}

export function createTranslateModule(): TranslateAppModule {
  const module = createModule() as TranslateAppModule;
  module.mount = (host) => mountTranslatePanel(host, module);
  return module;
}

export { AUTO_DETECT, LANGUAGES, languageName, type Language } from './languages';
export {
  collectTexts,
  hasSelectedText,
  isTranslatable,
  lookup,
  mapKeepingEdges,
  rewriteDocument,
  rewriteInlines,
  rewriteParagraph,
  rewriteSelection,
  selectionText,
  type MapText,
  type Translations,
} from './model';
export {
  SELECTION_DELAY_MS,
  SETTINGS_KEY,
  TRANSLATE_MODULE_ID,
  type TranslateCommandArgs,
  type TranslateMode,
  type TranslateModule,
  type TranslateOptions,
  type TranslateStatus,
} from './module';
export { mountTranslatePanel, type TranslatePanelHost } from './panel';
export {
  batches,
  createGoogleTranslator,
  GOOGLE_ENDPOINT,
  MAX_REQUEST_CHARS,
  parseGoogleResponse,
  TranslationError,
  type FetchLike,
  type TranslationResult,
  type Translator,
} from './service';

/** Voci della barra multifunzione, raccolte dalla UI (vedi `collectRibbonItems`). */
export const ribbonItems: readonly (RibbonItem & { readonly label: string; readonly options: readonly { label: string; args: unknown }[] })[] = [
  {
    tab: 'review',
    group: 'Lingua',
    commandId: 'translate.menu',
    label: 'Traduci',
    kind: 'dropdown',
    icon: '文A',
    options: [
      { label: 'Traduci selezione', args: { choice: 'selection' } },
      { label: 'Traduci documento', args: { choice: 'document' } },
    ],
  },
];
