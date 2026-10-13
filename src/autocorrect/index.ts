/**
 * Modulo correzione automatica: corregge mentre si scrive (maiuscola a
 * inizio frase, virgolette tipografiche, (c) → ©, -- → —, errori di battitura
 * frequenti in italiano), con le opzioni in Revisione > Correzione automatica.
 */

import type { RibbonItem } from '../core';
import { openAutoCorrectDialog } from './dialog';
import { createAutoCorrectModule as createModule, type AutoCorrectModule, type AutoCorrectOptions } from './module';
import './styles.css';

export function createAutoCorrectModule(options: AutoCorrectOptions = {}): AutoCorrectModule {
  return createModule({
    ...options,
    openDialog: options.openDialog ?? (typeof document !== 'undefined' ? openAutoCorrectDialog : undefined),
  });
}

/** Voci della barra multifunzione, raccolte dalla UI (vedi `collectRibbonItems`). */
export const ribbonItems: readonly (RibbonItem & { readonly label: string })[] = [
  { tab: 'review', group: 'Strumenti di correzione', commandId: 'autocorrect.options', label: 'Correzione automatica', kind: 'button', icon: 'A✓' },
];

export { AUTOCORRECT_MODULE_ID, replaceInParagraph, type AutoCorrectCommandArgs, type AutoCorrectModule, type AutoCorrectOptions } from './module';
export {
  ABBREVIATIONS,
  activeReplacements,
  correctWord,
  DEFAULT_REPLACEMENTS,
  DEFAULT_SETTINGS,
  isSentenceStart,
  looksLikeAddress,
  matchCase,
  smartQuote,
  tokenStart,
  type AutoCorrectSettings,
  type Correction,
  type Replacement,
} from './rules';
export { loadSettings, saveSettings, type AutoCorrectStorage } from './settings';
export { diffReplacements, openAutoCorrectDialog } from './dialog';
