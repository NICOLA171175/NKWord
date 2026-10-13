/**
 * Modulo Dettatura: Home > Detta, per scrivere parlando, come in Word.
 */

import type { RibbonItem } from '../core';
import { createDictationModule as createModule, type DictationModule } from './module';
import { mountDictation, type DictationHost } from './view';

export interface DictationAppModule extends DictationModule {
  /** Indicatore della dettatura (chiamato da src/ui/main.ts). */
  mount(host: DictationHost): () => void;
}

export function createDictationModule(): DictationAppModule {
  const module = createModule() as DictationAppModule;
  module.mount = (host) => mountDictation(host, module);
  return module;
}

export {
  detectEnvironment,
  dictationMode,
  errorMessage,
  type DesktopDictation,
  type DictationEnvironment,
  type DictationMode,
  type SpeechRecognitionCtor,
  type SpeechRecognitionLike,
  type SpeechResultEventLike,
} from './engine';
export { DICTATION_LANG, DICTATION_MODULE_ID, type DictationCommandArgs, type DictationMessage, type DictationModule } from './module';
export { capitalize, dictationPieces, VOICE_COMMANDS, type DictationPiece } from './text';
export { INFO_TIMEOUT_MS, mountDictation, type DictationHost } from './view';

/** Voci della barra multifunzione, raccolte dalla UI (vedi `collectRibbonItems`). */
export const ribbonItems: readonly (RibbonItem & { readonly label: string })[] = [
  { tab: 'home', group: 'Voce', commandId: 'dictation.toggle', label: 'Detta', kind: 'toggle', icon: '🎤' },
];
