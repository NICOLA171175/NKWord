/**
 * Modulo Lettura ad alta voce: Revisione > Leggi ad alta voce, come in Word.
 * Legge il documento dal cursore (o la selezione) evidenziando la parola letta.
 */

import type { RibbonItem } from '../core';
import { createReadAloudModule as createModule, type ReadAloudModule } from './module';
import { mountReadAloud, type ReadAloudHost } from './view';

export interface ReadAloudAppModule extends ReadAloudModule {
  /** Barra della lettura (chiamato da src/ui/main.ts). */
  mount(host: ReadAloudHost): () => void;
}

export function createReadAloudModule(): ReadAloudAppModule {
  const module = createModule() as ReadAloudAppModule;
  module.mount = (host) => mountReadAloud(host, module);
  return module;
}

export {
  READ_ALOUD_MODULE_ID,
  WORD_HIGHLIGHT,
  createReadAloudModule as createReadAloudCore,
  type ReadAloudCommandArgs,
  type ReadAloudEngine,
  type ReadAloudModule,
  type ReadAloudOptions,
  type ReadAloudStatus,
  type ReadAloudStorage,
  type SpeechLike,
  type UtteranceLike,
} from './module';
export * from './text';
export { mountReadAloud, rateLabel, type ReadAloudHost } from './view';

/** Voci della barra multifunzione, raccolte dalla UI (vedi `collectRibbonItems`). */
export const ribbonItems: readonly (RibbonItem & { readonly label: string })[] = [
  { tab: 'review', group: 'Voce', commandId: 'readaloud.toggle', label: 'Leggi ad alta voce', kind: 'toggle', icon: '🔊' },
];
