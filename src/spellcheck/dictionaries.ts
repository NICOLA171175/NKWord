/**
 * Lingue disponibili. I dizionari (Hunspell) si caricano solo quando
 * servono, in un pezzo separato del bundle: l'avvio dell'app non li attende.
 *
 * - Italiano: dizionario di LibreOffice/OpenOffice (GPL 3), vedi LICENSE-it_IT.txt.
 * - Inglese (Stati Uniti): SCOWL / wordlist.sourceforge.net, vedi LICENSE-en_US.txt.
 */

import type { HunspellData } from './hunspell';

export interface Language {
  readonly id: string;
  readonly label: string;
  load(): Promise<HunspellData>;
}

export const DEFAULT_LANGUAGE = 'it';

export const LANGUAGES: readonly Language[] = [
  {
    id: 'it',
    label: 'Italiano',
    load: async () => {
      const [aff, dic] = await Promise.all([import('./dictionaries/it_IT.aff?raw'), import('./dictionaries/it_IT.dic?raw')]);
      return { aff: aff.default, dic: dic.default };
    },
  },
  {
    id: 'en',
    label: 'Inglese (Stati Uniti)',
    load: async () => {
      const [aff, dic] = await Promise.all([import('./dictionaries/en_US.aff?raw'), import('./dictionaries/en_US.dic?raw')]);
      return { aff: aff.default, dic: dic.default };
    },
  },
];
