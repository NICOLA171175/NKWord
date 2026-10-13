/**
 * Opzioni della correzione automatica, ricordate fra una sessione e l'altra
 * nella memoria del browser (localStorage); se non è disponibile, valgono
 * solo finché l'app resta aperta.
 */

import { DEFAULT_SETTINGS, type AutoCorrectSettings, type Replacement } from './rules';

const KEY = 'nkoffice-correzione-automatica';

/** Dove si conservano le opzioni: basta un getItem/setItem (es. localStorage). */
export interface AutoCorrectStorage {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
}

export function browserStorage(): AutoCorrectStorage | null {
  try {
    return globalThis.localStorage ?? null;
  } catch {
    return null;
  }
}

const isReplacement = (r: unknown): r is Replacement =>
  typeof (r as Replacement)?.from === 'string' && typeof (r as Replacement).to === 'string' && (r as Replacement).from.length > 0;

/** Legge le opzioni salvate; quelle mancanti o rovinate prendono il valore predefinito. */
export function loadSettings(storage: AutoCorrectStorage | null): AutoCorrectSettings {
  try {
    const raw = storage?.getItem(KEY);
    if (!raw) return DEFAULT_SETTINGS;
    const saved = JSON.parse(raw) as Partial<Record<keyof AutoCorrectSettings, unknown>>;
    const flag = (k: 'twoInitialCaps' | 'capitalizeSentences' | 'capsLock' | 'smartQuotes' | 'replaceText') =>
      typeof saved[k] === 'boolean' ? (saved[k] as boolean) : DEFAULT_SETTINGS[k];
    return {
      twoInitialCaps: flag('twoInitialCaps'),
      capitalizeSentences: flag('capitalizeSentences'),
      capsLock: flag('capsLock'),
      smartQuotes: flag('smartQuotes'),
      replaceText: flag('replaceText'),
      quoteStyle: saved.quoteStyle === 'angle' ? 'angle' : 'curly',
      custom: Array.isArray(saved.custom) ? saved.custom.filter(isReplacement).map(({ from, to }) => ({ from, to })) : [],
      removed: Array.isArray(saved.removed) ? saved.removed.filter((s): s is string => typeof s === 'string') : [],
    };
  } catch {
    return DEFAULT_SETTINGS;
  }
}

export function saveSettings(storage: AutoCorrectStorage | null, settings: AutoCorrectSettings): void {
  try {
    storage?.setItem(KEY, JSON.stringify(settings));
  } catch {
    // Memoria del browser piena o bloccata: le opzioni valgono solo per questa sessione.
  }
}
