/**
 * Simboli usati di recente, come la riga in alto della finestra Simbolo di
 * Word. Si ricordano fra una sessione e l'altra nella memoria del browser
 * (localStorage); se non è disponibile, solo finché l'app resta aperta.
 */

export const MAX_RECENT_SYMBOLS = 20;
const KEY = 'nkoffice-simboli-recenti';

/** Dove si conservano i recenti: basta un getItem/setItem (es. localStorage). */
export interface SymbolStorage {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
}

/** I simboli predefiniti prima di averne usato qualcuno, come in Word. */
export const DEFAULT_RECENT: readonly string[] = ['€', '£', '¥', '©', '®', '™', '±', '≠', '≤', '≥', '÷', '×', '∞', 'µ', 'α', 'β', 'π', 'Ω', '∑', '☺'];

export function browserStorage(): SymbolStorage | null {
  try {
    return globalThis.localStorage ?? null;
  } catch {
    return null;
  }
}

export function loadRecent(storage: SymbolStorage | null): string[] {
  try {
    const raw = storage?.getItem(KEY);
    if (!raw) return [...DEFAULT_RECENT];
    const list = JSON.parse(raw) as unknown;
    if (!Array.isArray(list)) return [...DEFAULT_RECENT];
    return list.filter((c): c is string => typeof c === 'string' && c.length > 0 && c.length <= 4).slice(0, MAX_RECENT_SYMBOLS);
  } catch {
    return [...DEFAULT_RECENT];
  }
}

export function saveRecent(storage: SymbolStorage | null, list: readonly string[]): void {
  try {
    storage?.setItem(KEY, JSON.stringify(list));
  } catch {
    // Memoria del browser piena o bloccata: i recenti valgono solo per questa sessione.
  }
}

/** Mette `char` in testa all'elenco, senza doppioni, al massimo 20. */
export function pushRecent(list: readonly string[], char: string): string[] {
  return [char, ...list.filter((c) => c !== char)].slice(0, MAX_RECENT_SYMBOLS);
}
