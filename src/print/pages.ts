/**
 * Scelte della finestra Stampa che non dipendono dal DOM: quali pagine
 * stampare (tutte, la corrente, un intervallo come "1-3, 5") e
 * l'orientamento da usare.
 */

import type { DocumentModel } from '../core';

export type PageChoice = 'all' | 'current' | 'custom';
export type Orientation = 'portrait' | 'landscape';

export interface PrintSettings {
  readonly pages: PageChoice;
  /** Intervallo scritto dall'utente, usato con `pages: 'custom'`. */
  readonly range: string;
  readonly copies: number;
  /** `null`: come nel documento. */
  readonly orientation: Orientation | null;
}

export const DEFAULT_SETTINGS: PrintSettings = { pages: 'all', range: '', copies: 1, orientation: null };
export const MAX_COPIES = 999;

export type PageSelection = { readonly pages: readonly number[] } | { readonly error: string };

/**
 * Legge un intervallo come in Word: numeri e intervalli separati da virgole
 * o punti e virgola ("1-3, 5", "8-" fino all'ultima). Restituisce gli indici
 * (da 0) nell'ordine scritto, senza ripetizioni.
 */
export function parsePageRanges(text: string, total: number): PageSelection {
  const parts = text
    .split(/[,;]/)
    .map((p) => p.trim())
    .filter(Boolean);
  if (!parts.length) return { error: 'Scrivi le pagine da stampare, per esempio 1-3, 5.' };
  const out: number[] = [];
  const add = (n: number) => {
    if (!out.includes(n - 1)) out.push(n - 1);
  };
  for (const part of parts) {
    const m = /^(\d+)?\s*(?:([-–])\s*(\d+)?)?$/.exec(part);
    if (!m || (!m[1] && !m[3])) return { error: `"${part}" non è un numero di pagina valido.` };
    const from = Number(m[1] ?? 1);
    const to = m[2] ? Number(m[3] ?? total) : from;
    for (const n of [from, to]) {
      if (n < 1 || n > total) return { error: `La pagina ${n} non esiste: il documento ha ${total} ${total === 1 ? 'pagina' : 'pagine'}.` };
    }
    if (from <= to) for (let n = from; n <= to; n++) add(n);
    else for (let n = from; n >= to; n--) add(n);
  }
  return { pages: out };
}

/** Pagine da stampare secondo le impostazioni. `current` è l'indice della pagina col cursore. */
export function selectPages(settings: PrintSettings, total: number, current: number): PageSelection {
  if (total === 0) return { error: 'Il documento non ha pagine da stampare.' };
  if (settings.pages === 'current') return { pages: [Math.min(Math.max(current, 0), total - 1)] };
  if (settings.pages === 'custom') return parsePageRanges(settings.range, total);
  return { pages: Array.from({ length: total }, (_, i) => i) };
}

/** Il documento con tutte le sezioni nell'orientamento chiesto (larghezza e altezza scambiate). */
export function withOrientation(doc: DocumentModel, orientation: Orientation | null): DocumentModel {
  if (!orientation || doc.sections.every((s) => s.page.orientation === orientation)) return doc;
  return {
    ...doc,
    sections: doc.sections.map((s) => {
      if (s.page.orientation === orientation) return s;
      const turned = orientation === 'landscape' ? s.page.height > s.page.width : s.page.width > s.page.height;
      return { ...s, page: { ...s.page, orientation, ...(turned ? { width: s.page.height, height: s.page.width } : {}) } };
    }),
  };
}

export function clampCopies(n: number): number {
  return Number.isFinite(n) ? Math.min(MAX_COPIES, Math.max(1, Math.round(n))) : 1;
}
