/**
 * Calcoli di sola lettura della Lettura ad alta voce: quali pezzi di testo
 * leggere (il documento dal cursore in poi, oppure solo la selezione), dove
 * stanno le parole e quale voce scegliere. Nessun DOM, nessuna sintesi vocale.
 */

import { isCollapsed, orderedRange, paragraphs, paragraphText, type DocumentModel, type Id, type Selection } from '../core';

/** Un paragrafo (o il suo pezzo selezionato) da leggere. */
export interface ReadingChunk {
  readonly paragraphId: Id;
  /** Inizio e fine del pezzo nel paragrafo. */
  readonly start: number;
  readonly end: number;
  /**
   * Testo dell'intero paragrafo da pronunciare: tab, interruzioni e oggetti
   * (immagini, campi...) diventano spazi, così le posizioni restano uguali.
   */
  readonly text: string;
}

/** Velocità di lettura: 1 è quella normale della voce. */
export const MIN_RATE = 0.5;
export const MAX_RATE = 2;
export const RATE_STEP = 0.1;
export const DEFAULT_RATE = 1;

export function clampRate(rate: number): number {
  if (!Number.isFinite(rate)) return DEFAULT_RATE;
  const r = Math.min(MAX_RATE, Math.max(MIN_RATE, rate));
  return Math.round(r * 10) / 10;
}

/** Testo pronunciabile del paragrafo, con la stessa lunghezza di quello del modello. */
export function spokenText(raw: string): string {
  return raw.replace(/[\t\n\u000b￼­]/g, ' ');
}

const WORD_CHAR = /[\p{L}\p{N}'’_-]/u;
const isWordChar = (ch: string | undefined) => !!ch && WORD_CHAR.test(ch);

/**
 * La parola che contiene la posizione `index` (o la prima dopo, se lì c'è
 * uno spazio o una punteggiatura), come [inizio, fine]. null se non ce ne sono.
 */
export function wordAt(text: string, index: number, limit = text.length): [number, number] | null {
  let i = Math.max(0, Math.min(index, limit));
  // Dentro una parola: si torna all'inizio.
  if (isWordChar(text[i])) {
    while (i > 0 && isWordChar(text[i - 1])) i--;
  } else {
    while (i < limit && !isWordChar(text[i])) i++;
    if (i >= limit) return null;
  }
  let end = i;
  while (end < limit && isWordChar(text[end])) end++;
  return [i, end];
}

/** Inizio di tutte le parole del testo fino a `limit`, a partire da `from`. */
export function wordStarts(text: string, from = 0, limit = text.length): number[] {
  const out: number[] = [];
  let at = from;
  for (;;) {
    const w = wordAt(text, at, limit);
    if (!w) return out;
    out.push(w[0]);
    at = w[1];
  }
}

/**
 * Pezzi da leggere. Con una selezione si legge solo quella; altrimenti, come
 * in Word, dal cursore fino alla fine del documento. Paragrafi vuoti o senza
 * parole si saltano.
 */
export function readingChunks(doc: DocumentModel, selection: Selection): ReadingChunk[] {
  const selected = !isCollapsed(selection);
  const [from, to] = orderedRange(doc, selection);
  const out: ReadingChunk[] = [];
  let started = false;
  for (const { paragraph } of paragraphs(doc)) {
    const isFrom = paragraph.id === from.paragraphId;
    if (!started && !isFrom) continue;
    const text = spokenText(paragraphText(paragraph));
    let start = 0;
    if (!started) {
      started = true;
      // Dal cursore: si parte dall'inizio della parola in cui si trova.
      start = Math.min(from.offset, text.length);
      while (start > 0 && isWordChar(text[start - 1])) start--;
    }
    const isTo = selected && paragraph.id === to.paragraphId;
    const end = isTo ? Math.min(to.offset, text.length) : text.length;
    if (wordAt(text, start, end)) out.push({ paragraphId: paragraph.id, start, end, text });
    if (isTo) break;
  }
  return out;
}

/** Ciò che serve di una voce (una `SpeechSynthesisVoice` del browser). */
export interface VoiceLike {
  readonly name: string;
  readonly lang: string;
  readonly default?: boolean;
  readonly localService?: boolean;
}

export const isItalian = (v: VoiceLike) => /^it([-_]|$)/i.test(v.lang);

/**
 * La voce da usare: quella scelta dall'utente se c'è ancora, altrimenti una
 * italiana (meglio se installata sul computer, che funziona anche senza
 * Internet), altrimenti quella predefinita.
 */
export function pickVoice<V extends VoiceLike>(voices: readonly V[], preferred?: string | null): V | null {
  if (preferred) {
    const chosen = voices.find((v) => v.name === preferred);
    if (chosen) return chosen;
  }
  const italian = voices.filter(isItalian);
  return italian.find((v) => v.localService) ?? italian[0] ?? voices.find((v) => v.default) ?? voices[0] ?? null;
}

/** Le voci in ordine per il menu: prima le italiane, poi le altre, ognuna per nome. */
export function sortVoices<V extends VoiceLike>(voices: readonly V[]): V[] {
  return [...voices].sort((a, b) => Number(isItalian(b)) - Number(isItalian(a)) || a.name.localeCompare(b.name));
}

/** Etichetta della voce nel menu, es. "Microsoft Elsa (it-IT)". */
export function voiceLabel(v: VoiceLike): string {
  const name = v.name.replace(/\s*-\s*(Italian|Italiano).*$/i, '');
  return `${name} (${v.lang})`;
}
