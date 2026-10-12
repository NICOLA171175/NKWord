/**
 * Confini di carattere e di parola sul testo di un paragrafo
 * (`paragraphText`, dove ogni inline non testuale vale un carattere).
 */

type CharClass = 'word' | 'space' | 'other';

const WORD = /[\p{L}\p{N}\p{M}_]/u;
const SPACE = /\s/u;

function classOf(ch: string | undefined): CharClass | undefined {
  if (ch === undefined) return undefined;
  if (WORD.test(ch)) return 'word';
  if (SPACE.test(ch)) return 'space';
  return 'other';
}

const isHigh = (c: number) => c >= 0xd800 && c <= 0xdbff;
const isLow = (c: number) => c >= 0xdc00 && c <= 0xdfff;

/** Offset del carattere precedente, senza spezzare le coppie surrogate. */
export function prevChar(text: string, offset: number): number {
  if (offset <= 0) return 0;
  if (offset >= 2 && isLow(text.charCodeAt(offset - 1)) && isHigh(text.charCodeAt(offset - 2))) return offset - 2;
  return offset - 1;
}

export function nextChar(text: string, offset: number): number {
  if (offset >= text.length) return text.length;
  if (isHigh(text.charCodeAt(offset)) && isLow(text.charCodeAt(offset + 1))) return offset + 2;
  return offset + 1;
}

/** Ctrl+Freccia sinistra: inizio della parola precedente (salta gli spazi). */
export function wordLeft(text: string, offset: number): number {
  let i = offset;
  while (i > 0 && classOf(text[i - 1]) === 'space') i--;
  const cls = classOf(text[i - 1]);
  while (i > 0 && classOf(text[i - 1]) === cls) i--;
  return i;
}

/** Ctrl+Freccia destra: inizio della parola successiva, come in Word. */
export function wordRight(text: string, offset: number): number {
  let i = offset;
  const cls = classOf(text[i]);
  if (cls !== 'space') while (i < text.length && classOf(text[i]) === cls) i++;
  while (i < text.length && classOf(text[i]) === 'space') i++;
  return i;
}

/**
 * Parola che contiene il cursore, solo se il cursore è al suo interno
 * (lettere su entrambi i lati). Word applica la formattazione all'intera
 * parola in questo caso.
 */
export function wordAround(text: string, offset: number): [number, number] | null {
  if (classOf(text[offset - 1]) !== 'word' || classOf(text[offset]) !== 'word') return null;
  let a = offset;
  let b = offset;
  while (a > 0 && classOf(text[a - 1]) === 'word') a--;
  while (b < text.length && classOf(text[b]) === 'word') b++;
  return [a, b];
}

export function isWordChar(ch: string | undefined): boolean {
  return classOf(ch) === 'word';
}
