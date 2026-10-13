/**
 * Parole da controllare nel testo di un paragrafo (`paragraphText`).
 * Gli apostrofi fanno parte della parola ("l'albero", "po'"), i trattini la
 * dividono ("auto-ironia" → "auto", "ironia"). Si saltano indirizzi web e
 * email, parole con cifre e, come in Word, le parole tutte maiuscole.
 */

export interface WordSpan {
  readonly start: number;
  readonly end: number;
  readonly word: string;
}

export interface WordOptions {
  /** Salta le parole TUTTE MAIUSCOLE (sigle). Default: true. */
  readonly ignoreUppercase?: boolean;
}

const WORD = /[\p{L}\p{M}\p{N}_]+(?:['’][\p{L}\p{M}\p{N}_]+)*['’]?/gu;
const SKIP_SPANS = /\b(?:https?:\/\/|www\.)\S+|[^\s@]+@[^\s@]+\.[^\s@]+|\S+\.(?:com|it|org|net|eu|io)\b\S*/giu;
const DIGIT = /[\p{N}_]/u;

export function findWords(text: string, options: WordOptions = {}): WordSpan[] {
  const ignoreUppercase = options.ignoreUppercase ?? true;
  const skipped: [number, number][] = [];
  for (const m of text.matchAll(SKIP_SPANS)) skipped.push([m.index, m.index + m[0].length]);
  const out: WordSpan[] = [];
  for (const m of text.matchAll(WORD)) {
    const start = m.index;
    const end = start + m[0].length;
    const word = m[0];
    if (DIGIT.test(word)) continue;
    if (skipped.some(([a, b]) => start < b && end > a)) continue;
    const letters = word.replace(/['’]/g, '');
    if (ignoreUppercase && letters.length > 1 && letters === letters.toUpperCase() && letters !== letters.toLowerCase()) continue;
    out.push({ start, end, word });
  }
  return out;
}
