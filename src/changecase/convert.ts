/**
 * Conversione maiuscole/minuscole come il pulsante "Aa" di Word. Ogni
 * carattere resta della stessa lunghezza, così gli offset (e la selezione)
 * non cambiano.
 */

import {
  inlineLength,
  paragraphLength,
  paragraphText,
  replaceBlock,
  type EditorState,
  type Inline,
  type Paragraph,
  type Position,
  type Transaction,
} from '../core';
import { targetRange } from '../editing/format';
import { paragraphsBetween } from '../editing/range';

export type CaseMode = 'sentence' | 'lower' | 'upper' | 'title' | 'toggle';

const LETTER = /\p{L}/u;
const WORD = /[\p{L}\p{N}\p{M}'’]/u;

/** Sostituisce `ch` solo se la conversione non cambia la lunghezza (es. "ß" resta "ß"). */
const same = (ch: string, next: string) => (next.length === ch.length ? next : ch);

/**
 * Converte `text`. `sentenceStart` dice se il testo comincia una frase
 * (serve per "Tipo frase.").
 */
export function convertCase(text: string, mode: CaseMode, sentenceStart = true): string {
  let out = '';
  let startOfSentence = sentenceStart;
  let inWord = false;
  for (const ch of text) {
    const lower = same(ch, ch.toLowerCase());
    const upper = same(ch, ch.toUpperCase());
    const letter = LETTER.test(ch);
    switch (mode) {
      case 'lower':
        out += lower;
        break;
      case 'upper':
        out += upper;
        break;
      case 'toggle':
        out += ch === upper ? lower : upper;
        break;
      case 'title':
        out += letter ? (inWord ? lower : upper) : ch;
        break;
      case 'sentence':
        out += letter ? (startOfSentence ? upper : lower) : ch;
        break;
    }
    if (letter) startOfSentence = false;
    else if (/[.!?]/.test(ch)) startOfSentence = true;
    inWord = WORD.test(ch);
  }
  return out;
}

/** Maiusc+F3 come in Word: misto → MAIUSCOLO → minuscolo → Iniziali Maiuscole → MAIUSCOLO. */
export function nextCycleMode(text: string): CaseMode {
  const letters = [...text].filter((c) => LETTER.test(c)).join('');
  if (letters && letters === letters.toUpperCase() && letters !== letters.toLowerCase()) return 'lower';
  if (letters === letters.toLowerCase()) return 'title';
  return 'upper';
}

function pieces(state: EditorState, [start, end]: [Position, Position]) {
  return paragraphsBetween(state.doc, start, end).map(({ paragraph: p }) => ({
    p,
    from: p.id === start.paragraphId ? start.offset : 0,
    to: p.id === end.paragraphId ? end.offset : paragraphLength(p),
  }));
}

/** Testo selezionato (o della parola al cursore), paragrafi separati da '\n'. */
export function selectedText(state: EditorState): string {
  const range = targetRange(state);
  if (!range) return '';
  return pieces(state, range)
    .map(({ p, from, to }) => paragraphText(p).slice(from, to))
    .join('\n');
}

function withText(p: Paragraph, from: number, converted: string): Paragraph {
  let pos = 0;
  const children = p.children.map((inline): Inline => {
    const len = inlineLength(inline);
    const start = pos;
    pos += len;
    if (inline.type !== 'text') return inline;
    const a = Math.max(start, from);
    const b = Math.min(start + len, from + converted.length);
    if (a >= b) return inline;
    const text = inline.text.slice(0, a - start) + converted.slice(a - from, b - from) + inline.text.slice(b - start);
    return text === inline.text ? inline : { ...inline, text };
  });
  return { ...p, children };
}

/** Cambia maiuscole/minuscole della selezione o della parola al cursore. */
export function changeCase(state: EditorState, mode: CaseMode | 'cycle'): Transaction | null {
  const range = targetRange(state);
  if (!range) return null;
  const actual = mode === 'cycle' ? nextCycleMode(selectedText(state)) : mode;
  let doc = state.doc;
  for (const { p, from, to } of pieces(state, range)) {
    const text = paragraphText(p);
    const before = text.slice(0, from).trimEnd();
    const sentenceStart = before === '' || /[.!?]$/.test(before);
    const converted = convertCase(text.slice(from, to), actual, sentenceStart);
    const next = withText(p, from, converted);
    if (next !== p && next.children.some((c, i) => c !== p.children[i])) doc = replaceBlock(doc, p.id, [next]);
  }
  if (doc === state.doc) return null;
  return { state: { doc, selection: state.selection }, meta: { label: 'Maiuscole/minuscole' } };
}
