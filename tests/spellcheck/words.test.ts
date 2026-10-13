import { describe, expect, it } from 'vitest';
import { findWords } from '../../src/spellcheck';

const words = (text: string, upper?: boolean) => findWords(text, { ignoreUppercase: upper }).map((w) => w.word);

describe('findWords', () => {
  it("tiene insieme apostrofi ed elisioni, divide su spazi, trattini e punteggiatura", () => {
    expect(words("L'albero è alto, un po' storto: auto-ironia.")).toEqual(["L'albero", 'è', 'alto', 'un', "po'", 'storto', 'auto', 'ironia']);
    expect(words('dell’anno')).toEqual(['dell’anno']);
  });

  it('restituisce gli offset nel testo', () => {
    expect(findWords('ciao  mondo')).toEqual([
      { start: 0, end: 4, word: 'ciao' },
      { start: 6, end: 11, word: 'mondo' },
    ]);
  });

  it('salta numeri, parole con cifre, indirizzi web ed email', () => {
    expect(words('nel 2024 ho letto mp3 su www.esempio.it e https://x.org/a?b=c con mario.rossi@posta.it')).toEqual(['nel', 'ho', 'letto', 'su', 'e', 'con']);
  });

  it('salta le sigle in maiuscolo, a meno che non lo si chieda', () => {
    expect(words('La NASA e la Rai')).toEqual(['La', 'e', 'la', 'Rai']);
    expect(words('La NASA', false)).toEqual(['La', 'NASA']);
    expect(words('A casa')).toEqual(['A', 'casa']);
  });
});
