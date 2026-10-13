import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { Hunspell, matchCase, wordCase } from '../../src/spellcheck';
import { MINI_AFF, MINI_DIC } from './helpers';

const dir = new URL('../../src/spellcheck/dictionaries/', import.meta.url);
const real = (name: string) => new Hunspell({ aff: readFileSync(new URL(`${name}.aff`, dir), 'utf8'), dic: readFileSync(new URL(`${name}.dic`, dir), 'utf8') });

describe('Hunspell (dizionario minimo)', () => {
  const h = new Hunspell({ aff: MINI_AFF, dic: MINI_DIC });

  it('riconosce radici, suffissi, prefissi e la loro combinazione', () => {
    expect(h.size).toBe(6);
    for (const w of ['gatto', 'gatti', 'case', "l'amico", "l'amici", 'perché']) expect(h.correct(w)).toBe(true);
    // Il prefisso richiede una vocale; "case" non ha il flag L.
    for (const w of ["l'gatto", "l'casa", 'gatte', 'cas', 'perche']) expect(h.correct(w)).toBe(false);
  });

  it('accetta maiuscole a inizio frase e parole tutte maiuscole, non i nomi propri in minuscolo', () => {
    expect(h.correct('Gatto')).toBe(true);
    expect(h.correct('GATTI')).toBe(true);
    expect(h.correct("L'amico")).toBe(true);
    expect(h.correct('ROMA')).toBe(true);
    expect(h.correct('roma')).toBe(false);
    expect(h.correct('gAtto')).toBe(false);
  });

  it("tratta l'apostrofo tipografico come quello semplice", () => {
    expect(h.correct('l’amico')).toBe(true);
  });

  it('suggerisce accenti (MAP), errori tipici (REP), lettere scambiate o mancanti', () => {
    expect(h.suggest('perche')[0]).toBe('perché');
    expect(h.suggest("perche'")[0]).toBe('perché');
    expect(h.suggest('cassa')).toContain('casa');
    expect(h.suggest('gatot')[0]).toBe('gatto');
    expect(h.suggest('gato')[0]).toBe('gatto');
    expect(h.suggest('ciaoo')[0]).toBe('ciao');
    expect(h.suggest('roma')).toEqual(['Roma']);
    expect(h.suggest('ciaogatto')).toContain('ciao gatto');
    expect(h.suggest('zzzzzz')).toEqual([]);
  });

  it('rispetta le maiuscole della parola sbagliata', () => {
    expect(h.suggest('Gatot')[0]).toBe('Gatto');
    expect(h.suggest('GATOT')[0]).toBe('GATTO');
  });

  it('aggiunge e toglie parole', () => {
    const d = new Hunspell({ aff: MINI_AFF, dic: MINI_DIC });
    expect(d.correct('nikoffice')).toBe(false);
    d.add('nikoffice');
    expect(d.correct('nikoffice')).toBe(true);
    d.remove('nikoffice');
    expect(d.correct('nikoffice')).toBe(false);
  });

  it('riconosce maiuscole e minuscole', () => {
    expect(wordCase('casa')).toBe('lower');
    expect(wordCase('Casa')).toBe('capitalized');
    expect(wordCase('CASA')).toBe('upper');
    expect(wordCase('cAsa')).toBe('mixed');
    expect(matchCase('casa', 'Xyz')).toBe('Casa');
  });
});

describe('dizionario italiano', () => {
  const it_ = real('it_IT');

  it('accetta forme flesse, elisioni e parole accentate', () => {
    const ok = ['casa', 'case', 'perché', 'più', 'è', 'andavamo', 'mangiato', 'portandoglielo', "l'albero", "L'albero", "dell'anno", "po'", 'quest’anno', "c'è", 'Italia', 'ITALIA'];
    for (const w of ok) expect(it_.correct(w), w).toBe(true);
    for (const w of ['perche', 'ciaoo', 'mangaito', 'roma', "qual'è", 'propio']) expect(it_.correct(w), w).toBe(false);
  });

  it('propone le correzioni giuste per gli errori comuni', () => {
    expect(it_.suggest('perche')[0]).toBe('perché');
    expect(it_.suggest('perchè')[0]).toBe('perché');
    expect(it_.suggest('mangaito')[0]).toBe('mangiato');
    expect(it_.suggest('ciaoo')[0]).toBe('ciao');
    expect(it_.suggest('propio')[0]).toBe('proprio');
    expect(it_.suggest("qual'è")[0]).toBe('qual è');
    expect(it_.suggest('roma')).toEqual(['Roma']);
    expect(it_.suggest('erore')[0]).toBe('errore');
    expect(it_.suggest('aqua')).toContain('acqua');
  });
});

describe('dizionario inglese', () => {
  const en = real('en_US');
  it('controlla e corregge', () => {
    for (const w of ['hello', 'running', 'unhappy', "don't", 'cats']) expect(en.correct(w), w).toBe(true);
    expect(en.correct('teh')).toBe(false);
    expect(en.suggest('teh')[0]).toBe('the');
    expect(en.suggest('recieve')[0]).toBe('receive');
  });
});
