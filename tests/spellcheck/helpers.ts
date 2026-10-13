import type { Language, SpellStorage } from '../../src/spellcheck';

/** Mini dizionario Hunspell per i test: radici, un suffisso, un prefisso, REP e MAP. */
export const MINI_AFF = `SET UTF-8
TRY aeiouclrst'
NOSUGGEST !
REP 1
REP ss s
MAP 1
MAP eèé
SFX S Y 2
SFX S o i o
SFX S a e a
PFX L Y 1
PFX L 0 l' [aeiou]
`;

export const MINI_DIC = `6
gatto/SL
casa/S
amico/SL
perché
ciao
Roma
`;

export const miniLanguages: Language[] = [
  { id: 'it', label: 'Italiano', load: async () => ({ aff: MINI_AFF, dic: MINI_DIC }) },
  { id: 'xx', label: 'Prova', load: async () => ({ aff: 'SET UTF-8\n', dic: '1\nhello\n' }) },
];

export function memoryStorage(initial: string | null = null): SpellStorage & { value: string | null } {
  return {
    value: initial,
    read() {
      return this.value;
    },
    write(v) {
      this.value = v;
    },
  };
}
