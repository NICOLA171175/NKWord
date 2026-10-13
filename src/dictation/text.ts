/**
 * Dal parlato al testo: comandi vocali di punteggiatura ("punto",
 * "virgola", "a capo"...), spazi e maiuscole a inizio frase, come la
 * dettatura di Word in italiano.
 */

/** Un pezzo da inserire: testo, oppure un nuovo paragrafo ("a capo"). */
export type DictationPiece = { readonly kind: 'text'; readonly text: string } | { readonly kind: 'paragraph' };

/** Comandi vocali: frase detta → segno. `\n` è un nuovo paragrafo. */
export const VOICE_COMMANDS: readonly (readonly [string, string])[] = [
  ['punto e virgola', ';'],
  ['punto interrogativo', '?'],
  ['punto di domanda', '?'],
  ['punto esclamativo', '!'],
  ['puntini di sospensione', '...'],
  ['punto', '.'],
  ['virgola', ','],
  ['due punti', ':'],
  ['apri parentesi', '('],
  ['chiudi parentesi', ')'],
  ['apri virgolette', '"'],
  ['chiudi virgolette', '"'],
  ['nuovo paragrafo', '\n'],
  ['nuova riga', '\n'],
  ['a capo', '\n'],
];

/** Segni dopo i quali la parola seguente comincia con la maiuscola. */
const SENTENCE_END = /[.!?]$/;
/** Segni dopo i quali non si mette lo spazio (le virgolette solo se aperte). */
const NO_SPACE_AFTER = /[\s(«“‘]$/;

const normalize = (word: string) => word.toLocaleLowerCase('it-IT');

/** Prima lettera maiuscola (le altre restano come sono). */
export function capitalize(text: string): string {
  return text.charAt(0).toLocaleUpperCase('it-IT') + text.slice(1);
}

/** Il comando vocale che comincia alla parola `i`, con quante parole occupa. */
function commandAt(words: readonly string[], i: number): { sign: string; length: number } | null {
  for (const [phrase, sign] of VOICE_COMMANDS) {
    const parts = phrase.split(' ');
    if (parts.every((p, k) => words[i + k] !== undefined && normalize(words[i + k]!) === p)) return { sign, length: parts.length };
  }
  return null;
}

/**
 * Trasforma una frase riconosciuta nei pezzi da inserire. `before` è il
 * testo del paragrafo prima del cursore: serve per sapere se mettere lo
 * spazio e la maiuscola.
 */
export function dictationPieces(spoken: string, before = ''): DictationPiece[] {
  const words = spoken.trim().split(/\s+/).filter(Boolean);
  const out: DictationPiece[] = [];
  let context = before;
  let buffer = '';
  // Le virgolette dette si aprono e si chiudono a turno.
  let quoteOpen = (before.match(/"/g)?.length ?? 0) % 2 === 1;
  // Subito dopo una virgoletta aperta non va lo spazio.
  let glued = quoteOpen && before.endsWith('"');
  const needsSpace = () => !!context && !glued && !NO_SPACE_AFTER.test(context);
  const flush = () => {
    if (buffer) out.push({ kind: 'text', text: buffer });
    buffer = '';
  };
  const add = (text: string, glue = false) => {
    buffer += text;
    context += text;
    glued = glue;
  };

  for (let i = 0; i < words.length; ) {
    const command = commandAt(words, i);
    if (command) {
      i += command.length;
      const { sign } = command;
      if (sign === '\n') {
        flush();
        out.push({ kind: 'paragraph' });
        context = '';
        continue;
      }
      const opening = sign === '(' || (sign === '"' && !quoteOpen);
      if (sign === '"') quoteOpen = !quoteOpen;
      if (opening && needsSpace()) add(' ');
      add(sign, opening);
      continue;
    }
    let word = words[i]!;
    i += 1;
    const trimmed = context.trimEnd();
    if (!trimmed || SENTENCE_END.test(trimmed)) word = capitalize(word);
    // Il riconoscimento a volte mette da sé la punteggiatura: va attaccata.
    if (needsSpace() && !/^[.,;:!?)]/.test(word)) add(' ');
    add(word);
  }
  flush();
  return out;
}
