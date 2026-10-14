/**
 * Regole della correzione automatica, come in Word (File > Opzioni >
 * Strumenti di correzione > Opzioni correzione automatica). Funzioni pure
 * sul testo del paragrafo: dicono cosa sostituire, il modulo lo applica.
 */

export interface Replacement {
  /** Testo scritto, es. '(c)' o 'perchè'. */
  readonly from: string;
  /** Testo che lo sostituisce, es. '©' o 'perché'. */
  readonly to: string;
}

export interface AutoCorrectSettings {
  /** Correggi DUe INiziali MAiuscole. */
  readonly twoInitialCaps: boolean;
  /** Maiuscola all'inizio delle frasi. */
  readonly capitalizeSentences: boolean;
  /** Correggi l'uso accidentale del tasto BLOC MAIUSC (cIAO → Ciao). */
  readonly capsLock: boolean;
  /** Virgolette diritte → virgolette tipografiche. */
  readonly smartQuotes: boolean;
  /** Virgolette doppie: alte “ ” o basse « ». */
  readonly quoteStyle: 'curly' | 'angle';
  /** Sostituisci il testo durante la digitazione (elenco qui sotto). */
  readonly replaceText: boolean;
  /** Voci aggiunte dall'utente (o che cambiano una voce predefinita). */
  readonly custom: readonly Replacement[];
  /** Voci predefinite tolte dall'utente (il loro `from`). */
  readonly removed: readonly string[];
}

export const DEFAULT_SETTINGS: AutoCorrectSettings = {
  twoInitialCaps: true,
  capitalizeSentences: true,
  capsLock: true,
  smartQuotes: true,
  quoteStyle: 'curly',
  replaceText: true,
  custom: [],
  removed: [],
};

/** Simboli: si sostituiscono anche attaccati ad altro testo (es. "parola--parola"). */
const SYMBOLS: readonly Replacement[] = [
  { from: '<=>', to: '⇔' },
  { from: '==>', to: '⇒' },
  { from: '<==', to: '⇐' },
  { from: '-->', to: '→' },
  { from: '<--', to: '←' },
  { from: '...', to: '…' },
  { from: '--', to: '—' },
];

/** Voci predefinite dell'elenco, sostituite solo quando sono una parola intera. */
export const DEFAULT_REPLACEMENTS: readonly Replacement[] = [
  { from: '(c)', to: '©' },
  { from: '(r)', to: '®' },
  { from: '(tm)', to: '™' },
  { from: '(e)', to: '€' },
  { from: '1/2', to: '½' },
  { from: '1/4', to: '¼' },
  { from: '3/4', to: '¾' },
  { from: ':)', to: '☺' },
  { from: ':-)', to: '☺' },
  { from: ':(', to: '☹' },
  { from: ':-(', to: '☹' },
  ...SYMBOLS,
  // Errori di battitura frequenti in italiano.
  { from: 'perchè', to: 'perché' },
  { from: 'perche', to: 'perché' },
  { from: 'poichè', to: 'poiché' },
  { from: 'affinchè', to: 'affinché' },
  { from: 'benchè', to: 'benché' },
  { from: 'finchè', to: 'finché' },
  { from: 'giacchè', to: 'giacché' },
  { from: 'nonchè', to: 'nonché' },
  { from: 'cosicchè', to: 'cosicché' },
  { from: 'purchè', to: 'purché' },
  { from: 'sicchè', to: 'sicché' },
  { from: 'nè', to: 'né' },
  { from: 'pò', to: "po'" },
  { from: 'qual\'è', to: 'qual è' },
  { from: 'qualè', to: 'qual è' },
  { from: 'un\'altro', to: 'un altro' },
  { from: 'c\'e', to: 'c\'è' },
  { from: 'gia', to: 'già' },
  { from: 'piu', to: 'più' },
  { from: 'propio', to: 'proprio' },
  { from: 'sopratutto', to: 'soprattutto' },
  { from: 'daccordo', to: 'd\'accordo' },
  { from: 'pultroppo', to: 'purtroppo' },
  { from: 'purtoppo', to: 'purtroppo' },
  { from: 'efficente', to: 'efficiente' },
  { from: 'sufficente', to: 'sufficiente' },
  { from: 'deficente', to: 'deficiente' },
  { from: 'coscenza', to: 'coscienza' },
  { from: 'conoscienza', to: 'conoscenza' },
  { from: 'celebrale', to: 'cerebrale' },
  { from: 'metereologico', to: 'meteorologico' },
  { from: 'metereologia', to: 'meteorologia' },
  { from: 'aereoporto', to: 'aeroporto' },
  { from: 'areoporto', to: 'aeroporto' },
  { from: 'accellerare', to: 'accelerare' },
  { from: 'eccezzione', to: 'eccezione' },
  { from: 'eccezzionale', to: 'eccezionale' },
  { from: 'squola', to: 'scuola' },
  { from: 'ingegniere', to: 'ingegnere' },
  { from: 'cuoziente', to: 'quoziente' },
  { from: 'proffessore', to: 'professore' },
  { from: 'pressapoco', to: 'pressappoco' },
  { from: 'apposto', to: 'a posto' },
  { from: 'qualcun\'altro', to: 'qualcun altro' },
  { from: 'nessun\'altro', to: 'nessun altro' },
  { from: 'xchè', to: 'perché' },
  { from: 'xké', to: 'perché' },
];

/** Abbreviazioni dopo le quali il punto non chiude la frase ("Eccezioni" di Word). */
export const ABBREVIATIONS: ReadonlySet<string> = new Set([
  'es', 'ecc', 'etc', 'pag', 'pagg', 'p', 'pp', 'dott', 'dr', 'sig', 'sigg', 'sig.ra', 'prof', 'ing', 'avv', 'geom', 'rag',
  'arch', 'on', 'egr', 'gent', 'spett', 'art', 'artt', 'cap', 'capp', 'cfr', 'n', 'nr', 'num', 'tel', 'fig', 'figg', 'vol',
  'ca', 'circa', 'sez', 'par', 'cit', 'ed', 'vs', 'mr', 'mrs', 'ms', 'st', 'e.g', 'i.e', 'c.a', 's.p.a', 's.r.l', 'v', 'cod',
]);

/** Una sostituzione da fare nel paragrafo: il testo [start, end) diventa `text`. */
export interface Correction {
  readonly start: number;
  readonly end: number;
  readonly text: string;
}

const SPACE = /[\s ￼]/;
const OPENERS = '([{"\'“‘«„‚';
const CLOSERS = ')]}"\'”’»';
const TRAILING = '.,;:!?' + CLOSERS.replace("'", '').replace('’', '');

/** Elenco attivo: predefinite (meno quelle tolte) più quelle dell'utente. */
export function activeReplacements(settings: AutoCorrectSettings): Replacement[] {
  const removed = new Set(settings.removed.map(key));
  const custom = new Map(settings.custom.map((r) => [key(r.from), r]));
  const out = DEFAULT_REPLACEMENTS.filter((r) => !removed.has(key(r.from)) && !custom.has(key(r.from)));
  return [...out, ...custom.values()];
}

/** Chiave di confronto: minuscole, apostrofo tipografico come quello diritto. */
function key(s: string): string {
  return s.toLowerCase().replace(/’/g, "'");
}

/** Inizio della "parola" (tutto fino allo spazio precedente) che finisce a `end`. */
export function tokenStart(text: string, end: number): number {
  let i = end;
  while (i > 0 && !SPACE.test(text[i - 1]!)) i--;
  return i;
}

/** Indirizzi web, email e percorsi: li riconosce il modulo collegamenti, qui non si toccano. */
export function looksLikeAddress(token: string): boolean {
  return /:\/\/|^www\.|@|[\\/]\S*[\\/]|^[a-z]+:\S/i.test(token) || /\w\.\w+\//.test(token) || /\w\.(it|com|org|net|eu|edu|gov|info|io)\b/i.test(token);
}

const isSymbolEntry = (from: string) => !/[\p{L}\p{N}]/u.test(from);

/** Riporta sul testo nuovo maiuscole/minuscole di quello scritto (PERCHÈ → PERCHÉ). */
export function matchCase(typed: string, replacement: string): string {
  const letters = typed.replace(/[^\p{L}]/gu, '');
  if (!letters) return replacement;
  if (letters.length > 1 && letters === letters.toUpperCase() && letters !== letters.toLowerCase()) return replacement.toUpperCase();
  const first = typed.match(/\p{L}/u)?.[0];
  if (first && first !== first.toLowerCase()) return upperFirst(replacement);
  return replacement;
}

function upperFirst(s: string): string {
  const i = s.search(/\p{L}/u);
  return i < 0 ? s : s.slice(0, i) + s[i]!.toUpperCase() + s.slice(i + 1);
}

/** Toglie la punteggiatura intorno alla parola: '(perchè?' → ['(', 'perchè', '?']. */
function splitPunctuation(token: string): [string, string, string] {
  let a = 0;
  let b = token.length;
  while (a < b && OPENERS.includes(token[a]!)) a++;
  while (b > a && TRAILING.includes(token[b - 1]!)) b--;
  return [token.slice(0, a), token.slice(a, b), token.slice(b)];
}

function replaceWords(token: string, list: readonly Replacement[]): string {
  // Simboli anche dentro la parola (parola--parola), ma non nelle righe di trattini.
  let out = token;
  for (const r of list) {
    if (!isSymbolEntry(r.from) || !out.includes(r.from)) continue;
    if (r.from === '--' && /^-{3,}$/.test(out)) continue;
    out = out.split(r.from).join(r.to);
  }
  const lookup = new Map(list.filter((r) => !isSymbolEntry(r.from)).map((r) => [key(r.from), r.to]));
  const whole = lookup.get(key(out));
  if (whole !== undefined) return matchCase(out, whole);
  const [lead, core, tail] = splitPunctuation(out);
  const inner = lookup.get(key(core));
  return inner !== undefined && core ? lead + matchCase(core, inner) + tail : out;
}

function fixCaps(token: string, settings: AutoCorrectSettings): string {
  const [lead, core, tail] = splitPunctuation(token);
  if (!core || /[^\p{L}'’-]/u.test(core)) return token;
  // DUe INiziali → Due Iniziali (solo con almeno una minuscola dopo, così le sigle restano).
  if (settings.twoInitialCaps && /^\p{Lu}\p{Lu}\p{Ll}[\p{Ll}'’-]*$/u.test(core)) {
    return lead + core[0] + core.slice(1).toLowerCase() + tail;
  }
  // cIAO → Ciao: si era lasciato acceso il BLOC MAIUSC.
  if (settings.capsLock && /^\p{Ll}\p{Lu}{2,}$/u.test(core)) {
    return lead + core[0]!.toUpperCase() + core.slice(1).toLowerCase() + tail;
  }
  return token;
}

/** La parola che inizia a `start` è la prima di una frase? */
export function isSentenceStart(text: string, start: number): boolean {
  let i = start;
  while (i > 0 && SPACE.test(text[i - 1]!)) i--;
  if (i === 0) return true;
  // Le virgolette e parentesi chiuse dopo il punto non contano: «Vieni.» poi
  while (i > 0 && CLOSERS.includes(text[i - 1]!)) i--;
  const mark = text[i - 1];
  if (mark !== '.' && mark !== '!' && mark !== '?') return false;
  if (mark !== '.') return true;
  // Il punto di un'abbreviazione (es., ecc., dott.) o di un'iniziale (G. Verdi) non chiude la frase.
  const prevStart = tokenStart(text, i - 1);
  const prev = splitPunctuation(text.slice(prevStart, i - 1))[1].toLowerCase();
  if (text.slice(prevStart, i).endsWith('..')) return false;
  if (ABBREVIATIONS.has(prev) || /^\p{L}$/u.test(prev)) return false;
  return true;
}

function capitalizeSentence(text: string, start: number, token: string): string {
  const [lead, core] = splitPunctuation(token);
  if (!core || !/^\p{Ll}/u.test(core)) return token;
  // Nomi come iPhone o eBay restano come sono.
  if (/\p{Lu}/u.test(core.slice(1))) return token;
  if (!isSentenceStart(text, start)) return token;
  return lead + core[0]!.toUpperCase() + token.slice(lead.length + 1);
}

/**
 * Correzione della parola appena finita (che termina a `end`, prima dello
 * spazio o della punteggiatura scritta). `null` se non c'è nulla da fare.
 * La sostituzione è ristretta alla parte che cambia, così la formattazione
 * del resto della parola non si tocca.
 */
export function correctWord(text: string, end: number, settings: AutoCorrectSettings): Correction | null {
  const start = tokenStart(text, end);
  const token = text.slice(start, end);
  if (!token || looksLikeAddress(token)) return null;
  let out = token;
  if (settings.replaceText) {
    // Con le virgolette tipografiche anche gli apostrofi delle sostituzioni lo diventano (po' → po’).
    const list = activeReplacements(settings).map((r) => (settings.smartQuotes ? { ...r, to: r.to.replace(/'/g, '’') } : r));
    out = replaceWords(out, list);
  }
  // "parola - parola": il trattino isolato fra due parole diventa trattino medio.
  if (settings.replaceText && out === '-' && /\p{L}[.,;:!?]?\s+$/u.test(text.slice(0, start))) out = '–';
  // Il testo appena sostituito resta come scritto nell'elenco (es. NKWord).
  if (out === token) out = fixCaps(out, settings);
  if (settings.capitalizeSentences) out = capitalizeSentence(text, start, out);
  if (out === token) return null;
  // Solo la parte cambiata.
  let a = 0;
  while (a < token.length && a < out.length && token[a] === out[a]) a++;
  let b = 0;
  while (b < token.length - a && b < out.length - a && token[token.length - 1 - b] === out[out.length - 1 - b]) b++;
  return { start: start + a, end: end - b, text: out.slice(a, out.length - b) };
}

/**
 * Virgoletta tipografica per quella diritta appena scritta in `offset`:
 * aperta a inizio parola, chiusa (o apostrofo) altrimenti.
 */
export function smartQuote(text: string, offset: number, settings: AutoCorrectSettings): Correction | null {
  const ch = text[offset];
  if (ch !== '"' && ch !== "'") return null;
  const before = offset > 0 ? text[offset - 1]! : '';
  const opening = before === '' || SPACE.test(before) || '([{«“‘—–-/'.includes(before);
  let to: string;
  if (ch === "'") to = opening ? '‘' : '’';
  else if (settings.quoteStyle === 'angle') to = opening ? '«' : '»';
  else to = opening ? '“' : '”';
  return { start: offset, end: offset + 1, text: to };
}
