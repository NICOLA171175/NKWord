/**
 * Correttore compatibile con i dizionari Hunspell (.aff + .dic), ridotto a
 * quello che serve ai dizionari inclusi: prefissi e suffissi (anche
 * combinati), REP, MAP, TRY, NOSUGGEST, ONLYINCOMPOUND e FORBIDDENWORD.
 *
 * Invece di generare tutte le forme delle parole (milioni per l'italiano),
 * la verifica toglie prefissi e suffissi e cerca la radice nel dizionario,
 * come fa Hunspell: il caricamento resta veloce e la memoria contenuta.
 */

interface AffixRule {
  readonly flag: string;
  readonly cross: boolean;
  readonly strip: string;
  readonly add: string;
  readonly cond: RegExp | null;
}

interface AffixIndex {
  /** Regole raggruppate per testo aggiunto ('' per le regole che non aggiungono nulla). */
  readonly byAdd: Map<string, AffixRule[]>;
  maxAdd: number;
}

export interface HunspellData {
  readonly aff: string;
  readonly dic: string;
}

const APOSTROPHES = /[’ʼ]/g;

function emptyIndex(): AffixIndex {
  return { byAdd: new Map(), maxAdd: 0 };
}

function indexRule(index: AffixIndex, rule: AffixRule): void {
  let list = index.byAdd.get(rule.add);
  if (!list) index.byAdd.set(rule.add, (list = []));
  list.push(rule);
  index.maxAdd = Math.max(index.maxAdd, rule.add.length);
}

function condition(cond: string, suffix: boolean): RegExp | null {
  if (cond === '.' || cond === '') return null;
  return new RegExp(suffix ? `${cond}$` : `^${cond}`, 'u');
}

const KEY_ROWS = ['qwertyuiop', 'asdfghjkl', 'zxcvbnm'];

/** Tasti vicini su una tastiera QWERTY (errori di battitura). */
function nearKeys(a: string, b: string): boolean {
  const at = (ch: string): [number, number] | null => {
    for (let r = 0; r < KEY_ROWS.length; r++) {
      const c = KEY_ROWS[r]!.indexOf(ch);
      if (c >= 0) return [r, c];
    }
    return null;
  };
  const pa = at(a);
  const pb = at(b);
  if (!pa || !pb) return false;
  // Le righe sono sfalsate di mezzo tasto: la riga sotto è spostata a destra.
  const dr = pb[0] - pa[0];
  const dc = pb[1] - pa[1];
  if (dr === 0) return Math.abs(dc) === 1;
  if (dr === 1) return dc === 0 || dc === -1;
  if (dr === -1) return dc === 0 || dc === 1;
  return false;
}

const isUpper = (s: string) => s !== s.toLowerCase() && s === s.toUpperCase();
const capitalize = (s: string) => s.charAt(0).toUpperCase() + s.slice(1);

export type WordCase = 'lower' | 'capitalized' | 'upper' | 'mixed';

export function wordCase(word: string): WordCase {
  if (word === word.toLowerCase()) return 'lower';
  if (isUpper(word)) return word.length > 1 ? 'upper' : 'capitalized';
  const rest = word.slice(1);
  return isUpper(word.charAt(0)) && rest === rest.toLowerCase() ? 'capitalized' : 'mixed';
}

/** Applica a `word` la forma maiuscola/minuscola di `model`. */
export function matchCase(word: string, model: string): string {
  switch (wordCase(model)) {
    case 'upper':
      return word.toUpperCase();
    case 'capitalized':
      return capitalize(word);
    default:
      return word;
  }
}

export class Hunspell {
  private readonly words = new Map<string, string>();
  private readonly suffixes = emptyIndex();
  private readonly prefixes = emptyIndex();
  private readonly rep: [string, string][] = [];
  private readonly map: string[][] = [];
  private tryChars = '';
  private noSuggest = '';
  private onlyInCompound = '';
  private forbidden = '';
  private readonly cache = new Map<string, boolean>();

  constructor({ aff, dic }: HunspellData) {
    this.parseAff(aff);
    this.parseDic(dic);
  }

  /** Numero di radici nel dizionario. */
  get size(): number {
    return this.words.size;
  }

  /** Aggiunge una parola (senza flag) al dizionario. */
  add(word: string): void {
    const w = word.replace(APOSTROPHES, "'");
    if (!this.words.has(w)) this.words.set(w, '');
    this.cache.clear();
  }

  remove(word: string): void {
    this.words.delete(word.replace(APOSTROPHES, "'"));
    this.cache.clear();
  }

  /** La parola è scritta correttamente? Accetta maiuscole a inizio frase e parole tutte maiuscole. */
  correct(word: string): boolean {
    const w = word.replace(APOSTROPHES, "'");
    let ok = this.cache.get(w);
    if (ok === undefined) {
      ok = this.checkCased(w);
      if (this.cache.size > 20000) this.cache.clear();
      this.cache.set(w, ok);
    }
    return ok;
  }

  /** Suggerimenti di correzione, dal più probabile; al massimo `max`. */
  suggest(word: string, max = 5): string[] {
    const w = word.replace(APOSTROPHES, "'");
    const lower = w.toLowerCase();
    const found = new Map<string, number>();
    const offer = (candidate: string, score: number) => {
      if (candidate === w || candidate === '' || found.has(candidate)) return;
      if (!this.suggestible(candidate)) return;
      found.set(candidate, score);
    };

    // 0. Solo maiuscole/minuscole sbagliate (es. "roma" → "Roma").
    offer(capitalize(lower), 0);
    offer(lower, 0);
    if (found.size) return [...found.keys()].slice(0, max);

    // 1. Errori tipici del dizionario (REP) e accenti (MAP: e/è/é...).
    for (const [from, to] of this.rep) {
      for (let i = lower.indexOf(from); i >= 0; i = lower.indexOf(from, i + 1)) {
        offer(lower.slice(0, i) + to + lower.slice(i + from.length), 1);
      }
    }
    const variants = (s: string) => {
      const out: string[] = [];
      for (let i = 0; i < s.length; i++) {
        const group = this.map.find((g) => g.includes(s[i]!));
        if (!group) continue;
        for (const c of group) if (c !== s[i]) out.push(s.slice(0, i) + c + s.slice(i + 1));
      }
      return out;
    };
    for (const v of variants(lower)) offer(v, 1);
    // Accento scritto con l'apostrofo: "perche'" → "perché".
    if (lower.endsWith("'")) for (const v of variants(lower.slice(0, -1))) offer(v, 1);

    // Apostrofo al posto dello spazio: "qual'è" → "qual è".
    for (let i = lower.indexOf("'"); i > 0 && i < lower.length - 1; i = lower.indexOf("'", i + 1)) {
      offer(`${lower.slice(0, i)} ${lower.slice(i + 1)}`, 1);
    }

    // 2. Una sola modifica: lettere scambiate, mancanti, in più o sbagliate.
    // Una lettera doppia di troppo o un tasto vicino sono gli errori più comuni.
    const chars = this.tryChars;
    for (let i = 0; i < lower.length - 1; i++) {
      offer(lower.slice(0, i) + lower[i + 1] + lower[i] + lower.slice(i + 2), 2);
    }
    for (let i = 0; i < lower.length; i++) {
      offer(lower.slice(0, i) + lower.slice(i + 1), lower[i] === lower[i - 1] ? 2 : 3);
    }
    for (let i = 0; i <= lower.length; i++) {
      for (const c of chars) {
        const cur = lower[i];
        if (cur !== undefined && c !== cur) offer(lower.slice(0, i) + c + lower.slice(i + 1), nearKeys(cur, c) ? 3 : 3.5);
        // Doppia mancante ("erore" → "errore"): errore frequentissimo in italiano.
        offer(lower.slice(0, i) + c + lower.slice(i), c === lower[i - 1] || c === lower[i] ? 2 : 3);
      }
    }

    // 3. Due parole attaccate: "nonlo" → "non lo".
    for (let i = 1; i < lower.length; i++) {
      const a = lower.slice(0, i);
      const b = lower.slice(i);
      const short = (x: string) => x.length === 1 && !'aeioèy'.includes(x);
      if (short(a) || short(b)) continue;
      if (this.suggestible(a) && this.suggestible(b) && !found.has(`${a} ${b}`)) found.set(`${a} ${b}`, 4);
    }

    const keepsFirst = (s: string) => (s[0] === lower[0] ? 0 : 1);
    return [...found.entries()]
      .map(([s, score], i) => ({ s, score, i }))
      .sort((a, b) => a.score - b.score || keepsFirst(a.s) - keepsFirst(b.s) || Math.abs(a.s.length - lower.length) - Math.abs(b.s.length - lower.length) || a.i - b.i)
      .slice(0, max)
      .map(({ s }) => matchCase(s, w));
  }

  // -------------------------------------------------------------------------

  private suggestible(word: string): boolean {
    if (word.includes(' ')) return word.split(' ').every((p) => this.suggestible(p));
    if (!this.checkExact(word)) return false;
    const flags = this.words.get(word);
    return !(flags && this.noSuggest && flags.includes(this.noSuggest));
  }

  private checkCased(w: string): boolean {
    if (this.checkExact(w)) return true;
    switch (wordCase(w)) {
      case 'capitalized':
        return this.checkExact(w.toLowerCase());
      case 'upper': {
        const lower = w.toLowerCase();
        return this.checkExact(lower) || this.checkExact(capitalize(lower));
      }
      case 'mixed': {
        // Elisione a inizio frase: "L'Albero" non è valido, "L'albero" sì.
        const first = w.charAt(0);
        return isUpper(first) && this.checkExact(first.toLowerCase() + w.slice(1));
      }
      default:
        return false;
    }
  }

  private hasFlag(word: string, flag: string): boolean {
    const flags = this.words.get(word);
    return flags !== undefined && flags.includes(flag) && !(this.forbidden && flags.includes(this.forbidden));
  }

  private isRoot(word: string): boolean {
    const flags = this.words.get(word);
    if (flags === undefined) return false;
    if (this.onlyInCompound && flags.includes(this.onlyInCompound)) return false;
    return !(this.forbidden && flags.includes(this.forbidden));
  }

  /** Radici ottenute togliendo un suffisso valido: [radice, regola]. */
  private *stripSuffixes(word: string, crossOnly: boolean): Generator<[string, AffixRule]> {
    const max = Math.min(this.suffixes.maxAdd, word.length);
    for (let n = 0; n <= max; n++) {
      const rules = this.suffixes.byAdd.get(word.slice(word.length - n));
      if (!rules) continue;
      const base = word.slice(0, word.length - n);
      for (const rule of rules) {
        if (crossOnly && !rule.cross) continue;
        const stem = base + rule.strip;
        if (!stem || (rule.cond && !rule.cond.test(stem))) continue;
        yield [stem, rule];
      }
    }
  }

  private checkExact(word: string): boolean {
    if (!word) return false;
    if (this.isRoot(word)) return true;
    for (const [stem, rule] of this.stripSuffixes(word, false)) {
      if (this.hasFlag(stem, rule.flag)) return true;
    }
    const max = Math.min(this.prefixes.maxAdd, word.length);
    for (let n = 0; n <= max; n++) {
      const rules = this.prefixes.byAdd.get(word.slice(0, n));
      if (!rules) continue;
      const rest = word.slice(n);
      for (const rule of rules) {
        const stem = rule.strip + rest;
        if (!stem || (rule.cond && !rule.cond.test(stem))) continue;
        if (this.hasFlag(stem, rule.flag)) return true;
        if (!rule.cross) continue;
        for (const [root, sfx] of this.stripSuffixes(stem, true)) {
          if (this.hasFlag(root, sfx.flag) && this.hasFlag(root, rule.flag)) return true;
        }
      }
    }
    return false;
  }

  private parseAff(aff: string): void {
    const lines = aff.split(/\r?\n/);
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i]!.trim();
      if (!line || line.startsWith('#')) continue;
      const parts = line.split(/\s+/);
      const [key, a = '', b = '', c = ''] = parts;
      switch (key) {
        case 'FLAG':
          if (a === 'long' || a === 'num') throw new Error(`Formato di flag non supportato: ${a}`);
          break;
        case 'TRY': {
          const seen = new Set<string>();
          for (const ch of a.toLowerCase()) seen.add(ch);
          this.tryChars = [...seen].join('');
          break;
        }
        case 'NOSUGGEST':
          this.noSuggest = a;
          break;
        case 'ONLYINCOMPOUND':
          this.onlyInCompound = a;
          break;
        case 'FORBIDDENWORD':
          this.forbidden = a;
          break;
        case 'REP':
          if (b) this.rep.push([a.replace(/_/g, ' '), b.replace(/_/g, ' ')]);
          break;
        case 'MAP':
          if (b === '' && !/^\d+$/.test(a)) this.map.push([...a.replace(/[()]/g, '')]);
          break;
        case 'SFX':
        case 'PFX': {
          // Intestazione: "SFX A Y 12"; regole: "SFX A strip add cond".
          if (parts.length === 4 && /^[YN]$/.test(b) && /^\d+$/.test(c)) {
            const count = Number(c);
            const suffix = key === 'SFX';
            const index = suffix ? this.suffixes : this.prefixes;
            for (let j = 0; j < count && i + 1 < lines.length; j++) {
              const rule = lines[++i]!.trim().split(/\s+/);
              const strip = rule[2] === '0' ? '' : rule[2] ?? '';
              const add = (rule[3] ?? '').split('/')[0]!;
              indexRule(index, {
                flag: a,
                cross: b === 'Y',
                strip,
                add: add === '0' ? '' : add,
                cond: condition(rule[4] ?? '.', suffix),
              });
            }
          }
          break;
        }
        default:
          break;
      }
    }
  }

  private parseDic(dic: string): void {
    const lines = dic.split(/\r?\n/);
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i]!;
      // La prima riga è il numero di voci; il dizionario italiano usa '/' per i commenti.
      if (!line || line.startsWith('/') || line.startsWith('#') || (i === 0 && /^\d+\s*$/.test(line))) continue;
      const entry = line.split(/\s/)[0]!;
      const slash = entry.indexOf('/', 1);
      const word = (slash > 0 ? entry.slice(0, slash) : entry).replace(APOSTROPHES, "'");
      const flags = slash > 0 ? entry.slice(slash + 1) : '';
      const prev = this.words.get(word);
      this.words.set(word, prev ? prev + flags : flags);
    }
  }
}
