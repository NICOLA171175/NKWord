/**
 * Stato del controllo ortografico: lingua, dizionari caricati, parole
 * ignorate e dizionario personale. Non conosce né editor né DOM; trova gli
 * errori in un paragrafo e li tiene in cache finché il paragrafo (immutabile)
 * non cambia.
 */

import { Emitter, paragraphText, type DocumentModel, type Id, type Paragraph, paragraphs } from '../core';
import { DEFAULT_LANGUAGE, LANGUAGES, type Language } from './dictionaries';
import { Hunspell } from './hunspell';
import { findWords } from './words';

export interface Misspelling {
  readonly paragraphId: Id;
  readonly start: number;
  readonly end: number;
  readonly word: string;
}

export interface SpellSettings {
  /** Sottolinea gli errori mentre si scrive. */
  readonly enabled: boolean;
  readonly language: string;
  /** Salta le parole TUTTE MAIUSCOLE. */
  readonly ignoreUppercase: boolean;
}

/** Dove salvare impostazioni e dizionario personale (default: localStorage). */
export interface SpellStorage {
  read(): string | null;
  write(value: string): void;
}

export interface SpellCheckerOptions {
  readonly languages?: readonly Language[];
  /** `null` per non salvare nulla. */
  readonly storage?: SpellStorage | null;
  readonly settings?: Partial<SpellSettings>;
}

interface Saved {
  enabled?: boolean;
  language?: string;
  ignoreUppercase?: boolean;
  words?: string[];
}

const STORAGE_KEY = 'nikoffice.spellcheck';

export const localStorageSpellStorage: SpellStorage = {
  read() {
    try {
      return globalThis.localStorage?.getItem(STORAGE_KEY) ?? null;
    } catch {
      return null;
    }
  },
  write(value) {
    try {
      globalThis.localStorage?.setItem(STORAGE_KEY, value);
    } catch {
      // Archiviazione non disponibile (finestra privata...): si perde solo il salvataggio.
    }
  },
};

const normalize = (word: string) => word.replace(/[’ʼ]/g, "'");

export class SpellChecker {
  readonly events = new Emitter<{ change: void }>();
  readonly languages: readonly Language[];
  private _settings: SpellSettings;
  private readonly storage: SpellStorage | null;
  private readonly dictionaries = new Map<string, Hunspell>();
  private readonly loading = new Map<string, Promise<Hunspell | null>>();
  private readonly ignored = new Set<string>();
  private readonly personal = new Set<string>();
  private cache = new WeakMap<Paragraph, Misspelling[]>();

  constructor(options: SpellCheckerOptions = {}) {
    this.languages = options.languages ?? LANGUAGES;
    this.storage = options.storage === undefined ? localStorageSpellStorage : options.storage;
    let saved: Saved = {};
    try {
      saved = JSON.parse(this.storage?.read() ?? '{}') as Saved;
    } catch {
      saved = {};
    }
    for (const w of saved.words ?? []) if (typeof w === 'string') this.personal.add(normalize(w));
    const language = options.settings?.language ?? saved.language;
    this._settings = {
      enabled: options.settings?.enabled ?? saved.enabled ?? true,
      language: this.languages.some((l) => l.id === language) ? language! : (this.languages.find((l) => l.id === DEFAULT_LANGUAGE) ?? this.languages[0])?.id ?? DEFAULT_LANGUAGE,
      ignoreUppercase: options.settings?.ignoreUppercase ?? saved.ignoreUppercase ?? true,
    };
  }

  get settings(): SpellSettings {
    return this._settings;
  }

  /** Il dizionario della lingua corrente è caricato. */
  get ready(): boolean {
    return this.dictionaries.has(this._settings.language);
  }

  /** Parole aggiunte dall'utente, in ordine alfabetico. */
  get personalWords(): string[] {
    return [...this.personal].sort((a, b) => a.localeCompare(b));
  }

  /** Carica (una volta sola) il dizionario della lingua corrente. */
  async load(languageId = this._settings.language): Promise<boolean> {
    if (this.dictionaries.has(languageId)) return true;
    let pending = this.loading.get(languageId);
    if (!pending) {
      const lang = this.languages.find((l) => l.id === languageId);
      pending = lang
        ? lang.load().then(
            (data) => new Hunspell(data),
            () => null,
          )
        : Promise.resolve(null);
      this.loading.set(languageId, pending);
    }
    const dict = await pending;
    this.loading.delete(languageId);
    if (!dict) return false;
    this.dictionaries.set(languageId, dict);
    if (languageId === this._settings.language) this.changed();
    return true;
  }

  update(settings: Partial<SpellSettings>): void {
    const next = { ...this._settings, ...settings };
    if (!this.languages.some((l) => l.id === next.language)) return;
    const same = (Object.keys(next) as (keyof SpellSettings)[]).every((k) => next[k] === this._settings[k]);
    if (same) return;
    this._settings = next;
    this.save();
    this.changed();
    if (next.enabled) void this.load();
  }

  /** "Ignora tutto": non segnala più la parola fino alla chiusura dell'app. */
  ignore(word: string): void {
    this.ignored.add(normalize(word));
    this.changed();
  }

  /** "Aggiungi al dizionario": vale per tutte le lingue e resta salvata. */
  addWord(word: string): void {
    const w = normalize(word);
    if (this.personal.has(w)) return;
    this.personal.add(w);
    this.save();
    this.changed();
  }

  removeWord(word: string): void {
    if (!this.personal.delete(normalize(word))) return;
    this.save();
    this.changed();
  }

  /** Parola corretta? Se il dizionario non è ancora pronto non segnala nulla. */
  isCorrect(word: string): boolean {
    const dict = this.dictionaries.get(this._settings.language);
    if (!dict) return true;
    const w = normalize(word);
    if (this.ignored.has(w) || this.personal.has(w) || this.personal.has(w.toLowerCase())) return true;
    return dict.correct(w);
  }

  suggest(word: string, max = 5): string[] {
    return this.dictionaries.get(this._settings.language)?.suggest(word, max) ?? [];
  }

  /** Errori di un paragrafo (vuoto se il controllo è spento o il dizionario non è pronto). */
  misspellingsIn(p: Paragraph): readonly Misspelling[] {
    if (!this._settings.enabled || !this.ready) return [];
    let found = this.cache.get(p);
    if (!found) {
      found = [];
      for (const { start, end, word } of findWords(paragraphText(p), { ignoreUppercase: this._settings.ignoreUppercase })) {
        if (this.isCorrect(word)) continue;
        // Apostrofo finale usato come virgoletta: 'ciao' → si controlla "ciao".
        if (/['’]$/.test(word)) {
          const bare = word.slice(0, -1);
          if (this.isCorrect(bare)) continue;
          found.push({ paragraphId: p.id, start, end: end - 1, word: bare });
        } else found.push({ paragraphId: p.id, start, end, word });
      }
      this.cache.set(p, found);
    }
    return found;
  }

  /** Errori del corpo del documento, in ordine di lettura. */
  misspellings(doc: DocumentModel): Misspelling[] {
    const out: Misspelling[] = [];
    for (const { paragraph } of paragraphs(doc)) out.push(...this.misspellingsIn(paragraph));
    return out;
  }

  private changed(): void {
    this.cache = new WeakMap();
    this.events.emit('change', undefined);
  }

  private save(): void {
    const { enabled, language, ignoreUppercase } = this._settings;
    const saved: Saved = { enabled, language, ignoreUppercase, words: this.personalWords };
    this.storage?.write(JSON.stringify(saved));
  }
}
