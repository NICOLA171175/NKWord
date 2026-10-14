/**
 * Modulo Lettura ad alta voce, come Revisione > Leggi ad alta voce in Word:
 * legge il documento dal cursore in poi (o solo la selezione) con la sintesi
 * vocale del browser o del computer, evidenziando la parola che sta leggendo.
 * Non modifica il documento: alla fine il cursore resta dove si è fermato.
 *
 * - `readaloud.toggle` avvia la lettura o la ferma (Ctrl+Alt+Spazio).
 * - `readaloud.playPause` mette in pausa e riprende.
 * - `readaloud.next` / `readaloud.previous` saltano al paragrafo dopo o prima.
 * - `readaloud.stop` ferma (anche Esc).
 * - `readaloud.rate` e `readaloud.voice` cambiano velocità e voce.
 *
 * Ogni paragrafo è una frase da pronunciare a sé: così Avanti e Indietro sono
 * immediati e i browser non si interrompono sui testi lunghi. La pausa ferma
 * la voce e ricorda la parola: Riprendi riparte da lì (la pausa vera dei
 * browser non è affidabile).
 */

import { caret, findParagraph, firstParagraph, type AsyncCommand, type Editor, type EditorModule, type Id, type Selection } from '../core';
import { addViewExtension } from '../ui/extensions';
import { clampRate, DEFAULT_RATE, pickVoice, readingChunks, wordAt, type ReadingChunk, type VoiceLike } from './text';

export const READ_ALOUD_MODULE_ID = 'readaloud';
/** Etichetta delle transazioni del modulo (solo spostamenti del cursore). */
const META_LABEL = 'readaloud';
/** Colore dell'evidenziazione della parola letta, come il grigio-azzurro di Word. */
export const WORD_HIGHLIGHT = 'rgba(24, 90, 189, 0.28)';

export type ReadAloudStatus = 'stopped' | 'playing' | 'paused';

export interface ReadAloudCommandArgs {
  'readaloud.toggle': void;
  'readaloud.playPause': void;
  'readaloud.next': void;
  'readaloud.previous': void;
  'readaloud.stop': void;
  'readaloud.rate': { rate: number };
  'readaloud.voice': { name: string };
}

/** Ciò che serve della frase da pronunciare (`SpeechSynthesisUtterance`). */
export interface UtteranceLike {
  text: string;
  lang: string;
  rate: number;
  voice: VoiceLike | null;
  onboundary: ((event: { name?: string; charIndex: number }) => void) | null;
  onend: (() => void) | null;
  onerror: ((event: { error?: string }) => void) | null;
}

/** Ciò che serve della sintesi vocale (`window.speechSynthesis`). */
export interface SpeechLike {
  speak(utterance: UtteranceLike): void;
  cancel(): void;
  getVoices(): readonly VoiceLike[];
  addEventListener?(type: 'voiceschanged', listener: () => void): void;
  removeEventListener?(type: 'voiceschanged', listener: () => void): void;
}

export interface ReadAloudEngine {
  readonly speech: SpeechLike;
  createUtterance(text: string): UtteranceLike;
}

/** Dove si ricordano velocità e voce (predefinito: localStorage). */
export interface ReadAloudStorage {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
}

export interface ReadAloudOptions {
  /** Sintesi vocale; predefinita quella del browser. null: non disponibile. */
  readonly engine?: ReadAloudEngine | null;
  readonly storage?: ReadAloudStorage | null;
}

export interface ReadAloudModule extends EditorModule {
  readonly status: ReadAloudStatus;
  /** false se il browser non sa leggere ad alta voce. */
  readonly supported: boolean;
  readonly rate: number;
  /** Nome della voce scelta (null: scelta automatica). */
  readonly voiceName: string | null;
  /** La voce che verrà usata davvero. */
  readonly voice: VoiceLike | null;
  /** Parola che si sta leggendo, se c'è. */
  readonly word: { readonly paragraphId: Id; readonly start: number; readonly end: number } | null;
  /** Ultimo problema da mostrare all'utente (null se nessuno). */
  readonly message: string | null;
  voices(): readonly VoiceLike[];
  start(): void;
  stop(): void;
  toggle(): void;
  pause(): void;
  resume(): void;
  playPause(): void;
  next(): void;
  previous(): void;
  setRate(rate: number): void;
  setVoice(name: string | null): void;
  onChange(listener: () => void): () => void;
}

const STORAGE_KEY = 'nkoffice.readaloud';

function browserEngine(): ReadAloudEngine | null {
  const g = globalThis as { speechSynthesis?: SpeechLike; SpeechSynthesisUtterance?: new (text: string) => UtteranceLike };
  if (!g.speechSynthesis || !g.SpeechSynthesisUtterance) return null;
  const Utterance = g.SpeechSynthesisUtterance;
  return { speech: g.speechSynthesis, createUtterance: (text) => new Utterance(text) };
}

function browserStorage(): ReadAloudStorage | null {
  try {
    return globalThis.localStorage ?? null;
  } catch {
    return null;
  }
}

export function createReadAloudModule(options: ReadAloudOptions = {}): ReadAloudModule {
  const engine = options.engine === undefined ? browserEngine() : options.engine;
  const storage = options.storage === undefined ? browserStorage() : options.storage;

  let rate = DEFAULT_RATE;
  let voiceName: string | null = null;
  try {
    const saved = JSON.parse(storage?.getItem(STORAGE_KEY) ?? 'null') as { rate?: number; voice?: string } | null;
    if (typeof saved?.rate === 'number') rate = clampRate(saved.rate);
    if (typeof saved?.voice === 'string') voiceName = saved.voice;
  } catch {
    // Impostazioni illeggibili: si usano quelle normali.
  }
  const save = () => {
    try {
      storage?.setItem(STORAGE_KEY, JSON.stringify({ rate, ...(voiceName ? { voice: voiceName } : {}) }));
    } catch {
      // Spazio pieno o archivio bloccato: le impostazioni valgono fino alla chiusura.
    }
  };

  let editor: Editor | undefined;
  let status: ReadAloudStatus = 'stopped';
  let chunks: ReadingChunk[] = [];
  let index = 0;
  /** Da dove riprendere nel pezzo attuale (inizio della parola letta). */
  let resumeAt = 0;
  let word: ReadAloudModule['word'] = null;
  let message: string | null = null;
  /** Frase in corso: gli eventi di quelle annullate si ignorano. */
  let current: UtteranceLike | null = null;
  /** Cursore quando si è messo in pausa: se l'utente lo sposta, si riprende da lì. */
  let pausedSelection: Selection | null = null;
  const listeners = new Set<() => void>();
  const notify = () => [...listeners].forEach((l) => l());

  const voices = () => engine?.speech.getVoices() ?? [];
  const voice = () => pickVoice(voices(), voiceName);

  /** Sposta il cursore (senza Annulla) all'inizio della parola letta. */
  const placeCaret = (paragraphId: Id, offset: number) => {
    if (!editor || !findParagraph(editor.state.doc, paragraphId)) return;
    const sel = editor.state.selection;
    if (sel.anchor.paragraphId === paragraphId && sel.anchor.offset === offset && sel.focus.offset === offset && sel.focus.paragraphId === paragraphId) return;
    editor.apply({ state: { ...editor.state, selection: caret(paragraphId, offset) }, meta: { addToHistory: false, label: META_LABEL } });
  };

  const silence = () => {
    const was = current;
    current = null;
    if (was) engine?.speech.cancel();
  };

  const setWord = (chunk: ReadingChunk, at: number) => {
    const w = wordAt(chunk.text, at, chunk.end);
    if (!w) return;
    word = { paragraphId: chunk.paragraphId, start: w[0], end: w[1] };
    resumeAt = w[0];
    notify();
  };

  const finish = () => {
    silence();
    const last = word;
    status = 'stopped';
    word = null;
    chunks = [];
    pausedSelection = null;
    if (last) placeCaret(last.paragraphId, last.start);
    notify();
  };

  /** Legge il pezzo `i` dalla posizione `from` (dentro il pezzo). */
  const speak = (i: number, from?: number) => {
    silence();
    const chunk = chunks[i];
    if (!engine || !chunk) return finish();
    index = i;
    const start = Math.max(chunk.start, Math.min(from ?? chunk.start, chunk.end));
    if (!wordAt(chunk.text, start, chunk.end)) return speak(i + 1);
    status = 'playing';
    // Si evidenzia subito la prima parola: alcune voci non dicono dove sono.
    setWord(chunk, start);
    const u = engine.createUtterance(chunk.text.slice(start, chunk.end));
    const v = voice();
    u.lang = v?.lang || 'it-IT';
    if (v) u.voice = v;
    u.rate = rate;
    u.onboundary = (e) => {
      if (current !== u || (e.name && e.name !== 'word')) return;
      setWord(chunk, start + e.charIndex);
    };
    u.onend = () => {
      if (current !== u) return;
      current = null;
      speak(index + 1);
    };
    u.onerror = (e) => {
      if (current !== u || e.error === 'interrupted' || e.error === 'canceled') return;
      message =
        e.error === 'synthesis-unavailable' || e.error === 'voice-unavailable' || e.error === 'language-unavailable'
          ? 'Nessuna voce disponibile per leggere il testo. Installa una voce italiana nelle impostazioni del computer.'
          : 'La lettura si è interrotta per un problema della voce.';
      finish();
    };
    current = u;
    engine.speech.speak(u);
    notify();
  };

  const startFrom = (selection: Selection) => {
    if (!editor) return;
    message = null;
    if (!engine) {
      message = 'Questo browser non sa leggere ad alta voce. Prova con Chrome o Edge, oppure con il programma NKWord.';
      notify();
      return;
    }
    const { doc } = editor.state;
    chunks = readingChunks(doc, selection);
    // Cursore in fondo al documento (o in un'intestazione): si legge da capo.
    if (!chunks.length) chunks = readingChunks(doc, caret(firstParagraph(doc).id, 0));
    if (!chunks.length) {
      message = 'Il documento non contiene testo da leggere.';
      status = 'stopped';
      notify();
      return;
    }
    pausedSelection = null;
    speak(0);
  };

  const module: ReadAloudModule = {
    id: READ_ALOUD_MODULE_ID,
    get status() {
      return status;
    },
    get supported() {
      return !!engine;
    },
    get rate() {
      return rate;
    },
    get voiceName() {
      return voiceName;
    },
    get voice() {
      return voice();
    },
    get word() {
      return word;
    },
    get message() {
      return message;
    },
    voices,
    start() {
      if (editor) startFrom(editor.state.selection);
    },
    stop() {
      if (status !== 'stopped') finish();
    },
    toggle() {
      if (status === 'stopped') module.start();
      else module.stop();
    },
    pause() {
      if (status !== 'playing') return;
      silence();
      status = 'paused';
      if (word) placeCaret(word.paragraphId, word.start);
      pausedSelection = editor?.state.selection ?? null;
      notify();
    },
    resume() {
      if (status !== 'paused' || !editor) return;
      const sel = editor.state.selection;
      // Il cursore è stato spostato durante la pausa: si legge da lì, come in Word.
      if (pausedSelection && sel !== pausedSelection) return startFrom(sel);
      speak(index, resumeAt);
    },
    playPause() {
      if (status === 'playing') module.pause();
      else if (status === 'paused') module.resume();
      else module.start();
    },
    next() {
      if (status === 'stopped') return;
      if (index + 1 >= chunks.length) return finish();
      if (status === 'playing') speak(index + 1);
      else {
        index += 1;
        moveWhilePaused();
      }
    },
    previous() {
      if (status === 'stopped') return;
      // Come in Word: a metà paragrafo si torna al suo inizio, all'inizio al precedente.
      const chunk = chunks[index]!;
      const atStart = !word || word.start <= chunk.start;
      const target = atStart ? Math.max(0, index - 1) : index;
      if (status === 'playing') speak(target);
      else {
        index = target;
        moveWhilePaused();
      }
    },
    setRate(value) {
      const next = clampRate(value);
      if (next === rate) return;
      rate = next;
      save();
      if (status === 'playing') speak(index, resumeAt);
      else notify();
    },
    setVoice(name) {
      if (name === voiceName) return;
      voiceName = name;
      save();
      if (status === 'playing') speak(index, resumeAt);
      else notify();
    },
    onChange(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    activate(ed) {
      editor = ed;
      const command = <A>(id: string, label: string, run: (args: A) => void, isEnabled?: (args: A) => boolean): AsyncCommand<A> => ({
        id,
        label,
        async: true,
        ...(isEnabled ? { isEnabled: (_ctx, args: A) => isEnabled(args) } : {}),
        async run(_ctx, args) {
          run(args);
          return null;
        },
      });
      const active = () => status !== 'stopped';
      const commands: AsyncCommand<any>[] = [
        command<void>('readaloud.toggle', 'Leggi ad alta voce', () => module.toggle()),
        command<void>('readaloud.playPause', 'Riproduci o pausa', () => module.playPause()),
        command<void>('readaloud.next', 'Paragrafo successivo', () => module.next(), active),
        command<void>('readaloud.previous', 'Paragrafo precedente', () => module.previous(), active),
        command<void>('readaloud.stop', 'Interrompi lettura', () => module.stop(), active),
        command<{ rate: number }>('readaloud.rate', 'Velocità di lettura', (a) => module.setRate(a.rate), (a) => typeof a?.rate === 'number'),
        command<{ name: string }>('readaloud.voice', 'Voce', (a) => module.setVoice(a.name || null), (a) => typeof a?.name === 'string'),
      ];
      const offs = commands.map((c) => ed.commands.register(c));
      offs.push(
        addViewExtension(ed, {
          id: READ_ALOUD_MODULE_ID,
          marks: () =>
            word && status !== 'stopped'
              ? [{ ...word, kind: 'highlight', color: WORD_HIGHLIGHT, className: 'readaloud-word' }]
              : [],
          isActive: (commandId) => commandId === 'readaloud.toggle' && status !== 'stopped',
          onChange: (listener) => module.onChange(listener),
          keybindings: [
            { key: 'Ctrl-Alt-Space', commandId: 'readaloud.toggle' },
            { key: 'Escape', commandId: 'readaloud.stop' },
          ],
        }),
        ed.events.on('change', ({ state, previous, meta }) => {
          if (status === 'stopped' || meta.label === META_LABEL) return;
          // Il testo è cambiato: le posizioni non valgono più, si smette di leggere.
          if (state.doc !== previous.doc) return finish();
          // Clic altrove mentre legge: si riparte da lì, come in Word.
          if (status === 'playing' && state.selection !== previous.selection) startFrom(state.selection);
        }),
      );
      const onVoices = () => notify();
      engine?.speech.addEventListener?.('voiceschanged', onVoices);
      offs.push(() => engine?.speech.removeEventListener?.('voiceschanged', onVoices));
      return () => {
        silence();
        status = 'stopped';
        word = null;
        offs.forEach((off) => off());
        editor = undefined;
      };
    },
  };

  /** Avanti o Indietro durante la pausa: si evidenzia l'inizio del nuovo paragrafo. */
  function moveWhilePaused() {
    const chunk = chunks[index]!;
    setWord(chunk, chunk.start);
    placeCaret(chunk.paragraphId, resumeAt);
    pausedSelection = editor?.state.selection ?? null;
  }

  return module;
}
