/**
 * Modulo Dettatura, come Home > Detta in Word: si parla e il testo viene
 * scritto nel punto del cursore, con i comandi vocali di punteggiatura
 * ("punto", "virgola", "a capo"...). Si ferma con un altro clic su Detta
 * o con Esc.
 *
 * - `dictation.toggle` avvia o ferma la dettatura (vedi engine.ts per come
 *   funziona nel browser e nel programma installato).
 * - `dictation.stop` la ferma.
 */

import { findParagraph, isCollapsed, orderedRange, paragraphText, type AsyncCommand, type Editor, type EditorModule } from '../core';
import { addViewExtension } from '../ui/extensions';
import { detectEnvironment, dictationMode, errorMessage, type DictationEnvironment, type DictationMode, type SpeechRecognitionLike } from './engine';
import { dictationPieces } from './text';

export const DICTATION_MODULE_ID = 'dictation';
export const DICTATION_LANG = 'it-IT';
/** Riavvii di fila senza nessun risultato, oltre i quali ci si ferma. */
const MAX_EMPTY_RESTARTS = 8;

export interface DictationCommandArgs {
  'dictation.toggle': void;
  'dictation.stop': void;
}

export interface DictationMessage {
  readonly text: string;
  readonly kind: 'info' | 'error';
}

export interface DictationModule extends EditorModule {
  readonly mode: DictationMode;
  /** In ascolto con il riconoscimento del browser. */
  readonly listening: boolean;
  /** Ciò che si sta dicendo, non ancora definitivo. */
  readonly interim: string;
  /** Avviso da mostrare (es. microfono non consentito, come usare Win+H). */
  readonly message: DictationMessage | null;
  toggle(): Promise<void>;
  start(): Promise<void>;
  stop(): void;
  dismissMessage(): void;
  /** Inserisce una frase riconosciuta nel punto del cursore. */
  insert(spoken: string): void;
  /** Riporta il focus al documento (lo imposta la vista). */
  focusDocument?: () => void;
  onChange(listener: () => void): () => void;
}

export function createDictationModule(env: DictationEnvironment = detectEnvironment()): DictationModule {
  const mode = dictationMode(env);
  let editor: Editor | undefined;
  let recognition: SpeechRecognitionLike | null = null;
  let listening = false;
  let interim = '';
  let message: DictationMessage | null = null;
  let emptyRestarts = 0;
  const listeners = new Set<() => void>();
  const notify = () => [...listeners].forEach((l) => l());

  const setMessage = (next: DictationMessage | null) => {
    message = next;
    notify();
  };

  const startWeb = () => {
    const Ctor = env.recognition;
    if (!Ctor) return;
    const rec = new Ctor();
    rec.lang = DICTATION_LANG;
    rec.continuous = true;
    rec.interimResults = true;
    rec.onresult = (event) => {
      let pending = '';
      for (let i = event.resultIndex; i < event.results.length; i++) {
        const result = event.results[i]!;
        const transcript = result[0].transcript;
        if (result.isFinal) {
          emptyRestarts = 0;
          module.insert(transcript);
        } else pending += transcript;
      }
      interim = pending.trim();
      notify();
    };
    rec.onerror = (event) => {
      const text = errorMessage(event.error, env.windows);
      if (!text) return;
      listening = false;
      setMessage({ text, kind: 'error' });
    };
    rec.onend = () => {
      if (recognition !== rec) return;
      // Il browser smette da solo dopo un po' di silenzio: come in Word, si
      // continua ad ascoltare finché non lo si ferma.
      if (listening && emptyRestarts < MAX_EMPTY_RESTARTS) {
        emptyRestarts += 1;
        try {
          rec.start();
          return;
        } catch {
          // Si ferma qui sotto.
        }
      }
      recognition = null;
      listening = false;
      interim = '';
      notify();
    };
    recognition = rec;
    listening = true;
    emptyRestarts = 0;
    message = null;
    try {
      rec.start();
    } catch {
      recognition = null;
      listening = false;
      message = { text: errorMessage('', env.windows)!, kind: 'error' };
    }
    notify();
  };

  const module: DictationModule = {
    id: DICTATION_MODULE_ID,
    mode,
    get listening() {
      return listening;
    },
    get interim() {
      return interim;
    },
    get message() {
      return message;
    },
    async toggle() {
      if (listening) module.stop();
      else await module.start();
    },
    async start() {
      if (mode === 'web') {
        if (!listening) startWeb();
        return;
      }
      if (mode === 'windows') {
        module.focusDocument?.();
        setMessage({ text: 'Si apre la dettatura di Windows: parla pure, il testo va nel punto del cursore. Per fermarla premi di nuovo Detta o Win+H.', kind: 'info' });
        try {
          await env.desktop?.startSystemDictation?.();
          module.focusDocument?.();
        } catch {
          setMessage({ text: 'Non sono riuscito ad aprire la dettatura di Windows: prova a premere Win+H.', kind: 'error' });
        }
        return;
      }
      if (mode === 'mac') {
        module.focusDocument?.();
        setMessage({ text: 'Sul Mac premi due volte il tasto Fn (o il tasto con il microfono) per dettare con la Dettatura del Mac.', kind: 'info' });
        return;
      }
      setMessage({
        text: `Questo browser non ha il riconoscimento vocale: apri NKoffice con Chrome o Edge.${env.windows ? ' Oppure premi Win+H per la dettatura di Windows.' : ''}`,
        kind: 'error',
      });
    },
    stop() {
      if (!listening) return;
      listening = false;
      interim = '';
      // stop() lascia arrivare l'ultima frase già detta.
      recognition?.stop();
      notify();
    },
    dismissMessage() {
      if (message) setMessage(null);
    },
    insert(spoken) {
      if (!editor || !spoken.trim()) return;
      const { doc, selection } = editor.state;
      const start = isCollapsed(selection) ? selection.focus : orderedRange(doc, selection)[0];
      const p = findParagraph(doc, start.paragraphId)?.paragraph;
      const before = p ? paragraphText(p).slice(0, start.offset) : '';
      for (const piece of dictationPieces(spoken, before)) {
        if (piece.kind === 'paragraph') editor.execute('edit.splitParagraph');
        else editor.execute('edit.insertText', { text: piece.text });
      }
    },
    onChange(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    activate(ed) {
      editor = ed;
      const toggle: AsyncCommand<void> = {
        id: 'dictation.toggle',
        label: 'Detta',
        async: true,
        run: async () => {
          await module.toggle();
          return null;
        },
      };
      const stop: AsyncCommand<void> = {
        id: 'dictation.stop',
        label: 'Ferma dettatura',
        async: true,
        isEnabled: () => listening,
        run: async () => {
          module.stop();
          return null;
        },
      };
      const offs = [
        ed.commands.register(toggle),
        ed.commands.register(stop),
        addViewExtension(ed, {
          id: DICTATION_MODULE_ID,
          isActive: (commandId) => commandId === 'dictation.toggle' && listening,
          onChange: (l) => module.onChange(l),
        }),
      ];
      return () => {
        module.stop();
        recognition?.abort();
        recognition = null;
        offs.forEach((off) => off());
        editor = undefined;
      };
    },
  };
  return module;
}
