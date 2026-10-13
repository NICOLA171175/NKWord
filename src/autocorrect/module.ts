/**
 * Modulo correzione automatica (come in Word): mentre si scrive corregge la
 * parola appena finita (maiuscola a inizio frase, DUe INiziali, errori di
 * battitura frequenti, (c) → ©, -- → —) e trasforma le virgolette diritte
 * in tipografiche.
 *
 * - Ascolta i comandi di digitazione: dopo uno spazio, Tab, Invio o
 *   punteggiatura (, ; ! ?) esegue `autocorrect.apply` sulla parola prima.
 * - Ogni correzione è un passo a sé di Annulla: Ctrl+Z subito dopo riporta
 *   il testo come era stato scritto, come in Word.
 * - `autocorrect.options` apre la finestra delle opzioni.
 * - Gli indirizzi web ed email non si toccano: li gestisce il modulo
 *   collegamenti.
 */

import {
  caret,
  findParagraph,
  isCollapsed,
  normalizeInlines,
  orderedRange,
  textRun,
  updateParagraph,
  paragraphText,
  type AsyncCommand,
  type Command,
  type Editor,
  type EditorModule,
  type EditorState,
  type Inline,
  type Paragraph,
  type Position,
} from '../core';
import { correctWord, smartQuote, type AutoCorrectSettings, type Correction } from './rules';
import { browserStorage, loadSettings, saveSettings, type AutoCorrectStorage } from './settings';

export const AUTOCORRECT_MODULE_ID = 'autocorrect';

export interface AutoCorrectCommandArgs {
  /** Sostituisce [start, end) del paragrafo con `text` (passo di Annulla a sé). */
  'autocorrect.apply': Correction & { readonly paragraphId: string };
  /** Apre la finestra "Opzioni correzione automatica". */
  'autocorrect.options': void;
}

export interface AutoCorrectOptions {
  /** Dove ricordare le opzioni (predefinito: localStorage). */
  readonly storage?: AutoCorrectStorage | null;
  /** Finestra delle opzioni; restituisce le nuove opzioni o null se annullata. */
  readonly openDialog?: (settings: AutoCorrectSettings) => Promise<AutoCorrectSettings | null>;
}

export interface AutoCorrectModule extends EditorModule {
  /** Opzioni attuali. */
  readonly settings: AutoCorrectSettings;
  /** Cambia le opzioni (e le ricorda). */
  setSettings(settings: AutoCorrectSettings): void;
}

/** Caratteri che chiudono una parola e fanno scattare la correzione. */
const TRIGGERS = /[\s ,;!?]$/;

/** Sostituisce [start, end) del paragrafo; il testo nuovo prende il formato del primo carattere sostituito. */
export function replaceInParagraph(p: Paragraph, start: number, end: number, text: string): Paragraph {
  const out: Inline[] = [];
  let pos = 0;
  let inserted = false;
  const insert = (inline: Inline) => {
    if (inserted) return;
    inserted = true;
    if (text) out.push(textRun(text, inline.type === 'text' ? inline.format : {}));
  };
  for (const inline of p.children) {
    const len = inline.type === 'text' ? inline.text.length : 1;
    const a = pos;
    const b = pos + len;
    pos = b;
    if (b <= start || a >= end) {
      if (a >= end && !inserted) insert(inline);
      out.push(inline);
      continue;
    }
    if (inline.type !== 'text') throw new Error('La correzione attraversa un elemento non testuale');
    const before = inline.text.slice(0, Math.max(0, start - a));
    const after = inline.text.slice(Math.min(len, end - a));
    if (before) out.push({ ...inline, text: before });
    insert(inline);
    if (after) out.push({ ...inline, text: after });
  }
  if (!inserted && text) out.push(textRun(text));
  return { ...p, children: normalizeInlines(out) };
}

export function createAutoCorrectModule(options: AutoCorrectOptions = {}): AutoCorrectModule {
  const storage = options.storage === undefined ? browserStorage() : options.storage;
  let settings = loadSettings(storage);

  const apply: Command<AutoCorrectCommandArgs['autocorrect.apply']> = {
    id: 'autocorrect.apply',
    label: 'Correzione automatica',
    isEnabled: ({ state }, args) => !!args && !!findParagraph(state.doc, args.paragraphId),
    run: ({ state }, { paragraphId, start, end, text }) => {
      const p = findParagraph(state.doc, paragraphId)?.paragraph;
      if (!p || start < 0 || end < start || end > paragraphText(p).length) return null;
      const doc = updateParagraph(state.doc, replaceInParagraph(p, start, end, text));
      const delta = text.length - (end - start);
      const map = (pos: Position): Position => {
        if (pos.paragraphId !== paragraphId || pos.offset <= start) return pos;
        return { paragraphId, offset: pos.offset >= end ? pos.offset + delta : start + text.length };
      };
      const selection = { anchor: map(state.selection.anchor), focus: map(state.selection.focus) };
      // Niente mergeKey: la correzione è un passo di Annulla separato dalla digitazione.
      return { state: { ...state, doc, selection }, meta: { label: 'Correzione automatica' } };
    },
  };

  const openOptions: AsyncCommand = {
    id: 'autocorrect.options',
    label: 'Opzioni correzione automatica',
    async: true,
    isEnabled: () => !!options.openDialog,
    run: async () => {
      const next = await options.openDialog?.(settings);
      if (next) module.setSettings(next);
      return null;
    },
  };

  /** Parola appena finita: dove finiva il testo prima del carattere che l'ha chiusa. */
  const correctionAfter = (previous: EditorState, state: EditorState, commandId: string, args: unknown): (Correction & { paragraphId: string }) | null => {
    if (!isCollapsed(state.selection)) return null;
    const at = state.selection.focus;
    if (commandId === 'edit.insertText') {
      const typed = (args as { text?: string } | undefined)?.text ?? '';
      const p = findParagraph(state.doc, at.paragraphId)?.paragraph;
      if (!p || !typed) return null;
      const text = paragraphText(p);
      const last = typed[typed.length - 1]!;
      if (settings.smartQuotes && (last === '"' || last === "'") && text[at.offset - 1] === last) {
        const fix = smartQuote(text, at.offset - 1, settings);
        return fix && { ...fix, paragraphId: p.id };
      }
      if (!TRIGGERS.test(typed) || text[at.offset - 1] !== last) return null;
      const fix = correctWord(text, at.offset - 1, settings);
      return fix && { ...fix, paragraphId: p.id };
    }
    if (commandId === 'edit.insertTab' || commandId === 'edit.insertLineBreak') {
      const p = findParagraph(state.doc, at.paragraphId)?.paragraph;
      if (!p || at.offset < 1) return null;
      const fix = correctWord(paragraphText(p), at.offset - 1, settings);
      return fix && { ...fix, paragraphId: p.id };
    }
    if (commandId === 'edit.splitParagraph') {
      // La parola è alla fine del paragrafo di prima (che conserva l'id).
      const end = isCollapsed(previous.selection) ? previous.selection.focus : orderedRange(previous.doc, previous.selection)[0];
      const p = findParagraph(state.doc, end.paragraphId)?.paragraph;
      if (!p || end.paragraphId === at.paragraphId) return null;
      const text = paragraphText(p);
      if (end.offset !== text.length) return null;
      const fix = correctWord(text, end.offset, settings);
      return fix && { ...fix, paragraphId: p.id };
    }
    return null;
  };

  const module: AutoCorrectModule = {
    id: AUTOCORRECT_MODULE_ID,
    get settings() {
      return settings;
    },
    setSettings(next) {
      settings = next;
      saveSettings(storage, settings);
    },
    activate(editor: Editor) {
      // Ultimo cambiamento visto: serve lo stato di prima per sapere dove finiva la parola.
      let last: { previous: EditorState; state: EditorState } | null = null;
      const disposers = [
        editor.commands.register(apply),
        editor.commands.register(openOptions),
        editor.events.on('change', ({ previous, state }) => {
          last = { previous, state };
        }),
        editor.events.on('executed', ({ commandId, args }) => {
          const seen = last;
          last = null;
          // Solo se lo stato è ancora quello prodotto dal comando (nessun altro modulo l'ha già cambiato).
          if (!seen || seen.state !== editor.state || seen.state.doc === seen.previous.doc) return;
          const fix = correctionAfter(seen.previous, seen.state, commandId, args);
          if (fix) editor.execute('autocorrect.apply', fix);
        }),
      ];
      return () => disposers.forEach((d) => d());
    },
  };
  return module;
}
