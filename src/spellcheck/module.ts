/**
 * Modulo controllo ortografico: sottolinea in rosso le parole sbagliate,
 * propone correzioni col clic destro e aggiunge alla scheda Revisione i
 * pulsanti Ortografia (F7), Mostra errori e Lingua. Italiano predefinito.
 *
 * I comandi `spelling.*` non sono nel catalogo del core: la UI li collega
 * tramite `spellcheckRibbonItems` e l'estensione della vista pubblicata in
 * `activate` (segni, menu contestuale, F7, stato dei pulsanti).
 */

import {
  caret,
  comparePositions,
  findParagraph,
  isCollapsed,
  normalizeInlines,
  orderedRange,
  textRun,
  updateParagraph,
  type AsyncCommand,
  type Command,
  type Editor,
  type EditorModule,
  type EditorState,
  type Id,
  type Inline,
  type Paragraph,
  type Position,
  type RibbonItem,
  type Transaction,
} from '../core';
import { addViewExtension, type ContextMenuItem, type TextMark, type ViewExtension } from '../ui/extensions';
import { SpellChecker, type Misspelling, type SpellCheckerOptions } from './checker';
import { LANGUAGES } from './dictionaries';

/** Argomenti dei comandi del modulo. */
export interface SpellingCommandArgs {
  /** Va al prossimo errore dopo il cursore e lo seleziona (F7). */
  'spelling.next': void;
  /** Accende/spegne la sottolineatura degli errori. */
  'spelling.toggle': void;
  'spelling.setLanguage': { language: string };
  /** Sostituisce il testo di un paragrafo (es. con un suggerimento). */
  'spelling.replace': { paragraphId: Id; start: number; end: number; text: string };
  'spelling.ignoreAll': { word: string };
  'spelling.addToDictionary': { word: string };
}

export const SPELLCHECK_MODULE_ID = 'spellcheck';

/** Voci della scheda Revisione, raccolte dalla UI tramite l'esportazione `ribbonItems` di index.ts. */
export const spellcheckRibbonItems: readonly (RibbonItem & { readonly label: string; readonly options?: readonly { label: string; args: unknown }[] })[] = [
  { tab: 'review', group: 'Strumenti di correzione', commandId: 'spelling.next', label: 'Ortografia', icon: 'ABC✓', kind: 'button' },
  { tab: 'review', group: 'Strumenti di correzione', commandId: 'spelling.toggle', label: 'Mostra errori', icon: '〰', kind: 'toggle' },
  {
    tab: 'review',
    group: 'Lingua',
    commandId: 'spelling.setLanguage',
    label: 'Lingua',
    kind: 'dropdown',
    options: LANGUAGES.map((l) => ({ label: l.label, args: { language: l.id } })),
  },
];
export const SPELLING_ERROR_COLOR = '#e81123';

export interface SpellcheckModuleOptions extends SpellCheckerOptions {
  readonly checker?: SpellChecker;
  /** Messaggio quando non ci sono altri errori (default `window.alert`). */
  readonly notify?: (message: string) => void;
}

export interface SpellcheckModule extends EditorModule {
  readonly checker: SpellChecker;
}

/** Sostituisce [start, end) del paragrafo con `text`, con il formato del primo carattere sostituito. */
export function replaceInParagraph(p: Paragraph, start: number, end: number, text: string): Paragraph {
  const out: Inline[] = [];
  let pos = 0;
  let inserted = false;
  const insert = (format: Inline) => {
    if (inserted) return;
    inserted = true;
    if (text) out.push(textRun(text, format.type === 'text' ? format.format : {}));
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
    if (inline.type !== 'text') throw new Error('La sostituzione attraversa un elemento non testuale');
    const before = inline.text.slice(0, Math.max(0, start - a));
    const after = inline.text.slice(Math.min(len, end - a));
    if (before) out.push({ ...inline, text: before });
    insert(inline);
    if (after) out.push({ ...inline, text: after });
  }
  if (!inserted && text) out.push(textRun(text));
  return { ...p, children: normalizeInlines(out) };
}

function rangeText(p: Paragraph, start: number, end: number): string | null {
  let pos = 0;
  let s = '';
  for (const inline of p.children) {
    const len = inline.type === 'text' ? inline.text.length : 1;
    if (pos + len > start && pos < end) {
      if (inline.type !== 'text') return null;
      s += inline.text.slice(Math.max(0, start - pos), end - pos);
    }
    pos += len;
  }
  return s;
}

const contains = (m: Misspelling, pos: Position) => m.paragraphId === pos.paragraphId && m.start <= pos.offset && pos.offset <= m.end;

export function createSpellcheckModule(options: SpellcheckModuleOptions = {}): SpellcheckModule {
  const checker = options.checker ?? new SpellChecker(options);
  const notify = options.notify ?? ((m: string) => globalThis.alert?.(m));
  /** Fine della parola che si sta scrivendo: non la si segnala finché non la si lascia, come in Word. */
  let typing: Position | null = null;

  const next: AsyncCommand<void> = {
    id: 'spelling.next',
    label: 'Ortografia',
    async: true,
    run: async ({ state }) => {
      await checker.load();
      const { doc, selection } = state;
      const list = checker.misspellings(doc);
      if (!list.length) {
        notify('Controllo ortografia completato. Nessun errore trovato.');
        return null;
      }
      // Il primo errore che inizia dopo la selezione (o il cursore); poi si ricomincia dall'inizio.
      const [, to] = orderedRange(doc, selection);
      const found = list.find((m) => comparePositions(doc, { paragraphId: m.paragraphId, offset: m.start }, to) >= 0) ?? list[0]!;
      return {
        state: { ...state, selection: { anchor: { paragraphId: found.paragraphId, offset: found.start }, focus: { paragraphId: found.paragraphId, offset: found.end } } },
        meta: { addToHistory: false, label: 'Ortografia' },
      };
    },
  };

  const toggle: AsyncCommand<void> = {
    id: 'spelling.toggle',
    label: 'Mostra errori',
    async: true,
    run: async () => {
      checker.update({ enabled: !checker.settings.enabled });
      return null;
    },
  };

  const setLanguage: AsyncCommand<SpellingCommandArgs['spelling.setLanguage']> = {
    id: 'spelling.setLanguage',
    label: 'Lingua',
    async: true,
    isEnabled: (_ctx, args) => !!args && checker.languages.some((l) => l.id === args.language),
    run: async (_ctx, { language }) => {
      checker.update({ language });
      await checker.load();
      return null;
    },
  };

  const replace: Command<SpellingCommandArgs['spelling.replace']> = {
    id: 'spelling.replace',
    label: 'Correggi',
    isEnabled: ({ state }, args) => {
      if (!args) return false;
      const p = findParagraph(state.doc, args.paragraphId)?.paragraph;
      return !!p && args.start <= args.end && rangeText(p, args.start, args.end) !== null;
    },
    run: ({ state }, { paragraphId, start, end, text }): Transaction | null => {
      const p = findParagraph(state.doc, paragraphId)?.paragraph;
      if (!p) return null;
      const doc = updateParagraph(state.doc, replaceInParagraph(p, start, end, text));
      return { state: { doc, selection: caret(paragraphId, start + text.length) }, meta: { label: 'Correggi ortografia' } };
    },
  };

  const ignoreAll: AsyncCommand<SpellingCommandArgs['spelling.ignoreAll']> = {
    id: 'spelling.ignoreAll',
    label: 'Ignora tutto',
    async: true,
    isEnabled: (_ctx, args) => !!args?.word,
    run: async (_ctx, { word }) => {
      checker.ignore(word);
      return null;
    },
  };

  const addToDictionary: AsyncCommand<SpellingCommandArgs['spelling.addToDictionary']> = {
    id: 'spelling.addToDictionary',
    label: 'Aggiungi al dizionario',
    async: true,
    isEnabled: (_ctx, args) => !!args?.word,
    run: async (_ctx, { word }) => {
      checker.addWord(word);
      return null;
    },
  };

  const commands = [next, toggle, setLanguage, replace, ignoreAll, addToDictionary];

  return {
    id: SPELLCHECK_MODULE_ID,
    checker,
    activate(editor: Editor) {
      const run = (id: string, args?: unknown) => {
        if (editor.canExecute(id, args)) void editor.executeAsync(id, args);
      };

      const marks = (state: EditorState): TextMark[] =>
        checker
          .misspellings(state.doc)
          .filter((m) => !(typing && contains(m, typing)))
          .map((m) => ({ paragraphId: m.paragraphId, start: m.start, end: m.end, kind: 'squiggle', color: SPELLING_ERROR_COLOR, className: 'spelling-error' }));

      const contextMenu = (pos: Position, state: EditorState): ContextMenuItem[] => {
        const p = findParagraph(state.doc, pos.paragraphId)?.paragraph;
        const m = p && checker.misspellingsIn(p).find((x) => contains(x, pos));
        if (!m) return [];
        const suggestions = checker.suggest(m.word);
        const items: ContextMenuItem[] = suggestions.length
          ? suggestions.map((s) => ({ label: s, strong: true, run: () => run('spelling.replace', { paragraphId: m.paragraphId, start: m.start, end: m.end, text: s }) }))
          : [{ label: '(Nessun suggerimento)', disabled: true }];
        items.push(
          { label: 'Ignora tutto', separator: true, run: () => run('spelling.ignoreAll', { word: m.word }) },
          { label: 'Aggiungi al dizionario', run: () => run('spelling.addToDictionary', { word: m.word }) },
        );
        return items;
      };

      const extension: ViewExtension = {
        id: SPELLCHECK_MODULE_ID,
        marks,
        contextMenu,
        keybindings: [{ key: 'F7', commandId: 'spelling.next' }],
        isActive(commandId, args) {
          if (commandId === 'spelling.toggle') return checker.settings.enabled;
          if (commandId === 'spelling.setLanguage') return (args as { language?: string } | undefined)?.language === checker.settings.language;
          return false;
        },
        onChange: (listener) => checker.events.on('change', listener),
      };

      const disposers = commands.map((c) => editor.commands.register(c));
      disposers.push(
        addViewExtension(editor, extension),
        editor.events.on('change', ({ state, previous, meta }) => {
          const editingWord = meta.mergeKey === 'typing' || meta.mergeKey === 'deleteBackward';
          typing = state.doc !== previous.doc && editingWord && isCollapsed(state.selection) ? state.selection.focus : null;
        }),
      );
      if (checker.settings.enabled) void checker.load();
      return () => disposers.forEach((d) => d());
    },
  };
}
