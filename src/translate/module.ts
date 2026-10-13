/**
 * Modulo Traduttore, come Revisione > Traduci di Word.
 *
 * - `translate.selection` (Traduci selezione) apre il riquadro Traduttore
 *   e traduce il testo selezionato; "Inserisci" lo sostituisce con la
 *   traduzione.
 * - `translate.document` (Traduci documento) apre il riquadro sulla scheda
 *   Documento; "Traduci" traduce tutto il documento in un solo passo, che
 *   si può annullare con Ctrl+Z per tornare all'originale.
 * - `translate.menu` è il menu a tendina "Traduci" della barra.
 * - `translate.insert` / `translate.applyDocument` applicano un dizionario
 *   di traduzioni già pronto alla selezione o al documento (puri).
 *
 * Il servizio usato è in service.ts.
 */

import { caret, firstParagraph, type AnyCommand, type Editor, type EditorModule, type EditorState } from '../core';
import { addViewExtension } from '../ui/extensions';
import { AUTO_DETECT } from './languages';
import { collectTexts, hasSelectedText, lookup, rewriteDocument, rewriteSelection, selectionText, type Translations } from './model';
import { createGoogleTranslator, TranslationError, type Translator } from './service';

export const TRANSLATE_MODULE_ID = 'translate';
export const SETTINGS_KEY = 'nkoffice.traduttore';
/** Attesa dopo un cambio di selezione prima di tradurla. */
export const SELECTION_DELAY_MS = 500;

export type TranslateMode = 'selection' | 'document';

export interface TranslateCommandArgs {
  'translate.selection': void;
  'translate.document': void;
  'translate.menu': { readonly choice: TranslateMode };
  'translate.insert': { readonly translations: Translations };
  'translate.applyDocument': { readonly translations: Translations };
}

export type TranslateStatus = 'idle' | 'busy' | 'done' | 'error';

export interface TranslateModule extends EditorModule {
  readonly open: boolean;
  readonly mode: TranslateMode;
  /** Lingua di partenza ('auto' = Rileva automaticamente). */
  readonly from: string;
  /** Lingua di arrivo. */
  readonly to: string;
  readonly status: TranslateStatus;
  /** Avviso o errore da mostrare nel riquadro. */
  readonly message: string | null;
  /** Testo selezionato (scheda Selezione). */
  readonly source: string;
  /** Traduzione della selezione. */
  readonly result: string;
  /** Lingua riconosciuta dal servizio. */
  readonly detected: string | null;
  show(mode: TranslateMode): void;
  close(): void;
  setLanguages(languages: { from?: string; to?: string }): void;
  /** Traduce la selezione e mostra il risultato nel riquadro. */
  translateSelection(): Promise<void>;
  /** Sostituisce la selezione con la traduzione mostrata. */
  insert(): boolean;
  /** Traduce tutto il documento. */
  translateDocument(): Promise<boolean>;
  onChange(listener: () => void): () => void;
}

export interface TranslateOptions {
  readonly translator?: Translator;
  readonly storage?: Pick<Storage, 'getItem' | 'setItem'> | null;
  readonly delayMs?: number;
}

function defaultStorage(): Pick<Storage, 'getItem' | 'setItem'> | null {
  try {
    return globalThis.localStorage ?? null;
  } catch {
    return null;
  }
}

function errorText(error: unknown): string {
  return error instanceof TranslationError ? error.message : 'La traduzione non è riuscita. Riprova più tardi.';
}

export function createTranslateModule(options: TranslateOptions = {}): TranslateModule {
  const translator = options.translator ?? createGoogleTranslator();
  const storage = options.storage === undefined ? defaultStorage() : options.storage;
  const delayMs = options.delayMs ?? SELECTION_DELAY_MS;
  let editor: Editor | undefined;
  let open = false;
  let mode: TranslateMode = 'selection';
  let from = AUTO_DETECT;
  let to = 'en';
  let status: TranslateStatus = 'idle';
  let message: string | null = null;
  let source = '';
  let result = '';
  let detected: string | null = null;
  let translations: Translations = new Map();
  let translatedFor: EditorState | null = null;
  let request = 0;
  let controller: AbortController | null = null;
  let timer: ReturnType<typeof setTimeout> | undefined;
  const listeners = new Set<() => void>();
  const notify = () => [...listeners].forEach((l) => l());

  try {
    const saved = JSON.parse(storage?.getItem(SETTINGS_KEY) ?? 'null') as { from?: unknown; to?: unknown } | null;
    if (typeof saved?.from === 'string') from = saved.from;
    if (typeof saved?.to === 'string') to = saved.to;
  } catch {
    // Impostazioni illeggibili: si usano quelle predefinite.
  }

  const save = () => {
    try {
      storage?.setItem(SETTINGS_KEY, JSON.stringify({ from, to }));
    } catch {
      // Memoria del browser non disponibile: le lingue valgono finché l'app resta aperta.
    }
  };

  /** Avvia una nuova richiesta, annullando quella in corso. */
  const begin = () => {
    controller?.abort();
    controller = typeof AbortController === 'function' ? new AbortController() : null;
    return ++request;
  };

  const translate = async (texts: readonly string[]) => {
    const id = begin();
    status = 'busy';
    message = null;
    notify();
    try {
      const out = await translator.translate(texts, to, from, controller?.signal);
      if (id !== request) return null;
      detected = out.detected ?? null;
      return new Map(texts.map((t, i) => [t, out.texts[i] ?? t])) as Translations;
    } catch (error) {
      if (id !== request) return null;
      status = 'error';
      message = errorText(error);
      notify();
      return null;
    }
  };

  const scheduleSelection = () => {
    clearTimeout(timer);
    if (!open || mode !== 'selection' || !editor) return;
    if (translatedFor && translatedFor.doc === editor.state.doc && translatedFor.selection === editor.state.selection) return;
    timer = setTimeout(() => void module.translateSelection(), delayMs);
  };

  const module: TranslateModule = {
    id: TRANSLATE_MODULE_ID,
    get open() {
      return open;
    },
    get mode() {
      return mode;
    },
    get from() {
      return from;
    },
    get to() {
      return to;
    },
    get status() {
      return status;
    },
    get message() {
      return message;
    },
    get source() {
      return source;
    },
    get result() {
      return result;
    },
    get detected() {
      return detected;
    },

    show(next) {
      open = true;
      if (mode !== next) {
        mode = next;
        begin();
        status = 'idle';
        message = null;
      }
      notify();
      if (mode === 'selection') void module.translateSelection();
    },

    close() {
      clearTimeout(timer);
      begin();
      open = false;
      status = 'idle';
      notify();
    },

    setLanguages(next) {
      if (next.from) from = next.from;
      if (next.to) to = next.to;
      save();
      translatedFor = null;
      if (mode === 'selection' && open) void module.translateSelection();
      else {
        status = 'idle';
        message = null;
        notify();
      }
    },

    async translateSelection() {
      clearTimeout(timer);
      if (!editor) return;
      const state = editor.state;
      translatedFor = state;
      source = selectionText(state).trim();
      result = '';
      translations = new Map();
      if (!hasSelectedText(state)) {
        begin();
        status = 'idle';
        message = 'Seleziona del testo nel documento per vederne la traduzione.';
        notify();
        return;
      }
      const map = await translate(collectTexts((mapText) => rewriteSelection(state, mapText)));
      if (!map) return;
      translations = map;
      result = selectionText(state, lookup(map)).trim();
      status = 'done';
      notify();
    },

    insert() {
      if (!editor || status !== 'done' || mode !== 'selection') return false;
      const done = editor.execute('translate.insert', { translations });
      // La selezione ora è il testo tradotto: non c'è da ritradurla.
      if (done) translatedFor = editor.state;
      return done;
    },

    async translateDocument() {
      if (!editor) return false;
      const state = editor.state;
      const texts = collectTexts((mapText) => rewriteDocument(state.doc, mapText));
      if (!texts.length) {
        status = 'error';
        message = 'Il documento non contiene testo da tradurre.';
        notify();
        return false;
      }
      const map = await translate(texts);
      if (!map) return false;
      const done = editor.execute('translate.applyDocument', { translations: map });
      status = 'done';
      message = 'Documento tradotto. Per tornare al testo originale usa Annulla (Ctrl+Z).';
      notify();
      return done;
    },

    onChange(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },

    activate(ed) {
      editor = ed;
      const commands: AnyCommand[] = [
        {
          id: 'translate.selection',
          label: 'Traduci selezione',
          async: true,
          run: async () => {
            module.show('selection');
            return null;
          },
        },
        {
          id: 'translate.document',
          label: 'Traduci documento',
          async: true,
          run: async () => {
            module.show('document');
            return null;
          },
        },
        {
          id: 'translate.menu',
          label: 'Traduci',
          async: true,
          run: async (_ctx, args?: TranslateCommandArgs['translate.menu']) => {
            module.show(args?.choice ?? 'selection');
            return null;
          },
        },
        {
          id: 'translate.insert',
          label: 'Inserisci traduzione',
          isEnabled: ({ state }, args?: TranslateCommandArgs['translate.insert']) => !!args?.translations.size && hasSelectedText(state),
          run: ({ state }, args: TranslateCommandArgs['translate.insert']) => rewriteSelection(state, lookup(args.translations)),
        },
        {
          id: 'translate.applyDocument',
          label: 'Traduci documento',
          isEnabled: (_ctx, args?: TranslateCommandArgs['translate.applyDocument']) => !!args?.translations.size,
          run: ({ state }, args: TranslateCommandArgs['translate.applyDocument']) => {
            const doc = rewriteDocument(state.doc, lookup(args.translations));
            return { state: { doc, selection: caret(firstParagraph(doc).id, 0) }, meta: { label: 'Traduci documento' } };
          },
        },
      ];
      const disposers = commands.map((c) => ed.commands.register(c));
      disposers.push(
        addViewExtension(ed, {
          id: TRANSLATE_MODULE_ID,
          contextMenu: (_pos, state) =>
            hasSelectedText(state) ? [{ label: 'Traduci', separator: true, run: () => module.show('selection') }] : [],
        }),
      );
      const offChange = ed.events.on('change', () => scheduleSelection());
      return () => {
        clearTimeout(timer);
        begin();
        offChange();
        disposers.forEach((d) => d());
        editor = undefined;
      };
    },
  };
  return module;
}
