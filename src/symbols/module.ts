/**
 * Modulo Simbolo (Inserisci > Simbolo, come in Word).
 *
 * - `symbol.show` apre e chiude la tabella dei simboli (stato della vista,
 *   non entra in Annulla).
 * - `symbol.insert` scrive il carattere al cursore, con la formattazione del
 *   testo intorno; si annulla con Ctrl+Z come un'unica azione.
 * - I simboli inseriti finiscono in testa ai "Simboli usati di recente".
 */

import type { AsyncCommand, Command, Editor, EditorModule } from '../core';
import { insertText } from '../editing/text';
import { browserStorage, loadRecent, pushRecent, saveRecent, type SymbolStorage } from './recent';

export const SYMBOLS_MODULE_ID = 'symbols';

export interface SymbolCommandArgs {
  /** Apre o chiude la tabella dei simboli. */
  'symbol.show': void;
  /** Inserisce un simbolo al cursore (sostituisce la selezione). */
  'symbol.insert': { char: string };
}

export interface SymbolsOptions {
  /** Dove ricordare i simboli recenti (predefinito: localStorage). */
  readonly storage?: SymbolStorage | null;
}

export interface SymbolsModule extends EditorModule {
  /** Simboli usati di recente, il più recente per primo. */
  readonly recent: readonly string[];
  /** Avvisa quando cambia l'elenco dei recenti. */
  onChange(listener: () => void): () => void;
}

export function createSymbolsModule(options: SymbolsOptions = {}): SymbolsModule {
  const storage = options.storage === undefined ? browserStorage() : options.storage;
  let recent = loadRecent(storage);
  const listeners = new Set<() => void>();

  const show: AsyncCommand = {
    id: 'symbol.show',
    label: 'Simbolo',
    async: true,
    // L'apertura della tabella la gestisce l'interfaccia (vedi panel.ts).
    run: async () => null,
  };
  const insert: Command<{ char: string }> = {
    id: 'symbol.insert',
    label: 'Inserisci simbolo',
    isEnabled: (_ctx, args) => typeof args?.char === 'string' && args.char.length > 0,
    run: ({ state }, { char }) => {
      const tr = insertText(state, char);
      if (!tr) return null;
      // Ogni simbolo è un passo a sé di Annulla, non si fonde con la digitazione.
      const { mergeKey: _merge, ...meta } = tr.meta ?? {};
      return { ...tr, meta: { ...meta, label: 'Inserisci simbolo' } };
    },
  };

  return {
    id: SYMBOLS_MODULE_ID,
    get recent() {
      return recent;
    },
    onChange(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    activate(editor: Editor) {
      const offs = [show, insert].map((c) => editor.commands.register(c));
      offs.push(
        editor.events.on('executed', ({ commandId, args }) => {
          if (commandId !== 'symbol.insert') return;
          recent = pushRecent(recent, (args as { char: string }).char);
          saveRecent(storage, recent);
          for (const l of [...listeners]) l();
        }),
      );
      return () => offs.forEach((off) => off());
    },
  };
}
