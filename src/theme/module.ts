/**
 * Modulo Tema: interfaccia chiara, scura o automatica, come in Word.
 *
 * - `view.theme` cambia il tema (`{ theme: 'light' | 'dark' | 'auto' }`);
 *   è una scelta della vista, non entra in Annulla.
 * - Il tema effettivo si scrive in `data-theme` sull'elemento <html>: i fogli
 *   di stile lo usano per i colori scuri. Il foglio del documento resta
 *   bianco, come in Word.
 * - In "automatico" segue il sistema anche mentre l'app è aperta.
 */

import type { AsyncCommand, Editor, EditorModule } from '../core';
import { addViewExtension } from '../ui/extensions';
import {
  browserStorage,
  browserSystemTheme,
  isThemeChoice,
  loadTheme,
  resolveTheme,
  saveTheme,
  type ResolvedTheme,
  type SystemTheme,
  type ThemeChoice,
  type ThemeStorage,
} from './settings';

export const THEME_MODULE_ID = 'theme';

export interface ThemeCommandArgs {
  'view.theme': { theme: ThemeChoice };
}

export interface ThemeOptions {
  /** Dove ricordare la scelta (predefinito: localStorage). */
  readonly storage?: ThemeStorage | null;
  /** Preferenza del sistema (predefinito: prefers-color-scheme del browser). */
  readonly system?: SystemTheme;
  /** Elemento su cui scrivere `data-theme` (predefinito: <html>). */
  readonly root?: HTMLElement | null;
}

export interface ThemeModule extends EditorModule {
  /** La scelta dell'utente. */
  readonly choice: ThemeChoice;
  /** Il tema effettivo, dopo aver risolto "automatico". */
  readonly resolved: ResolvedTheme;
  setTheme(choice: ThemeChoice): void;
  onChange(listener: () => void): () => void;
}

export function createThemeModule(options: ThemeOptions = {}): ThemeModule {
  const storage = options.storage === undefined ? browserStorage() : options.storage;
  const system = options.system ?? browserSystemTheme();
  const root = options.root === undefined ? (globalThis.document?.documentElement ?? null) : options.root;
  let choice = loadTheme(storage);
  const listeners = new Set<() => void>();

  const resolved = () => resolveTheme(choice, system.dark);
  const apply = () => {
    if (!root) return;
    root.dataset.theme = resolved();
    root.style.colorScheme = resolved();
  };
  const notify = () => {
    apply();
    for (const l of [...listeners]) l();
  };
  // Subito, prima che compaia l'interfaccia: niente lampo bianco all'avvio.
  apply();

  const setTheme = (next: ThemeChoice) => {
    if (next === choice) return;
    choice = next;
    saveTheme(storage, choice);
    notify();
  };

  const command: AsyncCommand<{ theme: ThemeChoice }> = {
    id: 'view.theme',
    label: 'Tema',
    async: true,
    isEnabled: (_ctx, args) => isThemeChoice(args?.theme),
    run: async (_ctx, { theme }) => {
      setTheme(theme);
      return null;
    },
  };

  return {
    id: THEME_MODULE_ID,
    get choice() {
      return choice;
    },
    get resolved() {
      return resolved();
    },
    setTheme,
    onChange(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    activate(editor: Editor) {
      const offs = [
        editor.commands.register(command),
        system.onChange(() => {
          if (choice === 'auto') notify();
        }),
        addViewExtension(editor, {
          id: THEME_MODULE_ID,
          isActive: (commandId, args) => commandId === 'view.theme' && (args as { theme?: unknown } | undefined)?.theme === choice,
          onChange: (listener) => {
            listeners.add(listener);
            return () => listeners.delete(listener);
          },
        }),
      ];
      apply();
      return () => offs.forEach((off) => off());
    },
  };
}
