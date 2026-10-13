/**
 * Modulo Copia formato (il pennello della scheda Home di Word).
 *
 * - Un clic sul pennello copia la formattazione della selezione; la si
 *   applica selezionando altro testo con il mouse, poi il pennello si spegne.
 * - Doppio clic: il pennello resta attivo per più applicazioni, finché non
 *   si preme Esc o lo si clicca di nuovo.
 * - Ctrl+Maiusc+C copia il formato e Ctrl+Maiusc+V lo incolla sulla
 *   selezione, senza usare il mouse.
 *
 * Lo stato del pennello è della vista, non del documento: accenderlo e
 * spegnerlo non entra in Annulla. Applicare il formato invece sì.
 */

import type { AsyncCommand, Command, Editor, EditorModule } from '../core';
import { addViewExtension } from '../ui/extensions';
import { copyFormat, LABEL, pasteFormat, type CopiedFormat } from './painter';

export const FORMAT_PAINTER_MODULE_ID = 'formatpainter';

/** 'once': si spegne dopo un'applicazione; 'sticky': resta acceso (doppio clic). */
export type PainterMode = 'off' | 'once' | 'sticky';

export interface FormatPainterCommandArgs {
  /** Accende o spegne il pennello; `sticky` lo tiene acceso (doppio clic). */
  'formatPainter.toggle': { sticky?: boolean } | undefined;
  /** Applica la formattazione copiata alla selezione (o alla parola al cursore). */
  'formatPainter.apply': void;
  /** Spegne il pennello (Esc). */
  'formatPainter.cancel': void;
  /** Copia il formato senza accendere il pennello (Ctrl+Maiusc+C). */
  'formatPainter.copy': void;
  /** Incolla il formato copiato (Ctrl+Maiusc+V). */
  'formatPainter.paste': void;
}

export interface FormatPainterModule extends EditorModule {
  readonly mode: PainterMode;
  /** Ultima formattazione copiata, se c'è. */
  readonly copied: CopiedFormat | null;
  /** Avvisa quando il pennello si accende o si spegne. */
  onChange(listener: () => void): () => void;
}

export function createFormatPainterModule(): FormatPainterModule {
  let mode: PainterMode = 'off';
  let copied: CopiedFormat | null = null;
  const listeners = new Set<() => void>();
  const setMode = (next: PainterMode) => {
    if (next === mode) return;
    mode = next;
    for (const l of [...listeners]) l();
  };

  const toggle: AsyncCommand<{ sticky?: boolean } | undefined> = {
    id: 'formatPainter.toggle',
    label: LABEL,
    async: true,
    run: async ({ state }, args) => {
      if (args?.sticky) {
        copied = copyFormat(state);
        setMode('sticky');
      } else if (mode !== 'off') {
        setMode('off');
      } else {
        copied = copyFormat(state);
        setMode('once');
      }
      return null;
    },
  };
  const apply: Command = {
    id: 'formatPainter.apply',
    label: LABEL,
    isEnabled: () => mode !== 'off' && copied !== null,
    run: ({ state }) => pasteFormat(state, copied!),
  };
  const cancel: AsyncCommand = {
    id: 'formatPainter.cancel',
    label: 'Annulla Copia formato',
    async: true,
    isEnabled: () => mode !== 'off',
    run: async () => {
      setMode('off');
      return null;
    },
  };
  const copy: AsyncCommand = {
    id: 'formatPainter.copy',
    label: 'Copia formato',
    async: true,
    run: async ({ state }) => {
      copied = copyFormat(state);
      return null;
    },
  };
  const paste: Command = {
    id: 'formatPainter.paste',
    label: 'Incolla formato',
    isEnabled: () => copied !== null,
    run: ({ state }) => pasteFormat(state, copied!),
  };

  return {
    id: FORMAT_PAINTER_MODULE_ID,
    get mode() {
      return mode;
    },
    get copied() {
      return copied;
    },
    onChange(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    activate(editor: Editor) {
      const offs = [toggle, apply, cancel, copy, paste].map((c) => editor.commands.register(c));
      // Dopo un'applicazione con un solo clic il pennello si spegne.
      offs.push(
        editor.events.on('executed', ({ commandId }) => {
          if (commandId === 'formatPainter.apply' && mode === 'once') setMode('off');
        }),
        addViewExtension(editor, {
          id: FORMAT_PAINTER_MODULE_ID,
          isActive: (commandId) => commandId === 'formatPainter.toggle' && mode !== 'off',
          onChange: (listener) => {
            listeners.add(listener);
            return () => listeners.delete(listener);
          },
          keybindings: [
            { key: 'Mod-Shift-c', commandId: 'formatPainter.copy' },
            { key: 'Mod-Shift-v', commandId: 'formatPainter.paste' },
            { key: 'Escape', commandId: 'formatPainter.cancel' },
          ],
        }),
      );
      return () => {
        setMode('off');
        offs.forEach((off) => off());
      };
    },
  };
}
