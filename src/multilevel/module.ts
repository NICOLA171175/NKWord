/**
 * Comandi degli elenchi a più livelli. Non fanno parte del catalogo di
 * `src/core`: la barra multifunzione li trova tramite `ribbonItems` e il
 * menu del clic destro tramite l'estensione della vista.
 *
 * - `multilevel.apply`: applica uno stile pronto (`preset`), dei livelli
 *   dati (`levels`) o toglie l'elenco (`preset: 'none'`).
 * - `multilevel.restart` / `multilevel.continue`: Ricomincia da 1 e
 *   Continua numerazione, come nel menu del clic destro di Word.
 * - `multilevel.setValue`: Imposta valore di numerazione.
 * - `multilevel.defineDialog` / `multilevel.setValueDialog`: le finestre.
 * - `multilevel.menu`: il menu "Elenco a più livelli" della scheda Home.
 */

import type { AnyCommand, DocumentModel, Editor, EditorModule, EditorState, ListLevel, Transaction } from '../core';
import { addViewExtension, type ContextMenuItem } from '../ui/extensions';
import type { MultilevelDialogs } from './dialog';
import {
  applyLevels,
  applyPreset,
  canContinue,
  canRestart,
  continueNumbering,
  currentValue,
  focusItem,
  focusLevels,
  fullLevels,
  removeList,
  restartNumbering,
  type PresetId,
} from './model';

export const MULTILEVEL_MODULE_ID = 'multilevel';

export type MenuChoice = PresetId | 'none' | 'define' | 'restart' | 'continue' | 'setValue';

export interface MultilevelCommandArgs {
  'multilevel.apply': { readonly preset: PresetId | 'none' } | { readonly levels: readonly ListLevel[] };
  'multilevel.restart': void;
  'multilevel.continue': void;
  'multilevel.setValue': { readonly value: number };
  'multilevel.defineDialog': void;
  'multilevel.setValueDialog': void;
  'multilevel.menu': { readonly choice: MenuChoice };
}

export type MultilevelCommandId = keyof MultilevelCommandArgs;

export interface MultilevelOptions {
  readonly dialogs?: MultilevelDialogs;
}

const LABELS = {
  apply: 'Elenco a più livelli',
  restart: 'Ricomincia da 1',
  continue: 'Continua numerazione',
  setValue: 'Imposta valore di numerazione',
} as const;

function tr(state: EditorState, doc: DocumentModel | null, label: string): Transaction | null {
  return doc && { state: { ...state, doc }, meta: { label } };
}

export function createMultilevelCommands(editor: Editor, options: MultilevelOptions = {}): AnyCommand[] {
  const dialogs = options.dialogs;

  const apply = (state: EditorState, args: MultilevelCommandArgs['multilevel.apply'] | undefined) => {
    if (!args) return null;
    if ('levels' in args) return tr(state, applyLevels(state, fullLevels(args.levels)), LABELS.apply);
    if (args.preset === 'none') return tr(state, removeList(state), LABELS.apply);
    return tr(state, applyPreset(state, args.preset), LABELS.apply);
  };

  const define = async (): Promise<Transaction | null> => {
    if (!dialogs) return null;
    const at = focusLevels(editor.state);
    const levels = await dialogs.define({ levels: fullLevels(at?.levels), level: at?.level ?? 0 });
    return levels ? apply(editor.state, { levels }) : null;
  };

  const setValue = async (): Promise<Transaction | null> => {
    if (!dialogs || !focusItem(editor.state)) return null;
    const value = await dialogs.setValue({ value: currentValue(editor.state) });
    return value === null ? null : tr(editor.state, restartNumbering(editor.state, value), LABELS.setValue);
  };

  const hasDialogs = () => !!dialogs;
  const numberedFocus = (state: EditorState) => {
    const item = focusItem(state);
    return !!item && state.doc.lists[item.ref.listId]?.levels[0]?.format !== 'bullet';
  };

  return [
    { id: 'multilevel.apply', label: LABELS.apply, run: ({ state }, args) => apply(state, args) },
    { id: 'multilevel.restart', label: LABELS.restart, isEnabled: ({ state }) => canRestart(state), run: ({ state }) => tr(state, restartNumbering(state), LABELS.restart) },
    { id: 'multilevel.continue', label: LABELS.continue, isEnabled: ({ state }) => canContinue(state), run: ({ state }) => tr(state, continueNumbering(state), LABELS.continue) },
    {
      id: 'multilevel.setValue',
      label: LABELS.setValue,
      isEnabled: ({ state }) => numberedFocus(state),
      run: ({ state }, args: MultilevelCommandArgs['multilevel.setValue']) => tr(state, args ? restartNumbering(state, Math.max(0, Math.round(args.value))) : null, LABELS.setValue),
    },
    { id: 'multilevel.defineDialog', label: 'Definisci nuovo elenco a più livelli', async: true, isEnabled: hasDialogs, run: define },
    { id: 'multilevel.setValueDialog', label: LABELS.setValue, async: true, isEnabled: ({ state }) => hasDialogs() && numberedFocus(state), run: setValue },
    {
      id: 'multilevel.menu',
      label: LABELS.apply,
      async: true,
      run: async ({ state }, args: MultilevelCommandArgs['multilevel.menu']) => {
        switch (args?.choice) {
          case undefined:
            return null;
          case 'define':
            return define();
          case 'setValue':
            return setValue();
          case 'restart':
            return tr(state, restartNumbering(state), LABELS.restart);
          case 'continue':
            return tr(state, continueNumbering(state), LABELS.continue);
          default:
            return apply(state, { preset: args.choice });
        }
      },
    },
  ];
}

export function createMultilevelModule(options: MultilevelOptions = {}): EditorModule {
  return {
    id: MULTILEVEL_MODULE_ID,
    activate(editor) {
      const offs = createMultilevelCommands(editor, options).map((c) => editor.commands.register(c));
      offs.push(
        addViewExtension(editor, {
          id: MULTILEVEL_MODULE_ID,
          // Clic destro su una voce di elenco numerato, come in Word.
          contextMenu(pos, current): ContextMenuItem[] {
            // Il clic destro non sposta il cursore: si ragiona sul paragrafo cliccato.
            const moved = current.selection.focus.paragraphId !== pos.paragraphId;
            const state: EditorState = moved ? { ...current, selection: { anchor: pos, focus: pos } } : current;
            const item = focusItem(state);
            if (!item || state.doc.lists[item.ref.listId]?.levels[0]?.format === 'bullet') return [];
            const run = (id: MultilevelCommandId) => () => {
              if (moved) editor.execute('edit.setSelection', { selection: { anchor: pos, focus: pos } });
              void editor.executeAsync(id);
            };
            return [
              { label: LABELS.restart, run: run('multilevel.restart'), disabled: !canRestart(state), separator: true },
              { label: LABELS.continue, run: run('multilevel.continue'), disabled: !canContinue(state) },
              { label: `${LABELS.setValue}…`, run: run('multilevel.setValueDialog'), disabled: !options.dialogs },
            ];
          },
        }),
      );
      return () => offs.forEach((off) => off());
    },
  };
}
