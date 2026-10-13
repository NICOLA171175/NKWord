/**
 * Modulo Maiuscole/minuscole: il pulsante "Aa" della scheda Home di Word
 * (Tipo frase., minuscolo, MAIUSCOLO, Iniziali Maiuscole, aLTERNA) e
 * Maiusc+F3 che passa da una forma all'altra.
 */

import type { Command, Editor, EditorModule } from '../core';
import { addViewExtension } from '../ui/extensions';
import { changeCase, type CaseMode } from './convert';

export const CHANGE_CASE_MODULE_ID = 'changecase';

export interface ChangeCaseCommandArgs {
  /** Applica una forma alla selezione (o alla parola al cursore). */
  'changeCase.apply': { mode: CaseMode };
  /** Maiusc+F3: MAIUSCOLO → minuscolo → Iniziali Maiuscole. */
  'changeCase.cycle': void;
}

export function createChangeCaseModule(): EditorModule {
  const apply: Command<{ mode: CaseMode }> = {
    id: 'changeCase.apply',
    label: 'Maiuscole/minuscole',
    isEnabled: (_ctx, args) => !!args?.mode,
    run: ({ state }, args) => changeCase(state, args.mode),
  };
  const cycle: Command = {
    id: 'changeCase.cycle',
    label: 'Maiuscole/minuscole',
    run: ({ state }) => changeCase(state, 'cycle'),
  };
  return {
    id: CHANGE_CASE_MODULE_ID,
    activate(editor: Editor) {
      const offs = [apply, cycle].map((c) => editor.commands.register(c));
      offs.push(addViewExtension(editor, { id: CHANGE_CASE_MODULE_ID, keybindings: [{ key: 'Shift-F3', commandId: 'changeCase.cycle' }] }));
      return () => offs.forEach((off) => off());
    },
  };
}
