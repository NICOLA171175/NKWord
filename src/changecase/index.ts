/**
 * Modulo Maiuscole/minuscole della scheda Home, come in Word.
 */

import type { RibbonItem } from '../core';

export { createChangeCaseModule, CHANGE_CASE_MODULE_ID, type ChangeCaseCommandArgs } from './module';
export { changeCase, convertCase, nextCycleMode, type CaseMode } from './convert';

/** Voce della barra multifunzione, raccolta dalla UI (vedi `collectRibbonItems`). */
export const ribbonItems: readonly (RibbonItem & { readonly label: string; readonly options: readonly { label: string; args: unknown }[] })[] = [
  {
    tab: 'home',
    group: 'Carattere',
    commandId: 'changeCase.apply',
    label: 'Maiuscole/minuscole',
    kind: 'dropdown',
    icon: 'Aa',
    options: [
      { label: 'Tipo frase.', args: { mode: 'sentence' } },
      { label: 'minuscolo', args: { mode: 'lower' } },
      { label: 'MAIUSCOLO', args: { mode: 'upper' } },
      { label: 'Iniziali Maiuscole', args: { mode: 'title' } },
      { label: 'aLTERNA mAIUSC/minusc', args: { mode: 'toggle' } },
    ],
  },
];
