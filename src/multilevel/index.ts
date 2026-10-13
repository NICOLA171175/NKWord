/**
 * Modulo Elenchi a più livelli (Home > Paragrafo > Elenco a più livelli):
 * stili pronti come 1. 1.1. 1.1.1. o I. A. 1., formato di ogni livello,
 * Ricomincia da 1, Continua numerazione e Imposta valore. Tab e Maiusc+Tab
 * all'inizio di una voce scendono e salgono di livello (modulo strutture).
 */

import type { RibbonItem } from '../core';
import { domDialogs } from './dialog';
import { createMultilevelModule as createModule, type MenuChoice, type MultilevelOptions } from './module';
import { PRESETS } from './model';
import './styles.css';

/** Modulo con le finestre del browser, se c'è un documento. */
export function createMultilevelModule(options: MultilevelOptions = {}) {
  return createModule({ ...options, dialogs: options.dialogs ?? (typeof document !== 'undefined' ? domDialogs : undefined) });
}

const MENU: readonly { label: string; choice: MenuChoice }[] = [
  ...PRESETS.map((p) => ({ label: p.label, choice: p.id })),
  { label: 'Nessuno', choice: 'none' },
  { label: 'Definisci nuovo elenco a più livelli…', choice: 'define' },
  { label: 'Ricomincia da 1', choice: 'restart' },
  { label: 'Continua numerazione', choice: 'continue' },
  { label: 'Imposta valore di numerazione…', choice: 'setValue' },
];

/** Voci della barra multifunzione, raccolte dalla UI (vedi `collectRibbonItems`). */
export const ribbonItems: readonly (RibbonItem & { readonly label: string; readonly options: readonly { label: string; args: unknown }[] })[] = [
  {
    tab: 'home', group: 'Paragrafo', commandId: 'multilevel.menu', label: 'Elenco a più livelli', kind: 'dropdown', icon: '1.1',
    options: MENU.map((m) => ({ label: m.label, args: { choice: m.choice } })),
  },
];

export { MULTILEVEL_MODULE_ID, createMultilevelCommands, type MenuChoice, type MultilevelCommandArgs, type MultilevelCommandId, type MultilevelOptions } from './module';
export * from './model';
export { domDialogs, FORMAT_CHOICES, openDefineDialog, openSetValueDialog, type DefineDialogInput, type MultilevelDialogs } from './dialog';
