/**
 * Modulo Tema: Visualizza > Tema, per scegliere fra interfaccia chiara,
 * scura o automatica (segue il sistema). La scelta resta fra un avvio e
 * l'altro.
 */

import type { RibbonItem } from '../core';
import { THEME_CHOICES, THEME_LABELS } from './settings';
import './styles.css';

export { createThemeModule, THEME_MODULE_ID, type ThemeCommandArgs, type ThemeModule, type ThemeOptions } from './module';
export {
  browserSystemTheme,
  DEFAULT_THEME,
  isThemeChoice,
  loadTheme,
  resolveTheme,
  saveTheme,
  THEME_CHOICES,
  THEME_LABELS,
  type ResolvedTheme,
  type SystemTheme,
  type ThemeChoice,
  type ThemeStorage,
} from './settings';

/** Voci della barra multifunzione, raccolte dalla UI (vedi `collectRibbonItems`). */
export const ribbonItems: readonly (RibbonItem & { readonly label: string; readonly options: readonly { label: string; args: unknown }[] })[] = [
  {
    tab: 'view', group: 'Tema', commandId: 'view.theme', label: 'Tema', kind: 'dropdown', icon: '◐',
    options: THEME_CHOICES.map((theme) => ({ label: THEME_LABELS[theme], args: { theme } })),
  },
];
