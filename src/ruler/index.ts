/**
 * Modulo Righello: righello orizzontale con margini, rientri del paragrafo
 * e tabulazioni, come in Word (Visualizza > Righello).
 */

import type { RibbonItem } from '../core';
import './styles.css';

export { createRulerModule, RULER_MODULE_ID, type RulerCommandArgs, type RulerModule, type RulerOptions, type RulerStorage } from './module';
export { changeTabs, effectiveIndents, effectiveTabs, setIndents, setTabs, snap, SNAP_STEP, TAB_LABELS, TAB_TYPES, type IndentChange, type Indents, type TabChange } from './model';
export { RulerView } from './view';

/** Voci della barra multifunzione, raccolte dalla UI (vedi `collectRibbonItems`). */
export const ribbonItems: readonly (RibbonItem & { readonly label: string })[] = [
  { tab: 'view', group: 'Mostra', commandId: 'view.ruler', label: 'Righello', kind: 'toggle', icon: '📏' },
];
