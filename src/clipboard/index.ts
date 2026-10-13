/**
 * Modulo Appunti della scheda Home: Incolla, Taglia, Copia.
 */

import type { RibbonItem } from '../core';

export {
  browserClipboard,
  CLIPBOARD_MODULE_ID,
  createClipboardModule,
  type ClipboardCommandArgs,
  type ClipboardModuleOptions,
  type SystemClipboard,
} from './module';

/** Voci della barra multifunzione, raccolte dalla UI (vedi `collectRibbonItems`). */
export const ribbonItems: readonly (RibbonItem & { readonly label: string })[] = [
  { tab: 'home', group: 'Appunti', commandId: 'clipboard.paste', label: 'Incolla' },
  { tab: 'home', group: 'Appunti', commandId: 'clipboard.paste', label: 'Incolla solo testo', args: { textOnly: true } },
  { tab: 'home', group: 'Appunti', commandId: 'clipboard.cut', label: 'Taglia' },
  { tab: 'home', group: 'Appunti', commandId: 'clipboard.copy', label: 'Copia' },
];
