/**
 * Modulo intestazione e piè di pagina: Inserisci > Intestazione, Piè di
 * pagina e Numero di pagina. Il contenuto si scrive in una finestra (testo,
 * allineamento, numero di pagina) e compare in tutte le pagine.
 */

import type { EditorModule, RibbonItem } from '../core';
import { createHeaderFooterCommands, type HeaderFooterDialog } from './commands';
import { openHeaderFooterDialog } from './dialog';

export const HEADER_FOOTER_MODULE_ID = 'headerFooter';

export interface HeaderFooterOptions {
  /** Finestra di modifica (per default quella del browser). */
  readonly dialog?: HeaderFooterDialog;
}

export function createHeaderFooterModule(options: HeaderFooterOptions = {}): EditorModule {
  const dialog = options.dialog ?? (typeof document !== 'undefined' ? openHeaderFooterDialog : undefined);
  return {
    id: HEADER_FOOTER_MODULE_ID,
    activate(editor) {
      const disposers = createHeaderFooterCommands(editor, dialog).map((c) => editor.commands.register(c));
      return () => disposers.forEach((d) => d());
    },
  };
}

/** Voci della barra multifunzione, raccolte dalla UI (vedi `collectRibbonItems`). */
export const ribbonItems: readonly (RibbonItem & { readonly label: string; readonly options: readonly { label: string; args: unknown }[] })[] = [
  {
    tab: 'insert', group: 'Intestazione e piè di pagina', commandId: 'headerFooter.pageNumber', label: 'Numero di pagina', kind: 'dropdown', icon: '#',
    options: [
      { label: 'Inizio pagina', args: { position: 'header' } },
      { label: 'Fine pagina', args: { position: 'footer' } },
      { label: 'Posizione corrente', args: { position: 'here' } },
    ],
  },
];

export { createHeaderFooterCommands, currentHeaderFooter, type HeaderFooterCommandArgs, type HeaderFooterCommandId, type HeaderFooterDialog } from './commands';
export { openHeaderFooterDialog } from './dialog';
export * from './model';
