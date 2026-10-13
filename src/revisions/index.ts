/**
 * Modulo revisioni (scheda Revisione): rilevamento modifiche,
 * accetta/rifiuta e visualizzazione delle revisioni. Registra i comandi
 * `revision.*` (vedi `RevisionCommandArgs`).
 */

import type { RibbonItem } from '../core';
import type { RevisionCommandId } from './module';

type Item = RibbonItem & { readonly commandId: RevisionCommandId; readonly label: string; readonly prompt?: unknown };
const item = (group: string, commandId: RevisionCommandId, label: string, icon: string, extra: Partial<Item> = {}): Item => ({
  tab: 'review',
  group,
  commandId,
  label,
  icon,
  kind: 'button',
  ...extra,
});

/** Voci della barra multifunzione, raccolte dalla UI (vedi `collectRibbonItems`). */
export const ribbonItems: readonly Item[] = [
  item('Rilevamento', 'revision.toggleTracking', 'Rileva modifiche', '✎', { kind: 'toggle' }),
  item('Rilevamento', 'revision.setAuthor', 'Nome autore', '👤', {
    prompt: (ask: (message: string, initial?: string) => string | null) => {
      const author = ask('Nome con cui firmare le modifiche:');
      return author?.trim() ? { author } : undefined;
    },
  }),
  item('Revisioni', 'revision.previous', 'Precedente', '⇠'),
  item('Revisioni', 'revision.next', 'Successiva', '⇢'),
  item('Revisioni', 'revision.accept', 'Accetta', '✔'),
  item('Revisioni', 'revision.reject', 'Rifiuta', '✖'),
  item('Revisioni', 'revision.acceptAll', 'Accetta tutto', '✔✔'),
  item('Revisioni', 'revision.rejectAll', 'Rifiuta tutto', '✖✖'),
];

export * from './module';
export * from './review';
export * from './track';
export * from './markup';
export { tokenize, buildParagraphs, revisionOf, type Token } from './stream';
