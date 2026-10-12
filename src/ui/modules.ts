/**
 * Scoperta dei moduli dell'editor. Ogni modulo parallelo esporta un
 * `EditorModule` dal proprio `index.ts`; qui li si riconosce per forma, così
 * aggiungere un modulo non richiede di modificare la UI.
 */

import type { EditorModule } from '../core';

export function isEditorModule(value: unknown): value is EditorModule {
  return (
    typeof value === 'object' &&
    value !== null &&
    typeof (value as EditorModule).id === 'string' &&
    typeof (value as EditorModule).activate === 'function'
  );
}

/**
 * Raccoglie i moduli dalle esportazioni di più file (nell'ordine dato),
 * senza duplicati per id. Riconosce oggetti `EditorModule` e funzioni
 * `create…Module()` senza argomenti.
 */
export function collectModules(exportsList: readonly Record<string, unknown>[], exclude: readonly string[] = []): EditorModule[] {
  const seen = new Set(exclude);
  const out: EditorModule[] = [];
  const add = (m: unknown) => {
    if (isEditorModule(m) && !seen.has(m.id)) {
      seen.add(m.id);
      out.push(m);
    }
  };
  for (const exports of exportsList) {
    for (const [name, value] of Object.entries(exports)) {
      if (typeof value === 'function' && /^create\w*Module$/.test(name) && value.length === 0) {
        add((value as () => unknown)());
      } else add(value);
    }
  }
  return out;
}
