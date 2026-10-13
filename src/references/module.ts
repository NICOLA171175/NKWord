/**
 * Comandi di didascalie, riferimenti incrociati e indice analitico. Non
 * fanno parte del catalogo di `src/core`: la barra multifunzione li trova
 * tramite `ribbonItems` (vedi `index.ts`).
 *
 * I comandi `…Dialog` aprono la finestra corrispondente e poi eseguono il
 * comando puro con le scelte fatte; i comandi puri si possono usare anche
 * da soli (test, macro).
 */

import type { AsyncCommand, Command, DocumentModel, Editor, EditorModule, EditorState, Transaction } from '../core';
import { captionLabels, canInsertCaption, defaultPosition, insertCaption, nextCaptionNumber, renumberCaptions, type CaptionArgs, type CaptionPosition } from './captions';
import { pageMap } from './common';
import {
  canInsertCrossRef,
  crossRefTypingFilter,
  insertCrossRef,
  listTargets,
  parseCrossRef,
  updateCrossRefs,
  type CrossRef,
  type RefTarget,
} from './crossrefs';
import { findIndex, markEntry, suggestedEntry, writeIndex, type MarkArgs } from './indexing';

export const REFERENCES_MODULE_ID = 'references';

export interface ReferencesCommandArgs {
  'references.insertCaption': CaptionArgs;
  'references.captionDialog': void;
  'references.insertCrossRef': CrossRef | { ref: string };
  'references.crossRefDialog': void;
  'references.markEntry': MarkArgs;
  'references.markEntryDialog': void;
  /** Inserisce l'indice al cursore; se esiste già lo aggiorna. */
  'references.insertIndex': void;
  'references.updateIndex': void;
  /** Rinumera le didascalie e aggiorna riferimenti incrociati e indice (come F9 in Word). */
  'references.updateFields': void;
}

export type ReferencesCommandId = keyof ReferencesCommandArgs;

/** Finestre usate dai comandi `…Dialog`; l'interfaccia le fornisce (vedi `dialog.ts`). */
export interface ReferencesDialogs {
  caption(data: { labels: readonly string[]; next: (label: string) => number; defaultPosition: (label: string) => CaptionPosition }): Promise<CaptionArgs | null>;
  crossRef(data: { targets: readonly RefTarget[] }): Promise<CrossRef | null>;
  markEntry(data: { entry: string }): Promise<MarkArgs | null>;
}

export interface ReferencesOptions {
  readonly dialogs?: ReferencesDialogs;
}

/** Paragrafo del cursore, se è un blocco di primo livello (non in una tabella). */
function caretBlock(state: EditorState): { sectionIndex: number; index: number } | undefined {
  for (let s = 0; s < state.doc.sections.length; s++) {
    const index = state.doc.sections[s]!.blocks.findIndex((b) => b.type === 'paragraph' && b.id === state.selection.focus.paragraphId);
    if (index >= 0) return { sectionIndex: s, index };
  }
  return undefined;
}

export function createReferencesCommands(editor: Editor, options: ReferencesOptions = {}): (Command<any> | AsyncCommand<any>)[] {
  const layout = () => editor.services.layout;
  const pages = (doc: DocumentModel) => pageMap(doc, layout());
  const dialogs = options.dialogs;

  /** Aggiorna tutto ciò che dipende da numeri e pagine. */
  const refreshAll = (doc: DocumentModel, withIndex: boolean): DocumentModel => {
    let out = renumberCaptions(doc);
    out = updateCrossRefs(out, pages(out));
    const at = withIndex ? findIndex(out) : undefined;
    if (at) {
      out = writeIndex(out, at, layout());
      // L'indice può spostare le pagine: i numeri di pagina dei riferimenti si ricalcolano.
      out = updateCrossRefs(out, pages(out));
    }
    return out;
  };

  const keep = (state: EditorState, doc: DocumentModel, label: string): Transaction | null =>
    doc === state.doc ? null : { state: { doc, selection: state.selection }, meta: { label } };

  const insertCaptionCmd: Command<CaptionArgs> = {
    id: 'references.insertCaption',
    label: 'Inserisci didascalia',
    isEnabled: ({ state }, args) => canInsertCaption(state) && (!args || !!args.label?.trim()),
    run: ({ state }, args) => {
      const tr = insertCaption(state, args);
      if (!tr) return null;
      return { ...tr, state: { ...tr.state, doc: updateCrossRefs(tr.state.doc, pages(tr.state.doc)) } };
    },
  };

  const insertCrossRefCmd: Command<CrossRef | { ref: string }> = {
    id: 'references.insertCrossRef',
    label: 'Riferimento incrociato',
    isEnabled: ({ state }) => canInsertCrossRef(state),
    run: ({ state }, args) => {
      const ref = args && 'ref' in args ? parseCrossRef(args.ref) : args;
      return ref ? insertCrossRef(state, ref, pages(state.doc)) : null;
    },
  };

  const markEntryCmd: Command<MarkArgs> = {
    id: 'references.markEntry',
    label: 'Segna voce',
    run: ({ state }, args) => (args ? markEntry(state, args) : null),
  };

  const insertIndexCmd: Command<void> = {
    id: 'references.insertIndex',
    label: 'Inserisci indice',
    isEnabled: ({ state }) => !!findIndex(state.doc) || !!caretBlock(state),
    run: ({ state }) => {
      const existing = findIndex(state.doc);
      if (existing) return keep(state, writeIndex(state.doc, existing, layout()), 'Aggiorna indice');
      const at = caretBlock(state);
      if (!at) return null;
      const doc = writeIndex(state.doc, { sectionIndex: at.sectionIndex, start: at.index, end: at.index }, layout());
      return { state: { doc, selection: state.selection }, meta: { label: 'Inserisci indice' } };
    },
  };

  const updateIndexCmd: Command<void> = {
    id: 'references.updateIndex',
    label: 'Aggiorna indice',
    isEnabled: ({ state }) => !!findIndex(state.doc),
    run: ({ state }) => {
      const at = findIndex(state.doc);
      return at ? keep(state, writeIndex(state.doc, at, layout()), 'Aggiorna indice') : null;
    },
  };

  const updateFieldsCmd: Command<void> = {
    id: 'references.updateFields',
    label: 'Aggiorna tutto',
    run: ({ state }) => keep(state, refreshAll(state.doc, true), 'Aggiorna didascalie e riferimenti'),
  };

  // Le finestre: il comando puro si esegue sullo stato del momento in cui si conferma.
  const dialog = <T>(id: string, label: string, open: (state: EditorState) => Promise<T | null> | undefined, then: string): AsyncCommand => ({
    id,
    label,
    async: true,
    isEnabled: () => !!dialogs,
    run: async ({ state }) => {
      const choice = await open(state);
      if (choice) editor.execute(then, choice);
      return null;
    },
  });

  const captionDialog = dialog(
    'references.captionDialog',
    'Inserisci didascalia',
    (state) => dialogs?.caption({ labels: captionLabels(state.doc), next: (l) => nextCaptionNumber(editor.state.doc, l), defaultPosition }),
    'references.insertCaption',
  );
  const crossRefDialog = dialog(
    'references.crossRefDialog',
    'Riferimento incrociato',
    (state) => dialogs?.crossRef({ targets: listTargets(state.doc) }),
    'references.insertCrossRef',
  );
  const markEntryDialog = dialog(
    'references.markEntryDialog',
    'Segna voce',
    (state) => dialogs?.markEntry({ entry: suggestedEntry(state) }),
    'references.markEntry',
  );

  return [
    insertCaptionCmd,
    insertCrossRefCmd,
    markEntryCmd,
    insertIndexCmd,
    updateIndexCmd,
    updateFieldsCmd,
    captionDialog,
    crossRefDialog,
    markEntryDialog,
  ];
}

export function createReferencesModule(options: ReferencesOptions = {}): EditorModule {
  return {
    id: REFERENCES_MODULE_ID,
    activate(editor) {
      const offs = createReferencesCommands(editor, options).map((c) => editor.commands.register(c));
      offs.push(editor.addTransactionFilter(crossRefTypingFilter));
      return () => offs.forEach((off) => off());
    },
  };
}
