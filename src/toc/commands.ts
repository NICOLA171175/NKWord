/**
 * Comandi del sommario. Non fanno parte del catalogo di `src/core`
 * (`CommandArgs`): sono registrati per id come gli altri e la barra
 * multifunzione li trova tramite `ribbonItems` (vedi `index.ts`).
 */

import {
  caret,
  findParagraph,
  type Command,
  type DocumentModel,
  type Editor,
  type EditorState,
  type Selection,
} from '../core';
import { findToc, removeToc, writeToc } from './toc';

export interface TocCommandArgs {
  /** Inserisce il sommario prima del paragrafo del cursore; se esiste già lo aggiorna. */
  'toc.insert': { levels?: number } | void;
  /** Rigenera voci e numeri di pagina del sommario esistente. */
  'toc.update': void;
  'toc.remove': void;
}

export type TocCommandId = keyof TocCommandArgs;

/** Il cursore resta dov'era; se il suo paragrafo è sparito va al primo paragrafo dopo il sommario. */
function keepSelection(doc: DocumentModel, sel: Selection, fallback: { sectionIndex: number; index: number }): Selection {
  if (findParagraph(doc, sel.anchor.paragraphId) && findParagraph(doc, sel.focus.paragraphId)) return sel;
  const blocks = doc.sections[fallback.sectionIndex]!.blocks;
  const block = blocks[Math.min(fallback.index, blocks.length - 1)]!;
  const target = block.type === 'paragraph' ? block : undefined;
  if (target) return caret(target.id, 0);
  for (const s of doc.sections) for (const b of s.blocks) if (b.type === 'paragraph') return caret(b.id, 0);
  return sel;
}

function transaction(state: EditorState, doc: DocumentModel, fallback: { sectionIndex: number; index: number }, label: string) {
  return { state: { doc, selection: keepSelection(doc, state.selection, fallback) }, meta: { label } };
}

/** Paragrafo del cursore, se è un blocco di primo livello (non in una tabella). */
function caretBlock(state: EditorState): { sectionIndex: number; index: number } | undefined {
  const loc = findParagraph(state.doc, state.selection.focus.paragraphId);
  if (!loc || loc.path.length !== 1) return undefined;
  return { sectionIndex: loc.sectionIndex, index: loc.path[0]! };
}

export function createTocCommands(editor: Editor): Command<any>[] {
  const layout = () => editor.services.layout;

  const insert: Command<TocCommandArgs['toc.insert']> = {
    id: 'toc.insert',
    label: 'Sommario',
    isEnabled: ({ state }) => !!findToc(state.doc) || !!caretBlock(state),
    run: ({ state }, args) => {
      const levels = args?.levels;
      const existing = findToc(state.doc);
      if (existing) {
        const doc = writeToc(state.doc, existing, { levels: levels ?? existing.levels, layout: layout() });
        return transaction(state, doc, { sectionIndex: existing.sectionIndex, index: existing.start }, 'Aggiorna sommario');
      }
      const at = caretBlock(state);
      if (!at) return null;
      const doc = writeToc(state.doc, { sectionIndex: at.sectionIndex, start: at.index, end: at.index }, { levels, layout: layout() });
      return transaction(state, doc, at, 'Inserisci sommario');
    },
  };

  const update: Command<void> = {
    id: 'toc.update',
    label: 'Aggiorna sommario',
    isEnabled: ({ state }) => !!findToc(state.doc),
    run: ({ state }) => {
      const at = findToc(state.doc);
      if (!at) return null;
      const doc = writeToc(state.doc, at, { levels: at.levels, layout: layout() });
      if (sameToc(state.doc, doc, at)) return null;
      return transaction(state, doc, { sectionIndex: at.sectionIndex, index: at.start }, 'Aggiorna sommario');
    },
  };

  const remove: Command<void> = {
    id: 'toc.remove',
    label: 'Rimuovi sommario',
    isEnabled: ({ state }) => !!findToc(state.doc),
    run: ({ state }) => {
      const at = findToc(state.doc);
      if (!at) return null;
      const doc = removeToc(state.doc, at);
      return transaction(state, doc, { sectionIndex: at.sectionIndex, index: at.start }, 'Rimuovi sommario');
    },
  };

  return [insert, update, remove];
}

/** Il sommario rigenerato è identico al precedente (stesso testo, stessi stili)? */
function sameToc(before: DocumentModel, after: DocumentModel, at: { sectionIndex: number; start: number; end: number }): boolean {
  const now = findToc(after);
  if (!now || now.end - now.start !== at.end - at.start || before.styles !== after.styles) return false;
  const a = before.sections[at.sectionIndex]!.blocks.slice(at.start, at.end);
  const b = after.sections[now.sectionIndex]!.blocks.slice(now.start, now.end);
  const strip = (x: unknown) => JSON.stringify(x, (k, v) => (k === 'id' ? undefined : v));
  return strip(a) === strip(b);
}
