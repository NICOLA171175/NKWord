import {
  Editor,
  createDocument,
  createState,
  findParagraph,
  paragraphText,
  type Block,
  type DocumentModel,
  type Selection,
  type Table,
} from '../../src/core';
import { structuresModule } from '../../src/structures';

export function editorWith(blocks: Block[], selection?: Selection): Editor {
  const doc = createDocument(blocks);
  const state = createState(doc);
  const editor = new Editor(selection ? { ...state, selection } : state);
  editor.use(structuresModule);
  editor.events.on('error', (e) => {
    throw e.error;
  });
  return editor;
}

export function focusText(editor: Editor): string {
  return paragraphText(findParagraph(editor.state.doc, editor.state.selection.focus.paragraphId)!.paragraph);
}

export function onlyTable(doc: DocumentModel): Table {
  const t = doc.sections.flatMap((s) => s.blocks).find((b) => b.type === 'table');
  if (!t || t.type !== 'table') throw new Error('nessuna tabella');
  return t;
}

/** Testo di ogni cella, riga per riga (le celle 'continue' come '^'). */
export function grid(t: Table): string[][] {
  return t.rows.map((r) =>
    r.cells.map((c) => {
      const text = c.blocks.map((b) => (b.type === 'paragraph' ? paragraphText(b) : '[t]')).join('|');
      const span = c.colSpan && c.colSpan > 1 ? `*${c.colSpan}` : '';
      return (c.vMerge === 'continue' ? '^' : text) + span;
    }),
  );
}

/** Sposta la selezione senza passare dal modulo di editing. */
export function select(editor: Editor, selection: Selection): void {
  editor.apply({ state: { ...editor.state, selection }, meta: { addToHistory: false } });
}

export function range(a: string, ao: number, b: string, bo: number): Selection {
  return { anchor: { paragraphId: a, offset: ao }, focus: { paragraphId: b, offset: bo } };
}
