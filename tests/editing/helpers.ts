import {
  Editor,
  createDocument,
  createState,
  paragraphText,
  paragraphs,
  type Block,
  type Paragraph,
  type Selection,
} from '../../src/core';
import { editingModule } from '../../src/editing';

export function setup(blocks: Block[]): Editor {
  const editor = new Editor(createState(createDocument(blocks)));
  editor.use(editingModule);
  return editor;
}

export function paras(editor: Editor): Paragraph[] {
  return [...paragraphs(editor.state.doc)].map((l) => l.paragraph);
}

export function texts(editor: Editor): string[] {
  return paras(editor).map(paragraphText);
}

export function select(editor: Editor, a: [Paragraph, number], f: [Paragraph, number] = a): void {
  const selection: Selection = {
    anchor: { paragraphId: a[0].id, offset: a[1] },
    focus: { paragraphId: f[0].id, offset: f[1] },
  };
  editor.execute('edit.setSelection', { selection });
}

export function focus(editor: Editor): [string, number] {
  const f = editor.state.selection.focus;
  return [f.paragraphId, f.offset];
}
