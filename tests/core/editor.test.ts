import { describe, expect, it } from 'vitest';
import { Editor, createDocument, findParagraph, paragraph, paragraphText, textRun, updateParagraph, type Command } from '../../src/core';

const append: Command<string> = {
  id: 'test.append',
  label: 'Aggiungi',
  run: ({ state }, text) => {
    const p = findParagraph(state.doc, state.selection.focus.paragraphId)!.paragraph;
    const doc = updateParagraph(state.doc, { ...p, children: [...p.children, textRun(text)] });
    return { state: { ...state, doc }, meta: { mergeKey: 'typing' } };
  },
};

function text(editor: Editor): string {
  return paragraphText(findParagraph(editor.state.doc, editor.state.selection.focus.paragraphId)!.paragraph);
}

describe('Editor', () => {
  it('esegue comandi, annulla e ripete', () => {
    const editor = new Editor();
    editor.commands.register(append);
    let changes = 0;
    editor.events.on('docChange', () => changes++);

    expect(editor.canExecute('edit.undo')).toBe(false);
    editor.execute('test.append', 'a');
    editor.history.seal();
    editor.execute('test.append', 'b');
    expect(text(editor)).toBe('ab');
    expect(changes).toBe(2);

    expect(editor.undo()).toBe(true);
    expect(text(editor)).toBe('a');
    expect(editor.redo()).toBe(true);
    expect(text(editor)).toBe('ab');
    editor.undo();
    editor.undo();
    expect(text(editor)).toBe('');
    expect(editor.undo()).toBe(false);
  });

  it('fonde le battiture consecutive in un solo passo', () => {
    const editor = new Editor();
    editor.commands.register(append);
    editor.execute('test.append', 'x');
    editor.execute('test.append', 'y');
    editor.undo();
    expect(text(editor)).toBe('');
  });

  it('rifiuta comandi duplicati e sconosciuti, segnala errori', () => {
    const editor = new Editor();
    editor.commands.register(append);
    expect(() => editor.commands.register(append)).toThrow();
    expect(() => editor.execute('nessuno')).toThrow();
    const errors: unknown[] = [];
    editor.events.on('error', (e) => errors.push(e.error));
    editor.commands.register({ id: 'test.boom', label: 'x', run: () => { throw new Error('boom'); } });
    expect(editor.execute('test.boom')).toBe(false);
    expect(errors).toHaveLength(1);
  });

  it('attiva i moduli e li disattiva in ordine inverso', () => {
    const editor = new Editor();
    const log: string[] = [];
    editor.use({ id: 'a', activate: (e) => { e.commands.register({ ...append, id: 'a.cmd' }); return () => log.push('a'); } });
    editor.use({ id: 'b', activate: () => () => log.push('b') });
    expect(editor.commands.has('a.cmd')).toBe(true);
    editor.destroy();
    expect(log).toEqual(['b', 'a']);
  });

  it('reset sostituisce il documento e azzera la cronologia', () => {
    const editor = new Editor();
    editor.commands.register(append);
    editor.execute('test.append', 'a');
    const doc = createDocument([paragraph('nuovo')]);
    editor.reset({ doc, selection: editor.state.selection });
    expect(editor.history.canUndo).toBe(false);
    expect(editor.state.doc).toBe(doc);
  });
});
