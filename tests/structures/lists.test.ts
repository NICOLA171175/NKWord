import { describe, expect, it } from 'vitest';
import { caret, paragraph, paragraphs, type Paragraph } from '../../src/core';
import { editorWith, select } from './helpers';

function paras(editor: ReturnType<typeof editorWith>): Paragraph[] {
  return [...paragraphs(editor.state.doc)].map((l) => l.paragraph);
}

describe('elenchi', () => {
  it('attiva e disattiva l\'elenco puntato sui paragrafi selezionati', () => {
    const a = paragraph('a', { indentLeft: 36 });
    const b = paragraph('b');
    const c = paragraph('c');
    const editor = editorWith([a, b, c], { anchor: { paragraphId: a.id, offset: 0 }, focus: { paragraphId: b.id, offset: 1 } });

    expect(editor.isActive('list.toggleBullet')).toBe(false);
    expect(editor.execute('list.toggleBullet')).toBe(true);
    const [pa, pb, pc] = paras(editor);
    expect(pa!.format.list).toEqual({ listId: pb!.format.list!.listId, level: 0 });
    expect(pa!.format.indentLeft).toBeUndefined();
    expect(pc!.format.list).toBeUndefined();
    expect(editor.state.doc.lists[pa!.format.list!.listId]!.levels[0]!.format).toBe('bullet');
    expect(editor.isActive('list.toggleBullet')).toBe(true);
    expect(editor.isActive('list.toggleNumbered')).toBe(false);

    editor.execute('list.toggleBullet');
    expect(paras(editor).every((p) => !p.format.list)).toBe(true);
    editor.undo();
    expect(paras(editor)[0]!.format.list).toBeTruthy();
  });

  it('crea un elenco numerato con il formato scelto e lo toglie', () => {
    const a = paragraph('a');
    const b = paragraph('b');
    const editor = editorWith([a, b], { anchor: { paragraphId: a.id, offset: 0 }, focus: { paragraphId: a.id, offset: 0 } });
    editor.execute('list.toggleNumbered', { format: 'upperRoman' });
    const listId = paras(editor)[0]!.format.list!.listId;
    expect(editor.state.doc.lists[listId]!.levels[0]!.format).toBe('upperRoman');
    expect(editor.state.doc.lists[listId]!.levels[1]!.text).toBe('%2.');

    // Di nuovo con lo stesso formato: lo toglie.
    editor.execute('list.toggleNumbered', { format: 'upperRoman' });
    expect(paras(editor)[0]!.format.list).toBeNull();
  });

  it('riusa l\'elenco del paragrafo precedente dello stesso tipo', () => {
    const a = paragraph('a');
    const b = paragraph('b');
    const editor = editorWith([a, b], { anchor: { paragraphId: a.id, offset: 0 }, focus: { paragraphId: a.id, offset: 0 } });
    editor.execute('list.toggleNumbered');
    select(editor, caret(b.id, 0));
    editor.execute('list.toggleNumbered');
    const [pa, pb] = paras(editor);
    expect(pb!.format.list!.listId).toBe(pa!.format.list!.listId);
    expect(Object.keys(editor.state.doc.lists)).toHaveLength(1);
    // Un puntato dopo un numerato crea un elenco nuovo.
    editor.execute('list.toggleBullet');
    expect(paras(editor)[1]!.format.list!.listId).not.toBe(pa!.format.list!.listId);
  });

  it('cambia livello entro 0..8', () => {
    const a = paragraph('a');
    const editor = editorWith([a]);
    expect(editor.canExecute('list.increaseLevel')).toBe(false);
    editor.execute('list.toggleBullet');
    expect(editor.execute('list.decreaseLevel')).toBe(false);
    for (let i = 0; i < 10; i++) editor.execute('list.increaseLevel');
    expect(paras(editor)[0]!.format.list!.level).toBe(8);
    editor.execute('list.decreaseLevel');
    expect(paras(editor)[0]!.format.list!.level).toBe(7);
  });
});
