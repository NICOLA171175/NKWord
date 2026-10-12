import { describe, expect, it } from 'vitest';
import { paragraph, textRun, type TextRun } from '../../src/core';
import { matchesInText } from '../../src/editing';
import { paras, select, setup, texts } from './helpers';

describe('trova', () => {
  it('opzioni maiuscole/minuscole e parola intera', () => {
    expect(matchesInText('Casa casale CASA', 'casa')).toEqual([[0, 4], [5, 9], [12, 16]]);
    expect(matchesInText('Casa casale CASA', 'casa', { matchCase: true })).toEqual([[5, 9]]);
    expect(matchesInText('Casa casale CASA', 'casa', { wholeWord: true })).toEqual([[0, 4], [12, 16]]);
    expect(matchesInText('a.b a+b', '.b')).toEqual([[1, 3]]);
  });

  it('seleziona la prossima occorrenza e ricomincia dall\'inizio', () => {
    const a = paragraph('gatto e cane');
    const b = paragraph('altro gatto');
    const editor = setup([a, b]);
    editor.execute('edit.find', { query: 'gatto' });
    expect(editor.state.selection).toEqual({ anchor: { paragraphId: a.id, offset: 0 }, focus: { paragraphId: a.id, offset: 5 } });
    editor.execute('edit.find', { query: 'gatto' });
    expect(editor.state.selection.anchor).toEqual({ paragraphId: b.id, offset: 6 });
    editor.execute('edit.find', { query: 'gatto' });
    expect(editor.state.selection.anchor).toEqual({ paragraphId: a.id, offset: 0 });
    expect(editor.execute('edit.find', { query: 'topo' })).toBe(false);
    expect(editor.canExecute('edit.find', { query: '' })).toBe(false);
  });

  it('sostituisci: prima trova, poi sostituisce e passa alla successiva', () => {
    const p = paragraph('uno uno uno');
    const editor = setup([p]);
    editor.execute('edit.replace', { query: 'uno', replacement: 'due' });
    expect(texts(editor)).toEqual(['uno uno uno']);
    editor.execute('edit.replace', { query: 'uno', replacement: 'due' });
    expect(texts(editor)).toEqual(['due uno uno']);
    expect(editor.state.selection.anchor.offset).toBe(4);
  });

  it('sostituisci tutto mantiene il formato e si annulla in un passo', () => {
    const p = paragraph([textRun('A '), textRun('rosso', { bold: true }), textRun(' e rosso')]);
    const editor = setup([p, paragraph('rosso')]);
    select(editor, [p, 0]);
    editor.execute('edit.replace', { query: 'rosso', replacement: 'blu', all: true });
    expect(texts(editor)).toEqual(['A blu e blu', 'blu']);
    expect((paras(editor)[0]!.children as TextRun[])[1]).toMatchObject({ text: 'blu', format: { bold: true } });
    editor.undo();
    expect(texts(editor)).toEqual(['A rosso e rosso', 'rosso']);
  });
});
