import { describe, expect, it } from 'vitest';
import { paragraph, textRun, type TextRun } from '../../src/core';
import { paras, select, setup, texts } from './helpers';

const runs = (p: { children: readonly unknown[] }) => p.children as TextRun[];

describe('formattazione carattere', () => {
  it('grassetto sulla selezione, poi lo toglie', () => {
    const p = paragraph('uno due tre');
    const editor = setup([p]);
    select(editor, [p, 4], [p, 7]);
    expect(editor.isActive('format.bold')).toBe(false);
    editor.execute('format.bold');
    expect(runs(paras(editor)[0]!).map((r) => [r.text, r.format.bold])).toEqual([
      ['uno ', undefined],
      ['due', true],
      [' tre', undefined],
    ]);
    expect(editor.isActive('format.bold')).toBe(true);
    editor.execute('format.bold');
    // Tornato come prima: i run si riuniscono e non resta bold:false.
    expect(runs(paras(editor)[0]!)).toEqual([textRun('uno due tre')]);
  });

  it('selezione mista: il primo clic attiva per tutti', () => {
    const p = paragraph([textRun('ab', { italic: true }), textRun('cd')]);
    const editor = setup([p]);
    select(editor, [p, 0], [p, 4]);
    editor.execute('format.italic');
    expect(runs(paras(editor)[0]!)).toEqual([textRun('abcd', { italic: true })]);
  });

  it('con il cursore dentro una parola formatta la parola intera', () => {
    const p = paragraph('ciao mondo');
    const editor = setup([p]);
    select(editor, [p, 7]);
    editor.execute('format.underline');
    expect(runs(paras(editor)[0]!)[1]).toMatchObject({ text: 'mondo', format: { underline: 'single' } });
    expect(editor.state.selection.focus.offset).toBe(7);
  });

  it('con lo stile del titolo, togliere il grassetto scrive bold:false', () => {
    const p = paragraph('Titolo', {}, 'Heading1');
    const editor = setup([p]);
    select(editor, [p, 0], [p, 6]);
    expect(editor.isActive('format.bold')).toBe(true);
    editor.execute('format.bold');
    expect(runs(paras(editor)[0]!)[0]!.format).toEqual({ bold: false });
  });

  it('apice e pedice si escludono a vicenda', () => {
    const p = paragraph('x2');
    const editor = setup([p]);
    select(editor, [p, 1], [p, 2]);
    editor.execute('format.superscript');
    expect(runs(paras(editor)[0]!)[1]!.format.verticalAlign).toBe('superscript');
    editor.execute('format.subscript');
    expect(runs(paras(editor)[0]!)[1]!.format.verticalAlign).toBe('subscript');
    expect(editor.isActive('format.superscript')).toBe(false);
  });

  it('carattere, dimensione, colore, evidenziazione', () => {
    const p = paragraph('testo');
    const editor = setup([p]);
    select(editor, [p, 0], [p, 5]);
    editor.execute('format.fontFamily', { fontFamily: 'Arial' });
    editor.execute('format.fontSize', { fontSize: 14 });
    editor.execute('format.color', { color: '#ff0000' });
    editor.execute('format.highlight', { color: '#ffff00' });
    expect(runs(paras(editor)[0]!)[0]!.format).toEqual({
      fontFamily: 'Arial',
      fontSize: 14,
      color: '#ff0000',
      highlight: '#ffff00',
    });
    editor.execute('format.highlight', { color: null });
    expect(runs(paras(editor)[0]!)[0]!.format.highlight).toBeUndefined();
    expect(editor.canExecute('format.fontSize', { fontSize: 0 })).toBe(false);
  });

  it('aumenta e riduci dimensione seguono la scala di Word, run per run', () => {
    const p = paragraph([textRun('a'), textRun('b', { fontSize: 72 })]);
    const editor = setup([p]);
    select(editor, [p, 0], [p, 2]);
    editor.execute('format.growFont');
    expect(runs(paras(editor)[0]!).map((r) => r.format.fontSize)).toEqual([12, 80]);
    editor.execute('format.shrinkFont');
    // 11 è la dimensione dello stile: la formattazione diretta sparisce.
    expect(runs(paras(editor)[0]!).map((r) => r.format.fontSize)).toEqual([undefined, 72]);
  });

  it('format.setChar e cancella formattazione', () => {
    const p = paragraph([textRun('link', { link: 'https://x.it' }), textRun(' testo')], { align: 'center' }, 'Heading2');
    const editor = setup([p]);
    select(editor, [p, 0], [p, 10]);
    editor.execute('format.setChar', { format: { smallCaps: true, color: '#00ff00' } });
    expect(runs(paras(editor)[0]!)[1]!.format).toMatchObject({ smallCaps: true, color: '#00ff00' });
    editor.execute('format.clear');
    const q = paras(editor)[0]!;
    expect(q.styleId).toBeUndefined();
    expect(q.format).toEqual({});
    expect(runs(q).map((r) => r.format)).toEqual([{ link: 'https://x.it' }, {}]);
  });

  it('collegamenti: aggiunge, rimuove e inserisce l\'indirizzo se non c\'è testo', () => {
    const p = paragraph('vai qui');
    const editor = setup([p]);
    select(editor, [p, 5]);
    editor.execute('format.link', { url: 'https://a.it' });
    expect(runs(paras(editor)[0]!)[1]).toMatchObject({ text: 'qui', format: { link: 'https://a.it' } });
    editor.execute('format.link', { url: null });
    expect(runs(paras(editor)[0]!)).toEqual([textRun('vai qui')]);
    select(editor, [paras(editor)[0]!, 7]);
    editor.execute('edit.insertText', { text: ' ' });
    editor.execute('format.link', { url: 'https://b.it' });
    expect(texts(editor)).toEqual(['vai qui https://b.it']);
    expect(editor.state.selection.focus.offset).toBe(20);
  });

  it('con selezione vuota fuori da una parola il formato resta in attesa e non entra in cronologia', () => {
    const p = paragraph('a ');
    const editor = setup([p]);
    select(editor, [p, 2]);
    editor.execute('format.italic');
    expect(editor.state.pendingFormat).toEqual({ italic: true });
    expect(editor.history.canUndo).toBe(false);
    editor.execute('edit.moveCaret', { direction: 'left' });
    expect(editor.state.pendingFormat).toBeUndefined();
  });
});
