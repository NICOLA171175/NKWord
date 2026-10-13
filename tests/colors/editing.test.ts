import { describe, expect, it } from 'vitest';
import { createDocument, Editor, createState, paragraph, textRun, type Block, type TextRun } from '../../src/core';
import { DEFAULT_BORDER, editingModule, isBorderActive } from '../../src/editing';
import { paras, select, setup } from '../editing/helpers';

const runs = (p: { children: readonly unknown[] }) => p.children as TextRun[];

describe('colore del testo', () => {
  it('Automatico toglie il colore diretto e torna a quello dello stile', () => {
    const p = paragraph('uno due');
    const editor = setup([p]);
    select(editor, [p, 0], [p, 3]);
    editor.execute('format.color', { color: '#c00000' });
    expect(runs(paras(editor)[0]!)[0]!.format.color).toBe('#c00000');
    editor.execute('format.color', { color: null });
    expect(runs(paras(editor)[0]!).map((r) => r.format)).toEqual([{}]);
    expect('color' in runs(paras(editor)[0]!)[0]!.format).toBe(false);
  });

  it('con il cursore fuori da una parola diventa formato in attesa', () => {
    const p = paragraph('uno ');
    const editor = setup([p]);
    select(editor, [p, 4]);
    editor.execute('format.color', { color: '#0070c0' });
    editor.execute('format.color', { color: null });
    expect(editor.state.pendingFormat).toEqual({});
    editor.execute('format.highlight', { color: '#00ff00' });
    editor.execute('edit.insertText', { text: 'x' });
    expect(runs(paras(editor)[0]!).at(-1)).toMatchObject({ text: 'x', format: { highlight: '#00ff00' } });
  });
});

describe('sfondo del paragrafo', () => {
  it('colora tutti i paragrafi selezionati e si toglie con Nessun colore', () => {
    const a = paragraph('a');
    const b = paragraph('b');
    const c = paragraph('c');
    const editor = setup([a, b, c]);
    select(editor, [a, 0], [b, 1]);
    editor.execute('paragraph.shading', { color: '#d9d9d9' });
    expect(paras(editor).map((p) => p.format.shading)).toEqual(['#d9d9d9', '#d9d9d9', undefined]);
    editor.execute('paragraph.shading', { color: null });
    expect(paras(editor).map((p) => 'shading' in p.format)).toEqual([false, false, false]);
    expect(editor.execute('paragraph.shading', { color: null })).toBe(false);
  });

  it('Nessun colore copre lo sfondo dato dallo stile', () => {
    const base = createDocument([{ ...paragraph('a'), styleId: 'Box' }] as Block[]);
    const doc = {
      ...base,
      styles: { ...base.styles, styles: { ...base.styles.styles, Box: { id: 'Box', name: 'Box', type: 'paragraph' as const, paraFormat: { shading: '#ffff00' } } } },
    };
    const editor = new Editor(createState(doc));
    editor.use(editingModule);
    editor.execute('paragraph.shading', { color: null });
    expect(paras(editor)[0]!.format.shading).toBeNull();
  });
});

describe('bordi del paragrafo', () => {
  it('un lato si attiva e si disattiva su tutta la selezione', () => {
    const a = paragraph('a');
    const b = paragraph('b');
    const editor = setup([a, b]);
    select(editor, [a, 0], [b, 1]);
    editor.execute('paragraph.borders', { sides: 'bottom' });
    expect(paras(editor).map((p) => p.format.borders)).toEqual([{ bottom: DEFAULT_BORDER }, { bottom: DEFAULT_BORDER }]);
    expect(isBorderActive(editor.state, 'bottom')).toBe(true);
    editor.execute('paragraph.borders', { sides: 'top' });
    expect(paras(editor)[0]!.format.borders).toEqual({ top: DEFAULT_BORDER, bottom: DEFAULT_BORDER });
    editor.execute('paragraph.borders', { sides: 'bottom' });
    expect(paras(editor)[1]!.format.borders).toEqual({ top: DEFAULT_BORDER });
  });

  it('Bordi esterni mette i quattro lati, Nessun bordo li toglie', () => {
    const p = paragraph('x');
    const editor = setup([p]);
    editor.execute('paragraph.borders', { sides: 'outside' });
    expect(isBorderActive(editor.state, 'outside')).toBe(true);
    expect(Object.keys(paras(editor)[0]!.format.borders!).sort()).toEqual(['bottom', 'left', 'right', 'top']);
    editor.execute('paragraph.borders', { sides: 'none' });
    expect(paras(editor)[0]!.format.borders).toBeUndefined();
    expect(isBorderActive(editor.state, 'none')).toBe(true);
  });

  it('se metà selezione ha già il lato, lo aggiunge a tutti', () => {
    const a = paragraph('a', { borders: { left: DEFAULT_BORDER } });
    const b = paragraph('b');
    const editor = setup([a, b]);
    select(editor, [a, 0], [b, 1]);
    editor.execute('paragraph.borders', { sides: 'left' });
    expect(paras(editor).map((p) => p.format.borders)).toEqual([{ left: DEFAULT_BORDER }, { left: DEFAULT_BORDER }]);
  });

  it('Cancella formattazione toglie colori, sfondo e bordi ma, come Word, non l\'evidenziatore', () => {
    const p = paragraph([textRun('testo', { color: '#ff0000', highlight: '#ffff00' })], {
      shading: '#d9d9d9',
      borders: { top: DEFAULT_BORDER },
    });
    const editor = setup([p]);
    select(editor, [p, 0], [p, 5]);
    editor.execute('format.clear');
    const out = paras(editor)[0]!;
    expect(out.format).toEqual({});
    expect(runs(out)[0]!.format).toEqual({ highlight: '#ffff00' });
  });
});
