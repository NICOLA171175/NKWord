import { describe, expect, it } from 'vitest';
import { paragraph, type LayoutEngine } from '../../src/core';
import { focus, paras, select, setup } from './helpers';

describe('cursore e selezione', () => {
  it('si sposta per carattere attraversando i paragrafi', () => {
    const a = paragraph('ab');
    const b = paragraph('c');
    const editor = setup([a, b]);
    select(editor, [a, 2]);
    editor.execute('edit.moveCaret', { direction: 'right' });
    expect(focus(editor)).toEqual([b.id, 0]);
    editor.execute('edit.moveCaret', { direction: 'left' });
    expect(focus(editor)).toEqual([a.id, 2]);
    expect(editor.history.canUndo).toBe(false);
  });

  it('si sposta per parole come Word', () => {
    const p = paragraph('uno, due  tre');
    const editor = setup([p]);
    select(editor, [p, 0]);
    const stops: number[] = [];
    for (let i = 0; i < 4; i++) {
      editor.execute('edit.moveCaret', { direction: 'wordRight' });
      stops.push(focus(editor)[1]);
    }
    expect(stops).toEqual([3, 5, 10, 13]);
    editor.execute('edit.moveCaret', { direction: 'wordLeft' });
    expect(focus(editor)[1]).toBe(10);
  });

  it('estende la selezione e la chiude sul bordo con le frecce', () => {
    const p = paragraph('abcdef');
    const editor = setup([p]);
    select(editor, [p, 2]);
    editor.execute('edit.moveCaret', { direction: 'right', extend: true });
    editor.execute('edit.moveCaret', { direction: 'wordRight', extend: true });
    expect(editor.state.selection).toEqual({ anchor: { paragraphId: p.id, offset: 2 }, focus: { paragraphId: p.id, offset: 6 } });
    editor.execute('edit.moveCaret', { direction: 'left' });
    expect(focus(editor)).toEqual([p.id, 2]);
  });

  it('inizio/fine documento e seleziona tutto', () => {
    const a = paragraph('ab');
    const b = paragraph('cd');
    const editor = setup([a, b]);
    editor.execute('edit.moveCaret', { direction: 'docEnd' });
    expect(focus(editor)).toEqual([b.id, 2]);
    editor.execute('edit.moveCaret', { direction: 'docStart', extend: true });
    expect(editor.state.selection.anchor).toEqual({ paragraphId: b.id, offset: 2 });
    editor.execute('edit.selectAll');
    expect(editor.state.selection).toEqual({ anchor: { paragraphId: a.id, offset: 0 }, focus: { paragraphId: b.id, offset: 2 } });
  });

  it('inizio/fine riga: paragrafo senza impaginazione, riga visiva con impaginazione', () => {
    const p = paragraph('prima riga seconda');
    const editor = setup([p]);
    select(editor, [p, 13]);
    editor.execute('edit.moveCaret', { direction: 'lineStart' });
    expect(focus(editor)[1]).toBe(0);

    const line = (start: number, end: number) => ({ paragraphId: p.id, start, end, x: 0, y: 0, width: 0, height: 0, baseline: 0, runs: [] });
    const layout = {
      layout: (doc: unknown) => ({ doc, pages: [{ lines: [line(0, 11), line(11, 18)] }] }),
    } as unknown as LayoutEngine;
    editor.services.layout = layout;
    select(editor, [p, 13]);
    editor.execute('edit.moveCaret', { direction: 'lineStart' });
    expect(focus(editor)[1]).toBe(11);
    // Fine resta sulla riga (prima dello spazio finale), Maiusc+Fine seleziona anche lo spazio.
    select(editor, [p, 3]);
    editor.execute('edit.moveCaret', { direction: 'lineEnd' });
    expect(focus(editor)[1]).toBe(10);
    select(editor, [p, 3]);
    editor.execute('edit.moveCaret', { direction: 'lineEnd', extend: true });
    expect(focus(editor)[1]).toBe(11);
  });

  it('setSelection riporta gli offset dentro il paragrafo e ignora id sconosciuti', () => {
    const p = paragraph('abc');
    const editor = setup([p]);
    select(editor, [p, 99]);
    expect(focus(editor)).toEqual([p.id, 3]);
    expect(
      editor.execute('edit.setSelection', {
        selection: { anchor: { paragraphId: 'nessuno', offset: 0 }, focus: { paragraphId: p.id, offset: 0 } },
      }),
    ).toBe(false);
    expect(paras(editor)).toHaveLength(1);
  });
});
