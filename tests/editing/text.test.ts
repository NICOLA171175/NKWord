import { describe, expect, it } from 'vitest';
import { createDocument, paragraph, table, textRun, type TextRun } from '../../src/core';
import { focus, paras, select, setup, texts } from './helpers';

describe('digitazione', () => {
  it('inserisce testo al cursore e lo unisce in un passo di annullamento', () => {
    const editor = setup([paragraph('ciao')]);
    select(editor, [paras(editor)[0]!, 4]);
    editor.execute('edit.insertText', { text: ' a' });
    editor.execute('edit.insertText', { text: ' tutti' });
    expect(texts(editor)).toEqual(['ciao a tutti']);
    expect(focus(editor)[1]).toBe(12);
    editor.undo();
    expect(texts(editor)).toEqual(['ciao']);
  });

  it('sostituisce la selezione mantenendo il formato del primo carattere selezionato', () => {
    const p = paragraph([textRun('uno '), textRun('due', { bold: true })]);
    const editor = setup([p]);
    select(editor, [p, 4], [p, 7]);
    editor.execute('edit.insertText', { text: 'tre' });
    expect(texts(editor)).toEqual(['uno tre']);
    const runs = paras(editor)[0]!.children as TextRun[];
    expect(runs[1]).toMatchObject({ text: 'tre', format: { bold: true } });
  });

  it('gestisce a capo e tabulazioni nel testo', () => {
    const editor = setup([paragraph('AB', { align: 'center' })]);
    select(editor, [paras(editor)[0]!, 1]);
    editor.execute('edit.insertText', { text: 'x\ty\nz' });
    expect(texts(editor)).toEqual(['Ax\ty', 'zB']);
    expect(paras(editor)[1]!.format.align).toBe('center');
    expect(focus(editor)).toEqual([paras(editor)[1]!.id, 1]);
  });

  it('applica la formattazione in attesa al testo digitato', () => {
    const editor = setup([paragraph('a')]);
    select(editor, [paras(editor)[0]!, 1]);
    editor.execute('edit.insertText', { text: ' ' });
    editor.execute('format.bold');
    expect(editor.isActive('format.bold')).toBe(true);
    editor.execute('edit.insertText', { text: 'b' });
    const runs = paras(editor)[0]!.children as TextRun[];
    expect(runs.map((r) => [r.text, !!r.format.bold])).toEqual([['a ', false], ['b', true]]);
    expect(editor.state.pendingFormat).toBeUndefined();
  });
});

describe('Invio e interruzioni', () => {
  it('divide il paragrafo e porta il cursore nel nuovo', () => {
    const p = paragraph('primasecondo');
    const editor = setup([p]);
    select(editor, [p, 5]);
    editor.execute('edit.splitParagraph');
    expect(texts(editor)).toEqual(['prima', 'secondo']);
    expect(paras(editor)[0]!.id).toBe(p.id);
    expect(focus(editor)).toEqual([paras(editor)[1]!.id, 0]);
  });

  it('dopo un titolo il paragrafo successivo usa lo stile "next"', () => {
    const p = paragraph('Titolo', {}, 'Heading1');
    const editor = setup([p]);
    select(editor, [p, 6]);
    editor.execute('edit.splitParagraph');
    expect(paras(editor)[1]!.styleId).toBeUndefined();
    // In mezzo al titolo invece lo stile resta.
    select(editor, [paras(editor)[0]!, 3]);
    editor.execute('edit.splitParagraph');
    expect(paras(editor)[1]!.styleId).toBe('Heading1');
  });

  it('inserisce interruzione di riga e tabulazione come inline', () => {
    const p = paragraph('ab');
    const editor = setup([p]);
    select(editor, [p, 1]);
    editor.execute('edit.insertLineBreak');
    editor.execute('edit.insertTab');
    expect(texts(editor)).toEqual(['a\n\tb']);
    expect(focus(editor)[1]).toBe(3);
  });
});

describe('cancellazione', () => {
  it('Backspace cancella un carattere, o una parola con unit "word"', () => {
    const p = paragraph('ciao bel mondo');
    const editor = setup([p]);
    select(editor, [p, 14]);
    editor.execute('edit.deleteBackward');
    expect(texts(editor)).toEqual(['ciao bel mond']);
    editor.execute('edit.deleteBackward', { unit: 'word' });
    expect(texts(editor)).toEqual(['ciao bel ']);
    editor.execute('edit.deleteBackward', { unit: 'word' });
    expect(texts(editor)).toEqual(['ciao ']);
  });

  it('non spezza le coppie surrogate (emoji)', () => {
    const p = paragraph('a😀');
    const editor = setup([p]);
    select(editor, [p, 3]);
    editor.execute('edit.deleteBackward');
    expect(texts(editor)).toEqual(['a']);
  });

  it('Backspace a inizio paragrafo unisce al precedente', () => {
    const a = paragraph('uno');
    const b = paragraph('due');
    const editor = setup([a, b]);
    select(editor, [b, 0]);
    editor.execute('edit.deleteBackward');
    expect(texts(editor)).toEqual(['unodue']);
    expect(focus(editor)).toEqual([a.id, 3]);
    expect(editor.execute('edit.deleteBackward', {})).toBe(true);
    select(editor, [paras(editor)[0]!, 0]);
    expect(editor.execute('edit.deleteBackward', {})).toBe(false);
  });

  it('Canc a fine paragrafo unisce il successivo; Ctrl+Canc cancella la parola', () => {
    const a = paragraph('uno');
    const b = paragraph('due tre');
    const editor = setup([a, b]);
    select(editor, [a, 3]);
    editor.execute('edit.deleteForward');
    expect(texts(editor)).toEqual(['unodue tre']);
    select(editor, [paras(editor)[0]!, 3]);
    editor.execute('edit.deleteForward', { unit: 'word' });
    expect(texts(editor)).toEqual(['unotre']);
  });

  it('non unisce un paragrafo con il contenuto di una tabella adiacente', () => {
    const t = table(1, 1, 100);
    const after = paragraph('dopo');
    const editor = setup([paragraph('prima'), t, after]);
    select(editor, [after, 0]);
    expect(editor.execute('edit.deleteBackward', {})).toBe(false);
  });

  it('cancella una selezione su più paragrafi, tabelle comprese', () => {
    const a = paragraph('alfa');
    const c = paragraph('gamma');
    const editor = setup([a, table(2, 2, 100), paragraph('beta'), c]);
    select(editor, [c, 2], [a, 2]);
    editor.execute('edit.deleteSelection');
    expect(texts(editor)).toEqual(['almma']);
    expect(editor.state.doc.sections[0]!.blocks).toHaveLength(1);
    expect(focus(editor)).toEqual([a.id, 2]);
  });

  it('se la selezione entra in una tabella svuota solo il testo selezionato', () => {
    const t = table(1, 1, 100);
    const a = paragraph('alfa');
    const editor = setup([a, t]);
    const cell = paras(editor)[1]!;
    editor.execute('edit.setSelection', {
      selection: { anchor: { paragraphId: a.id, offset: 1 }, focus: { paragraphId: cell.id, offset: 0 } },
    });
    editor.execute('edit.insertText', { text: 'X' });
    expect(texts(editor)).toEqual(['aX', '']);
    expect(editor.state.doc.sections[0]!.blocks).toHaveLength(2);
  });
});

describe('incolla frammento', () => {
  it('fonde il primo e l\'ultimo paragrafo con il testo intorno e rigenera gli id', () => {
    const p = paragraph('[]');
    const editor = setup([p]);
    select(editor, [p, 1]);
    const fragment = createDocument([paragraph('uno'), paragraph('due'), paragraph('tre')]);
    editor.execute('edit.insertFragment', { fragment });
    expect(texts(editor)).toEqual(['[uno', 'due', 'tre]']);
    const ids = new Set(paras(editor).map((x) => x.id));
    for (const f of fragment.sections[0]!.blocks) expect(ids.has(f.id)).toBe(false);
    expect(focus(editor)).toEqual([paras(editor)[2]!.id, 3]);
  });

  it('un frammento di un solo paragrafo si inserisce in linea', () => {
    const p = paragraph('ac');
    const editor = setup([p]);
    select(editor, [p, 1]);
    editor.execute('edit.insertFragment', { fragment: createDocument([paragraph([textRun('b', { italic: true })])]) });
    expect(texts(editor)).toEqual(['abc']);
    expect(paras(editor)).toHaveLength(1);
    expect(focus(editor)).toEqual([p.id, 2]);
  });

  it('incolla una tabella dividendo il paragrafo', () => {
    const p = paragraph('ab');
    const editor = setup([p]);
    select(editor, [p, 1]);
    editor.execute('edit.insertFragment', { fragment: createDocument([table(1, 2, 100)]) });
    const blocks = editor.state.doc.sections[0]!.blocks;
    expect(blocks.map((b) => b.type)).toEqual(['paragraph', 'table', 'paragraph']);
    expect(texts(editor)).toEqual(['a', '', '', 'b']);
  });
});
