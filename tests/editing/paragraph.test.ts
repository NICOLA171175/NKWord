import { describe, expect, it } from 'vitest';
import { paragraph, textRun, type TextRun } from '../../src/core';
import { INDENT_STEP, isAlignActive } from '../../src/editing';
import { paras, select, setup } from './helpers';

describe('formattazione paragrafo', () => {
  it('allinea tutti i paragrafi selezionati', () => {
    const a = paragraph('a');
    const b = paragraph('b');
    const c = paragraph('c');
    const editor = setup([a, b, c]);
    select(editor, [a, 0], [b, 1]);
    editor.execute('paragraph.align', { align: 'justify' });
    expect(paras(editor).map((p) => p.format.align)).toEqual(['justify', 'justify', undefined]);
    expect(isAlignActive(editor.state, 'justify')).toBe(true);
  });

  it('rientro: passi di 1,25 cm, mai negativo', () => {
    const p = paragraph('x', { indentLeft: 10 });
    const editor = setup([p]);
    editor.execute('paragraph.indent');
    expect(paras(editor)[0]!.format.indentLeft).toBeCloseTo(INDENT_STEP);
    editor.execute('paragraph.indent');
    expect(paras(editor)[0]!.format.indentLeft).toBeCloseTo(2 * INDENT_STEP);
    editor.execute('paragraph.outdent');
    editor.execute('paragraph.outdent');
    expect(paras(editor)[0]!.format.indentLeft).toBe(0);
    expect(editor.execute('paragraph.outdent')).toBe(false);
  });

  it('interlinea, spaziatura e formato generico', () => {
    const editor = setup([paragraph('x')]);
    editor.execute('paragraph.lineSpacing', { lineSpacing: { rule: 'multiple', value: 1.5 } });
    editor.execute('paragraph.spacing', { before: 6 });
    editor.execute('paragraph.setFormat', { format: { keepWithNext: true } });
    expect(paras(editor)[0]!.format).toEqual({
      lineSpacing: { rule: 'multiple', value: 1.5 },
      spaceBefore: 6,
      keepWithNext: true,
    });
  });
});

describe('stili', () => {
  it('applica uno stile di paragrafo e torna a Normale', () => {
    const p = paragraph('x', { align: 'center', outlineLevel: 2 });
    const editor = setup([p]);
    editor.execute('style.apply', { styleId: 'Heading1' });
    expect(paras(editor)[0]!.styleId).toBe('Heading1');
    expect(paras(editor)[0]!.format).toEqual({ align: 'center' });
    editor.execute('style.apply', { styleId: 'Normal' });
    expect(paras(editor)[0]!.styleId).toBeUndefined();
  });

  it('applica uno stile di carattere alla selezione', () => {
    const p = paragraph('uno due');
    const editor = setup([p]);
    const doc = editor.state.doc;
    editor.reset({
      ...editor.state,
      doc: {
        ...doc,
        styles: {
          ...doc.styles,
          styles: { ...doc.styles.styles, Strong: { id: 'Strong', name: 'Forte', type: 'character', charFormat: { bold: true } } },
        },
      },
    });
    select(editor, [p, 4], [p, 7]);
    editor.execute('style.apply', { styleId: 'Strong' });
    const runs = paras(editor)[0]!.children as TextRun[];
    expect(runs[1]).toEqual({ ...textRun('due'), styleId: 'Strong' });
    expect(editor.isActive('format.bold')).toBe(true);
  });

  it('non accetta stili inesistenti', () => {
    const editor = setup([paragraph('x')]);
    expect(editor.canExecute('style.apply', { styleId: 'Boh' })).toBe(false);
  });
});
