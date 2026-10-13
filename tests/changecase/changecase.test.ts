import { describe, expect, it } from 'vitest';
import { createDocument, createState, Editor, paragraph, paragraphText, textRun, type Paragraph } from '../../src/core';
import { convertCase, createChangeCaseModule, nextCycleMode, ribbonItems } from '../../src/changecase';
import { editingModule } from '../../src/editing';

function setup(p: Paragraph, from: number, to = from) {
  const editor = new Editor(createState(createDocument([p])));
  editor.use(editingModule);
  editor.use(createChangeCaseModule());
  editor.execute('edit.setSelection', {
    selection: { anchor: { paragraphId: p.id, offset: from }, focus: { paragraphId: p.id, offset: to } },
  });
  return editor;
}
const text = (editor: Editor) => paragraphText(editor.state.doc.sections[0]!.blocks[0] as Paragraph);

describe('Maiuscole/minuscole', () => {
  it('converte come le voci del menu di Word', () => {
    expect(convertCase('ciao MONDO. come va?', 'sentence')).toBe('Ciao mondo. Come va?');
    expect(convertCase('Ciao Mondo', 'lower')).toBe('ciao mondo');
    expect(convertCase('Ciao mondo', 'upper')).toBe('CIAO MONDO');
    expect(convertCase("l'albero di natale", 'title')).toBe("L'albero Di Natale");
    expect(convertCase('Ciao', 'toggle')).toBe('cIAO');
    // Le conversioni che cambierebbero la lunghezza lasciano il carattere com'è.
    expect(convertCase('straße', 'upper')).toBe('STRAßE');
  });

  it('Maiusc+F3 passa da MAIUSCOLO a minuscolo a Iniziali Maiuscole', () => {
    expect(nextCycleMode('Ciao mondo')).toBe('upper');
    expect(nextCycleMode('CIAO MONDO')).toBe('lower');
    expect(nextCycleMode('ciao mondo')).toBe('title');
  });

  it('mantiene formattazione e selezione', () => {
    const p = paragraph([textRun('ciao '), textRun('mondo', { bold: true })]);
    const editor = setup(p, 0, 10);
    editor.execute('changeCase.apply', { mode: 'upper' });
    expect(text(editor)).toBe('CIAO MONDO');
    const out = editor.state.doc.sections[0]!.blocks[0] as Paragraph;
    expect(out.children[1]).toMatchObject({ text: 'MONDO', format: { bold: true } });
    expect(editor.state.selection.focus.offset).toBe(10);
    editor.execute('changeCase.cycle');
    expect(text(editor)).toBe('ciao mondo');
    editor.undo();
    expect(text(editor)).toBe('CIAO MONDO');
  });

  it('con il cursore dentro una parola cambia solo quella parola', () => {
    const p = paragraph('uno due tre');
    const editor = setup(p, 5);
    editor.execute('changeCase.apply', { mode: 'upper' });
    expect(text(editor)).toBe('uno DUE tre');
  });

  it('compare nel gruppo Carattere della scheda Home', () => {
    expect(ribbonItems[0]).toMatchObject({ tab: 'home', group: 'Carattere', commandId: 'changeCase.apply' });
  });
});
