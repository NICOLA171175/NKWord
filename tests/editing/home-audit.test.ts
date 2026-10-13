import { describe, expect, it } from 'vitest';
import { paragraph, textRun, type TextRun } from '../../src/core';
import { paras, select, setup } from './helpers';

const runs = (editor: ReturnType<typeof setup>, i = 0) =>
  paras(editor)[i]!.children.map((c) => (c as TextRun).type === 'text' ? [(c as TextRun).text, (c as TextRun).format] : c.type);

describe('controllo della scheda Home: comportamento come in Word', () => {
  it('colore Automatico dopo il rosso: il testo scritto torna al colore predefinito', () => {
    const p = paragraph('nero ');
    const editor = setup([p]);
    select(editor, [p, 5]);
    editor.execute('format.color', { color: '#ff0000' });
    editor.execute('edit.insertText', { text: 'rosso ' });
    editor.execute('format.color', { color: null });
    expect(editor.state.pendingFormat).toHaveProperty('color', undefined);
    editor.execute('edit.insertText', { text: 'auto' });
    expect(runs(editor)).toEqual([
      ['nero ', {}],
      ['rosso ', { color: '#ff0000' }],
      ['auto', {}],
    ]);
  });

  it('Nessun colore dell\'evidenziatore e grassetto spento non lasciano formati inutili', () => {
    const p = paragraph([textRun('giallo', { highlight: '#ffff00', bold: true })]);
    const editor = setup([p]);
    select(editor, [p, 6]);
    editor.execute('format.highlight', { color: null });
    editor.execute('format.bold');
    editor.execute('edit.insertText', { text: ' niente' });
    expect(runs(editor)).toEqual([
      ['giallo', { highlight: '#ffff00', bold: true }],
      [' niente', {}],
    ]);
  });

  it('Invio dopo Automatico: il nuovo paragrafo non è rosso', () => {
    const p = paragraph([textRun('rosso', { color: '#ff0000' })]);
    const editor = setup([p]);
    select(editor, [p, 5]);
    editor.execute('format.color', { color: null });
    editor.execute('edit.splitParagraph');
    editor.execute('edit.insertText', { text: 'x' });
    expect(runs(editor, 1)).toEqual([['x', {}]]);
  });

  it('Centra due volte riporta il paragrafo a sinistra (anche Destra e Giustifica)', () => {
    const p = paragraph('testo');
    const editor = setup([p]);
    select(editor, [p, 0]);
    editor.execute('paragraph.align', { align: 'center' });
    expect(paras(editor)[0]!.format.align).toBe('center');
    editor.execute('paragraph.align', { align: 'center' });
    expect(paras(editor)[0]!.format.align).toBe('left');
    editor.execute('paragraph.align', { align: 'justify' });
    editor.execute('paragraph.align', { align: 'justify' });
    expect(paras(editor)[0]!.format.align).toBe('left');
    editor.execute('paragraph.align', { align: 'left' });
    expect(paras(editor)[0]!.format.align).toBe('left');
  });

  it('Invio su una voce di elenco vuota chiude l\'elenco (o risale di livello)', () => {
    const a = paragraph('uno', { list: { listId: 'l', level: 1 } });
    const editor = setup([a]);
    select(editor, [a, 3]);
    editor.execute('edit.splitParagraph');
    expect(paras(editor)).toHaveLength(2);
    editor.execute('edit.splitParagraph');
    expect(paras(editor)).toHaveLength(2);
    expect(paras(editor)[1]!.format.list).toEqual({ listId: 'l', level: 0 });
    editor.execute('edit.splitParagraph');
    expect(paras(editor)).toHaveLength(2);
    expect(paras(editor)[1]!.format.list).toBeUndefined();
  });

  it('Backspace all\'inizio di una voce di elenco toglie il punto elenco', () => {
    const a = paragraph('uno', { list: { listId: 'l', level: 0 } });
    const editor = setup([a]);
    select(editor, [a, 0]);
    editor.execute('edit.deleteBackward');
    expect(paras(editor)[0]!.format.list).toBeUndefined();
    expect(runs(editor)).toEqual([['uno', {}]]);
  });

  it('Aumenta e Riduci rientro in un elenco cambiano il livello', () => {
    const a = paragraph('uno', { list: { listId: 'l', level: 0 } });
    const editor = setup([a]);
    select(editor, [a, 1]);
    editor.execute('paragraph.indent');
    expect(paras(editor)[0]!.format).toEqual({ list: { listId: 'l', level: 1 } });
    editor.execute('paragraph.outdent');
    expect(paras(editor)[0]!.format).toEqual({ list: { listId: 'l', level: 0 } });
  });

  it('Ctrl+Spazio toglie solo il formato del carattere e lascia lo stile del paragrafo', () => {
    const p = paragraph([textRun('Titolo', { bold: true, color: '#ff0000' })], { align: 'center' }, 'Heading1');
    const editor = setup([p]);
    select(editor, [p, 0], [p, 6]);
    editor.execute('format.clear', { charOnly: true });
    const out = paras(editor)[0]!;
    expect(out.styleId).toBe('Heading1');
    expect(out.format).toEqual({ align: 'center' });
    expect(runs(editor)).toEqual([['Titolo', {}]]);
  });
});
