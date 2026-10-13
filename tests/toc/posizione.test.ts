import { describe, expect, it } from 'vitest';
import { Editor, caret, createDocument, createState, paragraph, paragraphText, type Paragraph } from '../../src/core';
import { createReferencesModule, findIndex } from '../../src/references';
import { findToc, tocModule } from '../../src/toc';

// Controllo dei comandi (Riferimenti): sommario e indice vanno dove si trova il cursore, come in Word.
function editorAt(texts: string[], index: number, offset: number): Editor {
  const blocks = texts.map((t) => paragraph(t, {}, t.startsWith('#') ? 'Heading1' : undefined));
  const doc = createDocument(blocks);
  const editor = new Editor({ ...createState(doc), selection: caret(blocks[index]!.id, offset) });
  editor.use(tocModule);
  editor.use(createReferencesModule({ dialogs: undefined }));
  return editor;
}

const texts = (e: Editor) => e.state.doc.sections[0]!.blocks.map((b) => paragraphText(b as Paragraph));

describe('posizione di sommario e indice', () => {
  it('con il cursore alla fine di un paragrafo il sommario va dopo, non prima', () => {
    const editor = editorAt(['# Capitolo', 'Fine del testo'], 1, 'Fine del testo'.length);
    editor.execute('toc.insert');
    const at = findToc(editor.state.doc)!;
    expect(at.start).toBe(2);
    expect(texts(editor)[1]).toBe('Fine del testo');
    // In fondo al documento resta un paragrafo vuoto dove continuare a scrivere.
    expect(texts(editor).at(-1)).toBe('');
    expect(at.end).toBe(editor.state.doc.sections[0]!.blocks.length - 1);
  });

  it("con il cursore all'inizio del paragrafo il sommario va prima", () => {
    const editor = editorAt(['# Capitolo', 'Testo'], 1, 0);
    editor.execute('toc.insert');
    expect(findToc(editor.state.doc)!.start).toBe(1);
    expect(texts(editor).at(-1)).toBe('Testo');
  });

  it("l'indice segue la stessa regola", () => {
    const editor = editorAt(['Uno', 'Due'], 1, 3);
    editor.execute('references.insertIndex');
    expect(findIndex(editor.state.doc)!.start).toBe(2);
    expect(texts(editor).slice(0, 2)).toEqual(['Uno', 'Due']);
    expect(texts(editor).at(-1)).toBe('');
    editor.undo();
    expect(texts(editor)).toEqual(['Uno', 'Due']);
  });
});
