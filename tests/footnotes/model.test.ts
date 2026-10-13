import { describe, expect, it } from 'vitest';
import {
  COMMAND_IDS,
  COMMAND_OWNERS,
  Editor,
  caret,
  createDocument,
  createState,
  firstParagraph,
  footnoteNumbers,
  paragraph,
  paragraphText,
  textRun,
  type EditorState,
  type Paragraph,
} from '../../src/core';
import {
  FOOTNOTE_TEXT_STYLE,
  footnoteAtSelection,
  footnoteRefs,
  footnoteText,
  footnotesModule,
  insertFootnote,
  insertInlineAt,
} from '../../src/footnotes';

function editorWith(text: string, offset = text.length): Editor {
  const doc = createDocument([paragraph([textRun(text, { bold: true })])]);
  const state = createState(doc);
  const editor = new Editor({ ...state, selection: caret(firstParagraph(doc).id, offset) });
  editor.use(footnotesModule);
  return editor;
}

const body = (state: EditorState) => firstParagraph(state.doc);

describe('modulo note a piè di pagina', () => {
  it('registra tutti e soli i comandi assegnati a "footnotes"', () => {
    const editor = new Editor();
    const before = new Set(editor.commands.list().map((c) => c.id));
    editor.use(footnotesModule);
    const added = editor.commands.list().map((c) => c.id).filter((id) => !before.has(id)).sort();
    expect(added).toEqual(COMMAND_IDS.filter((id) => COMMAND_OWNERS[id] === 'footnotes').sort());
  });
});

describe('footnote.insert', () => {
  it('inserisce il richiamo al cursore e crea la nota con il suo testo', () => {
    const editor = editorWith('Ciao mondo', 4);
    expect(editor.execute('footnote.insert', { text: 'Una nota' })).toBe(true);
    const { doc, selection } = editor.state;
    const p = body(editor.state);
    expect(p.children.map((c) => c.type)).toEqual(['text', 'footnoteRef', 'text']);
    expect(paragraphText(p)).toBe('Ciao￼ mondo');
    const ref = p.children[1]!;
    if (ref.type !== 'footnoteRef') throw new Error('richiamo mancante');
    // Il richiamo prende il formato del testo che lo precede.
    expect(ref.format).toEqual({ bold: true });
    expect(footnoteText(doc, ref.noteId)).toBe('Una nota');
    expect(doc.footnotes![ref.noteId]!.blocks[0]).toMatchObject({ styleId: FOOTNOTE_TEXT_STYLE });
    expect(doc.styles.styles[FOOTNOTE_TEXT_STYLE]).toBeDefined();
    expect(selection.focus).toEqual({ paragraphId: p.id, offset: 5 });
  });

  it('con una selezione mette il richiamo alla fine e lascia il testo', () => {
    const editor = editorWith('Ciao mondo');
    const id = body(editor.state).id;
    editor.apply({ state: { ...editor.state, selection: { anchor: { paragraphId: id, offset: 9 }, focus: { paragraphId: id, offset: 0 } } } });
    editor.execute('footnote.insert', { text: 'x' });
    expect(paragraphText(body(editor.state))).toBe('Ciao mond￼o');
  });

  it('numera le note in ordine di lettura, anche se inserite in disordine', () => {
    const editor = editorWith('abcdef', 6);
    editor.execute('footnote.insert', { text: 'seconda' });
    editor.apply({ state: { ...editor.state, selection: caret(body(editor.state).id, 2) } });
    editor.execute('footnote.insert', { text: 'prima' });
    const { doc } = editor.state;
    const numbers = footnoteNumbers(doc);
    const refs = footnoteRefs(doc);
    expect(refs.map((r) => footnoteText(doc, r.noteId))).toEqual(['prima', 'seconda']);
    expect(refs.map((r) => numbers.get(r.noteId))).toEqual([1, 2]);
  });

  it('il testo su più righe diventa più paragrafi', () => {
    const editor = editorWith('a');
    editor.execute('footnote.insert', { text: 'riga 1\nriga 2' });
    const [ref] = footnoteRefs(editor.state.doc);
    expect(editor.state.doc.footnotes![ref!.noteId]!.blocks).toHaveLength(2);
    expect(footnoteText(editor.state.doc, ref!.noteId)).toBe('riga 1\nriga 2');
  });

  it('si annulla con un solo passo', () => {
    const editor = editorWith('a');
    const before = editor.state.doc;
    editor.execute('footnote.insert', { text: 'n' });
    editor.undo();
    expect(editor.state.doc).toBe(before);
  });
});

describe('footnote.edit e footnote.delete', () => {
  it('trovano la nota accanto al cursore', () => {
    const editor = editorWith('abc', 1);
    expect(editor.canExecute('footnote.edit', { text: 'x' })).toBe(false);
    editor.execute('footnote.insert', { text: 'vecchio' });
    const [ref] = footnoteRefs(editor.state.doc);
    // Subito dopo il richiamo (posizione dopo l'inserimento) e subito prima.
    expect(footnoteAtSelection(editor.state)).toBe(ref!.noteId);
    editor.apply({ state: { ...editor.state, selection: caret(body(editor.state).id, 1) } });
    expect(footnoteAtSelection(editor.state)).toBe(ref!.noteId);
    editor.apply({ state: { ...editor.state, selection: caret(body(editor.state).id, 4) } });
    expect(footnoteAtSelection(editor.state)).toBeUndefined();
  });

  it('modifica il testo della nota', () => {
    const editor = editorWith('abc');
    editor.execute('footnote.insert', { text: 'vecchio' });
    expect(editor.execute('footnote.edit', { text: 'nuovo' })).toBe(true);
    const [ref] = footnoteRefs(editor.state.doc);
    expect(footnoteText(editor.state.doc, ref!.noteId)).toBe('nuovo');
  });

  it('elimina richiamo e nota e tiene il cursore sullo stesso testo', () => {
    const editor = editorWith('abcdef', 3);
    editor.execute('footnote.insert', { text: 'n' });
    const [ref] = footnoteRefs(editor.state.doc);
    editor.apply({ state: { ...editor.state, selection: caret(body(editor.state).id, 6) } });
    expect(editor.execute('footnote.delete', { noteId: ref!.noteId })).toBe(true);
    expect(paragraphText(body(editor.state))).toBe('abcdef');
    expect(body(editor.state).children).toHaveLength(1);
    expect(editor.state.doc.footnotes).toEqual({});
    expect(editor.state.selection.focus.offset).toBe(5);
  });

  it('una nota senza richiamo (es. dopo aver cancellato il testo) non ha numero', () => {
    const editor = editorWith('abc');
    editor.execute('footnote.insert', { text: 'n' });
    const p = body(editor.state);
    const cleaned: Paragraph = { ...p, children: p.children.filter((c) => c.type !== 'footnoteRef') };
    const doc = { ...editor.state.doc, sections: [{ ...editor.state.doc.sections[0]!, blocks: [cleaned] }] };
    expect(footnoteNumbers(doc).size).toBe(0);
  });
});

describe('insertInlineAt', () => {
  it('divide il run e funziona agli estremi', () => {
    const p = paragraph('abc');
    const ref = { type: 'footnoteRef' as const, noteId: 'n', format: {} };
    expect(paragraphText(insertInlineAt(p, 0, ref))).toBe('￼abc');
    expect(paragraphText(insertInlineAt(p, 2, ref))).toBe('ab￼c');
    expect(paragraphText(insertInlineAt(p, 3, ref))).toBe('abc￼');
    expect(paragraphText(insertInlineAt(paragraph(), 0, ref))).toBe('￼');
  });

  it('insertFootnote non fa nulla se il cursore non è nel corpo', () => {
    const state = createState();
    expect(insertFootnote({ ...state, selection: caret('altrove', 0) })).toBeNull();
  });
});
