import { describe, expect, it } from 'vitest';
import {
  createDocument,
  createState,
  Editor,
  paragraph,
  textRun,
  type DocumentModel,
  type Paragraph,
  type Selection,
  type TextRun,
} from '../../src/core';
import { commentRanges, commentsAt, commentThreads, createCommentsModule, wordAt } from '../../src/comments';
import { editingModule } from '../../src/editing';

const NOW = new Date('2026-10-13T08:00:00Z');

function setup(doc: DocumentModel, selection?: Selection) {
  const state = createState(doc);
  const editor = new Editor(selection ? { ...state, selection } : state);
  const module = createCommentsModule({ author: 'Nicola Rossi', now: () => NOW });
  editor.use(module);
  return { editor, module };
}

const sel = (p: Paragraph, a: number, b: number, q: Paragraph = p): Selection => ({
  anchor: { paragraphId: p.id, offset: a },
  focus: { paragraphId: q.id, offset: b },
});

const runs = (doc: DocumentModel, i = 0) =>
  (doc.sections[0]!.blocks[i] as Paragraph).children.filter((c): c is TextRun => c.type === 'text');

describe('comment.insert', () => {
  it('ancora il commento alla selezione', () => {
    const p = paragraph([textRun('Ciao bel mondo')]);
    const { editor } = setup(createDocument([p]), sel(p, 5, 8));
    expect(editor.execute('comment.insert', { text: 'Da rivedere' })).toBe(true);
    const doc = editor.state.doc;
    const [c] = Object.values(doc.comments!);
    expect(c).toMatchObject({ author: 'Nicola Rossi', initials: 'NR', text: 'Da rivedere', date: '2026-10-13T08:00:00Z' });
    expect(runs(doc).map((r) => [r.text, r.format.commentIds])).toEqual([
      ['Ciao ', undefined],
      ['bel', c!.id],
      [' mondo', undefined],
    ]);
    expect(commentRanges(doc).get(c!.id)).toEqual({ from: { paragraphId: p.id, offset: 5 }, to: { paragraphId: p.id, offset: 8 } });
  });

  it('con il solo cursore commenta la parola, e si annulla', () => {
    const p = paragraph([textRun('uno due tre')]);
    const { editor } = setup(createDocument([p]), sel(p, 5, 5));
    editor.execute('comment.insert');
    const t = commentThreads(editor.state.doc);
    expect(t).toHaveLength(1);
    expect(t[0]!.range.from.offset).toBe(4);
    expect(t[0]!.range.to.offset).toBe(7);
    expect(t[0]!.comment.text).toBe('');
    editor.undo();
    expect(editor.state.doc.comments).toBeUndefined();
    expect(runs(editor.state.doc)).toHaveLength(1);
  });

  it('non è disponibile su un paragrafo vuoto', () => {
    const p = paragraph();
    const { editor } = setup(createDocument([p]), sel(p, 0, 0));
    expect(editor.canExecute('comment.insert')).toBe(false);
  });

  it('attraversa più paragrafi e si sovrappone ad altri commenti', () => {
    const a = paragraph([textRun('primo')]);
    const b = paragraph([textRun('secondo')]);
    const { editor } = setup(createDocument([a, b]), sel(a, 2, 3, b));
    editor.execute('comment.insert', { text: 'uno' });
    const first = Object.keys(editor.state.doc.comments!)[0]!;
    // Secondo commento sulla parola "secondo", già in parte commentata.
    editor.apply({ state: { ...editor.state, selection: sel(b, 0, 7) } });
    editor.execute('comment.insert', { text: 'due' });
    const doc = editor.state.doc;
    const second = Object.keys(doc.comments!).find((id) => id !== first)!;
    expect(runs(doc, 0).map((r) => r.format.commentIds)).toEqual([undefined, first]);
    expect(runs(doc, 1).map((r) => r.text)).toEqual(['sec', 'ondo']);
    expect(runs(doc, 1)[0]!.format.commentIds!.split(' ').sort()).toEqual([first, second].sort());
    expect(commentsAt(doc, { paragraphId: b.id, offset: 1 }).sort()).toEqual([first, second].sort());
    expect(commentThreads(doc).map((t) => t.comment.text)).toEqual(['uno', 'due']);
  });
});

describe('risposte, modifica, risoluzione, eliminazione', () => {
  function withComment() {
    const p = paragraph([textRun('testo commentato')]);
    const ctx = setup(createDocument([p]), sel(p, 0, 5));
    ctx.editor.execute('comment.insert', { text: 'Primo' });
    const id = Object.keys(ctx.editor.state.doc.comments!)[0]!;
    return { ...ctx, p, id };
  }

  it('aggiunge risposte alla discussione', () => {
    const { editor, id } = withComment();
    expect(editor.execute('comment.reply', { parentId: id, text: 'Risposta', author: 'Anna' })).toBe(true);
    const replyId = Object.keys(editor.state.doc.comments!).find((x) => x !== id)!;
    // Una risposta a una risposta finisce comunque nella stessa discussione.
    editor.execute('comment.reply', { parentId: replyId, text: 'Altra' });
    const [t] = commentThreads(editor.state.doc);
    expect(t!.replies.map((r) => [r.author, r.text, r.parentId])).toEqual([
      ['Anna', 'Risposta', id],
      ['Nicola Rossi', 'Altra', id],
    ]);
  });

  it('modifica e risolve', () => {
    const { editor, id } = withComment();
    editor.execute('comment.edit', { id, text: 'Corretto' });
    editor.execute('comment.resolve', { id, resolved: true });
    expect(editor.state.doc.comments![id]).toMatchObject({ text: 'Corretto', resolved: true });
    expect(editor.execute('comment.resolve', { id, resolved: true })).toBe(false);
  });

  it('elimina il commento al cursore con le risposte e l\'ancoraggio', () => {
    const { editor, id, p } = withComment();
    editor.execute('comment.reply', { parentId: id, text: 'r' });
    editor.apply({ state: { ...editor.state, selection: sel(p, 2, 2) } });
    expect(editor.execute('comment.delete')).toBe(true);
    expect(editor.state.doc.comments).toBeUndefined();
    expect(runs(editor.state.doc)).toEqual([expect.objectContaining({ text: 'testo commentato', format: {} })]);
    editor.undo();
    expect(Object.keys(editor.state.doc.comments!)).toHaveLength(2);
  });

  it('elimina solo una risposta', () => {
    const { editor, id } = withComment();
    editor.execute('comment.reply', { parentId: id, text: 'r' });
    const replyId = Object.keys(editor.state.doc.comments!).find((x) => x !== id)!;
    editor.execute('comment.delete', { id: replyId });
    expect(Object.keys(editor.state.doc.comments!)).toEqual([id]);
    expect(commentThreads(editor.state.doc)).toHaveLength(1);
  });
});

describe('navigazione e integrazione con l\'editing', () => {
  it('passa al commento successivo e precedente, ricominciando dai capi', () => {
    const p = paragraph([textRun('alfa beta gamma')]);
    const { editor } = setup(createDocument([p]), sel(p, 0, 4));
    editor.execute('comment.insert', { text: '1' });
    editor.apply({ state: { ...editor.state, selection: sel(p, 11, 15) } });
    editor.execute('comment.insert', { text: '2' });
    editor.apply({ state: { ...editor.state, selection: sel(p, 6, 6) } });
    const at = () => editor.state.selection.anchor.offset;
    editor.execute('comment.next');
    expect(at()).toBe(11);
    editor.execute('comment.next');
    expect(at()).toBe(0);
    editor.execute('comment.previous');
    expect(at()).toBe(11);
    expect(editor.history.canUndo).toBe(true); // solo per gli inserimenti, non per gli spostamenti
    editor.undo();
    expect(Object.keys(editor.state.doc.comments!)).toHaveLength(1);
  });

  it('"Cancella formattazione" conserva i commenti e il testo cancellato li nasconde', () => {
    const p = paragraph([textRun('grassetto', { bold: true })]);
    const editor = new Editor({ ...createState(createDocument([p])), selection: sel(p, 0, 9) });
    editor.use(editingModule);
    editor.use(createCommentsModule());
    editor.execute('comment.insert', { text: 'x' });
    editor.execute('format.clear');
    const id = Object.keys(editor.state.doc.comments!)[0]!;
    expect(runs(editor.state.doc)[0]!.format).toEqual({ commentIds: id });
    editor.execute('edit.deleteSelection');
    expect(commentThreads(editor.state.doc)).toHaveLength(0);
    editor.undo();
    expect(commentThreads(editor.state.doc)).toHaveLength(1);
  });

  it('wordAt riconosce la parola anche ai bordi', () => {
    const p = paragraph([textRun("l'albero, verde")]);
    expect(wordAt(p, 0)).toEqual([0, 8]);
    expect(wordAt(p, 8)).toEqual([0, 8]);
    expect(wordAt(p, 9)).toBeNull();
    expect(wordAt(p, 15)).toEqual([10, 15]);
  });
});
