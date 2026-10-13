import { describe, expect, it } from 'vitest';
import { Editor, caret, createDocument, createState, paragraph } from '../../src/core';
import { commentRanges, commentTarget, createCommentsModule, nearestWord } from '../../src/comments';

// Controllo dei comandi (Revisione > Nuovo commento): senza selezione si commenta la parola più vicina.
describe('nuovo commento con il cursore fuori da una parola', () => {
  it('dopo il punto finale commenta la parola precedente', () => {
    const p = paragraph('Fine del testo.');
    const state = { ...createState(createDocument([p])), selection: caret(p.id, 15) };
    expect(commentTarget(state)).toEqual({ anchor: { paragraphId: p.id, offset: 9 }, focus: { paragraphId: p.id, offset: 14 } });
    const editor = new Editor(state);
    editor.use(createCommentsModule({ author: 'Nicola' }));
    expect(editor.canExecute('comment.insert')).toBe(true);
    editor.execute('comment.insert');
    expect(Object.keys(editor.state.doc.comments ?? {})).toHaveLength(1);
  });

  it("all'inizio prima di uno spazio usa la parola successiva; un paragrafo vuoto non ha niente da commentare", () => {
    expect(nearestWord(paragraph('  ciao'), 0)).toEqual([2, 6]);
    const empty = paragraph('');
    expect(commentTarget({ ...createState(createDocument([empty])), selection: caret(empty.id, 0) })).toBeNull();
    expect(commentRanges).toBeTypeOf('function');
  });
});
