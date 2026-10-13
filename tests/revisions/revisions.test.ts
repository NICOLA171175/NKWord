import { describe, expect, it } from 'vitest';
import {
  Editor,
  createDocument,
  createState,
  paragraph,
  paragraphText,
  paragraphs,
  textRun,
  type Block,
  type Paragraph,
  type Revision,
} from '../../src/core';
import { editingModule } from '../../src/editing';
import * as revisions from '../../src/revisions';
import { createRevisionsModule, revisionSpans, ribbonItems } from '../../src/revisions';
import { collectRibbonItems } from '../../src/ui/modules';

const DATE = new Date('2026-10-13T08:00:00Z');

function setup(blocks: Block[], author = 'Nicola') {
  const editor = new Editor(createState(createDocument(blocks)));
  editor.use(editingModule);
  const module = createRevisionsModule({ author, now: () => DATE });
  editor.use(module);
  return { editor, module };
}

const paras = (e: Editor): Paragraph[] => [...paragraphs(e.state.doc)].map((l) => l.paragraph);
const texts = (e: Editor) => paras(e).map(paragraphText);

/** Testo con le revisioni in evidenza: [+inserito+] e [-eliminato-]; ¶ marca un segno di paragrafo revisionato. */
function markup(e: Editor): string {
  return paras(e)
    .map((p) => {
      let s = '';
      for (const c of p.children) {
        const t = c.type === 'text' ? c.text : c.type === 'tab' ? '\t' : '?';
        s += c.revision ? (c.revision.kind === 'insert' ? `[+${t}+]` : `[-${t}-]`) : t;
      }
      if (p.revision) s += p.revision.kind === 'insert' ? '[+¶+]' : '[-¶-]';
      return s;
    })
    .join(' | ');
}

function caretAt(e: Editor, p: Paragraph, offset: number) {
  e.execute('edit.setSelection', { selection: { anchor: { paragraphId: p.id, offset }, focus: { paragraphId: p.id, offset } } });
}

function selectRange(e: Editor, a: [Paragraph, number], f: [Paragraph, number]) {
  e.execute('edit.setSelection', {
    selection: { anchor: { paragraphId: a[0].id, offset: a[1] }, focus: { paragraphId: f[0].id, offset: f[1] } },
  });
}

function type(e: Editor, text: string) {
  for (const ch of text) e.execute('edit.insertText', { text: ch });
}

describe('modulo revisioni', () => {
  it('registra i comandi di tutte le voci della barra multifunzione', () => {
    const { editor } = setup([paragraph('x')]);
    expect(ribbonItems.length).toBeGreaterThan(0);
    for (const item of ribbonItems) {
      expect(item.tab).toBe('review');
      expect(editor.commands.has(item.commandId)).toBe(true);
    }
  });

  it("la barra multifunzione dell'app trova le voci del modulo", () => {
    expect(collectRibbonItems([revisions]).map((i) => i.commandId)).toEqual(ribbonItems.map((i) => i.commandId));
  });

  it('senza rilevamento le modifiche sono normali', () => {
    const { editor } = setup([paragraph('ciao')]);
    caretAt(editor, paras(editor)[0]!, 4);
    type(editor, '!');
    editor.execute('edit.deleteBackward', {});
    editor.execute('edit.deleteBackward', {});
    expect(markup(editor)).toBe('cia');
  });

  it('il pulsante è un interruttore', () => {
    const { editor, module } = setup([paragraph('x')]);
    expect(editor.isActive('revision.toggleTracking')).toBe(false);
    editor.execute('revision.toggleTracking');
    expect(editor.isActive('revision.toggleTracking')).toBe(true);
    expect(module.controller.tracking).toBe(true);
    expect(editor.history.canUndo).toBe(false);
  });
});

describe('rilevamento modifiche', () => {
  function tracked(text: string) {
    const ctx = setup([paragraph(text)]);
    ctx.editor.execute('revision.toggleTracking');
    return ctx;
  }

  it('il testo digitato diventa un unico inserimento con autore e data', () => {
    const { editor } = tracked('Ciao mondo');
    caretAt(editor, paras(editor)[0]!, 4);
    type(editor, ' bel');
    expect(markup(editor)).toBe('Ciao[+ bel+] mondo');
    const run = paras(editor)[0]!.children[1]!;
    expect(run.revision).toEqual<Revision>({ kind: 'insert', author: 'Nicola', date: DATE.toISOString() });
    expect(editor.state.selection.focus.offset).toBe(8);
  });

  it('Backspace barra il carattere e lascia il cursore prima', () => {
    const { editor } = tracked('abc');
    caretAt(editor, paras(editor)[0]!, 3);
    editor.execute('edit.deleteBackward', {});
    expect(markup(editor)).toBe('ab[-c-]');
    expect(editor.state.selection.focus.offset).toBe(2);
    editor.execute('edit.deleteBackward', {});
    expect(markup(editor)).toBe('a[-bc-]');
    expect(editor.state.selection.focus.offset).toBe(1);
  });

  it('Canc barra il carattere e porta il cursore dopo', () => {
    const { editor } = tracked('abc');
    caretAt(editor, paras(editor)[0]!, 0);
    editor.execute('edit.deleteForward', {});
    editor.execute('edit.deleteForward', {});
    expect(markup(editor)).toBe('[-ab-]c');
    expect(editor.state.selection.focus.offset).toBe(2);
  });

  it('Backspace su testo già eliminato lo salta senza cambiare il documento', () => {
    const { editor } = tracked('abc');
    caretAt(editor, paras(editor)[0]!, 3);
    editor.execute('edit.deleteBackward', {});
    const doc = editor.state.doc;
    caretAt(editor, paras(editor)[0]!, 3);
    editor.execute('edit.deleteBackward', {});
    expect(editor.state.doc).toBe(doc);
    expect(editor.state.selection.focus.offset).toBe(2);
  });

  it('cancellare un proprio inserimento lo toglie davvero', () => {
    const { editor } = tracked('ab');
    caretAt(editor, paras(editor)[0]!, 2);
    type(editor, 'xy');
    editor.execute('edit.deleteBackward', {});
    expect(markup(editor)).toBe('ab[+x+]');
  });

  it('il carattere nuovo è quello prima del cursore anche se uguale al successivo', () => {
    const { editor } = tracked('ab');
    caretAt(editor, paras(editor)[0]!, 2);
    type(editor, 'b');
    expect(markup(editor)).toBe('ab[+b+]');
    caretAt(editor, paras(editor)[0]!, 1);
    editor.execute('edit.deleteForward', {});
    expect(markup(editor)).toBe('a[-b-][+b+]');
  });

  it('scrivere sopra una selezione: eliminato poi inserito', () => {
    const { editor } = tracked('il gatto nero');
    const p = paras(editor)[0]!;
    selectRange(editor, [p, 3], [p, 8]);
    type(editor, 'cane');
    expect(markup(editor)).toBe('il [-gatto-][+cane+] nero');
    expect(editor.state.selection.focus.offset).toBe(12);
  });

  it('Invio inserisce un segno di paragrafo', () => {
    const { editor } = tracked('abcdef');
    caretAt(editor, paras(editor)[0]!, 3);
    editor.execute('edit.splitParagraph');
    expect(markup(editor)).toBe('abc[+¶+] | def');
    expect(editor.state.selection.focus).toEqual({ paragraphId: paras(editor)[1]!.id, offset: 0 });
  });

  it('unire due paragrafi lascia il segno eliminato e i paragrafi separati', () => {
    const { editor } = tracked('');
    editor.reset(createState(createDocument([paragraph('abc'), paragraph('def')])));
    caretAt(editor, paras(editor)[1]!, 0);
    editor.execute('edit.deleteBackward', {});
    expect(markup(editor)).toBe('abc[-¶-] | def');
    const ids = paras(editor).map((p) => p.id);
    expect(new Set(ids).size).toBe(2);
  });

  it('annulla toglie la modifica rilevata in un solo passo', () => {
    const { editor } = tracked('abc');
    caretAt(editor, paras(editor)[0]!, 3);
    editor.execute('edit.deleteBackward', {});
    editor.undo();
    expect(markup(editor)).toBe('abc');
    editor.redo();
    expect(markup(editor)).toBe('ab[-c-]');
  });

  it('due autori hanno revisioni distinte', () => {
    const { editor } = tracked('x');
    caretAt(editor, paras(editor)[0]!, 1);
    type(editor, 'a');
    editor.execute('revision.setAuthor', { author: 'Giulia' });
    type(editor, 'b');
    const runs = paras(editor)[0]!.children;
    expect(runs.map((r) => r.revision?.author)).toEqual([undefined, 'Nicola', 'Giulia']);
  });

  it('rileva le modifiche anche dentro le tabelle', () => {
    const { editor } = setup([paragraph('prima')]);
    // Tabella creata a mano per non dipendere dal modulo strutture.
    const cellPara = paragraph('cella');
    const doc = createDocument([
      paragraph('prima'),
      {
        type: 'table',
        id: 't1',
        columnWidths: [100],
        rows: [{ id: 'r1', cells: [{ id: 'c1', blocks: [cellPara], format: {} }] }],
        format: {},
      },
    ]);
    editor.reset(createState(doc));
    editor.execute('revision.toggleTracking');
    caretAt(editor, cellPara, 5);
    type(editor, '!');
    expect(markup(editor)).toBe('prima | cella[+!+]');
  });
});

describe('accetta e rifiuta', () => {
  const ins = (text: string, author = 'A'): ReturnType<typeof textRun> => ({
    ...textRun(text),
    revision: { kind: 'insert', author },
  });
  const del = (text: string, author = 'A'): ReturnType<typeof textRun> => ({
    ...textRun(text),
    revision: { kind: 'delete', author },
  });

  function withRevisions() {
    const p1: Paragraph = { ...paragraph([textRun('Il '), del('gatto'), ins('cane'), textRun(' dorme')]) };
    const p2: Paragraph = { ...paragraph([textRun('Fine'), ins('!')]) };
    return setup([p1, p2]);
  }

  it('elenca le revisioni in ordine', () => {
    const { editor } = withRevisions();
    const spans = revisionSpans(editor.state.doc);
    expect(spans.map((s) => s.revision.kind)).toEqual(['delete', 'insert', 'insert']);
  });

  it('accetta tutto', () => {
    const { editor } = withRevisions();
    expect(editor.execute('revision.acceptAll')).toBe(true);
    expect(markup(editor)).toBe('Il cane dorme | Fine!');
    expect(editor.canExecute('revision.acceptAll')).toBe(false);
  });

  it('rifiuta tutto', () => {
    const { editor } = withRevisions();
    editor.execute('revision.rejectAll');
    expect(markup(editor)).toBe('Il gatto dorme | Fine');
  });

  it('accetta solo la revisione al cursore', () => {
    const { editor } = withRevisions();
    caretAt(editor, paras(editor)[0]!, 9); // dentro "cane"
    expect(editor.execute('revision.accept')).toBe(true);
    expect(markup(editor)).toBe('Il [-gatto-]cane dorme | Fine[+!+]');
  });

  it('rifiuta le revisioni nella selezione', () => {
    const { editor } = withRevisions();
    const p = paras(editor)[0]!;
    selectRange(editor, [p, 0], [p, 12]);
    editor.execute('revision.reject');
    expect(markup(editor)).toBe('Il gatto dorme | Fine[+!+]');
  });

  it('senza revisioni al cursore Accetta è disabilitato', () => {
    const { editor } = withRevisions();
    caretAt(editor, paras(editor)[0]!, 1);
    expect(editor.canExecute('revision.accept')).toBe(false);
  });

  it('accettare un segno di paragrafo eliminato unisce i paragrafi', () => {
    const p1: Paragraph = { ...paragraph('abc'), revision: { kind: 'delete', author: 'A' } };
    const { editor } = setup([p1, paragraph('def')]);
    editor.execute('revision.acceptAll');
    expect(texts(editor)).toEqual(['abcdef']);
  });

  it("l'ultimo segno di paragrafo perde solo la revisione", () => {
    const p1: Paragraph = { ...paragraph('abc'), revision: { kind: 'insert', author: 'A' } };
    const { editor } = setup([p1]);
    editor.execute('revision.rejectAll');
    expect(markup(editor)).toBe('abc');
  });

  it('rifiutare un Invio rilevato riunisce il paragrafo', () => {
    const { editor } = setup([paragraph('abcdef')]);
    editor.execute('revision.toggleTracking');
    caretAt(editor, paras(editor)[0]!, 3);
    editor.execute('edit.splitParagraph');
    editor.execute('revision.rejectAll');
    expect(markup(editor)).toBe('abcdef');
  });

  it('accetta e rifiuta non vengono rilevati come modifiche', () => {
    const { editor } = withRevisions();
    editor.execute('revision.toggleTracking');
    editor.execute('revision.acceptAll');
    expect(markup(editor)).toBe('Il cane dorme | Fine!');
    editor.undo();
    expect(markup(editor)).toBe('Il [-gatto-][+cane+] dorme | Fine[+!+]');
  });

  it('successiva e precedente selezionano le revisioni, ricominciando', () => {
    const { editor } = withRevisions();
    const [p1, p2] = paras(editor);
    caretAt(editor, p1!, 0);
    editor.execute('revision.next');
    expect(editor.state.selection).toEqual({ anchor: { paragraphId: p1!.id, offset: 3 }, focus: { paragraphId: p1!.id, offset: 8 } });
    editor.execute('revision.next');
    expect(editor.state.selection.anchor).toEqual({ paragraphId: p1!.id, offset: 8 });
    editor.execute('revision.next');
    expect(editor.state.selection.anchor).toEqual({ paragraphId: p2!.id, offset: 4 });
    editor.execute('revision.next');
    expect(editor.state.selection.anchor).toEqual({ paragraphId: p1!.id, offset: 3 });
    editor.execute('revision.previous');
    expect(editor.state.selection.anchor).toEqual({ paragraphId: p2!.id, offset: 4 });
  });
});
