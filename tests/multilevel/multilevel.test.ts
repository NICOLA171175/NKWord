import { describe, expect, it } from 'vitest';
import { Editor, createDocument, createState, paragraph, paragraphText, paragraphs, type Paragraph, type Selection } from '../../src/core';
import { editingModule } from '../../src/editing';
import { exportDocx, importDocx } from '../../src/io';
import { readZip } from '../../src/io/util/zip';
import { computeMarkers } from '../../src/layout';
import { buildPattern, createMultilevelModule, markerText, parsePattern, presetById, type MultilevelDialogs } from '../../src/multilevel';
import { structuresModule } from '../../src/structures';

function setup(blocks: Paragraph[], dialogs?: MultilevelDialogs): Editor {
  const editor = new Editor(createState(createDocument(blocks)));
  editor.use(editingModule);
  editor.use(structuresModule);
  editor.use(createMultilevelModule({ dialogs }));
  editor.events.on('error', (e) => {
    throw e.error;
  });
  return editor;
}

function at(editor: Editor, p: Paragraph, offset = 0, end?: [Paragraph, number]): void {
  const selection: Selection = { anchor: { paragraphId: p.id, offset }, focus: end ? { paragraphId: end[0].id, offset: end[1] } : { paragraphId: p.id, offset } };
  editor.execute('edit.setSelection', { selection });
}

const paras = (editor: Editor) => [...paragraphs(editor.state.doc)].map((l) => l.paragraph);
const markers = (editor: Editor) => {
  const m = computeMarkers(editor.state.doc);
  return paras(editor).map((p) => m.get(p.id)?.text ?? '');
};

/** Cinque voci: livelli 0, 1, 2, 1, 0. */
function outline(dialogs?: MultilevelDialogs) {
  const ps = ['a', 'b', 'c', 'd', 'e'].map((t) => paragraph(t));
  const editor = setup(ps, dialogs);
  at(editor, ps[0]!, 0, [ps[4]!, 1]);
  return { editor, ps };
}

async function levelsBy(editor: Editor, ps: Paragraph[], levels: number[]) {
  for (const [i, l] of levels.entries()) {
    at(editor, ps[i]!);
    for (let k = 0; k < l; k++) expect(editor.execute('list.increaseLevel')).toBe(true);
  }
}

describe('stili pronti', () => {
  it('1. 1.1. 1.1.1. con Tab per scendere di livello', async () => {
    const { editor, ps } = outline();
    await editor.executeAsync('multilevel.menu', { choice: 'outline' });
    await levelsBy(editor, ps, [0, 1, 2, 1, 0]);
    expect(markers(editor)).toEqual(['1.', '1.1.', '1.1.1.', '1.2.', '2.']);
  });

  it('I. A. 1. e 1) a) i)', async () => {
    const { editor, ps } = outline();
    await editor.executeAsync('multilevel.apply', { preset: 'roman' });
    await levelsBy(editor, ps, [0, 1, 2, 1, 0]);
    expect(markers(editor)).toEqual(['I.', 'A.', '1.', 'B.', 'II.']);
    at(editor, ps[0]!, 0, [ps[4]!, 1]);
    // Tutti nello stesso elenco: cambia l'aspetto dell'intero elenco, livelli compresi.
    await editor.executeAsync('multilevel.apply', { preset: 'parens' });
    expect(markers(editor)).toEqual(['1)', 'a)', 'i)', 'b)', '2)']);
  });

  it('Maiusc+Tab risale; Tab in mezzo al testo inserisce una tabulazione', async () => {
    const { editor, ps } = outline();
    await editor.executeAsync('multilevel.apply', { preset: 'outline' });
    await levelsBy(editor, ps, [0, 1, 2, 1, 0]);
    at(editor, ps[2]!);
    editor.execute('list.decreaseLevel');
    expect(markers(editor)).toEqual(['1.', '1.1.', '1.2.', '1.3.', '2.']);
    at(editor, ps[0]!, 1);
    expect(editor.canExecute('list.increaseLevel')).toBe(false);
  });

  it('Nessuno toglie l\'elenco', async () => {
    const { editor } = outline();
    await editor.executeAsync('multilevel.apply', { preset: 'outline' });
    await editor.executeAsync('multilevel.apply', { preset: 'none' });
    expect(paras(editor).every((p) => p.format.list === null)).toBe(true);
  });
});

describe('numerazione', () => {
  it('Ricomincia da 1, Continua numerazione e Imposta valore', async () => {
    const { editor, ps } = outline();
    await editor.executeAsync('multilevel.apply', { preset: 'outline' });
    at(editor, ps[0]!);
    expect(editor.canExecute('multilevel.restart')).toBe(false);
    at(editor, ps[2]!);
    expect(editor.execute('multilevel.restart')).toBe(true);
    expect(markers(editor)).toEqual(['1.', '2.', '1.', '2.', '3.']);
    expect(editor.canExecute('multilevel.continue')).toBe(true);
    expect(editor.execute('multilevel.continue')).toBe(true);
    expect(markers(editor)).toEqual(['1.', '2.', '3.', '4.', '5.']);
    at(editor, ps[3]!);
    expect(editor.execute('multilevel.setValue', { value: 10 })).toBe(true);
    expect(markers(editor)).toEqual(['1.', '2.', '3.', '10.', '11.']);
    editor.undo();
    expect(markers(editor)).toEqual(['1.', '2.', '3.', '4.', '5.']);
  });

  it('Imposta valore dalla finestra, partendo dal numero attuale', async () => {
    let shown = 0;
    const dialogs: MultilevelDialogs = {
      define: async () => null,
      setValue: async ({ value }) => ((shown = value), 7),
    };
    const { editor, ps } = outline(dialogs);
    await editor.executeAsync('multilevel.apply', { preset: 'letters' });
    at(editor, ps[1]!);
    await editor.executeAsync('multilevel.menu', { choice: 'setValue' });
    expect(shown).toBe(2);
    expect(markers(editor)).toEqual(['A.', 'G.', 'H.', 'I.', 'J.']);
  });
});

describe('formato per livello', () => {
  it('scompone e ricompone il testo del livello', () => {
    expect(parsePattern('%1.%2.%3.', 2)).toEqual({ before: '', parents: true, after: '.' });
    expect(parsePattern('(%4)', 3)).toEqual({ before: '(', parents: false, after: ')' });
    expect(parsePattern('Capitolo %1', 0)).toEqual({ before: 'Capitolo ', parents: false, after: '' });
    expect(buildPattern({ before: 'Art. ', parents: true, after: ')' }, 2)).toBe('Art. %1.%2.%3)');
    const levels = presetById('outline')!.levels();
    expect(markerText(levels, 2, [2, 3, 1])).toBe('2.3.1.');
  });

  it('la finestra "Definisci" cambia i livelli dell\'elenco', async () => {
    const dialogs: MultilevelDialogs = {
      define: async ({ levels, level }) => {
        expect(level).toBe(1);
        return levels.map((l, i) => (i === 1 ? { ...l, format: 'upperLetter', text: 'Sezione %1.%2' } : l));
      },
      setValue: async () => null,
    };
    const { editor, ps } = outline(dialogs);
    await editor.executeAsync('multilevel.apply', { preset: 'outline' });
    await levelsBy(editor, ps, [0, 1, 1, 0, 1]);
    at(editor, ps[1]!);
    await editor.executeAsync('multilevel.defineDialog');
    expect(markers(editor)).toEqual(['1.', 'Sezione 1.A', 'Sezione 1.B', '2.', 'Sezione 2.A']);
  });
});

describe('DOCX', () => {
  it('salva e riapre numerazione e livelli (numbering.xml a più livelli)', async () => {
    const { editor, ps } = outline();
    await editor.executeAsync('multilevel.apply', { preset: 'outline' });
    await levelsBy(editor, ps, [0, 1, 2, 1, 0]);
    at(editor, ps[3]!);
    editor.execute('multilevel.restart');
    const before = markers(editor);
    const bytes = await exportDocx(editor.state.doc);
    const numbering = new TextDecoder().decode((await readZip(bytes)).get('word/numbering.xml')!);
    expect(numbering).toContain('<w:multiLevelType w:val="multilevel"/>');
    expect(numbering).toContain('w:val="%1.%2.%3."');
    const doc = await importDocx(bytes);
    const m = computeMarkers(doc);
    const reopened = [...paragraphs(doc)].map((l) => m.get(l.paragraph.id)?.text ?? '');
    expect(reopened).toEqual(before);
    expect([...paragraphs(doc)].map((l) => [paragraphText(l.paragraph), l.paragraph.format.list?.level])).toEqual([
      ['a', 0], ['b', 1], ['c', 2], ['d', 1], ['e', 0],
    ]);
  });
});
