import { describe, expect, it } from 'vitest';
import {
  caret,
  createDocument,
  createState,
  Editor,
  findParagraph,
  paragraph,
  paragraphText,
  paragraphs,
  table,
  textRun,
  type Block,
  type DocumentModel,
  type Paragraph,
  type TextRun,
} from '../../src/core';
import { editingModule } from '../../src/editing';
import { footnotesModule } from '../../src/footnotes';
import { exportDocx, importDocx } from '../../src/io';
import {
  buildEntries,
  collectCaptions,
  collectMarks,
  createReferencesModule,
  crossRefCategories,
  findIndex,
  indexLetter,
  listTargets,
  MISSING_TARGET,
  NO_INDEX_ENTRIES,
  parseCrossRef,
  ribbonItems,
  splitEntry,
} from '../../src/references';
import * as referencesExports from '../../src/references';
import { tocAnchor } from '../../src/toc';
import { collectModules, collectRibbonItems } from '../../src/ui/modules';

const h1 = (text: string) => paragraph(text, {}, 'Heading1');
const body = (doc: DocumentModel) => doc.sections[0]!.blocks;
const texts = (doc: DocumentModel) => [...paragraphs(doc)].map(({ paragraph: p }) => paragraphText(p));
const byText = (doc: DocumentModel, text: string) => [...paragraphs(doc)].find(({ paragraph: p }) => paragraphText(p) === text)!.paragraph;

function setup(blocks: Block[], at?: { paragraphId: string; offset: number }) {
  const doc = createDocument(blocks);
  const first = blocks.find((b): b is Paragraph => b.type === 'paragraph')!;
  const editor = new Editor({ ...createState(doc), selection: caret(at?.paragraphId ?? first.id, at?.offset ?? 0) });
  editor.use(editingModule);
  editor.use(footnotesModule);
  editor.use(createReferencesModule({ dialogs: undefined }));
  return editor;
}

const moveTo = (editor: Editor, paragraphId: string, offset: number) =>
  editor.execute('edit.setSelection', { selection: caret(paragraphId, offset) });

describe('didascalie', () => {
  it('inserisce "Figura 1" sotto il paragrafo e "Tabella 1" sopra la tabella', () => {
    const img = paragraph('[immagine]');
    const t = table(2, 2, 400);
    const editor = setup([img, t, paragraph('fine')]);
    editor.execute('references.insertCaption', { label: 'Figura', text: 'Il grafico' });
    expect(texts(editor.state.doc)[1]).toBe('Figura 1: Il grafico');
    expect(editor.state.doc.styles.styles.Caption?.name).toBe('caption');

    // Cursore dentro la tabella: la didascalia va sopra la tabella intera.
    const cellPara = (t.rows[0]!.cells[0]!.blocks[0] as Paragraph).id;
    moveTo(editor, cellPara, 0);
    editor.execute('references.insertCaption', { label: 'Tabella' });
    const kinds = body(editor.state.doc).map((b) => (b.type === 'table' ? '[tabella]' : paragraphText(b)));
    expect(kinds).toEqual(['[immagine]', 'Figura 1: Il grafico', 'Tabella 1', '[tabella]', 'fine']);
  });

  it('rinumera le didascalie quando se ne aggiunge una prima delle altre', () => {
    const a = paragraph('primo');
    const b = paragraph('secondo');
    const editor = setup([a, b], { paragraphId: b.id, offset: 0 });
    editor.execute('references.insertCaption', { label: 'Figura', text: 'Seconda' });
    moveTo(editor, a.id, 0);
    editor.execute('references.insertCaption', { label: 'Figura', text: 'Prima' });
    expect(collectCaptions(editor.state.doc).map((c) => c.text)).toEqual(['Figura 1: Prima', 'Figura 2: Seconda']);
    // Le etichette diverse contano da sole.
    editor.execute('references.insertCaption', { label: 'Tabella', position: 'below' });
    expect(collectCaptions(editor.state.doc).map((c) => `${c.label} ${c.number}`)).toEqual(['Figura 1', 'Tabella 1', 'Figura 2']);
    // Annulla toglie l'ultima didascalia in un passo solo.
    editor.undo();
    expect(collectCaptions(editor.state.doc)).toHaveLength(2);
  });

  it('rifiuta etichette vuote o con spazi', () => {
    const editor = setup([paragraph('x')]);
    expect(editor.execute('references.insertCaption', { label: '  ' })).toBe(false);
    expect(editor.execute('references.insertCaption', { label: 'Due parole' })).toBe(false);
    expect(editor.execute('references.insertCaption', { label: 'Grafico' })).toBe(true);
    expect(texts(editor.state.doc)).toEqual(['x', 'Grafico 1']);
  });
});

describe('riferimenti incrociati', () => {
  it('si aggiornano quando cambia la numerazione e portano alla destinazione', () => {
    const fig = paragraph('[immagine]');
    const text = paragraph('Come mostra ');
    const editor = setup([text, fig], { paragraphId: fig.id, offset: 0 });
    editor.execute('references.insertCaption', { label: 'Figura', text: 'Vendite' });
    const caption = collectCaptions(editor.state.doc)[0]!;

    moveTo(editor, text.id, 'Come mostra '.length);
    editor.execute('references.insertCrossRef', { kind: 'label', type: 'caption', id: caption.paragraphId });
    expect(paragraphText(findParagraph(editor.state.doc, text.id)!.paragraph)).toBe('Come mostra Figura 1');
    const run = findParagraph(editor.state.doc, text.id)!.paragraph.children[1] as TextRun;
    expect(run.format.link).toBe(tocAnchor(caption.paragraphId));
    expect(parseCrossRef(run.format.crossRef)).toEqual({ kind: 'label', type: 'caption', id: caption.paragraphId });

    // Una nuova figura in testa sposta "Vendite" a Figura 2: il riferimento segue.
    moveTo(editor, text.id, 0);
    editor.execute('references.insertCaption', { label: 'Figura', position: 'above' });
    expect(paragraphText(findParagraph(editor.state.doc, text.id)!.paragraph)).toBe('Come mostra Figura 2');
  });

  it('mostra titolo, pagina, sopra/sotto e numero di nota', () => {
    const title = h1('Introduzione');
    const p = paragraph('Vedi ');
    const editor = setup([title, p], { paragraphId: p.id, offset: 5 });
    editor.execute('footnote.insert', { text: 'Fonte ISTAT' });
    const note = listTargets(editor.state.doc).find((t) => t.type === 'note')!;
    expect(note.display).toBe('1 Fonte ISTAT');

    moveTo(editor, p.id, 5);
    for (const kind of ['text', 'page', 'position'] as const) {
      editor.execute('references.insertCrossRef', { kind, type: 'heading', id: title.id });
      editor.execute('edit.insertText', { text: ' ' });
    }
    editor.execute('references.insertCrossRef', { kind: 'number', type: 'note', id: note.id });
    expect(paragraphText(findParagraph(editor.state.doc, p.id)!.paragraph).replace('￼', '[n]')).toBe('Vedi Introduzione 1 sopra 1[n]');
  });

  it('il testo scritto subito dopo non diventa parte del riferimento', () => {
    const title = h1('Metodo');
    const p = paragraph('Vedi ');
    const editor = setup([title, p], { paragraphId: p.id, offset: 5 });
    editor.execute('references.insertCrossRef', { kind: 'text', type: 'heading', id: title.id });
    editor.execute('edit.insertText', { text: ' a pagina' });
    const runs = findParagraph(editor.state.doc, p.id)!.paragraph.children as TextRun[];
    expect(runs.map((r) => [r.text, !!r.format.crossRef, !!r.format.link])).toEqual([
      ['Vedi ', false, false],
      ['Metodo', true, true],
      [' a pagina', false, false],
    ]);
    // Aggiornando, il testo scritto resta.
    editor.execute('references.updateFields');
    expect(paragraphText(findParagraph(editor.state.doc, p.id)!.paragraph)).toBe('Vedi Metodo a pagina');
  });

  it('cambia testo se il titolo cambia, e segnala la destinazione eliminata', () => {
    const title = h1('Bozza');
    const p = paragraph('Vedi ');
    const editor = setup([title, p], { paragraphId: p.id, offset: 5 });
    editor.execute('references.insertCrossRef', { ref: `text:heading:${title.id}` });
    moveTo(editor, title.id, 5);
    editor.execute('edit.insertText', { text: ' finale' });
    editor.execute('references.updateFields');
    expect(paragraphText(findParagraph(editor.state.doc, p.id)!.paragraph)).toBe('Vedi Bozza finale');

    editor.execute('edit.setSelection', { selection: { anchor: { paragraphId: title.id, offset: 0 }, focus: { paragraphId: p.id, offset: 0 } } });
    editor.execute('edit.deleteBackward');
    editor.execute('references.updateFields');
    expect(texts(editor.state.doc)).toEqual([`Vedi ${MISSING_TARGET}`]);
  });

  it('propone i tipi di destinazione come Word', () => {
    const editor = setup([h1('A'), paragraph('x')], undefined);
    moveTo(editor, body(editor.state.doc)[1]!.id, 1);
    editor.execute('references.insertCaption', { label: 'Grafico' });
    const cats = crossRefCategories(listTargets(editor.state.doc)).map((c) => c.label);
    expect(cats).toEqual(['Titolo', 'Figura', 'Tabella', 'Equazione', 'Grafico', 'Nota a piè di pagina']);
  });
});

describe('indice analitico', () => {
  it('segna una voce sulla selezione e con "Segna tutto" ogni occorrenza della parola', () => {
    const a = paragraph('Il Rinascimento italiano');
    const b = paragraph('Dopo il Rinascimento e il Rinascimentale');
    const editor = setup([a, b]);
    editor.execute('edit.setSelection', { selection: { anchor: { paragraphId: a.id, offset: 3 }, focus: { paragraphId: a.id, offset: 15 } } });
    editor.execute('references.markEntry', { entry: 'Rinascimento', all: true });
    const marks = collectMarks(editor.state.doc);
    // "Rinascimentale" non è la stessa parola.
    expect(marks.map((m) => [m.entry, m.paragraphId])).toEqual([
      ['Rinascimento', a.id],
      ['Rinascimento', b.id],
    ]);
    // Il segno non cambia il testo.
    expect(texts(editor.state.doc)).toEqual(['Il Rinascimento italiano', 'Dopo il Rinascimento e il Rinascimentale']);
  });

  it('con il cursore in una parola segna quella parola', () => {
    const a = paragraph('arte moderna');
    const editor = setup([a], { paragraphId: a.id, offset: 7 });
    editor.execute('references.markEntry', { entry: 'arte:moderna' });
    const runs = findParagraph(editor.state.doc, a.id)!.paragraph.children as TextRun[];
    expect(runs.map((r) => [r.text, r.format.indexEntry])).toEqual([
      ['arte ', undefined],
      ['moderna', 'arte:moderna'],
    ]);
  });

  it('ordina le voci per lettera con sottovoci e pagine', () => {
    const pages = new Map([
      ['p1', 1],
      ['p2', 3],
      ['p3', 2],
    ]);
    const groups = buildEntries(
      [
        { entry: 'zucchero', paragraphId: 'p1' },
        { entry: 'Àncora', paragraphId: 'p2' },
        { entry: 'albero:radici', paragraphId: 'p3' },
        { entry: 'albero', paragraphId: 'p2' },
        { entry: 'Albero', paragraphId: 'p1' },
        { entry: '3D', paragraphId: 'p1' },
      ],
      pages,
    );
    expect(groups.map((g) => [g.letter, g.entries.map((e) => `${'  '.repeat(e.level - 1)}${e.text} ${e.pages.map((p) => p.page).join(',')}`)])).toEqual([
      ['#', ['3D 1']],
      ['A', ['albero 1,3', '  radici 2', 'Àncora 3']],
      ['Z', ['zucchero 1']],
    ]);
    expect(indexLetter('èlite')).toBe('E');
    expect(splitEntry(' voce :  sotto ')).toEqual(['voce', 'sotto']);
  });

  it("inserisce l'indice al cursore e lo aggiorna con le nuove voci", () => {
    const a = paragraph('Galileo e Keplero');
    const end = paragraph('');
    const editor = setup([a, end], { paragraphId: a.id, offset: 0 });
    editor.execute('references.markEntry', { entry: 'Galileo' });
    moveTo(editor, end.id, 0);
    editor.execute('references.insertIndex');
    const at = findIndex(editor.state.doc)!;
    const index = () => body(editor.state.doc).slice(findIndex(editor.state.doc)!.start, findIndex(editor.state.doc)!.end).map((b) => paragraphText(b as Paragraph));
    expect(at).toBeDefined();
    expect(index()).toEqual(['G', 'Galileo, 1']);
    const pageRun = (body(editor.state.doc)[at.start + 1] as Paragraph).children.at(-1) as TextRun;
    expect(pageRun.format.link).toBe(tocAnchor(a.id));

    // Nuova voce e aggiornamento: l'indice resta uno solo.
    editor.execute('edit.setSelection', { selection: { anchor: { paragraphId: a.id, offset: 10 }, focus: { paragraphId: a.id, offset: 17 } } });
    editor.execute('references.markEntry', { entry: 'Keplero' });
    editor.execute('references.updateIndex');
    expect(index()).toEqual(['G', 'Galileo, 1', 'K', 'Keplero, 1']);
    editor.execute('references.insertIndex');
    expect(body(editor.state.doc).filter((b) => b.type === 'paragraph' && paragraphText(b) === 'G')).toHaveLength(1);
  });

  it('senza voci segnate lo dice, e la pagina segue le interruzioni', () => {
    const editor = setup([paragraph('')]);
    editor.execute('references.insertIndex');
    expect(texts(editor.state.doc)[0]).toBe(NO_INDEX_ENTRIES);

    const a = paragraph('Dante');
    const b = paragraph([textRun('prima'), { type: 'break', kind: 'page' }]);
    const c = paragraph('Petrarca');
    const ed = setup([a, b, c, paragraph('')], { paragraphId: c.id, offset: 3 });
    ed.execute('references.markEntry', { entry: 'Petrarca' });
    moveTo(ed, body(ed.state.doc)[3]!.id, 0);
    ed.execute('references.insertIndex');
    expect(texts(ed.state.doc).slice(3)).toEqual(['P', 'Petrarca, 2', '']);
  });
});

describe('DOCX', () => {
  it('salva le voci come campi XE e le rilegge', async () => {
    const a = paragraph('Il porto di Genova');
    const editor = setup([a]);
    editor.execute('edit.setSelection', { selection: { anchor: { paragraphId: a.id, offset: 12 }, focus: { paragraphId: a.id, offset: 18 } } });
    editor.execute('references.markEntry', { entry: 'Genova:porto' });
    const data = await exportDocx(editor.state.doc);
    const back = await importDocx(data);
    expect(collectMarks(back).map((m) => m.entry)).toEqual(['Genova:porto']);
    expect(texts(back)).toEqual(['Il porto di Genova']);
  });

  it('riconosce didascalie e riferimenti dopo il salvataggio', async () => {
    const fig = paragraph('[immagine]');
    const text = paragraph('Vedi ');
    const editor = setup([fig, text], { paragraphId: fig.id, offset: 0 });
    editor.execute('references.insertCaption', { label: 'Figura', text: 'Mappa' });
    const caption = collectCaptions(editor.state.doc)[0]!;
    moveTo(editor, text.id, 5);
    editor.execute('references.insertCrossRef', { kind: 'label', type: 'caption', id: caption.paragraphId });
    const back = await importDocx(await exportDocx(editor.state.doc));
    expect(collectCaptions(back).map((c) => c.text)).toEqual(['Figura 1: Mappa']);
    expect(byText(back, 'Vedi Figura 1')).toBeDefined();
  });
});

describe('modulo', () => {
  it('compare nella scheda Riferimenti ed è trovato dalla UI', () => {
    const modules = collectModules([referencesExports]);
    expect(modules.map((m) => m.id)).toContain('references');
    const items = collectRibbonItems([referencesExports]);
    expect(items.map((i) => i.commandId)).toEqual(ribbonItems.map((i) => i.commandId));
    expect(new Set(items.map((i) => i.tab))).toEqual(new Set(['references']));
  });
});
