import { describe, expect, it } from 'vitest';
import {
  Editor,
  caret,
  createDocument,
  createState,
  paragraph,
  paragraphText,
  table,
  type Block,
  type DocumentModel,
  type Paragraph,
  type TextRun,
} from '../../src/core';
import { createApproxMeasurer, createLayoutModule } from '../../src/layout';
import {
  NO_ENTRIES_TEXT,
  TOC_TITLE,
  collectHeadings,
  findToc,
  ribbonItems,
  tocAnchor,
  tocEntryLevel,
  tocModule,
} from '../../src/toc';

const h = (level: number, text: string) => paragraph(text, {}, `Heading${level}`);
const pageBreak = (text = '') => paragraph([...(text ? [{ type: 'text' as const, text, format: {} }] : []), { type: 'break', kind: 'page' }]);

function editorWith(blocks: Block[], caretAt = 0, withLayout = false): Editor {
  const doc = createDocument(blocks);
  const target = doc.sections[0]!.blocks[caretAt] as Paragraph;
  const editor = new Editor({ ...createState(doc), selection: caret(target.id, 0) });
  if (withLayout) editor.use(createLayoutModule({ measurer: createApproxMeasurer() }));
  editor.use(tocModule);
  editor.events.on('error', (e) => {
    throw e.error;
  });
  return editor;
}

const body = (doc: DocumentModel) => doc.sections[0]!.blocks as readonly Paragraph[];
const texts = (doc: DocumentModel) => body(doc).map((p) => paragraphText(p));
const tocParagraphs = (doc: DocumentModel) => {
  const at = findToc(doc)!;
  return body(doc).slice(at.start, at.end);
};

describe('titoli', () => {
  it('raccoglie i titoli fino al livello scelto, anche nelle tabelle, saltando quelli vuoti', () => {
    const t = table(1, 1, 100);
    const inCell = { ...t, rows: [{ ...t.rows[0]!, cells: [{ ...t.rows[0]!.cells[0]!, blocks: [h(2, 'In tabella')] }] }] };
    const doc = createDocument([h(1, 'Uno'), paragraph('testo'), h(2, '  Due\tbis '), h(1, ''), inCell, paragraph('Quattro', {}, 'Heading4')]);
    // Heading4 non esiste nel foglio di stile predefinito: livello da formato diretto.
    const withL4: DocumentModel = {
      ...doc,
      sections: [{ ...doc.sections[0]!, blocks: [...doc.sections[0]!.blocks, paragraph('Quattro bis', { outlineLevel: 4 })] }],
    };
    expect(collectHeadings(withL4).map((x) => [x.level, x.text])).toEqual([
      [1, 'Uno'],
      [2, 'Due bis'],
      [2, 'In tabella'],
    ]);
    expect(collectHeadings(withL4, 4).map((x) => x.text)).toContain('Quattro bis');
  });
});

describe('toc.insert', () => {
  it('inserisce titolo e voci prima del paragrafo del cursore', () => {
    const editor = editorWith([paragraph(''), h(1, 'Introduzione'), paragraph('bla'), h(2, 'Dettagli'), h(3, 'Minuzie')]);
    const caretBefore = editor.state.selection;
    expect(editor.execute('toc.insert')).toBe(true);
    const doc = editor.state.doc;
    expect(texts(doc).slice(0, 4)).toEqual([TOC_TITLE, 'Introduzione\t1', 'Dettagli\t1', 'Minuzie\t1']);
    expect(body(doc).slice(0, 4).map((p) => p.styleId)).toEqual(['TOCHeading', 'TOC1', 'TOC2', 'TOC3']);
    expect(doc.styles.styles.TOC2?.name).toBe('toc 2');
    // Il paragrafo del cursore resta dopo il sommario.
    expect(editor.state.selection).toEqual(caretBefore);
    expect(body(doc)[4]!.id).toBe(caretBefore.focus.paragraphId);

    const entry = body(doc)[1]!;
    const width = doc.sections[0]!.page.width - doc.sections[0]!.page.margins.left - doc.sections[0]!.page.margins.right;
    expect(entry.format.tabs).toEqual([{ position: width, align: 'right', leader: 'dot' }]);
    const links = entry.children.filter((c): c is TextRun => c.type === 'text').map((c) => c.format.link);
    expect(links).toEqual([tocAnchor(body(doc)[5]!.id), tocAnchor(body(doc)[5]!.id)]);
  });

  it('stima le pagine dalle interruzioni senza impaginazione', () => {
    const editor = editorWith([paragraph(''), h(1, 'A'), pageBreak('fine A'), h(1, 'B'), paragraph('x', { pageBreakBefore: true }), h(2, 'C')]);
    editor.execute('toc.insert');
    expect(texts(editor.state.doc).slice(1, 4)).toEqual(['A\t1', 'B\t2', 'C\t3']);
  });

  it('usa l\'impaginazione per i numeri di pagina', () => {
    const filler = Array.from({ length: 80 }, (_, i) => paragraph(`riga ${i}`));
    const editor = editorWith([paragraph(''), h(1, 'Inizio'), ...filler, h(1, 'Dopo')], 0, true);
    editor.execute('toc.insert');
    const layout = editor.services.layout!.layout(editor.state.doc);
    const pageOf = (id: string) => layout.pages.find((p) => p.lines.some((l) => l.paragraphId === id))!.number;
    const doc = editor.state.doc;
    const dopo = body(doc).find((p) => paragraphText(p) === 'Dopo')!;
    expect(pageOf(dopo.id)).toBeGreaterThan(1);
    expect(texts(doc)[2]).toBe(`Dopo\t${pageOf(dopo.id)}`);
  });

  it('senza titoli scrive il segnaposto, e non si attiva in una tabella', () => {
    const editor = editorWith([paragraph('solo testo')]);
    editor.execute('toc.insert');
    expect(texts(editor.state.doc).slice(0, 2)).toEqual([TOC_TITLE, NO_ENTRIES_TEXT]);

    const t = table(1, 1, 100);
    const doc = createDocument([t]);
    const cellPara = t.rows[0]!.cells[0]!.blocks[0]!;
    const inTable = new Editor({ ...createState(doc), selection: caret(cellPara.id, 0) });
    inTable.use(tocModule);
    expect(inTable.canExecute('toc.insert')).toBe(false);
  });

  it('se il sommario esiste già lo aggiorna invece di duplicarlo', () => {
    const editor = editorWith([paragraph(''), h(1, 'A')]);
    editor.execute('toc.insert');
    editor.execute('toc.insert');
    expect(texts(editor.state.doc).filter((t) => t === TOC_TITLE)).toHaveLength(1);
  });
});

describe('toc.update e toc.remove', () => {
  it('aggiorna voci e numeri mantenendo gli id dei paragrafi del sommario', () => {
    const editor = editorWith([paragraph(''), h(1, 'A'), paragraph('x'), h(1, 'B')]);
    expect(editor.canExecute('toc.update')).toBe(false);
    editor.execute('toc.insert');
    const before = tocParagraphs(editor.state.doc).map((p) => p.id);
    // Nessun cambiamento: nessuna transazione.
    expect(editor.execute('toc.update')).toBe(false);

    // Nuovo titolo, interruzione di pagina prima di B.
    const doc = editor.state.doc;
    const blocks = body(doc);
    const xIndex = blocks.findIndex((p) => paragraphText(p) === 'x');
    const next = [...blocks.slice(0, xIndex), h(2, 'A.1'), pageBreak('x'), ...blocks.slice(xIndex + 1)];
    editor.apply({ state: { ...editor.state, doc: { ...doc, sections: [{ ...doc.sections[0]!, blocks: next }] } } });

    expect(editor.execute('toc.update')).toBe(true);
    const toc = tocParagraphs(editor.state.doc);
    expect(toc.map(paragraphText)).toEqual([TOC_TITLE, 'A\t1', 'A.1\t1', 'B\t2']);
    expect(toc.slice(0, 3).map((p) => p.id)).toEqual(before);
    editor.undo();
    expect(tocParagraphs(editor.state.doc).map(paragraphText)).toEqual([TOC_TITLE, 'A\t1', 'B\t1']);
  });

  it('rimuove il sommario e sposta il cursore che c\'era dentro', () => {
    const editor = editorWith([paragraph(''), h(1, 'A')]);
    editor.execute('toc.insert');
    const inside = tocParagraphs(editor.state.doc)[1]!;
    editor.execute('edit.undo');
    editor.redo();
    editor.apply({ state: { ...editor.state, selection: caret(inside.id, 2) } });
    expect(editor.execute('toc.remove')).toBe(true);
    expect(findToc(editor.state.doc)).toBeUndefined();
    expect(texts(editor.state.doc)).toEqual(['', 'A']);
    expect(editor.state.selection.focus.paragraphId).toBe(body(editor.state.doc)[0]!.id);
    expect(editor.canExecute('toc.remove')).toBe(false);
  });

  it('riconosce un sommario di Word con stili localizzati e li riusa', () => {
    const base = createDocument();
    const styles = {
      ...base.styles,
      styles: {
        ...base.styles.styles,
        Sommario1: { id: 'Sommario1', name: 'toc 1', type: 'paragraph' as const },
        Titolosommario: { id: 'Titolosommario', name: 'TOC Heading', type: 'paragraph' as const, basedOn: 'Heading1' },
      },
    };
    const blocks = [
      paragraph('Sommario', {}, 'Titolosommario'),
      paragraph('Vecchio\t3', {}, 'Sommario1'),
      h(1, 'Nuovo'),
    ];
    const doc: DocumentModel = { ...base, styles, sections: [{ ...base.sections[0]!, blocks }] };
    expect(tocEntryLevel(styles, 'Sommario1')).toBe(1);
    // Il titolo del sommario (basato su Titolo 1) non finisce nel sommario.
    expect(collectHeadings(doc).map((x) => x.text)).toEqual(['Nuovo']);
    const editor = new Editor(createState(doc));
    editor.use(tocModule);
    editor.execute('toc.update');
    const toc = tocParagraphs(editor.state.doc);
    expect(toc.map(paragraphText)).toEqual(['Sommario', 'Nuovo\t1']);
    expect(toc.map((p) => p.styleId)).toEqual(['Titolosommario', 'Sommario1']);
  });
});

describe('barra multifunzione', () => {
  it('dichiara le voci della scheda Riferimenti per i comandi del modulo', () => {
    const editor = editorWith([paragraph('')]);
    for (const item of ribbonItems) {
      expect(item.tab).toBe('references');
      expect(editor.commands.has(item.commandId)).toBe(true);
    }
  });
});
