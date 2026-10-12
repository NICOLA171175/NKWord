import { describe, expect, it } from 'vitest';
import {
  COMMAND_IDS,
  COMMAND_OWNERS,
  Editor,
  PAGE_SIZES,
  caret,
  paragraph,
  paragraphText,
  textRun,
  type Paragraph,
} from '../../src/core';
import { structuresModule } from '../../src/structures';
import { editorWith, range } from './helpers';

const blocks = (e: Editor) => e.state.doc.sections.flatMap((s) => s.blocks) as Paragraph[];

describe('modulo strutture', () => {
  it('registra esattamente i comandi di cui è proprietario', () => {
    const editor = new Editor();
    const before = new Set(editor.commands.list().map((c) => c.id));
    const dispose = editor.use(structuresModule);
    const added = editor.commands.list().map((c) => c.id).filter((id) => !before.has(id)).sort();
    const owned = COMMAND_IDS.filter((id) => COMMAND_OWNERS[id] === 'structures').sort();
    expect(added).toEqual(owned);
    dispose();
    expect(editor.commands.list().map((c) => c.id).sort()).toEqual([...before].sort());
  });
});

describe('insert.*', () => {
  it('inserisce un\'immagine al cursore sostituendo la selezione', () => {
    const p = paragraph('abcdef');
    const editor = editorWith([p], range(p.id, 1, p.id, 3));
    expect(editor.canExecute('insert.image', { src: '', width: 1, height: 1 })).toBe(false);
    editor.execute('insert.image', { src: 'data:image/png;base64,xx', width: 100, height: 50, alt: 'logo' });
    const out = blocks(editor)[0]!;
    expect(paragraphText(out)).toBe('a￼def');
    expect(out.children[1]).toEqual({ type: 'image', src: 'data:image/png;base64,xx', width: 100, height: 50, alt: 'logo' });
    expect(editor.state.selection.focus).toEqual({ paragraphId: p.id, offset: 2 });
  });

  it('elimina una selezione su più paragrafi prima di inserire', () => {
    const a = paragraph('abc');
    const b = paragraph('def');
    const editor = editorWith([a, b], range(b.id, 1, a.id, 1));
    editor.execute('insert.field', { kind: 'date' });
    expect(blocks(editor)).toHaveLength(1);
    expect(paragraphText(blocks(editor)[0]!)).toBe('a￼ef');
  });

  it('i campi prendono la formattazione del testo vicino', () => {
    const p = paragraph([textRun('Pagina ', { bold: true })]);
    const editor = editorWith([p], caret(p.id, 7));
    editor.execute('insert.field', { kind: 'page' });
    expect(blocks(editor)[0]!.children[1]).toEqual({ type: 'field', kind: 'page', format: { bold: true } });
  });

  it('l\'interruzione di pagina chiude il paragrafo e sposta il resto in uno nuovo', () => {
    const p = paragraph('primadopo');
    const editor = editorWith([p], caret(p.id, 5));
    editor.execute('insert.pageBreak');
    const [x, y] = blocks(editor);
    expect(x!.children).toEqual([textRun('prima'), { type: 'break', kind: 'page' }]);
    expect(paragraphText(y!)).toBe('dopo');
    expect(editor.state.selection.focus).toEqual({ paragraphId: y!.id, offset: 0 });
  });

  it('l\'interruzione di sezione divide la sezione ereditando la pagina', () => {
    const a = paragraph('uno');
    const b = paragraph('duetre');
    const c = paragraph('quattro');
    const editor = editorWith([a, b, c], caret(b.id, 3));
    editor.execute('page.setup', { page: { orientation: 'landscape' } });
    editor.execute('insert.sectionBreak', { type: 'continuous' });
    const [s1, s2] = editor.state.doc.sections;
    expect(s1!.blocks.map((x) => paragraphText(x as Paragraph))).toEqual(['uno', 'due']);
    expect(s2!.blocks.map((x) => paragraphText(x as Paragraph))).toEqual(['tre', 'quattro']);
    expect(s2!.breakType).toBe('continuous');
    expect(s2!.page).toEqual(s1!.page);
    expect(s2!.id).not.toBe(s1!.id);
  });

  it('non inserisce interruzioni di sezione dentro una tabella', () => {
    const editor = editorWith([paragraph()]);
    editor.execute('table.insert', { rows: 1, cols: 1 });
    expect(editor.canExecute('insert.sectionBreak', { type: 'nextPage' })).toBe(false);
  });
});

describe('page.setup', () => {
  it('cambia orientamento scambiando le dimensioni e unisce i margini', () => {
    const editor = editorWith([paragraph('x')]);
    editor.execute('page.setup', { page: { orientation: 'landscape', margins: { top: 10 } as never } });
    const page = editor.state.doc.sections[0]!.page;
    expect(page.orientation).toBe('landscape');
    expect(page.width).toBeCloseTo(PAGE_SIZES.A4.height);
    expect(page.height).toBeCloseTo(PAGE_SIZES.A4.width);
    expect(page.margins.top).toBe(10);
    expect(page.margins.left).toBeGreaterThan(0);
    editor.undo();
    expect(editor.state.doc.sections[0]!.page.orientation).toBe('portrait');
  });

  it('rifiuta impostazioni impossibili', () => {
    const editor = editorWith([paragraph('x')]);
    expect(editor.canExecute('page.setup', { page: { width: -1 } })).toBe(false);
    expect(editor.canExecute('page.setup', { page: { columns: { count: 0, gap: 10 } } })).toBe(false);
    expect(editor.execute('page.setup', { page: { ...PAGE_SIZES.Letter } })).toBe(true);
  });

  it('si applica a tutte le sezioni selezionate', () => {
    const a = paragraph('uno');
    const b = paragraph('due');
    const editor = editorWith([a, b], caret(b.id, 0));
    editor.execute('insert.sectionBreak', { type: 'nextPage' });
    const s2 = editor.state.doc.sections[1];
    editor.apply({ state: { ...editor.state, selection: range(a.id, 0, s2!.blocks[0]!.id, 0) }, meta: { addToHistory: false } });
    editor.execute('page.setup', { page: { columns: { count: 2, gap: 12 } } });
    expect(editor.state.doc.sections.every((s) => s.page.columns?.count === 2)).toBe(true);
  });
});
