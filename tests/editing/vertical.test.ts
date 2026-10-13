import { describe, expect, it } from 'vitest';
import { Editor, createDocument, createState, paragraph, paragraphs, table, type Block, type Paragraph } from '../../src/core';
import { editingModule } from '../../src/editing';
import { createApproxMeasurer, createLayoutModule } from '../../src/layout';
import { defaultKeybindings, normalizeKey } from '../../src/ui/keymap';
import { focus, select } from './helpers';

/** Paragrafo senza spaziatura: righe alte uguali, conti semplici. */
const plain = (text: string, format = {}) => paragraph(text, { spaceAfter: 0, lineSpacing: { rule: 'multiple', value: 1 }, ...format });

function setup(blocks: Block[], lists = {}): Editor {
  const doc = { ...createDocument(blocks), lists };
  const editor = new Editor(createState(doc));
  editor.use(editingModule);
  editor.use(createLayoutModule({ measurer: createApproxMeasurer() }));
  return editor;
}

const move = (editor: Editor, direction: string, extend = false, distance?: number) =>
  editor.execute('edit.moveCaret', { direction: direction as 'lineDown', extend, distance });

/** Pagina e x del cursore. */
function caretAt(editor: Editor) {
  const engine = editor.services.layout!;
  return engine.caretRect(engine.layout(editor.state.doc), editor.state.selection.focus)!;
}

/** Testo lungo che va a capo più volte (parole di 4 lettere). */
const long = (words: number) => Array.from({ length: words }, (_, i) => `w${String(i).padStart(3, '0')}`).join(' ');

describe('frecce Su e Giù', () => {
  it('vanno alla riga sotto e sopra nella stessa colonna', () => {
    const p = plain(long(60));
    const editor = setup([p]);
    select(editor, [p, 7]);
    const start = caretAt(editor);
    move(editor, 'lineDown');
    const down = caretAt(editor);
    expect(focus(editor)[0]).toBe(p.id);
    expect(down.rect.y).toBeGreaterThan(start.rect.y);
    expect(Math.abs(down.rect.x - start.rect.x)).toBeLessThan(6);
    move(editor, 'lineUp');
    expect(focus(editor)).toEqual([p.id, 7]);
  });

  it('passano da un paragrafo all\'altro e ricordano la colonna di partenza', () => {
    const a = plain('una riga abbastanza lunga');
    const b = plain('corta');
    const c = plain('ancora una riga lunga come la prima');
    const editor = setup([a, b, c]);
    select(editor, [a, 20]);
    move(editor, 'lineDown');
    expect(focus(editor)).toEqual([b.id, 5]);
    move(editor, 'lineDown');
    expect(focus(editor)).toEqual([c.id, 20]);
    move(editor, 'lineUp');
    move(editor, 'lineUp');
    expect(focus(editor)).toEqual([a.id, 20]);
  });

  it('sulla prima riga Su va all\'inizio, sull\'ultima Giù va alla fine', () => {
    const a = plain('prima');
    const b = plain('ultima');
    const editor = setup([a, b]);
    select(editor, [a, 3]);
    move(editor, 'lineUp');
    expect(focus(editor)).toEqual([a.id, 0]);
    select(editor, [b, 3]);
    move(editor, 'lineDown');
    expect(focus(editor)).toEqual([b.id, 6]);
  });

  it('con Maiusc estendono la selezione', () => {
    const a = plain('alfa');
    const b = plain('beta');
    const editor = setup([a, b]);
    select(editor, [a, 2]);
    move(editor, 'lineDown', true);
    expect(editor.state.selection).toEqual({ anchor: { paragraphId: a.id, offset: 2 }, focus: { paragraphId: b.id, offset: 2 } });
    // Senza Maiusc, Su parte dall'inizio della selezione.
    move(editor, 'lineUp');
    expect(focus(editor)).toEqual([a.id, 0]);
  });

  it('in una tabella vanno alla cella sopra o sotto nella stessa colonna', () => {
    const before = plain('prima della tabella');
    const t = table(2, 2, 400);
    const after = plain('dopo');
    const editor = setup([before, t, after]);
    const cell = (r: number, c: number) => t.rows[r]!.cells[c]!.blocks[0] as Paragraph;
    // Testo più lungo nella cella a destra: Giù non deve finirci dentro.
    const all = [...paragraphs(editor.state.doc)].map((l) => l.paragraph);
    expect(all).toContain(cell(0, 0));
    select(editor, [before, 0]);
    move(editor, 'lineDown');
    expect(focus(editor)[0]).toBe(cell(0, 0).id);
    move(editor, 'lineDown');
    expect(focus(editor)[0]).toBe(cell(1, 0).id);
    move(editor, 'lineDown');
    expect(focus(editor)[0]).toBe(after.id);
    move(editor, 'lineUp');
    expect(focus(editor)[0]).toBe(cell(1, 0).id);
    select(editor, [cell(1, 1), 0]);
    move(editor, 'lineUp');
    expect(focus(editor)[0]).toBe(cell(0, 1).id);
  });

  it('entrano in un paragrafo rientrato anche se la colonna è più a sinistra', () => {
    const a = plain('normale');
    const b = plain('rientrato', { indentLeft: 72 });
    const c = plain('normale');
    const editor = setup([a, b, c]);
    select(editor, [a, 0]);
    move(editor, 'lineDown');
    expect(focus(editor)).toEqual([b.id, 0]);
    move(editor, 'lineDown');
    expect(focus(editor)).toEqual([c.id, 0]);
  });

  it('in tabella scelgono la cella della colonna anche se la vicina ha più righe', () => {
    const t = table(1, 2, 400);
    const after = plain('dopo la tabella, una riga lunga per arrivare alla seconda colonna');
    const editor = setup([t, after]);
    const left = t.rows[0]!.cells[0]!.blocks[0] as Paragraph;
    const right = t.rows[0]!.cells[1]!.blocks[0] as Paragraph;
    // La cella di sinistra va a capo più volte, quella di destra no.
    editor.execute('edit.setSelection', { selection: { anchor: { paragraphId: left.id, offset: 0 }, focus: { paragraphId: left.id, offset: 0 } } });
    editor.execute('edit.insertText', { text: long(20) });
    select(editor, [after, 50]);
    expect(caretAt(editor).rect.x).toBeGreaterThan(262);
    move(editor, 'lineUp');
    expect(focus(editor)[0]).toBe(right.id);
  });

  it('negli elenchi la colonna conta dal testo, non dal pallino', () => {
    const lists = { B: { id: 'B', levels: [{ format: 'bullet', text: '•', start: 1, indentLeft: 36, hanging: 18 }] } };
    const a = plain('elemento uno', { list: { listId: 'B', level: 0 } });
    const b = plain('elemento due', { list: { listId: 'B', level: 0 } });
    const editor = setup([a, b], lists);
    select(editor, [a, 0]);
    move(editor, 'lineDown');
    expect(focus(editor)).toEqual([b.id, 0]);
    select(editor, [a, 9]);
    move(editor, 'lineDown');
    expect(focus(editor)).toEqual([b.id, 9]);
  });

  it('attraversano il confine fra due pagine', () => {
    const paras = Array.from({ length: 80 }, (_, i) => plain(`riga ${i}`));
    const editor = setup(paras);
    const engine = editor.services.layout!;
    const layout = engine.layout(editor.state.doc);
    expect(layout.pages.length).toBeGreaterThan(1);
    const lastOfPage1 = layout.pages[0]!.lines.at(-1)!;
    const firstOfPage2 = layout.pages[1]!.lines[0]!;
    editor.execute('edit.setSelection', { selection: { anchor: { paragraphId: lastOfPage1.paragraphId, offset: 2 }, focus: { paragraphId: lastOfPage1.paragraphId, offset: 2 } } });
    move(editor, 'lineDown');
    expect(focus(editor)).toEqual([firstOfPage2.paragraphId, 2]);
    expect(caretAt(editor).page).toBe(1);
    move(editor, 'lineUp');
    expect(focus(editor)).toEqual([lastOfPage1.paragraphId, 2]);
  });
});

describe('Ctrl+Su e Ctrl+Giù', () => {
  it('vanno all\'inizio del paragrafo, del precedente e del seguente', () => {
    const a = plain('alfa');
    const b = plain('beta gamma');
    const editor = setup([a, b]);
    select(editor, [b, 5]);
    move(editor, 'paragraphUp');
    expect(focus(editor)).toEqual([b.id, 0]);
    move(editor, 'paragraphUp');
    expect(focus(editor)).toEqual([a.id, 0]);
    move(editor, 'paragraphDown', true);
    expect(editor.state.selection.anchor).toEqual({ paragraphId: a.id, offset: 0 });
    expect(focus(editor)).toEqual([b.id, 0]);
    move(editor, 'paragraphDown');
    expect(focus(editor)).toEqual([b.id, 10]);
  });
});

describe('Fine su una riga che va a capo', () => {
  it('lascia il cursore sulla riga; Maiusc+Fine seleziona anche lo spazio', () => {
    const p = plain(long(60));
    const editor = setup([p]);
    const engine = editor.services.layout!;
    const first = engine.layout(editor.state.doc).pages[0]!.lines[0]!;
    select(editor, [p, 2]);
    move(editor, 'lineEnd');
    expect(focus(editor)[1]).toBe(first.end - 1);
    expect(caretAt(editor).rect.y).toBe(first.y);
    select(editor, [p, 2]);
    move(editor, 'lineEnd', true);
    expect(focus(editor)[1]).toBe(first.end);
  });

  it('dopo un a capo manuale (Maiusc+Invio) resta prima dell\'a capo', () => {
    const p = plain([{ type: 'text', text: 'abc' }, { type: 'break', kind: 'line' }, { type: 'text', text: 'de' }] as never);
    const editor = setup([p]);
    select(editor, [p, 1]);
    move(editor, 'lineEnd');
    expect(focus(editor)[1]).toBe(3);
  });
});

describe('PagSu e PagGiù', () => {
  it('spostano il cursore di una schermata mantenendo la colonna', () => {
    const paras = Array.from({ length: 120 }, (_, i) => plain(`riga numero ${i}`));
    const editor = setup(paras);
    select(editor, [paras[0]!, 5]);
    const y0 = caretAt(editor).rect.y;
    move(editor, 'pageDown', false, 200);
    const after = caretAt(editor);
    expect(after.page).toBe(0);
    expect(after.rect.y - y0).toBeGreaterThan(180);
    expect(after.rect.y - y0).toBeLessThan(220);
    expect(focus(editor)[1]).toBe(5);
    move(editor, 'pageDown', false, 2000);
    expect(caretAt(editor).page).toBeGreaterThan(0);
    move(editor, 'pageUp', false, 100000);
    expect(focus(editor)).toEqual([paras[0]!.id, 5]);
  });
});

describe('scorciatoie', () => {
  it('frecce, Ctrl+frecce e PagSu/PagGiù hanno un comando, anche con Maiusc', () => {
    const keys = new Map(defaultKeybindings(false).map((b) => [normalizeKey(b.key, false), b]));
    expect(keys.get('ArrowUp')?.args).toEqual({ direction: 'lineUp' });
    expect(keys.get('Shift-ArrowDown')?.args).toEqual({ direction: 'lineDown', extend: true });
    expect(keys.get('Ctrl-ArrowDown')?.args).toEqual({ direction: 'paragraphDown' });
    expect(keys.get('Ctrl-Shift-ArrowUp')?.args).toEqual({ direction: 'paragraphUp', extend: true });
    expect(keys.get('PageDown')?.args).toEqual({ direction: 'pageDown' });
    expect(keys.get('Shift-PageUp')?.args).toEqual({ direction: 'pageUp', extend: true });
  });
});
