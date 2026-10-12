import { describe, expect, it } from 'vitest';
import {
  caret,
  cmToPt,
  findParagraph,
  paragraph,
  paragraphText,
  table,
  type Table,
} from '../../src/core';
import { tableOps } from '../../src/structures';
import { editorWith, focusText, grid, onlyTable, range, select } from './helpers';

/** Tabella con il testo indicato in ogni cella. */
function tbl(texts: string[][]): Table {
  const t = table(texts.length, texts[0]!.length, 300);
  return {
    ...t,
    rows: t.rows.map((r, i) => ({
      ...r,
      cells: r.cells.map((c, j) => ({ ...c, blocks: [paragraph(texts[i]![j]!)] })),
    })),
  };
}

function cellParagraph(t: Table, r: number, c: number): string {
  return t.rows[r]!.cells[c]!.blocks[0]!.id;
}

describe('table.insert', () => {
  it('inserisce una tabella dividendo il paragrafo al cursore', () => {
    const p = paragraph('primasecondo');
    const editor = editorWith([p], caret(p.id, 5));
    expect(editor.canExecute('table.insert', { rows: 0, cols: 2 })).toBe(false);
    expect(editor.execute('table.insert', { rows: 2, cols: 3 })).toBe(true);
    const blocks = editor.state.doc.sections[0]!.blocks;
    expect(blocks.map((b) => b.type)).toEqual(['paragraph', 'table', 'paragraph']);
    expect(paragraphText(blocks[0] as never)).toBe('prima');
    expect(paragraphText(blocks[2] as never)).toBe('secondo');
    const t = onlyTable(editor.state.doc);
    expect(grid(t)).toEqual([['', '', ''], ['', '', '']]);
    // Larghezza: A4 meno i margini predefiniti (2 cm + 2 cm).
    const total = t.columnWidths.reduce((a, b) => a + b, 0);
    expect(total).toBeCloseTo(editor.state.doc.sections[0]!.page.width - cmToPt(4));
    expect(editor.state.selection.focus.paragraphId).toBe(cellParagraph(t, 0, 0));
    expect(editor.canExecute('table.delete')).toBe(true);
    editor.undo();
    expect(editor.state.doc.sections[0]!.blocks).toHaveLength(1);
  });

  it('a inizio paragrafo mette la tabella prima; in una cella crea una tabella annidata', () => {
    const p = paragraph('testo');
    const editor = editorWith([p], caret(p.id, 0));
    editor.execute('table.insert', { rows: 1, cols: 2 });
    const blocks = editor.state.doc.sections[0]!.blocks;
    expect(blocks.map((b) => b.type)).toEqual(['table', 'paragraph']);
    expect(blocks[1]!.id).toBe(p.id);

    editor.execute('table.insert', { rows: 1, cols: 1 });
    const outer = onlyTable(editor.state.doc);
    const inner = outer.rows[0]!.cells[0]!.blocks[0] as Table;
    expect(inner.type).toBe('table');
    expect(inner.columnWidths[0]).toBeCloseTo(outer.columnWidths[0]!);
  });
});

describe('righe e colonne', () => {
  it('inserisce righe sopra e sotto copiando la struttura', () => {
    const t = tbl([['a', 'b'], ['c', 'd']]);
    const editor = editorWith([t], caret(cellParagraph(t, 0, 1), 0));
    editor.execute('table.insertRowBelow');
    expect(grid(onlyTable(editor.state.doc))).toEqual([['a', 'b'], ['', ''], ['c', 'd']]);
    // Il cursore va nella nuova riga, stessa colonna.
    const loc = findParagraph(editor.state.doc, editor.state.selection.focus.paragraphId)!;
    expect(loc.path).toEqual([0, 1, 1, 0]);
    editor.execute('table.insertRowAbove');
    expect(grid(onlyTable(editor.state.doc))).toEqual([['a', 'b'], ['', ''], ['', ''], ['c', 'd']]);
  });

  it('inserisce tante righe quante sono selezionate', () => {
    const t = tbl([['a', 'b'], ['c', 'd']]);
    const editor = editorWith([t], range(cellParagraph(t, 0, 0), 0, cellParagraph(t, 1, 1), 0));
    editor.execute('table.insertRowBelow');
    expect(grid(onlyTable(editor.state.doc))).toEqual([['a', 'b'], ['c', 'd'], ['', ''], ['', '']]);
  });

  it('inserisce colonne mantenendo la larghezza totale', () => {
    const t = tbl([['a', 'b'], ['c', 'd']]);
    const editor = editorWith([t], caret(cellParagraph(t, 1, 0), 0));
    editor.execute('table.insertColumnRight');
    const out = onlyTable(editor.state.doc);
    expect(grid(out)).toEqual([['a', '', 'b'], ['c', '', 'd']]);
    expect(out.columnWidths).toHaveLength(3);
    expect(out.columnWidths.reduce((x, y) => x + y, 0)).toBeCloseTo(300);
    editor.execute('table.insertColumnLeft');
    expect(grid(onlyTable(editor.state.doc))).toEqual([['a', '', '', 'b'], ['c', '', '', 'd']]);
  });

  it('allarga una cella unita orizzontalmente invece di spezzarla', () => {
    const t = tbl([['a', 'b', 'c'], ['d', 'e', 'f']]);
    const merged = tableOps.mergeCells(t, 0, 0, 0, 1)!;
    expect(grid(merged)).toEqual([['a|b*2', 'c'], ['d', 'e', 'f']]);
    const out = tableOps.insertColumns(merged, 0, 0, true);
    expect(grid(out)).toEqual([['a|b*3', 'c'], ['d', '', 'e', 'f']]);
  });

  it('elimina righe e colonne; eliminando tutto elimina la tabella', () => {
    const t = tbl([['a', 'b'], ['c', 'd']]);
    const after = paragraph('dopo');
    const editor = editorWith([t, after], caret(cellParagraph(t, 0, 0), 0));
    editor.execute('table.deleteRow');
    expect(grid(onlyTable(editor.state.doc))).toEqual([['c', 'd']]);
    expect(focusText(editor)).toBe('c');
    editor.execute('table.deleteColumn');
    expect(grid(onlyTable(editor.state.doc))).toEqual([['d']]);
    expect(onlyTable(editor.state.doc).columnWidths).toEqual([150]);
    editor.execute('table.deleteColumn');
    expect(editor.state.doc.sections[0]!.blocks.map((b) => b.type)).toEqual(['paragraph']);
    expect(editor.state.selection.focus.paragraphId).toBe(after.id);
    expect(editor.canExecute('table.deleteRow')).toBe(false);
  });

  it('table.delete lascia un paragrafo vuoto se la tabella era l\'unico blocco', () => {
    const t = tbl([['a']]);
    const editor = editorWith([t], caret(cellParagraph(t, 0, 0), 0));
    editor.execute('table.delete');
    const blocks = editor.state.doc.sections[0]!.blocks;
    expect(blocks).toHaveLength(1);
    expect(blocks[0]!.type).toBe('paragraph');
    expect(editor.state.selection.focus.paragraphId).toBe(blocks[0]!.id);
  });

  it('table.delete porta il cursore alla fine del paragrafo precedente', () => {
    const before = paragraph('prima');
    const t = tbl([['a']]);
    const editor = editorWith([before, t], caret(cellParagraph(t, 0, 0), 0));
    editor.execute('table.delete');
    expect(editor.state.selection.focus).toEqual({ paragraphId: before.id, offset: 5 });
  });
});

describe('unione e divisione', () => {
  it('unisce un rettangolo di celle e conserva il contenuto', () => {
    const t = tbl([['a', 'b', 'c'], ['d', '', 'f'], ['g', 'h', 'i']]);
    const editor = editorWith([t], caret(cellParagraph(t, 0, 0), 0));
    expect(editor.canExecute('table.mergeCells')).toBe(false);
    select(editor, range(cellParagraph(t, 0, 0), 0, cellParagraph(t, 1, 1), 0));
    expect(editor.execute('table.mergeCells')).toBe(true);
    const out = onlyTable(editor.state.doc);
    expect(grid(out)).toEqual([['a|b|d*2', 'c'], ['^*2', 'f'], ['g', 'h', 'i']]);
    expect(out.rows[0]!.cells[0]!.vMerge).toBe('restart');
    expect(focusText(editor)).toBe('a');
  });

  it('rifiuta unioni che taglierebbero una cella unita', () => {
    const t = tableOps.mergeCells(tbl([['a', 'b'], ['c', 'd']]), 0, 0, 0, 1)!;
    // Rettangolo che prende solo metà della cella larga.
    expect(tableOps.canMerge(t, 0, 1, 0, 0)).toBe(false);
    const v = tableOps.mergeCells(tbl([['a', 'b'], ['c', 'd'], ['e', 'f']]), 0, 1, 0, 0)!;
    // Partire dalla seconda riga di un'unione verticale non è valido.
    expect(tableOps.canMerge(v, 1, 2, 0, 0)).toBe(false);
    // Fermarsi a metà di un'unione verticale nemmeno.
    expect(tableOps.canMerge(v, 0, 0, 0, 1)).toBe(false);
  });

  it('divide una cella in colonne aggiungendo colonne di griglia', () => {
    const t = tbl([['a', 'b'], ['c', 'd']]);
    const editor = editorWith([t], caret(cellParagraph(t, 0, 0), 0));
    expect(editor.canExecute('table.splitCell', { rows: 1, cols: 1 })).toBe(false);
    editor.execute('table.splitCell', { rows: 1, cols: 3 });
    const out = onlyTable(editor.state.doc);
    expect(grid(out)).toEqual([['a', '', '', 'b'], ['c*3', 'd']]);
    expect(out.columnWidths.reduce((x, y) => x + y, 0)).toBeCloseTo(300);
    expect(focusText(editor)).toBe('a');
  });

  it('divide una cella in righe unendo verticalmente le altre celle della riga', () => {
    const t = tbl([['a', 'b']]);
    const out = tableOps.splitCell(t, 0, 0, 2, 1)!;
    expect(grid(out)).toEqual([['a', 'b'], ['', '^']]);
    expect(out.rows[0]!.cells[1]!.vMerge).toBe('restart');
  });

  it('dividere una cella unita la riporta allo stato iniziale', () => {
    const t = tbl([['a', 'b'], ['c', 'd']]);
    const merged = tableOps.mergeCells(t, 0, 1, 0, 1)!;
    expect(grid(merged)).toEqual([['a|b|c|d*2'], ['^*2']]);
    const back = tableOps.splitCell(merged, 0, 0, 2, 2)!;
    expect(grid(back)).toEqual([['a|b|c|d', ''], ['', '']]);
    expect(back.rows.flatMap((r) => r.cells).every((c) => !c.vMerge && !c.colSpan)).toBe(true);
    // Una cella unita su 2 righe non si divide in 3 righe.
    expect(tableOps.splitCell(merged, 0, 0, 3, 1)).toBeNull();
  });

  it('eliminando la prima riga di un\'unione verticale la cella sotto la riapre', () => {
    const t = tableOps.mergeCells(tbl([['a', 'b'], ['c', 'd'], ['e', 'f']]), 0, 2, 0, 0)!;
    const out = tableOps.deleteRows(t, 0, 0)!;
    expect(out.rows[0]!.cells[0]!.vMerge).toBe('restart');
    expect(out.rows[1]!.cells[0]!.vMerge).toBe('continue');
    const last = tableOps.deleteRows(out, 1, 1)!;
    expect(last.rows[0]!.cells[0]!.vMerge).toBeUndefined();
  });
});
