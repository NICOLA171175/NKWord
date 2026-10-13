import { describe, expect, it } from 'vitest';
import { caret, createState, firstParagraph, paragraph, textRun, type DocumentModel, type Section } from '../../src/core';
import { insertFootnote } from '../../src/footnotes';
import { FOOTNOTE_SEPARATOR_SPACE } from '../../src/layout';
import { doc, engine, lineText, pageTexts, plain } from '../layout/helpers';

/** Pagina piccola: 10 righe da 11 pt di corpo. */
const small: Section['page'] = {
  width: 300,
  height: 150,
  orientation: 'portrait',
  margins: { top: 20, bottom: 20, left: 20, right: 20, header: 5, footer: 5 },
};

/** Aggiunge una nota alla fine del paragrafo `index`. */
function withNote(d: DocumentModel, index: number, text: string): DocumentModel {
  const p = d.sections[0]!.blocks[index]!;
  const state = { ...createState(d), selection: caret(p.id, p.type === 'paragraph' ? p.children.reduce((n, c) => n + (c.type === 'text' ? c.text.length : 1), 0) : 0) };
  return insertFootnote(state, text)!.state.doc;
}

describe('note a piè di pagina nel layout', () => {
  it('mostra il numero in apice nel testo e la nota in fondo alla pagina', () => {
    const d = withNote(doc([plain('Testo')], { page: small }), 0, 'La nota');
    const page = engine.layout(d).pages[0]!;
    const [line] = page.lines;
    expect(lineText(line!)).toBe('Testo1');
    const ref = line!.runs[line!.runs.length - 1]!;
    expect(ref).toMatchObject({ kind: 'field', text: '1', start: 5, end: 6 });
    expect(ref.format.verticalAlign).toBe('superscript');

    expect(page.footnotes?.map(lineText)).toEqual(['1 La nota']);
    const noteLine = page.footnotes![0]!;
    const bottom = page.contentRect.y + page.contentRect.height;
    // La nota finisce sul fondo dell'area del corpo, sopra il margine inferiore.
    expect(noteLine.y + noteLine.height).toBeCloseTo(bottom);
    expect(noteLine.x).toBeCloseTo(small.margins.left);
    // Corpo 10 pt (stile Testo nota): più bassa di una riga del corpo da 11 pt con il richiamo.
    expect(noteLine.height).toBeLessThan(line!.height);
    // Riga separatrice sopra la nota.
    const sep = page.decorations.find((dd) => dd.kind === 'border' && dd.rect.y < noteLine.y && dd.rect.y > noteLine.y - FOOTNOTE_SEPARATOR_SPACE);
    expect(sep).toBeDefined();
  });

  it('numera 1, 2, 3 nell\'ordine del documento e ricomincia il conto per ogni layout', () => {
    let d = doc([plain('a'), plain('b'), plain('c')], { page: small });
    d = withNote(d, 2, 'terza');
    d = withNote(d, 0, 'prima');
    d = withNote(d, 1, 'seconda');
    const page = engine.layout(d).pages[0]!;
    expect(page.lines.map(lineText)).toEqual(['a1', 'b2', 'c3']);
    expect(page.footnotes!.map(lineText)).toEqual(['1 prima', '2 seconda', '3 terza']);
  });

  it('accorcia il corpo: il testo che non ci sta passa alla pagina dopo', () => {
    const lines = Array.from({ length: 10 }, (_, i) => plain(`riga ${i}`));
    const without = engine.layout(doc(lines, { page: small }));
    expect(pageTexts(without)[0]).toHaveLength(10);
    const d = withNote(doc(lines, { page: small }), 0, 'nota');
    const layout = engine.layout(d);
    const first = layout.pages[0]!;
    // Nota + spazio del separatore: qualche riga in meno sulla prima pagina.
    const n = first.lines.length;
    expect(n).toBeLessThan(10);
    expect(n).toBeGreaterThan(5);
    expect(layout.pages[1]!.lines.map(lineText)).toEqual(lines.slice(n).map((_, i) => `riga ${n + i}`));
    const last = first.lines[first.lines.length - 1]!;
    expect(last.y + last.height).toBeLessThanOrEqual(first.footnotes![0]!.y - FOOTNOTE_SEPARATOR_SPACE + 0.01);
    expect(layout.pages[1]!.footnotes).toBeUndefined();
  });

  it('una riga con il richiamo va alla pagina dopo insieme alla sua nota', () => {
    const lines = Array.from({ length: 10 }, (_, i) => plain(`riga ${i}`));
    const d = withNote(doc(lines, { page: small }), 9, 'nota');
    const layout = engine.layout(d);
    expect(layout.pages[0]!.lines).toHaveLength(9);
    expect(layout.pages[0]!.footnotes).toBeUndefined();
    expect(layout.pages[1]!.lines.map(lineText)).toEqual(['riga 91']);
    expect(layout.pages[1]!.footnotes!.map(lineText)).toEqual(['1 nota']);
  });

  it('impagina anche le note richiamate in una tabella', async () => {
    const { table } = await import('../../src/core');
    const t = table(1, 1, 200);
    const cellPara = paragraph([textRun('cella')]);
    const withCell = { ...t, rows: [{ ...t.rows[0]!, cells: [{ ...t.rows[0]!.cells[0]!, blocks: [cellPara] }] }] };
    const d0 = doc([withCell], { page: small });
    const state = { ...createState(d0), selection: caret(cellPara.id, 5) };
    const d = insertFootnote(state, 'in tabella')!.state.doc;
    const page = engine.layout(d).pages[0]!;
    expect(page.lines.map(lineText)).toEqual(['cella1']);
    expect(page.footnotes!.map(lineText)).toEqual(['1 in tabella']);
  });

  it('un documento senza note non ha l\'area delle note', () => {
    const page = engine.layout(doc([plain('x')])).pages[0]!;
    expect(page.footnotes).toBeUndefined();
    expect(firstParagraph(doc([plain('x')])).children).toHaveLength(1);
  });
});
