import { describe, expect, it } from 'vitest';
import {
  caret,
  comparePositions,
  createDocument,
  findParagraph,
  normalizeInlines,
  orderedRange,
  paragraph,
  paragraphLength,
  paragraphText,
  paragraphs,
  replaceBlock,
  resolveCharFormat,
  resolveParaFormat,
  table,
  textRun,
  updateParagraph,
} from '../../src/core';

describe('model', () => {
  it('calcola lunghezza e testo con inline non testuali', () => {
    const p = paragraph([textRun('ab'), { type: 'tab' }, textRun('c'), { type: 'break', kind: 'line' }]);
    expect(paragraphLength(p)).toBe(5);
    expect(paragraphText(p)).toBe('ab\tc\n');
  });

  it('normalizza i run', () => {
    const out = normalizeInlines([textRun('a', { bold: true }), textRun(''), textRun('b', { bold: true }), textRun('c')]);
    expect(out).toEqual([textRun('ab', { bold: true }), textRun('c')]);
  });

  it('attraversa i paragrafi dentro le tabelle in ordine', () => {
    const t = table(2, 2, 400);
    const doc = createDocument([paragraph('prima'), t, paragraph('dopo')]);
    const list = [...paragraphs(doc)];
    expect(list).toHaveLength(6);
    expect(paragraphText(list[0]!.paragraph)).toBe('prima');
    expect(paragraphText(list[5]!.paragraph)).toBe('dopo');
    expect(list[2]!.path).toEqual([1, 0, 1, 0]);
  });

  it('aggiorna un paragrafo annidato senza toccare il resto', () => {
    const t = table(1, 2, 400);
    const before = paragraph('x');
    const doc = createDocument([before, t]);
    const target = t.rows[0]!.cells[1]!.blocks[0]!;
    if (target.type !== 'paragraph') throw new Error();
    const next = updateParagraph(doc, { ...target, children: [textRun('ciao')] });
    expect(paragraphText(findParagraph(next, target.id)!.paragraph)).toBe('ciao');
    expect(next.sections[0]!.blocks[0]).toBe(before);
    expect(doc).not.toBe(next);
    expect(() => replaceBlock(doc, 'inesistente', [])).toThrow();
  });

  it('confronta posizioni e ordina la selezione', () => {
    const a = paragraph('uno');
    const b = paragraph('due');
    const doc = createDocument([a, b]);
    expect(comparePositions(doc, { paragraphId: b.id, offset: 0 }, { paragraphId: a.id, offset: 3 })).toBeGreaterThan(0);
    const [from, to] = orderedRange(doc, { anchor: { paragraphId: b.id, offset: 1 }, focus: { paragraphId: a.id, offset: 2 } });
    expect(from.paragraphId).toBe(a.id);
    expect(to.paragraphId).toBe(b.id);
    expect(caret(a.id, 1).anchor).toEqual(caret(a.id, 1).focus);
  });

  it('risolve la cascata degli stili', () => {
    const doc = createDocument();
    const h = paragraph('Titolo', { align: 'center' }, 'Heading1');
    const run = textRun('Titolo', { italic: true });
    expect(resolveParaFormat(doc.styles, h)).toMatchObject({ align: 'center', keepWithNext: true, spaceAfter: 3 });
    expect(resolveCharFormat(doc.styles, h, run)).toMatchObject({ fontFamily: 'Calibri', fontSize: 16, bold: true, italic: true });
  });
});
