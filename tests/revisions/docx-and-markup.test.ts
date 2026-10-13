import { describe, expect, it } from 'vitest';
import {
  Editor,
  createDocument,
  createState,
  paragraph,
  paragraphText,
  paragraphs,
  textRun,
  type Paragraph,
  type TextRun,
} from '../../src/core';
import { exportDocx, importDocx } from '../../src/io';
import { readZip } from '../../src/io/util/zip';
import { createLayoutModule } from '../../src/layout';
import { createApproxMeasurer } from '../../src/layout/measure';
import { authorColor, createRevisionsModule } from '../../src/revisions';

const date = '2026-10-13T08:00:00Z';

function revisedDoc() {
  const p1: Paragraph = {
    ...paragraph([
      textRun('Il '),
      { ...textRun('gatto'), revision: { kind: 'delete', author: 'Nicola', date } },
      { ...textRun('cane', { bold: true }), revision: { kind: 'insert', author: 'Giulia & co', date } },
      textRun(' dorme'),
    ]),
    revision: { kind: 'insert', author: 'Nicola', date },
  };
  return createDocument([p1, paragraph('Fine')]);
}

describe('revisioni in DOCX', () => {
  it('scrive w:ins, w:del, w:delText e il segno di paragrafo inserito', async () => {
    const bytes = await exportDocx(revisedDoc());
    const xml = new TextDecoder().decode((await readZip(bytes)).get('word/document.xml')!);
    expect(xml).toMatch(/<w:del w:id="\d+" w:author="Nicola" w:date="2026-10-13T08:00:00Z"><w:r><w:delText xml:space="preserve">gatto<\/w:delText><\/w:r><\/w:del>/);
    expect(xml).toMatch(/<w:ins w:id="\d+" w:author="Giulia &amp; co" w:date="[^"]+"><w:r><w:rPr><w:b\/><\/w:rPr><w:t xml:space="preserve">cane<\/w:t><\/w:r><\/w:ins>/);
    expect(xml).toMatch(/<w:pPr><w:rPr><w:ins w:id="\d+" w:author="Nicola" w:date="[^"]+"\/><\/w:rPr><\/w:pPr>/);
    const ids = [...xml.matchAll(/<w:(?:ins|del) w:id="(\d+)"/g)].map((m) => m[1]);
    expect(new Set(ids).size).toBe(ids.length);
  });

  it('rilegge le revisioni scritte', async () => {
    const doc = await importDocx(await exportDocx(revisedDoc()));
    const [p1, p2] = [...paragraphs(doc)].map((l) => l.paragraph);
    expect(paragraphText(p1!)).toBe('Il gattocane dorme');
    const runs = p1!.children as TextRun[];
    expect(runs.find((r) => r.text === 'gatto')!.revision).toEqual({ kind: 'delete', author: 'Nicola', date });
    expect(runs.find((r) => r.text === 'cane')!.revision).toEqual({ kind: 'insert', author: 'Giulia & co', date });
    expect(runs.find((r) => r.text === 'cane')!.format.bold).toBe(true);
    expect(runs.find((r) => r.text === ' dorme')!.revision).toBeUndefined();
    expect(p1!.revision).toEqual({ kind: 'insert', author: 'Nicola', date });
    expect(p2!.revision).toBeUndefined();
  });
});

describe('evidenziazione delle revisioni', () => {
  function layoutRuns(order: 'layout-first' | 'revisions-first') {
    const editor = new Editor(createState(revisedDoc()));
    const layout = createLayoutModule({ measurer: createApproxMeasurer() });
    const revisions = createRevisionsModule();
    if (order === 'layout-first') {
      editor.use(layout);
      editor.use(revisions);
    } else {
      editor.use(revisions);
      editor.use(layout);
    }
    const result = editor.services.layout!.layout(editor.state.doc);
    return result.pages[0]!.lines.flatMap((l) => l.runs);
  }

  for (const order of ['layout-first', 'revisions-first'] as const) {
    it(`inserimenti sottolineati ed eliminazioni barrate nel colore dell'autore (${order})`, () => {
      const runs = layoutRuns(order);
      const gatto = runs.find((r) => r.text.startsWith('gatto'))!;
      const cane = runs.find((r) => r.text.startsWith('cane'))!;
      const dorme = runs.find((r) => r.text.includes('dorme'))!;
      expect(gatto.format).toMatchObject({ strike: true, color: authorColor('Nicola') });
      expect(cane.format).toMatchObject({ underline: 'single', bold: true, color: authorColor('Giulia & co') });
      expect(dorme.format.strike).toBeFalsy();
      expect(dorme.format.underline).toBeFalsy();
    });
  }

  it('alla disattivazione il motore torna quello originale', () => {
    const editor = new Editor();
    editor.use(createLayoutModule({ measurer: createApproxMeasurer() }));
    const plain = editor.services.layout;
    const off = editor.use(createRevisionsModule());
    expect(editor.services.layout).not.toBe(plain);
    off();
    expect(editor.services.layout).toBe(plain);
  });
});
