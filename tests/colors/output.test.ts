import { describe, expect, it } from 'vitest';
import { paragraph, textRun, type BorderLine, type DocumentModel, type Paragraph, type TextRun } from '../../src/core';
import { documentToHtml, exportDocx, htmlToDocument, importDocx } from '../../src/io';
import { readZip } from '../../src/io/util/zip';
import { layoutToPdf } from '../../src/pdf';
import { doc, engine, measurer, plain } from '../layout/helpers';
import { plainParagraphs } from '../io/fixtures';

const line: BorderLine = { style: 'single', width: 0.5, color: '#000000' };
const box = { top: line, right: line, bottom: line, left: line };
const runs = (p: Paragraph) => p.children.filter((c): c is TextRun => c.type === 'text');

function colorful(): DocumentModel {
  return doc([
    paragraph([textRun('rosso', { color: '#c00000' }), textRun(' evidenziato', { highlight: '#00ffff' }), textRun(' tema', { highlight: '#dae3f3' })]),
    paragraph('riquadro', { shading: '#fff2cc', borders: box }),
    paragraph('solo sotto', { borders: { bottom: { ...line, color: '#4472c4', width: 1.5 } } }),
  ]);
}

describe('impaginazione', () => {
  it('paragrafi consecutivi con gli stessi bordi formano un solo riquadro', () => {
    const page = engine.layout(doc([plain('uno', { borders: box, spaceAfter: 6 }), plain('due', { borders: box, spaceAfter: 6 }), plain('tre')])).pages[0]!;
    const borders = page.decorations.filter((d) => d.kind === 'border');
    // Un bordo sopra, uno sotto e i due lati di ciascun paragrafo.
    const horizontal = borders.filter((d) => d.rect.width > d.rect.height);
    expect(horizontal).toHaveLength(2);
    const [first, second] = page.lines;
    expect(horizontal[0]!.rect.y).toBeLessThan(first!.y);
    expect(horizontal[1]!.rect.y).toBeGreaterThan(second!.y + second!.height);
    // I lati coprono anche lo spazio fra i due paragrafi.
    const left = borders.filter((d) => d.rect.height > d.rect.width && d.rect.x < first!.x).sort((a, b) => a.rect.y - b.rect.y);
    expect(left).toHaveLength(2);
    expect(left[1]!.rect.y).toBeLessThanOrEqual(left[0]!.rect.y + left[0]!.rect.height);
    expect(left[1]!.rect.y + left[1]!.rect.height).toBeGreaterThan(second!.y + second!.height);
  });

  it('il testo resta staccato dal bordo (spazio di Word: 4 pt ai lati, 1 pt sopra e sotto)', () => {
    const page = engine.layout(doc([plain('testo', { borders: box })])).pages[0]!;
    const l = page.lines[0]!;
    const left = page.decorations.find((d) => d.kind === 'border' && d.rect.height > d.rect.width && d.rect.x < l.x)!;
    expect(l.x - (left.rect.x + left.rect.width / 2)).toBeCloseTo(4);
    const top = page.decorations.find((d) => d.kind === 'border' && d.rect.width > d.rect.height)!;
    expect(l.y - (top.rect.y + top.rect.height / 2)).toBeCloseTo(1);
  });

  it('lo sfondo di paragrafi uguali è continuo, uno diverso no', () => {
    const page = engine.layout(doc([plain('a', { shading: '#ffff00', spaceAfter: 10 }), plain('b', { shading: '#ffff00', spaceAfter: 10 }), plain('c', { shading: '#00ff00' })])).pages[0]!;
    const fills = page.decorations.filter((d) => d.kind === 'fill');
    const [a, b, c] = page.lines;
    expect(fills[1]!.rect.y).toBeCloseTo(a!.y + a!.height);
    expect(fills[1]!.rect.y + fills[1]!.rect.height).toBeCloseTo(b!.y + b!.height);
    expect(fills[2]!.rect.y).toBeCloseTo(c!.y);
  });
});

describe('DOCX', () => {
  it('scrive e rilegge colore, evidenziatore, sfondo e bordi', async () => {
    const bytes = await exportDocx(colorful());
    const xml = new TextDecoder().decode((await readZip(bytes)).get('word/document.xml'));
    expect(xml).toContain('<w:color w:val="C00000"/>');
    expect(xml).toContain('<w:highlight w:val="cyan"/>');
    expect(xml).toContain('<w:shd w:val="clear" w:color="auto" w:fill="DAE3F3"/>');
    expect(xml).toContain('<w:top w:val="single" w:sz="4" w:space="1" w:color="000000"/>');
    expect(xml).toContain('<w:left w:val="single" w:sz="4" w:space="4" w:color="000000"/>');
    const [p1, p2, p3] = plainParagraphs(await importDocx(bytes));
    expect(runs(p1!).map((r) => [r.format.color, r.format.highlight])).toEqual([
      ['#c00000', undefined],
      [undefined, '#00ffff'],
      [undefined, '#dae3f3'],
    ]);
    expect(p2!.format).toMatchObject({ shading: '#fff2cc', borders: box });
    expect(p3!.format.borders).toEqual({ bottom: { style: 'single', width: 1.5, color: '#4472c4' } });
  });
});

describe('HTML', () => {
  it('esporta e rilegge colori, sfondo e bordi', () => {
    const html = documentToHtml(colorful());
    expect(html).toContain('color:#c00000');
    expect(html).toContain('background-color:#00ffff');
    expect(html).toMatch(/background-color:#fff2cc[^"]*border-top:0\.5pt solid #000000[^"]*padding:1pt 4pt/);
    const [p1, p2, p3] = plainParagraphs(htmlToDocument(html));
    expect(runs(p1!)[0]!.format.color).toBe('#c00000');
    expect(runs(p1!)[1]!.format.highlight).toBe('#00ffff');
    expect(p2!.format).toMatchObject({ shading: '#fff2cc', borders: box });
    expect(p3!.format.borders?.bottom).toEqual({ style: 'single', width: 1.5, color: '#4472c4' });
  });
});

describe('PDF', () => {
  it('disegna testo colorato, evidenziatore, sfondo e bordi', async () => {
    const bytes = await layoutToPdf(engine.layout(colorful()), { measurer, now: () => new Date(Date.UTC(2026, 9, 13)) });
    const pdf = Array.from(bytes, (c) => String.fromCharCode(c)).join('');
    const rg = (hex: string) => [1, 3, 5].map((i) => String(Math.round((parseInt(hex.slice(i, i + 2), 16) / 255) * 1000) / 1000)).join(' ') + ' rg';
    for (const c of ['#c00000', '#00ffff', '#fff2cc', '#4472c4']) expect(pdf).toContain(rg(c));
  });
});
