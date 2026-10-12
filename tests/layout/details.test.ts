import { describe, expect, it } from 'vitest';
import { paragraph, textRun } from '../../src/core';
import { computeMarkers, createApproxMeasurer, createLayoutEngine, formatNumber } from '../../src/layout';
import { doc, engine, lineText, plain } from './helpers';

describe('dettagli', () => {
  it('formatta i numeri degli elenchi', () => {
    expect([1, 4, 9, 14].map((n) => formatNumber(n, 'upperRoman'))).toEqual(['I', 'IV', 'IX', 'XIV']);
    expect([1, 26, 27].map((n) => formatNumber(n, 'lowerLetter'))).toEqual(['a', 'z', 'aa']);
    expect(formatNumber(3, 'bullet')).toBe('');
  });

  it('i punti elenco usano il testo del livello', () => {
    const d = doc([plain('x', { list: { listId: 'B', level: 0 } })]);
    const markers = computeMarkers({
      ...d,
      lists: { B: { id: 'B', levels: [{ format: 'bullet', text: '•', start: 1, indentLeft: 18, hanging: 18 }] } },
    });
    expect([...markers.values()].map((m) => m.text)).toEqual(['•']);
  });

  it("un'immagine alza la riga", () => {
    const p = paragraph([textRun('a'), { type: 'image', src: 'img', width: 50, height: 40 }], { spaceAfter: 0 });
    const [line] = engine.layout(doc([p])).pages[0]!.lines;
    expect(line!.height).toBeGreaterThanOrEqual(40);
    const img = line!.runs.find((r) => r.kind === 'image')!;
    expect(img).toMatchObject({ src: 'img', width: 50, height: 40, start: 1, end: 2 });
  });

  it('campi data e ora con la data iniettata', () => {
    const p = paragraph([{ type: 'field', kind: 'date', format: {} }]);
    const l = createLayoutEngine(createApproxMeasurer(), { now: () => new Date(2026, 9, 12) }).layout(doc([p]));
    expect(lineText(l.pages[0]!.lines[0]!)).toBe('12/10/2026');
  });

  it('lo sfondo del paragrafo copre le sue righe', () => {
    const p = plain('testo', { shading: '#ffff00', indentLeft: 10 });
    const page = engine.layout(doc([p])).pages[0]!;
    const fill = page.decorations.find((d) => d.kind === 'fill')!;
    expect(fill.color).toBe('#ffff00');
    expect(fill.rect.x).toBeCloseTo(page.contentRect.x + 10);
    expect(fill.rect.height).toBeCloseTo(page.lines[0]!.height);
  });

  it('restituisce la stessa impaginazione per lo stesso documento', () => {
    const d = doc([plain('a')]);
    expect(engine.layout(d)).toBe(engine.layout(d));
  });

  it('maiuscolo e apice', () => {
    const p = paragraph([textRun('ciao', { allCaps: true }), textRun('2', { verticalAlign: 'superscript' })]);
    const [line] = engine.layout(doc([p])).pages[0]!.lines;
    expect(lineText(line!)).toBe('CIAO2');
    const sup = line!.runs[1]!;
    expect(sup.width).toBeCloseTo(11 * 0.65 * 0.5);
  });
});
