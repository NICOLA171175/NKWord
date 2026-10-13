import { describe, expect, it } from 'vitest';
import { caret, createDocument, createState, Editor, paragraph, type DocumentModel, type Drawing } from '../../src/core';
import {
  chartPrimitives,
  createDrawing,
  createDrawingsModule,
  drawingList,
  drawingPaths,
  fromLineEnds,
  gridToChart,
  chartToGrid,
  lineEnds,
  parseNumber,
  reanchorOrphans,
  removeDrawing,
  reorderDrawing,
  resizeGeometry,
  ribbonItems,
  sampleChart,
  setDrawing,
} from '../../src/drawings';
import { axisRange, niceStep } from '../../src/drawings/chart';
import { wrapText } from '../../src/drawings/pdf';
import { exportDocx, importDocx } from '../../src/io';
import { readZip } from '../../src/io/util/zip';
import { createApproxMeasurer, createLayoutEngine } from '../../src/layout';
import { layoutToPdf } from '../../src/pdf';

const measurer = createApproxMeasurer();
const engine = createLayoutEngine(measurer);
const dec = new TextDecoder();
const latin1 = (b: Uint8Array) => Array.from(b, (c) => String.fromCharCode(c)).join('');

function setup(texts = ['primo', 'secondo']) {
  const doc = createDocument(texts.map((t) => paragraph(t)));
  const ids = doc.sections[0]!.blocks.map((b) => b.id);
  const editor = new Editor({ ...createState(doc), selection: caret(ids[0]!, 0) });
  const module = createDrawingsModule();
  editor.use(module);
  return { editor, module, ids };
}

const only = (doc: DocumentModel) => drawingList(doc)[0]!;

describe('modello', () => {
  it('crea forme con i colori e le dimensioni di Word', () => {
    const rect = createDrawing({ kind: 'shape', shape: 'rect' }, 'p', 10, 20);
    expect(rect).toMatchObject({ kind: 'shape', shape: 'rect', x: 10, y: 20, width: 108, height: 72, fill: '#4472c4' });
    expect(rect.id).toMatch(/^d/);
    const line = createDrawing({ kind: 'shape', shape: 'arrow' }, 'p', 0, 0);
    expect(line).toMatchObject({ height: 0, fill: null });
    const box = createDrawing({ kind: 'textbox' }, 'p', 0, 0);
    expect(box.text).toMatchObject({ text: '', align: 'left', verticalAlign: 'top' });
    const chart = createDrawing({ kind: 'chart', type: 'pie' }, 'p', 0, 0);
    expect(chart.chart?.type).toBe('pie');
  });

  it('riordina, elimina e riancora i disegni', () => {
    let doc = createDocument([paragraph('a'), paragraph('b')]);
    const [a, b] = doc.sections[0]!.blocks.map((x) => x.id);
    const d1 = createDrawing({ kind: 'shape', shape: 'rect' }, a!, 0, 0);
    const d2 = createDrawing({ kind: 'shape', shape: 'ellipse' }, b!, 0, 0);
    doc = setDrawing(setDrawing(doc, d1), d2);
    expect(drawingList(reorderDrawing(doc, d2.id, 'back')).map((d) => d.id)).toEqual([d2.id, d1.id]);
    expect(drawingList(reorderDrawing(doc, d1.id, 'front')).map((d) => d.id)).toEqual([d2.id, d1.id]);
    const without = removeDrawing(removeDrawing(doc, d1.id), d2.id);
    expect(without.drawings).toBeUndefined();
    // Il paragrafo b sparisce: la forma passa ad a.
    const merged = { ...doc, sections: [{ ...doc.sections[0]!, blocks: [doc.sections[0]!.blocks[0]!] }] };
    expect(reanchorOrphans(merged, a!).drawings![d2.id]!.anchor).toBe(a);
    expect(reanchorOrphans(doc, a!)).toBe(doc);
  });

  it('converte gli estremi di una linea e ridimensiona dalle maniglie', () => {
    const g = fromLineEnds([100, 50], [20, 80]);
    expect(g).toEqual({ x: 20, y: 50, width: 80, height: 30, flipH: true, flipV: false });
    expect(lineEnds(g)).toEqual([[100, 50], [20, 80]]);
    const box = createDrawing({ kind: 'shape', shape: 'rect' }, 'p', 100, 100);
    expect(resizeGeometry(box, 'se', 10, 20)).toMatchObject({ x: 100, y: 100, width: 118, height: 92 });
    expect(resizeGeometry(box, 'nw', 10, 20)).toMatchObject({ x: 110, y: 120, width: 98, height: 52 });
    // Non scende sotto la dimensione minima.
    expect(resizeGeometry(box, 'e', -500, 0).width).toBe(4);
    // Con Maiusc mantiene le proporzioni.
    const kept = resizeGeometry(box, 'se', 108, 0, true);
    expect(kept.width / kept.height).toBeCloseTo(108 / 72, 5);
    const line = createDrawing({ kind: 'shape', shape: 'line' }, 'p', 0, 0);
    expect(resizeGeometry(line, 'end', -200, 30)).toMatchObject({ x: -92, y: 0, width: 92, height: 30, flipH: true });
  });
});

describe('comandi', () => {
  it('inserisce, modifica, riordina ed elimina con Annulla', () => {
    const { editor, module, ids } = setup();
    expect(editor.canExecute('drawing.update', { changes: { fill: '#ff0000' } })).toBe(false);
    expect(editor.execute('drawing.insert', { kind: 'shape', shape: 'ellipse' })).toBe(true);
    const d = only(editor.state.doc);
    expect(d.anchor).toBe(ids[0]);
    expect(module.selection.id).toBe(d.id);
    expect(editor.execute('drawing.update', { changes: { fill: '#ff0000', x: 200 } })).toBe(true);
    expect(only(editor.state.doc)).toMatchObject({ fill: '#ff0000', x: 200, kind: 'shape' });
    editor.execute('drawing.insert', { kind: 'textbox' });
    const box = drawingList(editor.state.doc)[1]!;
    editor.execute('drawing.order', { id: box.id, to: 'back' });
    expect(drawingList(editor.state.doc)[0]!.id).toBe(box.id);
    expect(module.selection.id).toBe(box.id);
    editor.execute('drawing.delete', { id: box.id });
    expect(drawingList(editor.state.doc)).toHaveLength(1);
    expect(module.selection.id).toBeNull();
    module.selection.select(d.id);
    editor.execute('drawing.delete');
    expect(editor.state.doc.drawings).toBeUndefined();
    expect(module.selection.id).toBeNull();
    editor.undo();
    expect(only(editor.state.doc).fill).toBe('#ff0000');
  });

  it('mette la forma nuova sotto il cursore quando c\'è l\'impaginazione', () => {
    const { editor } = setup();
    editor.services.layout = engine;
    editor.execute('drawing.insert', { kind: 'shape', shape: 'rect' });
    const d = only(editor.state.doc);
    const margins = editor.state.doc.sections[0]!.page.margins;
    expect(d.x).toBeCloseTo(margins.left, 0);
    expect(d.y).toBeGreaterThan(margins.top);
  });

  it('se il paragrafo di ancoraggio sparisce la forma passa al paragrafo del cursore', () => {
    const { editor, ids } = setup();
    editor.apply({ state: { ...editor.state, selection: caret(ids[1]!, 0) } });
    editor.execute('drawing.insert', { kind: 'shape', shape: 'rect' });
    const doc = editor.state.doc;
    const first = doc.sections[0]!.blocks[0]!;
    // Come unire i due paragrafi con Backspace.
    editor.apply({ state: { doc: { ...doc, sections: [{ ...doc.sections[0]!, blocks: [first] }] }, selection: caret(first.id, 5) } });
    expect(only(editor.state.doc).anchor).toBe(first.id);
  });

  it('dichiara i pulsanti della scheda Inserisci', () => {
    const labels = ribbonItems.map((i) => i.label);
    expect(labels).toEqual(expect.arrayContaining(['Forme', 'Grafico', 'Casella di testo', 'Riempimento', 'Contorno', 'Elimina forma']));
    expect(ribbonItems.every((i) => i.tab === 'insert')).toBe(true);
    const shapes = ribbonItems.find((i) => i.label === 'Forme')!;
    expect(shapes.options!.length).toBeGreaterThan(10);
  });
});

describe('geometria e grafici', () => {
  it('disegna le frecce con le punte', () => {
    const arrow = { ...createDrawing({ kind: 'shape', shape: 'doubleArrow' }, 'p', 0, 0), width: 100, height: 0 };
    const paths = drawingPaths(arrow);
    expect(paths).toHaveLength(3);
    expect(paths[1]!.fill).toBe(arrow.stroke);
    const ellipse = drawingPaths(createDrawing({ kind: 'shape', shape: 'ellipse' }, 'p', 0, 0))[0]!;
    expect(ellipse.segments.filter((s) => s.op === 'C')).toHaveLength(4);
  });

  it('calcola assi rotondi e una barra per valore', () => {
    expect(niceStep(4.5)).toBe(1);
    expect(niceStep(230)).toBe(50);
    expect(axisRange(sampleChart('column'))).toEqual({ min: 0, max: 5, step: 1 });
    const chart = sampleChart('column');
    const prims = chartPrimitives(chart, 360, 216);
    const bars = prims.filter((p) => p.t === 'rect' && p.w > 7);
    expect(bars).toHaveLength(12);
    const texts = prims.filter((p) => p.t === 'text').map((p) => (p as { text: string }).text);
    expect(texts).toEqual(expect.arrayContaining(['Titolo del grafico', 'Categoria 1', 'Serie 3', '5']));
    const pie = chartPrimitives(sampleChart('pie'), 300, 200).filter((p) => p.t === 'path');
    expect(pie).toHaveLength(4);
    const line = chartPrimitives(sampleChart('line'), 300, 200).filter((p) => p.t === 'path' && p.width > 2);
    expect(line).toHaveLength(3);
  });

  it('legge e scrive la tabella dei dati', () => {
    expect(parseNumber('3,5')).toBe(3.5);
    expect(parseNumber('1.234,5')).toBe(1234.5);
    expect(parseNumber('2.25')).toBe(2.25);
    expect(parseNumber('abc')).toBe(0);
    const chart = sampleChart('column');
    const grid = chartToGrid(chart);
    expect(grid.cells[0]).toEqual(['', 'Serie 1', 'Serie 2', 'Serie 3']);
    expect(grid.cells[1]).toEqual(['Categoria 1', '4,3', '2,4', '2']);
    const edited = gridToChart({ ...grid, type: 'bar', cells: [['', 'Vendite'], ['Nord', '10,5'], ['', '7']] });
    expect(edited).toEqual({ type: 'bar', title: 'Titolo del grafico', categories: ['Nord', 'Categoria 2'], series: [{ name: 'Vendite', values: [10.5, 7] }] });
  });

  it('va a capo per parole', () => {
    const w = (s: string) => s.length;
    expect(wrapText('uno due tre', 7, w)).toEqual(['uno due', 'tre']);
    expect(wrapText('a\nb', 10, w)).toEqual(['a', 'b']);
    expect(wrapText('lunghissima', 4, w)).toEqual(['lung', 'hiss', 'ima']);
  });
});

function docWithDrawings(): DocumentModel {
  let doc = createDocument([paragraph('Testo'), paragraph('Altro')]);
  const [a, b] = doc.sections[0]!.blocks.map((x) => x.id);
  const shape: Drawing = {
    ...createDrawing({ kind: 'shape', shape: 'star5' }, a!, 100, 150),
    text: { text: 'Ciao\nmondo', format: { fontSize: 14, color: '#ffffff', bold: true }, align: 'center', verticalAlign: 'middle' },
  };
  const box: Drawing = { ...createDrawing({ kind: 'textbox' }, b!, 300, 200), text: { text: 'Nota <importante> & basta', format: { fontSize: 11, color: '#000000' }, align: 'left', verticalAlign: 'top' } };
  const arrow: Drawing = { ...createDrawing({ kind: 'shape', shape: 'arrow' }, a!, 80, 400), ...fromLineEnds([200, 400], [80, 460]) };
  const chart: Drawing = { ...createDrawing({ kind: 'chart', type: 'bar' }, b!, 72, 500), chart: { ...sampleChart('bar'), series: [{ name: 'Vendite', values: [1, 2.5, 3, 4], color: '#ff0000' }] } };
  const pie = createDrawing({ kind: 'chart', type: 'pie' }, b!, 72, 100);
  for (const d of [shape, box, arrow, chart, pie]) doc = setDrawing(doc, d);
  return doc;
}

describe('DOCX', () => {
  it('scrive forme ancorate alla pagina e un grafico con i suoi dati', async () => {
    const zip = await readZip(await exportDocx(docWithDrawings()));
    const xml = dec.decode(zip.get('word/document.xml'));
    expect(xml).toContain('xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"');
    expect(xml.match(/<wp:anchor /g)).toHaveLength(5);
    expect(xml).toContain('<wp:positionH relativeFrom="page"><wp:posOffset>1270000</wp:posOffset>');
    expect(xml).toContain('<a:prstGeom prst="star5">');
    expect(xml).toContain('<wps:cNvSpPr txBox="1"/>');
    expect(xml).toContain('Nota &lt;importante&gt; &amp; basta');
    expect(xml).toContain('<a:tailEnd type="triangle"/>');
    expect(xml).toMatch(/<a:xfrm flipH="1">/);
    // Il primo paragrafo comincia con i suoi disegni.
    expect(xml.indexOf('prst="star5"')).toBeLessThan(xml.indexOf('>Testo<'));
    const chart = dec.decode(zip.get('word/charts/chart1.xml'));
    expect(chart).toContain('<c:barDir val="bar"/>');
    expect(chart).toContain('<c:v>Vendite</c:v>');
    expect(chart).toContain('<c:pt idx="1"><c:v>2.5</c:v></c:pt>');
    expect(dec.decode(zip.get('word/charts/chart2.xml'))).toContain('<c:pieChart>');
    expect(dec.decode(zip.get('word/_rels/document.xml.rels'))).toContain('Target="charts/chart1.xml"');
    expect(dec.decode(zip.get('[Content_Types].xml'))).toContain('PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"');
  });

  it('riapre le proprie forme e i grafici uguali', async () => {
    const doc = docWithDrawings();
    const back = await importDocx(await exportDocx(doc));
    const before = drawingList(doc);
    const after = drawingList(back);
    expect(after).toHaveLength(before.length);
    const paraIds = back.sections[0]!.blocks.map((b) => b.id);
    before.forEach((d, i) => {
      const r = after[i]!;
      const { id: _a, anchor: anchorBefore, ...restBefore } = d;
      const { id: _b, anchor, ...restAfter } = r;
      expect(anchor).toBe(paraIds[doc.sections[0]!.blocks.findIndex((b) => b.id === anchorBefore)]);
      // Un ribaltamento assente equivale a false.
      const clean = (o: object) => Object.fromEntries(Object.entries(o).filter(([, v]) => v !== false));
      expect(clean(restAfter)).toEqual(clean(restBefore));
    });
  });

  it('legge le forme salvate da Word dentro mc:AlternateContent, con posizione dal margine', async () => {
    const doc = createDocument([paragraph('x')]);
    const zip = await readZip(await exportDocx(setDrawing(doc, createDrawing({ kind: 'shape', shape: 'hexagon' }, doc.sections[0]!.blocks[0]!.id, 100, 100))));
    let xml = dec.decode(zip.get('word/document.xml'));
    xml = xml
      .replace('xmlns:wps=', 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wps=')
      .replace('<w:drawing>', '<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>')
      .replace('</w:drawing>', '</w:drawing></mc:Choice><mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent>')
      .replace('<wp:positionH relativeFrom="page">', '<wp:positionH relativeFrom="margin">');
    const { writeZip } = await import('../../src/io/util/zip');
    const files = [...zip].map(([name, data]) => [name, name === 'word/document.xml' ? new TextEncoder().encode(xml) : data] as [string, Uint8Array]);
    const back = await importDocx(await writeZip(files));
    const d = only(back);
    expect(d.shape).toBe('hexagon');
    expect(d.x).toBeCloseTo(100 + back.sections[0]!.page.margins.left, 3);
    expect(d.y).toBeCloseTo(100, 3);
  });
});

describe('PDF', () => {
  it('disegna forme, testo delle forme e grafici sulla pagina dell\'ancoraggio', async () => {
    const doc = docWithDrawings();
    const pdf = latin1(await layoutToPdf(engine.layout(doc), { measurer, now: () => new Date(0) }));
    // Ogni disegno ha il suo sistema di coordinate con l'asse y verso il basso.
    expect(pdf.match(/q 1 0 0 -1 /g)).toHaveLength(5);
    expect(pdf).toContain('q 1 0 0 -1 100 691.89 cm');
    // Riempimento blu della stella e testo bianco in grassetto.
    expect(pdf).toContain('0.267 0.447 0.769 rg');
    expect(pdf).toContain('/Helvetica-Bold');
    const hex = (s: string) => `<${Array.from(s, (c) => c.charCodeAt(0).toString(16).padStart(2, '0')).join('')}>`;
    expect(pdf).toContain(hex('Ciao'));
    expect(pdf).toContain(hex('Vendite'));
    // Fette della torta: curve di Bézier.
    expect(pdf).toMatch(/ c [\d. ]+ c /);
  });
});
