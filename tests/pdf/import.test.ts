import { describe, expect, it } from 'vitest';
import { createDocument, Editor, paragraph, textRun, type Block, type DocumentModel, type Paragraph, type TextRun } from '../../src/core';
import { createIoModule } from '../../src/io';
import { toDataUri } from '../../src/io/util/encoding';
import { createApproxMeasurer, createLayoutEngine } from '../../src/layout';
import { createPdfModule, encodePng, fontFamilyFromPdfName, importPdf, layoutToPdf } from '../../src/pdf';

const measurer = createApproxMeasurer();
const engine = createLayoutEngine(measurer);

async function roundTrip(doc: DocumentModel): Promise<DocumentModel> {
  return importPdf(await layoutToPdf(engine.layout(doc), { measurer }));
}

const paragraphs = (doc: DocumentModel) => doc.sections[0]!.blocks.filter((b): b is Paragraph => b.type === 'paragraph');
const textOf = (p: Paragraph) => p.children.map((c) => (c.type === 'text' ? c.text : '')).join('');
const runs = (p: Paragraph) => p.children.filter((c): c is TextRun => c.type === 'text');

describe('apertura di un PDF', () => {
  it('ricostruisce paragrafi, grassetto, corsivo e corpo', async () => {
    const long = 'Questa frase è abbastanza lunga da andare a capo più volte nella pagina. '.repeat(4).trim();
    const doc = await roundTrip(
      createDocument([
        paragraph([textRun('Titolo grande', { fontSize: 20, bold: true })]),
        paragraph([textRun('Testo normale, '), textRun('grassetto', { bold: true }), textRun(' e '), textRun('corsivo', { italic: true }), textRun('. ' + long)]),
        paragraph('Secondo paragrafo.'),
      ]),
    );
    const ps = paragraphs(doc);
    expect(ps.map(textOf)).toEqual(['Titolo grande', `Testo normale, grassetto e corsivo. ${long}`, 'Secondo paragrafo.']);
    expect(runs(ps[0]!)[0]!.format).toMatchObject({ bold: true, fontSize: 20 });
    const second = runs(ps[1]!);
    expect(second.find((r) => r.text.trim() === 'grassetto')?.format.bold).toBe(true);
    expect(second.find((r) => r.text.trim() === 'corsivo')?.format.italic).toBe(true);
    expect(second[0]!.format).toMatchObject({ fontSize: 11 });
    expect(second[0]!.format.bold).toBeUndefined();
  });

  it('mantiene le pagine e le dimensioni del foglio', async () => {
    const blocks: Block[] = Array.from({ length: 70 }, (_, i) => paragraph(`Riga ${i + 1}`));
    const source = createDocument(blocks);
    const pages = engine.layout(source).pages.length;
    expect(pages).toBeGreaterThan(1);
    const doc = await roundTrip(source);
    const ps = paragraphs(doc);
    expect(ps.map(textOf)).toEqual(blocks.map((_, i) => `Riga ${i + 1}`));
    expect(ps.filter((p) => p.format.pageBreakBefore)).toHaveLength(pages - 1);
    const page = doc.sections[0]!.page;
    expect(page.width).toBeCloseTo(595.3, 0);
    expect(page.height).toBeCloseTo(841.9, 0);
    expect(page.margins.left).toBeCloseTo(source.sections[0]!.page.margins.left, 0);
  });

  it('riconosce il testo centrato', async () => {
    const doc = await roundTrip(createDocument([paragraph('In mezzo', { align: 'center' }), paragraph('A sinistra')]));
    const [a, b] = paragraphs(doc);
    expect(a!.format.align).toBe('center');
    expect(b!.format.align).toBe('left');
  });

  it('estrae le immagini come PNG', async () => {
    const px = new Uint8Array(4 * 4 * 3).fill(200);
    const png = await encodePng(4, 4, 2, px);
    const src = toDataUri(png!, 'image/png');
    const doc = await roundTrip(
      createDocument([paragraph('Prima'), paragraph([{ type: 'image', src, width: 100, height: 50 }]), paragraph('Dopo')]),
    );
    const ps = paragraphs(doc);
    expect(ps.map((p) => p.children[0]?.type)).toEqual(['text', 'image', 'text']);
    const img = ps[1]!.children[0]!;
    expect(img).toMatchObject({ type: 'image', width: 100, height: 50 });
    const key = img.type === 'image' ? img.src : '';
    expect(doc.media[key]).toMatch(/^data:image\/png;base64,/);
  });

  it('dà un errore chiaro per un file non valido', async () => {
    await expect(importPdf(new TextEncoder().encode('non sono un pdf'))).rejects.toThrow(/non è un PDF valido/);
  });

  it('ricava la famiglia dal nome del carattere', () => {
    expect(fontFamilyFromPdfName('ABCDEF+TimesNewRomanPS-BoldMT')).toBe('Times New Roman');
    expect(fontFamilyFromPdfName('Helvetica-Oblique')).toBe('Arial');
    expect(fontFamilyFromPdfName('SourceSansPro-Regular')).toBe('Source Sans Pro');
  });

  it('file.open apre un .pdf con il modulo pdf attivo', async () => {
    const editor = new Editor();
    editor.use(createIoModule());
    editor.use(createPdfModule({ rasterize: async () => null }));
    await Promise.resolve();
    const data = await layoutToPdf(engine.layout(createDocument([paragraph('Ciao dal PDF')])), { measurer });
    expect(editor.services.formats?.byExtension('relazione.pdf')?.import).toBeTypeOf('function');
    expect(await editor.executeAsync('file.open', { data, formatId: 'pdf' })).toBe(true);
    expect(paragraphs(editor.state.doc).map(textOf)).toEqual(['Ciao dal PDF']);
  });
});
