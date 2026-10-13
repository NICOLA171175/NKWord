import { describe, expect, it } from 'vitest';
import { createDocument, Editor, paragraph, textRun, type Block, type DocumentModel } from '../../src/core';
import { createIoModule, type SaveRequest } from '../../src/io';
import { createApproxMeasurer, createLayoutEngine, createLayoutModule } from '../../src/layout';
import { createPdfFormat, createPdfModule, encodeWinAnsi, familyFor, layoutToPdf, loadImage, standardFont } from '../../src/pdf';
import { RIBBON_ITEMS } from '../../src/ui/ribbon-items';

const measurer = createApproxMeasurer();
const engine = createLayoutEngine(measurer);
const now = () => new Date(Date.UTC(2026, 9, 13, 8, 0, 0));
const latin1 = (b: Uint8Array) => Array.from(b, (c) => String.fromCharCode(c)).join('');
const num = (n: number) => String(Math.round(n * 1000) / 1000);
const hex = (s: string) => Array.from(encodeWinAnsi(s), (b) => b.toString(16).padStart(2, '0')).join('');

function plain(text: string): Block {
  return paragraph(text, { spaceAfter: 0, lineSpacing: { rule: 'multiple', value: 1 } });
}

async function pdfOf(doc: DocumentModel): Promise<string> {
  return latin1(await layoutToPdf(engine.layout(doc), { measurer, now }));
}

/** Ogni voce della tabella xref punta davvero all'inizio del suo oggetto. */
function checkXref(pdf: string): number {
  const start = Number(/startxref\n(\d+)/.exec(pdf)![1]);
  expect(pdf.slice(start, start + 4)).toBe('xref');
  const [, first, count] = /xref\n(\d+) (\d+)\n/.exec(pdf.slice(start))!;
  const entries = pdf.slice(start).split('\n').slice(2, 2 + Number(count));
  entries.slice(1).forEach((e, i) => {
    const offset = Number(e.slice(0, 10));
    expect(pdf.slice(offset).startsWith(`${Number(first) + i + 1} 0 obj`)).toBe(true);
  });
  return Number(count) - 1;
}

describe('caratteri standard', () => {
  it('ha le larghezze per tutti i caratteri ASCII e codifica in WinAnsi', () => {
    for (const f of [{}, { bold: true }, { fontFamily: 'Times New Roman' }, { fontFamily: 'Georgia', bold: true }]) {
      const font = standardFont(f);
      expect(font.widths).toHaveLength(224);
      expect(font.widths.every((w) => w > 0)).toBe(true);
    }
    expect(standardFont({ fontFamily: 'Calibri', bold: true, italic: true }).name).toBe('Helvetica-BoldOblique');
    expect(standardFont({ fontFamily: 'Times New Roman', italic: true }).name).toBe('Times-Italic');
    expect(familyFor('Courier New')).toBe('Courier');
    expect(Array.from(encodeWinAnsi('aè€✓'))).toEqual([0x61, 0xe8, 0x80, 0x3f]);
    // à ha la larghezza di a.
    const h = standardFont({});
    expect(h.widths[0xe0 - 32]).toBe(h.widths[0x61 - 32]);
  });
});

describe('layoutToPdf', () => {
  it('scrive un PDF valido con una pagina per ogni pagina del layout', async () => {
    const blocks = Array.from({ length: 80 }, (_, i) => plain(`Paragrafo numero ${i + 1}`));
    const doc = createDocument(blocks);
    const layout = engine.layout(doc);
    expect(layout.pages.length).toBeGreaterThan(1);
    const pdf = await pdfOf(doc);
    expect(pdf.startsWith('%PDF-1.4')).toBe(true);
    expect(pdf.trimEnd().endsWith('%%EOF')).toBe(true);
    checkXref(pdf);
    expect(pdf.match(/\/Type \/Page\b/g)).toHaveLength(layout.pages.length);
    expect(pdf).toContain(`/Count ${layout.pages.length}`);
    const page = layout.pages[0]!;
    expect(pdf).toContain(`/MediaBox [0 0 ${num(page.width)} ${num(page.height)}]`);
    expect(pdf).toContain(`<${hex('Paragrafo numero 80')}> Tj`);
    expect(pdf).toContain('/BaseFont /Helvetica /Encoding /WinAnsiEncoding');
  });

  it('mette ogni pezzo di testo dove lo mette il layout, alla stessa larghezza', async () => {
    const doc = createDocument([plain('Ciao mondo')]);
    const line = engine.layout(doc).pages[0]!.lines[0]!;
    const run = line.runs[0]!;
    const page = engine.layout(doc).pages[0]!;
    const pdf = await pdfOf(doc);
    const m = /([\d.]+) Tz 0 0 0 rg 1 0 0 1 ([\d.]+) ([\d.]+) Tm/.exec(pdf)!;
    expect(Number(m[2])).toBeCloseTo(line.x + run.x, 2);
    expect(Number(m[3])).toBeCloseTo(page.height - (line.y + line.baseline), 2);
    // La scala orizzontale riporta la larghezza Helvetica a quella misurata.
    const font = standardFont({});
    const helvetica = Array.from(encodeWinAnsi('Ciao mondo')).reduce((n, b) => n + font.widths[b - 32]!, 0) * 11 / 1000;
    expect((helvetica * Number(m[1])) / 100).toBeCloseTo(measurer.measure('Ciao mondo', {}).width, 1);
  });

  it('disegna colori, evidenziazione, sottolineatura e collegamenti', async () => {
    const p = paragraph('', { spaceAfter: 0 });
    const doc = createDocument([
      {
        ...p,
        children: [
          textRun('rosso ', { color: '#ff0000' }),
          textRun('giallo ', { highlight: 'yellow' }),
          textRun('sotto ', { underline: 'single', bold: true }),
          textRun('sito', { link: 'https://example.com/a b' }),
        ],
      },
    ]);
    const pdf = await pdfOf(doc);
    expect(pdf).toContain('1 0 0 rg 1 0 0 1');
    expect(pdf).toMatch(/1 1 0 rg [\d.]+ [\d.]+ [\d.]+ [\d.]+ re f/);
    expect(pdf).toMatch(/q 0 0 0 RG [\d.]+ w \[\] 0 d [\d.]+ [\d.]+ m [\d.]+ [\d.]+ l S Q/);
    expect(pdf).toContain('/BaseFont /Helvetica-Bold');
    expect(pdf).toContain('/Subtype /Link');
    expect(pdf).toContain('/URI (https://example.com/a%20b)');
    checkXref(pdf);
  });

  it('non rende cliccabili i collegamenti non sicuri', async () => {
    const p = paragraph('', { spaceAfter: 0 });
    const doc = createDocument([{ ...p, children: [textRun('x', { link: 'javascript:alert(1)' })] }]);
    expect(await pdfOf(doc)).not.toContain('/Subtype /Link');
  });

  it('scrive titolo e autore nei metadati', async () => {
    const doc: DocumentModel = { ...createDocument([plain('x')]), metadata: { title: 'Relazione è pronta', author: 'Nicola' } };
    const pdf = await pdfOf(doc);
    expect(pdf).toContain('/Author (Nicola)');
    expect(pdf).toContain('/Title <FEFF');
    expect(pdf).toContain('/CreationDate (D:20261013080000Z)');
  });

  it('disegna intestazioni, piè di pagina e numeri di pagina', async () => {
    const d = createDocument(Array.from({ length: 80 }, (_, i) => plain(`Riga ${i}`)));
    const s = d.sections[0]!;
    const footer = paragraph('', { spaceAfter: 0 });
    const doc: DocumentModel = {
      ...d,
      sections: [{ ...s, footer: { default: [{ ...footer, children: [textRun('Pagina '), { type: 'field', kind: 'page', format: {} }] }] } }],
    };
    const layout = engine.layout(doc);
    const pdf = await pdfOf(doc);
    expect(pdf).toContain(`<${hex('Pagina')}> Tj`);
    for (const p of layout.pages) expect(pdf).toContain(`<${hex(String(p.number))}> Tj`);
  });
});

async function deflate(data: Uint8Array): Promise<Uint8Array> {
  const out = new Response(new Blob([data as BlobPart]).stream().pipeThrough(new CompressionStream('deflate')));
  return new Uint8Array(await out.arrayBuffer());
}

function chunk(type: string, data: Uint8Array): number[] {
  const len = data.length;
  // Il CRC non viene controllato dal lettore: basta un valore qualsiasi.
  return [len >>> 24, (len >>> 16) & 255, (len >>> 8) & 255, len & 255, ...Array.from(type, (c) => c.charCodeAt(0)), ...data, 0, 0, 0, 0];
}

/** PNG 2×2 non compresso nei filtri (filtro 0 su ogni riga). */
async function png(colorType: 2 | 6, pixels: number[]): Promise<string> {
  const bpp = colorType === 6 ? 4 : 3;
  const raw: number[] = [];
  for (let y = 0; y < 2; y++) raw.push(0, ...pixels.slice(y * 2 * bpp, (y + 1) * 2 * bpp));
  const ihdr = new Uint8Array([0, 0, 0, 2, 0, 0, 0, 2, 8, colorType, 0, 0, 0]);
  const bytes = new Uint8Array([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a,
    ...chunk('IHDR', ihdr),
    ...chunk('IDAT', await deflate(new Uint8Array(raw))),
    ...chunk('IEND', new Uint8Array()),
  ]);
  return `data:image/png;base64,${btoa(latin1(bytes))}`;
}

describe('immagini', () => {
  it('incorpora un PNG RGB così com\'è, con i predittori', async () => {
    const img = (await loadImage(await png(2, [255, 0, 0, 0, 255, 0, 0, 0, 255, 255, 255, 255])))!;
    expect(img.width).toBe(2);
    expect(img.dict).toContain('/Predictor 15 /Colors 3 /BitsPerComponent 8 /Columns 2');
    expect(img.alpha).toBeUndefined();
  });

  it('separa il canale alfa di un PNG RGBA in una maschera', async () => {
    const img = (await loadImage(await png(6, [255, 0, 0, 255, 0, 255, 0, 128, 0, 0, 255, 0, 9, 9, 9, 255])))!;
    expect(img.alpha).toBeDefined();
    const inflate = async (d: Uint8Array) =>
      new Uint8Array(await new Response(new Blob([d as BlobPart]).stream().pipeThrough(new DecompressionStream('deflate'))).arrayBuffer());
    expect(Array.from(await inflate(img.alpha!))).toEqual([255, 128, 0, 255]);
    expect(Array.from(await inflate(img.data))).toEqual([255, 0, 0, 0, 255, 0, 0, 0, 255, 9, 9, 9]);
  });

  it('legge le dimensioni di un JPEG', async () => {
    const jpeg = new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 0, 4, 0, 0, 0xff, 0xc0, 0, 17, 8, 0, 16, 0, 32, 3, 1, 0x22, 0, 2, 0x11, 1, 3, 0x11, 1, 0xff, 0xd9]);
    const img = (await loadImage(`data:image/jpeg;base64,${btoa(latin1(jpeg))}`))!;
    expect([img.width, img.height]).toEqual([32, 16]);
    expect(img.dict).toContain('/DCTDecode');
  });

  it('disegna l\'immagine nel PDF e un riquadro per i formati sconosciuti', async () => {
    const src = await png(2, [0, 0, 0, 255, 255, 255, 0, 0, 0, 255, 255, 255]);
    const p = paragraph('', { spaceAfter: 0 });
    const doc = createDocument([
      { ...p, children: [{ type: 'image', src: 'img1', width: 40, height: 30 }, textRun(' '), { type: 'image', src: 'data:image/gif;base64,R0lG', width: 20, height: 20 }] },
    ]);
    const pdf = await pdfOf({ ...doc, media: { img1: src } });
    expect(pdf).toContain('/Subtype /Image /Width 2 /Height 2');
    expect(pdf).toMatch(/q 40 0 0 30 [\d.]+ [\d.]+ cm \/Im1 Do Q/);
    expect(pdf).toMatch(/0\.6 0\.6 0\.6 RG 0\.5 w [\d.]+ [\d.]+ 20 20 re S/);
    checkXref(pdf);
  });
});

describe('modulo pdf', () => {
  it('aggiunge il formato PDF e File > Esporta > PDF salva il file', async () => {
    const saved: SaveRequest[] = [];
    const editor = new Editor();
    editor.use(createLayoutModule({ measurer }));
    editor.use(createIoModule({ save: (f) => void saved.push(f) }));
    const dispose = editor.use(createPdfModule({ measurer }));
    expect(editor.services.formats!.get('pdf')?.mimeType).toBe('application/pdf');
    expect(editor.services.formats!.byExtension('relazione.pdf')?.id).toBe('pdf');

    expect(editor.canExecute('file.export', { formatId: 'pdf' })).toBe(true);
    // Esportare non cambia il documento: il comando non restituisce transazioni.
    await editor.executeAsync('file.export', { formatId: 'pdf' });
    expect(saved).toHaveLength(1);
    expect(saved[0]!.fileName).toBe('Documento.pdf');
    expect(saved[0]!.mimeType).toBe('application/pdf');
    expect(latin1(saved[0]!.data as Uint8Array).startsWith('%PDF-')).toBe(true);

    dispose();
    expect(editor.services.formats!.get('pdf')).toBeUndefined();
  });

  it('funziona anche se il modulo io si attiva dopo', async () => {
    const editor = new Editor();
    editor.use(createPdfModule({ measurer }));
    editor.use(createIoModule({ save: () => {} }));
    await Promise.resolve();
    expect(editor.services.formats!.get('pdf')).toBeDefined();
  });

  it('senza modulo layout impagina da sé', async () => {
    const data = await createPdfFormat({ measurer, now }).export!(createDocument([plain('Solo')]));
    expect(latin1(data as Uint8Array)).toContain(`<${hex('Solo')}> Tj`);
  });

  it('compare nel menu File, gruppo Esporta', () => {
    expect(RIBBON_ITEMS.some((i) => i.tab === 'file' && i.commandId === 'file.export' && (i.args as { formatId: string }).formatId === 'pdf')).toBe(true);
  });
});
