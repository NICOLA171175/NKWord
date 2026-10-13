/**
 * Apertura di un PDF come documento modificabile, come fa Word: il PDF non
 * descrive paragrafi ma solo pezzi di testo in posizioni fisse, quindi il
 * risultato è una ricostruzione approssimata.
 *
 * - Il testo si legge con pdf.js (pdfjs-dist), caricato solo quando serve.
 * - I pezzi sulla stessa linea di base formano una riga; le righe si uniscono
 *   in paragrafi se sono vicine, hanno lo stesso corpo e la precedente arriva
 *   fino al margine destro (altrimenti era l'ultima riga di un paragrafo).
 * - Corpo, famiglia, grassetto e corsivo si deducono dal carattere; il
 *   colore del testo non si recupera.
 * - Le immagini diventano PNG in un paragrafo proprio, al loro posto.
 * - Ogni pagina del PDF ricomincia su una pagina nuova; dimensioni e margini
 *   della prima pagina diventano quelli del documento.
 *
 * I PDF scansionati (solo immagini) si aprono come immagini: il testo non
 * viene riconosciuto.
 */

import {
  createDocument,
  defaultPageSetup,
  paragraph,
  section,
  textRun,
  type Alignment,
  type Block,
  type CharFormat,
  type DocumentModel,
  type FileData,
  type Inline,
  type PageSetup,
} from '../core';
import { dataToBytes, toDataUri } from '../io/util/encoding';
import { crc32 } from '../io/util/zip';

// ---------------------------------------------------------------------------
// Tipi minimi di pdf.js usati qui (la libreria si carica in modo dinamico).

type Matrix = readonly [number, number, number, number, number, number];

interface PdfTextItem {
  readonly str: string;
  readonly transform: number[];
  readonly width: number;
  readonly height: number;
  readonly fontName: string;
  readonly hasEOL: boolean;
}

interface PdfImageData {
  readonly width: number;
  readonly height: number;
  readonly kind?: number;
  readonly data?: Uint8Array | Uint8ClampedArray;
}

interface PdfObjects {
  get(id: string, callback: (data: unknown) => void): void;
  has(id: string): boolean;
}

interface PdfPage {
  readonly view: number[];
  readonly rotate: number;
  readonly objs: PdfObjects;
  readonly commonObjs: PdfObjects;
  getTextContent(): Promise<{ items: readonly (PdfTextItem | { type: string })[] }>;
  getOperatorList(): Promise<{ fnArray: number[]; argsArray: unknown[][] }>;
  cleanup(): void;
}

interface PdfDocument {
  readonly numPages: number;
  getPage(n: number): Promise<PdfPage>;
  getMetadata(): Promise<{ info?: Record<string, unknown> }>;
  destroy(): Promise<void>;
}

interface PdfJs {
  readonly OPS: Record<string, number>;
  readonly GlobalWorkerOptions: { workerSrc: string };
  getDocument(params: Record<string, unknown>): { promise: Promise<PdfDocument> };
}

let pdfjsPromise: Promise<PdfJs> | undefined;

/** pdf.js nel browser usa un worker; in Node (test) il worker finto interno. */
function loadPdfjs(): Promise<PdfJs> {
  pdfjsPromise ??= (async () => {
    const lib = (await import('pdfjs-dist/legacy/build/pdf.mjs')) as unknown as PdfJs;
    const isNode = typeof process !== 'undefined' && !!process.versions?.node && typeof window === 'undefined';
    if (!isNode && !lib.GlobalWorkerOptions.workerSrc) {
      const worker = await import('pdfjs-dist/legacy/build/pdf.worker.min.mjs?url');
      lib.GlobalWorkerOptions.workerSrc = worker.default;
    }
    return lib;
  })();
  pdfjsPromise.catch(() => (pdfjsPromise = undefined));
  return pdfjsPromise;
}

// ---------------------------------------------------------------------------
// Pezzi di testo, righe e immagini di una pagina

interface Piece {
  readonly text: string;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly size: number;
  readonly format: CharFormat;
}

interface Line {
  readonly pieces: Piece[];
  readonly x: number;
  right: number;
  readonly y: number;
  size: number;
}

interface PageImage {
  readonly src: string;
  readonly x: number;
  /** Bordo superiore, nelle coordinate del PDF (y verso l'alto). */
  readonly top: number;
  readonly width: number;
  readonly height: number;
}

const multiply = (m: Matrix, n: Matrix): Matrix => [
  m[0] * n[0] + m[2] * n[1],
  m[1] * n[0] + m[3] * n[1],
  m[0] * n[2] + m[2] * n[3],
  m[1] * n[2] + m[3] * n[3],
  m[0] * n[4] + m[2] * n[5] + m[4],
  m[1] * n[4] + m[3] * n[5] + m[5],
];

const round = (n: number, step = 0.5) => Math.round(n / step) * step;

const objectData = (objs: PdfObjects, id: string) =>
  new Promise<unknown>((resolve) => {
    try {
      objs.get(id, resolve);
    } catch {
      resolve(undefined);
    }
  });

/** Nome PostScript senza il prefisso del sottoinsieme (`ABCDEF+Arial-Bold` → `Arial-Bold`). */
const baseFontName = (name: string) => name.replace(/^[A-Z]{6}\+/, '');

const FAMILY_ALIASES: Record<string, string> = {
  helvetica: 'Arial',
  arial: 'Arial',
  arialmt: 'Arial',
  times: 'Times New Roman',
  timesroman: 'Times New Roman',
  timesnewroman: 'Times New Roman',
  timesnewromanps: 'Times New Roman',
  courier: 'Courier New',
  couriernew: 'Courier New',
  couriernewps: 'Courier New',
  dejavusans: 'DejaVu Sans',
  dejavuserif: 'DejaVu Serif',
};

/** Famiglia leggibile dal nome PostScript, es. `TimesNewRomanPS-BoldMT` → `Times New Roman`. */
export function fontFamilyFromPdfName(name: string): string | undefined {
  const family = baseFontName(name)
    .split(/[-,]/)[0]!
    .replace(/(PSMT|MT|PS)$/, '')
    .replace(/(Bold|Italic|Oblique|Regular|Light|Medium|Black|Semibold)+$/i, '');
  if (!family) return undefined;
  const alias = FAMILY_ALIASES[family.toLowerCase()];
  if (alias) return alias;
  // CamelCase → parole: 'SourceSansPro' → 'Source Sans Pro'.
  return family.replace(/([a-z])([A-Z])/g, '$1 $2').replace(/([A-Z]+)([A-Z][a-z])/g, '$1 $2');
}

interface FontInfo {
  readonly name?: string;
  readonly bold?: boolean;
  readonly black?: boolean;
  readonly italic?: boolean;
}

function charFormatFor(font: FontInfo | undefined, size: number): CharFormat {
  const name = font?.name ? baseFontName(font.name) : '';
  const style = name.split(/[-,]/).slice(1).join('-') || name;
  const bold = !!font?.bold || !!font?.black || /bold|black|heavy|semibold|demi/i.test(style);
  const italic = !!font?.italic || /italic|oblique|kursiv/i.test(style);
  const fontFamily = name ? fontFamilyFromPdfName(name) : undefined;
  return {
    ...(fontFamily ? { fontFamily } : {}),
    fontSize: round(size),
    ...(bold ? { bold: true } : {}),
    ...(italic ? { italic: true } : {}),
  };
}

/** Immagini della pagina con la matrice corrente al momento del disegno. */
async function scanOperators(page: PdfPage, OPS: Record<string, number>) {
  const ops = await page.getOperatorList();
  const images: { id: string | PdfImageData; ctm: Matrix }[] = [];
  let ctm: Matrix = [1, 0, 0, 1, 0, 0];
  const stack: Matrix[] = [];
  for (let i = 0; i < ops.fnArray.length; i++) {
    const fn = ops.fnArray[i];
    const args = ops.argsArray[i] ?? [];
    switch (fn) {
      case OPS.save:
        stack.push(ctm);
        break;
      case OPS.restore:
        ctm = stack.pop() ?? ctm;
        break;
      case OPS.transform:
        ctm = multiply(ctm, args as unknown as Matrix);
        break;
      case OPS.paintFormXObjectBegin:
        stack.push(ctm);
        if (Array.isArray(args[0])) ctm = multiply(ctm, args[0] as unknown as Matrix);
        break;
      case OPS.paintFormXObjectEnd:
        ctm = stack.pop() ?? ctm;
        break;
      case OPS.paintImageXObject:
        if (typeof args[0] === 'string') images.push({ id: args[0], ctm });
        break;
      case OPS.paintInlineImageXObject:
        if (args[0] && typeof args[0] === 'object') images.push({ id: args[0] as PdfImageData, ctm });
        break;
    }
  }
  return images;
}

async function pageImages(page: PdfPage, found: { id: string | PdfImageData; ctm: Matrix }[]): Promise<PageImage[]> {
  const out: PageImage[] = [];
  const cache = new Map<string, string | null>();
  for (const { id, ctm } of found) {
    const width = Math.hypot(ctm[0], ctm[1]);
    const height = Math.hypot(ctm[2], ctm[3]);
    // Puntini, linee e simili non sono immagini da mostrare.
    if (width < 8 || height < 8) continue;
    let src: string | null | undefined;
    if (typeof id === 'string') {
      src = cache.get(id);
      if (src === undefined) {
        const objs = id.startsWith('g_') ? page.commonObjs : page.objs;
        src = await imageToPng((await objectData(objs, id)) as PdfImageData | undefined);
        cache.set(id, src);
      }
    } else {
      src = await imageToPng(id);
    }
    if (!src) continue;
    const xs = [ctm[4], ctm[4] + ctm[0], ctm[4] + ctm[2], ctm[4] + ctm[0] + ctm[2]];
    const ys = [ctm[5], ctm[5] + ctm[1], ctm[5] + ctm[3], ctm[5] + ctm[1] + ctm[3]];
    out.push({ src, x: Math.min(...xs), top: Math.max(...ys), width, height });
  }
  return out;
}

/** Testo della pagina in righe, nell'ordine del flusso del PDF. */
async function pageLines(page: PdfPage): Promise<Line[]> {
  const content = await page.getTextContent();
  const fonts = new Map<string, FontInfo | undefined>();
  const lines: Line[] = [];
  let line: Line | undefined;
  for (const raw of content.items) {
    if (!('str' in raw)) continue;
    const item = raw as PdfTextItem;
    const [a, b, c, d, x, y] = item.transform as [number, number, number, number, number, number];
    const size = Math.hypot(c, d) || Math.hypot(a, b);
    // Testo ruotato (scritte verticali a margine): si ignora.
    if (Math.abs(b) > 0.01 * Math.abs(a) + 0.01) continue;
    if (!fonts.has(item.fontName)) {
      fonts.set(item.fontName, (await objectData(page.commonObjs, item.fontName)) as FontInfo | undefined);
    }
    const text = item.str.replace(/[\u0000-\u0008\u000b-\u001f­]/g, '').replace(/[\t\n\r]/g, ' ')
      // Punti elenco dei caratteri simbolo (area privata Unicode).
      .replace(/[\ue000-\uf8ff]/g, '•');
    const newLine = !line || Math.abs(y - line.y) > Math.max(1, size * 0.4) || x < line.right - size * 2;
    if (newLine && text.trim()) {
      line = { pieces: [], x, right: x, y, size };
      lines.push(line);
    }
    if (!line || !text) continue;
    line.pieces.push({ text, x, y, width: item.width, size, format: charFormatFor(fonts.get(item.fontName), size) });
    if (text.trim()) {
      line.right = Math.max(line.right, x + item.width);
      line.size = Math.max(line.size, size);
    }
  }
  return lines;
}

// ---------------------------------------------------------------------------
// Righe → paragrafi

/** Inline della riga, con gli spazi dove i pezzi sono distanti. */
function lineInlines(line: Line): Inline[] {
  const out: Inline[] = [];
  let end: number | undefined;
  let lastText = '';
  for (const p of line.pieces) {
    let text = p.text;
    if (end !== undefined && p.x - end > p.size * 0.15 && !/\s$/.test(lastText) && !/^\s/.test(text)) text = ' ' + text;
    if (!text) continue;
    out.push(textRun(text, p.format));
    lastText = text;
    end = p.x + p.width;
  }
  return out;
}

function trimInlines(inlines: Inline[]): Inline[] {
  const out = inlines.map((i) => (i.type === 'text' ? { ...i, text: i.text.replace(/\s+/g, ' ') } : i));
  const first = out[0];
  if (first?.type === 'text') out[0] = { ...first, text: first.text.trimStart() };
  const last = out[out.length - 1];
  if (last?.type === 'text') out[out.length - 1] = { ...last, text: last.text.trimEnd() };
  return out.filter((i) => i.type !== 'text' || i.text);
}

const sameFormat = (a: CharFormat, b: CharFormat) => JSON.stringify(a) === JSON.stringify(b);

/** Unisce i pezzi adiacenti con la stessa formattazione. */
function mergeRuns(inlines: Inline[]): Inline[] {
  const out: Inline[] = [];
  for (const i of inlines) {
    const prev = out[out.length - 1];
    if (i.type === 'text' && prev?.type === 'text' && sameFormat(prev.format, i.format)) out[out.length - 1] = { ...prev, text: prev.text + i.text };
    else out.push(i);
  }
  return out;
}

interface Box {
  readonly left: number;
  readonly right: number;
}

interface ParaDraft {
  readonly lines: Line[];
  readonly top: number;
}

/** Raggruppa le righe in paragrafi. `box` è l'area del testo della pagina. */
function groupLines(lines: readonly Line[], box: Box): ParaDraft[] {
  const width = box.right - box.left;
  const paras: ParaDraft[] = [];
  let current: Line[] = [];
  const flush = () => {
    if (current.length) paras.push({ lines: current, top: current[0]!.y + current[0]!.size });
    current = [];
  };
  for (const line of lines) {
    const prev = current[current.length - 1];
    if (prev) {
      const gap = prev.y - line.y;
      const sameSize = Math.abs(prev.size - line.size) <= Math.max(0.6, prev.size * 0.1);
      const close = gap > 0 && gap < Math.max(prev.size, line.size) * 1.6;
      const prevFull = prev.right >= box.right - Math.max(prev.size * 4, width * 0.12);
      const indented = line.x > box.left + line.size && line.x > prev.x + line.size * 0.5;
      if (!(sameSize && close && prevFull && !indented)) flush();
    }
    current.push(line);
  }
  flush();
  return paras;
}

function alignmentOf(draft: ParaDraft, box: Box): Alignment | undefined {
  const { lines } = draft;
  const centered = lines.every((l) => {
    const left = l.x - box.left;
    const right = box.right - l.right;
    return left > l.size * 2 && Math.abs(left - right) < Math.max(4, l.size * 0.8);
  });
  if (centered) return 'center';
  const flushRight = lines.every((l) => Math.abs(box.right - l.right) < 3 && l.x - box.left > l.size * 4);
  if (flushRight) return 'right';
  // Giustificato: tutte le righe tranne l'ultima arrivano al margine destro.
  if (lines.length >= 3 && lines.slice(0, -1).every((l) => Math.abs(box.right - l.right) < 2)) return 'justify';
  return undefined;
}

function draftToParagraph(draft: ParaDraft, box: Box, pageBreakBefore: boolean, spaceAfter: number): Block {
  const inlines: Inline[] = [];
  draft.lines.forEach((line, i) => {
    const runs = lineInlines(line);
    if (i > 0 && runs.length) {
      const prev = inlines[inlines.length - 1];
      // "parola-" a fine riga seguita da minuscola: sillabazione, si ricompone.
      const first = runs[0];
      if (prev?.type === 'text' && /\p{L}-\s*$/u.test(prev.text) && first?.type === 'text' && /^\s*\p{Ll}/u.test(first.text)) {
        inlines[inlines.length - 1] = { ...prev, text: prev.text.replace(/-\s*$/, '') };
      } else if (prev?.type === 'text' && !/\s$/.test(prev.text)) {
        inlines[inlines.length - 1] = { ...prev, text: prev.text + ' ' };
      }
    }
    inlines.push(...runs);
  });
  const first = draft.lines[0]!;
  const second = draft.lines[1];
  const align = alignmentOf(draft, box);
  const left = round((second ?? first).x - box.left, 1);
  const firstLine = second ? round(first.x - second.x, 1) : 0;
  const lineGaps = draft.lines.slice(1).map((l, i) => draft.lines[i]!.y - l.y);
  const avgGap = lineGaps.length ? lineGaps.reduce((s, g) => s + g, 0) / lineGaps.length : 0;
  const size = Math.max(...draft.lines.map((l) => l.size));
  const multiple = avgGap ? round(avgGap / (size * 1.15), 0.05) : 1;
  return paragraph(mergeRuns(trimInlines(inlines)), {
    ...(align ? { align } : { align: 'left' }),
    ...(align !== 'center' && align !== 'right' && left > 2 ? { indentLeft: left } : {}),
    ...(align !== 'center' && align !== 'right' && Math.abs(firstLine) > 2 ? { indentFirstLine: firstLine } : {}),
    spaceBefore: 0,
    spaceAfter: round(Math.max(0, spaceAfter), 1),
    lineSpacing: { rule: 'multiple', value: Math.min(3, Math.max(1, multiple)) },
    ...(pageBreakBefore ? { pageBreakBefore: true } : {}),
  });
}

function imageParagraph(image: PageImage, media: Record<string, string>, box: Box, pageBreakBefore: boolean): Block {
  const key = `pdf-image-${Object.keys(media).length + 1}.png`;
  media[key] = image.src;
  const maxWidth = Math.max(box.right - box.left, 72);
  const scale = Math.min(1, maxWidth / image.width);
  const offset = image.x - box.left;
  const center = Math.abs(offset - (box.right - (image.x + image.width))) < 12 && offset > 12;
  return paragraph(
    [{ type: 'image', src: key, width: round(image.width * scale, 0.1), height: round(image.height * scale, 0.1) }],
    {
      align: center ? 'center' : 'left',
      ...(!center && offset > 2 && image.width + offset <= maxWidth ? { indentLeft: round(offset, 1) } : {}),
      spaceBefore: 0,
      spaceAfter: 6,
      lineSpacing: { rule: 'multiple', value: 1 },
      ...(pageBreakBefore ? { pageBreakBefore: true } : {}),
    },
  );
}

// ---------------------------------------------------------------------------
// Immagini → PNG

/** Comprime con zlib (CompressionStream 'deflate'), disponibile in browser e Node ≥ 18. */
async function zlib(data: Uint8Array): Promise<Uint8Array> {
  const stream = new Blob([data as BlobPart]).stream().pipeThrough(new CompressionStream('deflate'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

function pngChunk(type: string, data: Uint8Array): Uint8Array {
  const out = new Uint8Array(12 + data.length);
  const view = new DataView(out.buffer);
  view.setUint32(0, data.length);
  for (let i = 0; i < 4; i++) out[4 + i] = type.charCodeAt(i);
  out.set(data, 8);
  view.setUint32(8 + data.length, crc32(out.subarray(4, 8 + data.length)));
  return out;
}

/** Codifica in PNG i pixel decodificati da pdf.js (1 bit, RGB o RGBA). */
export async function encodePng(width: number, height: number, kind: number, data: Uint8Array | Uint8ClampedArray): Promise<Uint8Array | null> {
  // kind: 1 = scala di grigi a 1 bit, 2 = RGB 8 bit, 3 = RGBA 8 bit.
  const channels = kind === 3 ? 4 : kind === 2 ? 3 : 1;
  const rowBytes = width * channels;
  if (!width || !height || (kind === 1 ? data.length < ((width + 7) >> 3) * height : data.length < rowBytes * height)) return null;
  const raw = new Uint8Array((rowBytes + 1) * height);
  for (let y = 0; y < height; y++) {
    const dst = y * (rowBytes + 1) + 1;
    if (kind === 1) {
      const src = y * ((width + 7) >> 3);
      for (let x = 0; x < width; x++) raw[dst + x] = (data[src + (x >> 3)]! >> (7 - (x & 7))) & 1 ? 255 : 0;
    } else {
      raw.set(data.subarray(y * rowBytes, (y + 1) * rowBytes), dst);
    }
  }
  const header = new Uint8Array(13);
  const hv = new DataView(header.buffer);
  hv.setUint32(0, width);
  hv.setUint32(4, height);
  header[8] = 8;
  header[9] = channels === 4 ? 6 : channels === 3 ? 2 : 0;
  const parts = [
    new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    pngChunk('IHDR', header),
    pngChunk('IDAT', await zlib(raw)),
    pngChunk('IEND', new Uint8Array()),
  ];
  const out = new Uint8Array(parts.reduce((s, p) => s + p.length, 0));
  let o = 0;
  for (const p of parts) {
    out.set(p, o);
    o += p.length;
  }
  return out;
}

async function imageToPng(img: PdfImageData | undefined): Promise<string | null> {
  if (!img?.data || !img.width || !img.height) return null;
  const png = await encodePng(img.width, img.height, img.kind ?? 2, img.data);
  return png ? toDataUri(png, 'image/png') : null;
}

// ---------------------------------------------------------------------------
// Documento

function pageSetupFrom(view: readonly number[], rotate: number, box: { left: number; right: number; top: number; bottom: number } | undefined): PageSetup {
  const base = defaultPageSetup();
  let width = (view[2] ?? 0) - (view[0] ?? 0);
  let height = (view[3] ?? 0) - (view[1] ?? 0);
  if (rotate % 180 !== 0) [width, height] = [height, width];
  if (!(width > 72 && height > 72)) return base;
  const clamp = (n: number) => round(Math.min(144, Math.max(18, n)), 1);
  const margins = box
    ? {
        ...base.margins,
        left: clamp(box.left),
        right: clamp(width - box.right),
        top: clamp(height - box.top),
        bottom: clamp(box.bottom),
      }
    : base.margins;
  return { ...base, width: round(width, 0.01), height: round(height, 0.01), orientation: width > height ? 'landscape' : 'portrait', margins };
}

export interface PdfImportOptions {
  /** Al più quante pagine leggere (default: tutte). */
  readonly maxPages?: number;
}

/** Converte un PDF in un documento modificabile. */
export async function importPdf(data: FileData, options: PdfImportOptions = {}): Promise<DocumentModel> {
  const bytes = dataToBytes(data);
  const pdfjs = await loadPdfjs();
  let pdf: PdfDocument;
  try {
    pdf = await pdfjs.getDocument({
      // pdf.js si prende il buffer: gli si passa una copia.
      data: new Uint8Array(bytes),
      isEvalSupported: false,
      isOffscreenCanvasSupported: false,
      disableFontFace: true,
      useSystemFonts: false,
      verbosity: 0,
    }).promise;
  } catch (e) {
    const name = (e as { name?: string })?.name;
    if (name === 'PasswordException') throw new Error('Il PDF è protetto da password e non si può aprire.');
    throw new Error('Il file non è un PDF valido o è danneggiato.', { cause: e });
  }
  try {
    const blocks: Block[] = [];
    const media: Record<string, string> = {};
    const count = Math.min(pdf.numPages, options.maxPages ?? Infinity);
    const pages: { lines: Line[]; images: PageImage[]; view: number[]; rotate: number }[] = [];
    for (let n = 1; n <= count; n++) {
      const p = await pdf.getPage(n);
      const found = await scanOperators(p, pdfjs.OPS);
      pages.push({ lines: await pageLines(p), images: await pageImages(p, found), view: p.view, rotate: p.rotate });
      p.cleanup();
    }

    // Area del testo di tutto il documento: fa da margini e da riferimento per
    // rientri, allineamento e righe "piene".
    const extents = pages.flatMap(({ lines, images }) => [
      ...lines.map((l) => ({ l: l.x, r: l.right, t: l.y + l.size, b: l.y - l.size * 0.25 })),
      ...images.map((i) => ({ l: i.x, r: i.x + i.width, t: i.top, b: i.top - i.height })),
    ]);
    const first = pages[0];
    const pageWidth = first ? (first.view[2] ?? 0) - (first.view[0] ?? 0) : 0;
    let bbox: { left: number; right: number; top: number; bottom: number } | undefined;
    if (extents.length) {
      const left = Math.max(0, Math.min(...extents.map((e) => e.l)));
      let right = Math.max(...extents.map((e) => e.r));
      if (pageWidth > 0) {
        right = Math.min(pageWidth, right);
        // Nessuna riga arriva al margine destro (testo breve): margini simmetrici.
        if (right - left < (pageWidth - 2 * left) * 0.7) right = pageWidth - left;
      }
      bbox = {
        left,
        right,
        top: Math.max(...extents.map((e) => e.t)),
        bottom: Math.max(0, Math.min(...extents.map((e) => e.b))),
      };
    }
    const page = first ? pageSetupFrom(first.view, first.rotate, bbox) : defaultPageSetup();
    const box: Box = bbox ? { left: bbox.left, right: bbox.right } : { left: page.margins.left, right: page.width - page.margins.right };

    pages.forEach(({ lines, images }, n) => {
      const items: { top: number; bottom: number; make: (pageBreak: boolean, spaceAfter: number) => Block }[] = [
        ...groupLines(lines, box).map((d) => {
          const last = d.lines[d.lines.length - 1]!;
          return { top: d.top, bottom: last.y - last.size * 0.25, make: (pb: boolean, sa: number) => draftToParagraph(d, box, pb, sa) };
        }),
        ...images.map((img) => ({ top: img.top, bottom: img.top - img.height, make: (pb: boolean) => imageParagraph(img, media, box, pb) })),
      ];
      // Le immagini vanno al loro posto fra i paragrafi, dall'alto in basso.
      const ordered = images.length ? sortByTop(items, items.length - images.length) : items;
      ordered.forEach((item, i) => {
        const next = ordered[i + 1];
        const gap = next ? item.bottom - next.top : 8;
        blocks.push(item.make(n > 0 && i === 0, Math.min(gap, 36)));
      });
      if (!ordered.length && n > 0) blocks.push(paragraph([], { pageBreakBefore: true }));
    });
    const info = (await pdf.getMetadata().catch(() => ({ info: undefined }))).info;
    const rawTitle = typeof info?.Title === 'string' ? info.Title.trim() : '';
    const title = rawTitle && !/^untitled$/i.test(rawTitle) ? rawTitle : undefined;
    const doc = createDocument();
    return {
      ...doc,
      sections: [section(blocks, page)],
      media,
      metadata: { ...doc.metadata, ...(title ? { title } : {}) },
    };
  } finally {
    await pdf.destroy();
  }
}

/**
 * Inserisce le immagini fra i paragrafi secondo la posizione verticale,
 * senza riordinare il testo (che segue l'ordine di lettura del PDF, colonne
 * comprese).
 */
function sortByTop<T extends { top: number }>(items: readonly T[], textCount: number): T[] {
  const text = items.slice(0, textCount);
  const images = [...items.slice(textCount)].sort((a, b) => b.top - a.top);
  const out: T[] = [];
  let i = 0;
  for (const t of text) {
    while (i < images.length && images[i]!.top >= t.top) out.push(images[i++]!);
    out.push(t);
  }
  out.push(...images.slice(i));
  return out;
}
