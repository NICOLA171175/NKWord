/**
 * Da `DocumentLayout` a PDF. Si disegna esattamente la geometria calcolata
 * dal modulo layout (pagine, righe, run, decorazioni), quindi il PDF ha le
 * stesse interruzioni di riga e di pagina che si vedono a schermo.
 */

import type { CharFormat, DocumentLayout, DocumentMetadata, GlyphRun, LineBox, PageLayout, TextMeasurer } from '../core';
import { normalizeColor } from '../io/util/color';
import { drawingPdfOps } from '../drawings/pdf';
import { baselineShift, effectiveFontSize, SMALL_CAPS_SCALE } from '../layout';
import { encodedWidth, encodeWinAnsi, standardFont, type StandardFont } from './fonts';
import { loadImage, type ImageRasterizer, type PdfImage } from './images';
import { hexString, num, PdfWriter, pdfName, pdfString, ref } from './writer';

export interface PdfRenderOptions {
  /** Lo stesso misuratore del layout: serve a sapere quanto è largo ogni pezzo di testo. */
  readonly measurer: TextMeasurer;
  /** Per le immagini che non sono JPEG o PNG. */
  readonly rasterize?: ImageRasterizer;
  /** Data di creazione scritta nei metadati (default: adesso). */
  readonly now?: () => Date;
}

function rgb(color: string | null | undefined, fallback: string | undefined = '#000000'): string | null {
  const hex = normalizeColor(color ?? undefined) ?? fallback;
  if (!hex) return null;
  const v = (i: number) => num(parseInt(hex.slice(i, i + 2), 16) / 255);
  return `${v(1)} ${v(3)} ${v(5)}`;
}

function pdfDate(d: Date): string {
  const p = (n: number) => String(n).padStart(2, '0');
  return `D:${d.getUTCFullYear()}${p(d.getUTCMonth() + 1)}${p(d.getUTCDate())}${p(d.getUTCHours())}${p(d.getUTCMinutes())}${p(d.getUTCSeconds())}Z`;
}

/** Collegamenti cliccabili solo verso indirizzi sicuri. */
function linkTarget(link: string | undefined): string | null {
  if (!link || !/^(https?:|mailto:)/i.test(link.trim())) return null;
  return encodeURI(decodeURIComponentSafe(link.trim()));
}

function decodeURIComponentSafe(s: string): string {
  try {
    return decodeURI(s);
  } catch {
    return s;
  }
}

/** Paragrafo di destinazione di un collegamento interno (`#_Toc<id>`), `''` per l'inizio del documento. */
function internalTarget(link: string | undefined): string | null {
  if (!link?.startsWith('#')) return null;
  if (link.toLowerCase() === '#_top') return '';
  return link.startsWith('#_Toc') ? link.slice(5) : null;
}

/** Pagina e altezza (dall'alto) della prima riga di ogni paragrafo. */
function paragraphPositions(layout: DocumentLayout): Map<string, { page: number; top: number }> {
  const out = new Map<string, { page: number; top: number }>();
  for (const page of layout.pages) {
    for (const line of page.lines) if (!out.has(line.paragraphId)) out.set(line.paragraphId, { page: page.index, top: line.y });
  }
  return out;
}

const isTextRun = (r: GlyphRun) => r.kind === 'text' || r.kind === 'field' || r.kind === 'listMarker';

class PageCanvas {
  readonly ops: string[] = [];
  readonly links: { rect: [number, number, number, number]; uri: string }[] = [];
  /** Collegamenti a un punto del documento: id del paragrafo ('' = inizio). */
  readonly gotos: { rect: [number, number, number, number]; paragraphId: string }[] = [];

  constructor(
    private readonly page: PageLayout,
    private readonly fonts: FontTable,
    private readonly images: Map<string, string>,
    private readonly measurer: TextMeasurer,
  ) {}

  /** y dall'alto (layout) → y dal basso (PDF). */
  private y(top: number): number {
    return this.page.height - top;
  }

  fillRect(x: number, top: number, w: number, h: number, color: string): void {
    this.ops.push(`${color} rg ${num(x)} ${num(this.y(top + h))} ${num(w)} ${num(h)} re f`);
  }

  line(line: LineBox): void {
    const last = line.runs[line.runs.length - 1];
    for (const run of line.runs) {
      if (run.kind === 'image') this.image(run, line);
      else if (isTextRun(run)) this.text(run, line, run === last);
    }
  }

  private image(run: GlyphRun, line: LineBox): void {
    const x = line.x + run.x;
    const h = run.height ?? line.height;
    const bottom = line.y + line.baseline;
    const name = run.src ? this.images.get(run.src) : undefined;
    if (name) {
      this.ops.push(`q ${num(run.width)} 0 0 ${num(h)} ${num(x)} ${num(this.y(bottom))} cm ${name} Do Q`);
    } else {
      // Formato non gestito: un riquadro grigio al posto dell'immagine.
      this.fillRect(x, bottom - h, run.width, h, '0.9 0.9 0.9');
      this.ops.push(`0.6 0.6 0.6 RG 0.5 w ${num(x)} ${num(this.y(bottom))} ${num(run.width)} ${num(h)} re S`);
    }
  }

  private text(run: GlyphRun, line: LineBox, isLast: boolean): void {
    const f = run.format;
    const x = line.x + run.x;
    const body = run.text.replace(/\s+$/, '');
    // Larghezza vera del testo (senza spazi finali) secondo il misuratore del layout.
    const natural = body ? Math.min(run.width, this.measurer.measure(body, f).width) : 0;
    const decoWidth = isLast ? natural : run.width;

    const highlight = f.highlight ? rgb(f.highlight, undefined) : null;
    if (highlight && run.width > 0) this.fillRect(x, line.y, isLast ? natural : run.width, line.height, highlight);
    if (!body) return;

    let size = effectiveFontSize(f);
    if (f.smallCaps) size *= SMALL_CAPS_SCALE;
    const shown = f.smallCaps ? body.toUpperCase() : body;
    const font = standardFont(f);
    const bytes = encodeWinAnsi(shown);
    const pdfWidth = encodedWidth(font, bytes, size);
    const scale = pdfWidth > 0 && natural > 0 ? Math.max(20, Math.min(500, (natural / pdfWidth) * 100)) : 100;
    const baseline = line.y + line.baseline - baselineShift(f);
    const color = rgb(f.color)!;
    this.ops.push(
      `BT ${this.fonts.use(font)} ${num(size)} Tf ${num(scale)} Tz ${color} rg 1 0 0 1 ${num(x)} ${num(this.y(baseline))} Tm ${hexString(bytes)} Tj ET`,
    );

    const full = f.fontSize ?? 11;
    const thickness = Math.max(0.5, full * 0.06);
    if (f.underline) this.rule(x, baseline + full * 0.12, decoWidth, thickness, color, f.underline);
    if (f.strike) this.rule(x, baseline - size * 0.3, decoWidth, thickness, color, 'single');

    const rect: [number, number, number, number] = [x, this.y(line.y + line.height), x + decoWidth, this.y(line.y)];
    const uri = linkTarget(f.link);
    if (uri) this.links.push({ rect, uri });
    const target = internalTarget(f.link);
    if (target !== null) this.gotos.push({ rect, paragraphId: target });
  }

  private rule(x: number, top: number, width: number, thickness: number, color: string, style: NonNullable<CharFormat['underline']>): void {
    const y = num(this.y(top));
    const dash = style === 'dotted' ? `[${num(thickness)} ${num(thickness * 2)}] 0 d` : style === 'dashed' ? `[${num(thickness * 4)} ${num(thickness * 2)}] 0 d` : '[] 0 d';
    const seg = (yy: string) => `${num(x)} ${yy} m ${num(x + width)} ${yy} l S`;
    let path = seg(y);
    if (style === 'double') path += ` ${seg(num(this.y(top + thickness * 2)))}`;
    this.ops.push(`q ${color} RG ${num(thickness)} w ${dash} ${path} Q`);
  }
}

/** Caratteri usati, con il nome di risorsa assegnato (/F1, /F2...). */
class FontTable {
  readonly names = new Map<string, string>();
  use(font: StandardFont): string {
    let name = this.names.get(font.name);
    if (!name) {
      name = `/F${this.names.size + 1}`;
      this.names.set(font.name, name);
    }
    return name;
  }
}

function imageSources(layout: DocumentLayout): Set<string> {
  const out = new Set<string>();
  for (const p of layout.pages) for (const lines of [p.header, p.lines, p.footer]) for (const l of lines) for (const r of l.runs) if (r.kind === 'image' && r.src) out.add(r.src);
  return out;
}

function info(writer: PdfWriter, meta: DocumentMetadata, now: Date): number {
  const entries = [`/Producer ${pdfString('NKoffice')}`, `/Creator ${pdfString('NKoffice')}`, `/CreationDate ${pdfString(pdfDate(now))}`];
  if (meta.title) entries.push(`/Title ${pdfString(meta.title)}`);
  if (meta.author) entries.push(`/Author ${pdfString(meta.author)}`);
  if (meta.subject) entries.push(`/Subject ${pdfString(meta.subject)}`);
  if (meta.keywords?.length) entries.push(`/Keywords ${pdfString(meta.keywords.join(', '))}`);
  return writer.add(`<< ${entries.join(' ')} >>`);
}

/** Scrive l'impaginazione come file PDF. */
export async function layoutToPdf(layout: DocumentLayout, options: PdfRenderOptions): Promise<Uint8Array> {
  const writer = new PdfWriter();
  const catalogId = writer.alloc();
  const pagesId = writer.alloc();
  const fonts = new FontTable();

  // Immagini: una volta sola per sorgente, anche se compaiono su più pagine.
  const images = new Map<string, string>();
  const xobjects: string[] = [];
  for (const src of imageSources(layout)) {
    const img: PdfImage | null = await loadImage(layout.doc.media[src] ?? src, options.rasterize);
    if (!img) continue;
    const base = `/Type /XObject /Subtype /Image /Width ${img.width} /Height ${img.height}`;
    const smask = img.alpha
      ? ` /SMask ${ref(writer.addStream(`/Type /XObject /Subtype /Image /Width ${img.width} /Height ${img.height} /ColorSpace /DeviceGray /BitsPerComponent 8${img.alphaFilter ? ` /Filter ${img.alphaFilter}` : ''}`, img.alpha))}`
      : '';
    const id = writer.addStream(`${base} ${img.dict}${smask}`, img.data);
    const name = `/Im${images.size + 1}`;
    images.set(src, name);
    xobjects.push(`${name} ${ref(id)}`);
  }

  const pageIds: number[] = [];
  const contents: { id: number; canvas: PageCanvas }[] = [];
  for (const page of layout.pages) {
    const canvas = new PageCanvas(page, fonts, images, options.measurer);
    for (const d of page.decorations) {
      const color = rgb(d.color, undefined);
      if (color) canvas.fillRect(d.rect.x, d.rect.y, d.rect.width, d.rect.height, color);
    }
    for (const lines of [page.header, page.lines, page.footnotes ?? [], page.footer]) for (const l of lines) canvas.line(l);
    // Forme, caselle di testo e grafici: davanti al testo.
    canvas.ops.push(...drawingPdfOps(layout, page.index, (font) => fonts.use(font)));
    contents.push({ id: writer.addStream('', canvas.ops.join('\n')), canvas });
    pageIds.push(writer.alloc());
  }

  // Le risorse dei caratteri sono condivise da tutte le pagine.
  const fontRefs: string[] = [];
  for (const [fontName, resName] of fonts.names) {
    const id = writer.add(`<< /Type /Font /Subtype /Type1 /BaseFont ${pdfName(fontName)} /Encoding /WinAnsiEncoding >>`);
    fontRefs.push(`${resName} ${ref(id)}`);
  }
  const resources = `<< /ProcSet [/PDF /Text /ImageB /ImageC /ImageI]${fontRefs.length ? ` /Font << ${fontRefs.join(' ')} >>` : ''}${xobjects.length ? ` /XObject << ${xobjects.join(' ')} >>` : ''} >>`;
  const resourcesId = writer.add(resources);

  const positions = paragraphPositions(layout);
  layout.pages.forEach((page, i) => {
    const { id: contentId, canvas } = contents[i]!;
    const annots = canvas.links.map(({ rect, uri }) =>
      writer.add(`<< /Type /Annot /Subtype /Link /Rect [${rect.map(num).join(' ')}] /Border [0 0 0] /A << /S /URI /URI ${pdfString(uri)} >> >>`),
    );
    for (const { rect, paragraphId } of canvas.gotos) {
      const to = paragraphId ? positions.get(paragraphId) : { page: 0, top: 0 };
      const toPage = to && layout.pages[to.page];
      if (!to || !toPage) continue;
      const dest = `[${ref(pageIds[to.page]!)} /XYZ null ${num(toPage.height - to.top)} null]`;
      annots.push(writer.add(`<< /Type /Annot /Subtype /Link /Rect [${rect.map(num).join(' ')}] /Border [0 0 0] /Dest ${dest} >>`));
    }
    writer.set(
      pageIds[i]!,
      `<< /Type /Page /Parent ${ref(pagesId)} /MediaBox [0 0 ${num(page.width)} ${num(page.height)}] /Resources ${ref(resourcesId)} /Contents ${ref(contentId)}${annots.length ? ` /Annots [${annots.map(ref).join(' ')}]` : ''} >>`,
    );
  });

  writer.set(pagesId, `<< /Type /Pages /Kids [${pageIds.map(ref).join(' ')}] /Count ${pageIds.length} >>`);
  const title = layout.doc.metadata.title;
  writer.set(catalogId, `<< /Type /Catalog /Pages ${ref(pagesId)}${title ? ' /ViewerPreferences << /DisplayDocTitle true >>' : ''} >>`);
  return writer.finish(catalogId, info(writer, layout.doc.metadata, options.now?.() ?? new Date()));
}
