/**
 * Motore di impaginazione: distribuisce i blocchi del documento sulle
 * pagine (sezioni, colonne, intestazioni e piè di pagina, interruzioni,
 * controllo righe isolate, "mantieni con il successivo") e implementa la
 * geometria del cursore (`caretRect`, `hitTest`, `selectionRects`).
 */

import type {
  Block,
  Decoration,
  DocumentLayout,
  DocumentModel,
  HeaderFooter,
  Id,
  LayoutEngine,
  LineBox,
  PageLayout,
  Paragraph,
  Points,
  Rect,
  Section,
  StyleSheet,
  Table,
  TextMeasurer,
} from '../core';
import { layoutFlow, layoutTable, paragraphDecorations, translateDecoration, translateLine, type LayoutContext, type RowBox } from './flow';
import { createGeometry } from './geometry';
import { computeMarkers } from './lists';
import { hasFields, layoutParagraph, type FieldContext, type MarkerInfo, type ParaBox } from './paragraph';

export interface LayoutOptions {
  /** Ora usata dai campi data/ora (iniettabile nei test). */
  readonly now?: () => Date;
}

/** Il motore espone anche il misuratore, che serve al rendering. */
export interface NikLayoutEngine extends LayoutEngine {
  readonly measurer: TextMeasurer;
}

interface Frame {
  readonly x: Points;
  readonly y: Points;
  readonly width: Points;
  readonly bottom: Points;
}

interface MutablePage {
  index: number;
  number: number;
  sectionId: Id;
  width: Points;
  height: Points;
  contentRect: Rect;
  lines: LineBox[];
  decorations: Decoration[];
  header: LineBox[];
  footer: LineBox[];
  /** Il corpo della pagina contiene già qualcosa. */
  hasBody: boolean;
  /** Punto più basso raggiunto dal corpo (per le sezioni continue). */
  maxY: Points;
}

function columnFrames(section: Section, top: Points, bottom: Points): Frame[] {
  const { width, margins } = section.page;
  const contentWidth = width - margins.left - margins.right;
  const count = Math.max(1, section.page.columns?.count ?? 1);
  const gap = count > 1 ? section.page.columns?.gap ?? 36 : 0;
  const colWidth = (contentWidth - gap * (count - 1)) / count;
  return Array.from({ length: count }, (_, i) => ({
    x: margins.left + i * (colWidth + gap),
    y: top,
    width: colWidth,
    bottom,
  }));
}

function pickHeaderFooter(hf: HeaderFooter | undefined, firstOfSection: boolean, titlePage: boolean, number: number) {
  if (!hf) return [];
  if (firstOfSection && titlePage) return hf.first ?? [];
  if (number % 2 === 0 && hf.even) return hf.even;
  return hf.default ?? [];
}

class Paginator implements LayoutContext {
  readonly pages: MutablePage[] = [];
  readonly markers: ReadonlyMap<Id, MarkerInfo>;
  private page!: MutablePage;
  private section!: Section;
  private frames: Frame[] = [];
  private col = 0;
  private y = 0;
  usedNumPages = false;

  constructor(
    readonly doc: DocumentModel,
    readonly measurer: TextMeasurer,
    private readonly numPages: number,
    private readonly now: Date,
    private readonly cache: (p: Paragraph, width: Points, fields: FieldContext, marker?: MarkerInfo) => ParaBox,
  ) {
    this.markers = computeMarkers(doc);
  }

  get fields(): FieldContext {
    return { page: this.page?.number ?? 1, numPages: this.numPages, now: this.now };
  }

  paragraph(p: Paragraph, width: Points, fields: FieldContext): ParaBox {
    if (!this.usedNumPages && p.children.some((c) => c.type === 'field' && c.kind === 'numPages')) this.usedNumPages = true;
    return this.cache(p, width, fields, this.markers.get(p.id));
  }

  run(): MutablePage[] {
    this.doc.sections.forEach((s, si) => {
      this.startSection(s, si === 0);
      s.blocks.forEach((b, bi) => {
        if (b.type === 'paragraph') this.placeParagraph(b, s.blocks[bi + 1]);
        else this.placeTable(b);
      });
    });
    return this.pages;
  }

  // -- Pagine e colonne ------------------------------------------------------

  private get frame(): Frame {
    return this.frames[this.col]!;
  }

  private atTop(): boolean {
    return this.y <= this.frame.y + 0.01;
  }

  private newPage(section: Section, number: number, firstOfSection: boolean): void {
    const { width, height, margins } = section.page;
    const page: MutablePage = {
      index: this.pages.length,
      number,
      sectionId: section.id,
      width,
      height,
      contentRect: { x: 0, y: 0, width: 0, height: 0 },
      lines: [],
      decorations: [],
      header: [],
      footer: [],
      hasBody: false,
      maxY: 0,
    };
    this.pages.push(page);
    this.page = page;
    this.section = section;
    const contentWidth = width - margins.left - margins.right;
    const titlePage = !!section.titlePage;

    const header = layoutFlow(this, pickHeaderFooter(section.header, firstOfSection, titlePage, number), contentWidth);
    page.header = header.lines.map((l) => translateLine(l, margins.left, margins.header));
    page.decorations.push(...header.decorations.map((d) => translateDecoration(d, margins.left, margins.header)));
    const footer = layoutFlow(this, pickHeaderFooter(section.footer, firstOfSection, titlePage, number), contentWidth);
    const footerTop = height - margins.footer - footer.height;
    page.footer = footer.lines.map((l) => translateLine(l, margins.left, footerTop));
    page.decorations.push(...footer.decorations.map((d) => translateDecoration(d, margins.left, footerTop)));

    // Un'intestazione più alta del margine spinge in basso il corpo, come in Word.
    const top = Math.max(margins.top, header.lines.length ? margins.header + header.height : 0);
    const bottom = Math.min(height - margins.bottom, footer.lines.length ? footerTop : height);
    page.contentRect = { x: margins.left, y: top, width: contentWidth, height: Math.max(0, bottom - top) };
    page.maxY = top;
    this.frames = columnFrames(section, top, bottom);
    this.col = 0;
    this.y = top;
  }

  private nextPage(): void {
    this.newPage(this.section, this.page.number + 1, false);
  }

  private nextFrame(): void {
    if (this.col + 1 < this.frames.length) {
      this.col++;
      this.y = this.frame.y;
    } else this.nextPage();
  }

  private startSection(s: Section, first: boolean): void {
    if (first) {
      this.newPage(s, s.pageNumberStart ?? 1, true);
      return;
    }
    const prev = this.page;
    const type = s.breakType ?? 'nextPage';
    if (type === 'continuous' && prev.width === s.page.width && prev.height === s.page.height) {
      this.section = s;
      this.frames = columnFrames(s, prev.maxY, prev.contentRect.y + prev.contentRect.height);
      this.col = 0;
      this.y = prev.maxY;
      return;
    }
    let number = s.pageNumberStart ?? prev.number + 1;
    const wantEven = type === 'evenPage';
    if ((type === 'evenPage' || type === 'oddPage') && (number % 2 === 0) !== wantEven) {
      // Pagina bianca per arrivare a una pagina pari/dispari.
      this.newPage(this.section, prev.number + 1, false);
      if (s.pageNumberStart === undefined) number++;
    }
    this.newPage(s, number, true);
  }

  // -- Paragrafi ---------------------------------------------------------------

  private pushLine(line: LineBox): void {
    this.page.lines.push(line);
    this.page.hasBody = true;
    this.y += line.height;
    this.page.maxY = Math.max(this.page.maxY, this.y);
  }

  private firstLineHeight(b: Block): Points {
    if (b.type === 'paragraph') {
      const box = this.paragraph(b, this.frame.width, this.fields);
      return box.spaceBefore + (box.lines[0]?.height ?? 0);
    }
    return layoutTable(this, b, this.frame.width).rows[0]?.height ?? 0;
  }

  private placeParagraph(p: Paragraph, next: Block | undefined): void {
    let box = this.paragraph(p, this.frame.width, this.fields);
    const fmt = box.format;
    if (fmt.pageBreakBefore && this.page.hasBody) {
      this.nextPage();
      box = this.paragraph(p, this.frame.width, this.fields);
    }
    const frameHeight = this.frame.bottom - this.frame.y;
    const before = () => (this.atTop() ? 0 : box.spaceBefore);
    const remaining = this.frame.bottom - this.y;
    if (!this.atTop()) {
      const total = before() + box.height;
      if (fmt.keepLinesTogether && total > remaining && box.height <= frameHeight) this.nextFrame();
      else if (fmt.keepWithNext && next) {
        // Semplificazione: si sposta l'intero paragrafo (i titoli sono di solito di una riga).
        const whole = total + box.spaceAfter + this.firstLineHeight(next);
        if (whole > remaining && whole <= frameHeight) this.nextFrame();
      }
    }

    this.y += before();
    const lines = box.lines;
    const n = lines.length;
    const widowControl = fmt.widowControl !== false && n > 1;
    let firstOnFrame = 0;
    let chunkTop = this.y;
    const closeChunk = () => {
      if (this.y > chunkTop) {
        this.page.decorations.push(...paragraphDecorations(fmt, this.frame.x, this.frame.width, chunkTop, this.y - chunkTop));
      }
    };
    const popLine = () => {
      const l = this.page.lines.pop()!;
      this.y -= l.height;
    };

    let i = 0;
    while (i < n) {
      const l = lines[i]!;
      if (this.y + l.height > this.frame.bottom + 0.01 && !this.atTop()) {
        const onFrame = i - firstOnFrame;
        if (widowControl && onFrame === 1 && firstOnFrame === 0 && chunkTop > this.frame.y + 0.01) {
          // Riga orfana: la prima riga da sola in fondo alla pagina.
          popLine();
          i--;
        } else if (widowControl && n - i === 1 && onFrame >= 3) {
          // Riga vedova: l'ultima riga da sola in cima alla pagina successiva.
          popLine();
          i--;
        }
        closeChunk();
        this.nextFrame();
        firstOnFrame = i;
        chunkTop = this.y;
        continue;
      }
      const lb: LineBox = {
        paragraphId: p.id,
        start: l.start,
        end: l.end,
        x: this.frame.x + l.x,
        y: this.y,
        width: l.width,
        height: l.height,
        baseline: l.baseline,
        runs: l.runs,
      };
      this.pushLine(lb);
      i++;
      if (l.breakAfter) {
        closeChunk();
        if (l.breakAfter === 'page') this.nextPage();
        else this.nextFrame();
        firstOnFrame = i;
        chunkTop = this.y;
      }
    }
    closeChunk();
    this.y += box.spaceAfter;
    this.page.maxY = Math.max(this.page.maxY, Math.min(this.y, this.frame.bottom));
  }

  // -- Tabelle -----------------------------------------------------------------

  private placeRow(row: RowBox): void {
    const dx = this.frame.x;
    const dy = this.y;
    this.page.decorations.push(...row.decorations.map((d) => translateDecoration(d, dx, dy)));
    this.page.lines.push(...row.lines.map((l) => translateLine(l, dx, dy)));
    this.page.hasBody = true;
    this.y += row.height;
    this.page.maxY = Math.max(this.page.maxY, this.y);
  }

  private placeTable(t: Table): void {
    const box = layoutTable(this, t, this.frame.width);
    let headerCount = 0;
    while (headerCount < box.rows.length && box.rows[headerCount]!.isHeader) headerCount++;
    box.rows.forEach((row, r) => {
      if (this.y + row.height > this.frame.bottom + 0.01 && !this.atTop()) {
        this.nextFrame();
        // Le righe di intestazione si ripetono in cima a ogni pagina.
        if (r >= headerCount) for (let h = 0; h < headerCount; h++) this.placeRow(box.rows[h]!);
      }
      this.placeRow(row);
    });
  }
}

function toPageLayout(p: MutablePage): PageLayout {
  return {
    index: p.index,
    number: p.number,
    sectionId: p.sectionId,
    width: p.width,
    height: p.height,
    contentRect: p.contentRect,
    lines: p.lines,
    decorations: p.decorations,
    header: p.header,
    footer: p.footer,
  };
}

export function createLayoutEngine(measurer: TextMeasurer, options: LayoutOptions = {}): NikLayoutEngine {
  const now = options.now ?? (() => new Date());
  // Cache delle righe: il modello è immutabile, quindi un paragrafo (stesso
  // oggetto) con gli stessi stili e la stessa larghezza dà le stesse righe.
  const cache = new WeakMap<StyleSheet, WeakMap<Paragraph, Map<string, ParaBox>>>();
  const cachedParagraph = (doc: DocumentModel) => (p: Paragraph, width: Points, fields: FieldContext, marker?: MarkerInfo) => {
    if (hasFields(p)) return layoutParagraph(doc, p, width, measurer, fields, marker);
    let byPara = cache.get(doc.styles);
    if (!byPara) cache.set(doc.styles, (byPara = new WeakMap()));
    let byKey = byPara.get(p);
    if (!byKey) byPara.set(p, (byKey = new Map()));
    const key = `${width}|${marker ? JSON.stringify(marker) : ''}`;
    let box = byKey.get(key);
    if (!box) {
      box = layoutParagraph(doc, p, width, measurer, fields, marker);
      byKey.set(key, box);
    }
    return box;
  };

  const geometry = createGeometry(measurer);
  // Più moduli (UI, navigazione) chiedono l'impaginazione dello stesso documento.
  const memo = new WeakMap<DocumentModel, DocumentLayout>();

  return {
    measurer,
    layout(doc) {
      const cached = memo.get(doc);
      if (cached) return cached;
      const date = now();
      let numPages = 1;
      let pages: MutablePage[] = [];
      // Il numero totale di pagine serve ai campi "numPages": si ripete
      // l'impaginazione finché il totale non è stabile (di solito 1-2 volte).
      for (let attempt = 0; attempt < 3; attempt++) {
        const pag = new Paginator(doc, measurer, numPages, date, cachedParagraph(doc));
        pages = pag.run();
        if (!pag.usedNumPages || pages.length === numPages) break;
        numPages = pages.length;
      }
      const result: DocumentLayout = { doc, pages: pages.map(toPageLayout) };
      memo.set(doc, result);
      return result;
    },
    caretRect: geometry.caretRect,
    hitTest: geometry.hitTest,
    selectionRects: geometry.selectionRects,
  };
}

export type { DocumentLayout };
