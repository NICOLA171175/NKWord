/**
 * Motore di impaginazione: distribuisce i blocchi del documento sulle
 * pagine (sezioni, colonne, intestazioni e piè di pagina, interruzioni,
 * controllo righe isolate, "mantieni con il successivo") e implementa la
 * geometria del cursore (`caretRect`, `hitTest`, `selectionRects`).
 *
 * Le note a piè di pagina occupano il fondo della pagina in cui compare il
 * loro richiamo: lo spazio del corpo si accorcia di conseguenza, e se una
 * riga non ci sta più insieme alle sue note passa alla pagina successiva.
 */

import { footnoteNumbers, paragraphs } from '../core';
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
import { layoutFlow, layoutTable, paragraphDecorations, translateDecoration, translateLine, type Flow, type LayoutContext, type RowBox } from './flow';
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

/** Spazio sopra le note a piè di pagina, con la riga separatrice a metà. */
export const FOOTNOTE_SEPARATOR_SPACE: Points = 12;
/** Lunghezza massima della riga separatrice (5 cm circa, come in Word). */
export const FOOTNOTE_SEPARATOR_WIDTH: Points = 144;

/** Nota collocata su una pagina; `line` è la riga del corpo che la richiama. */
interface PageNote {
  readonly noteId: Id;
  readonly flow: Flow;
  readonly line?: LineBox;
}

function notesHeight(notes: readonly PageNote[]): Points {
  return notes.length ? FOOTNOTE_SEPARATOR_SPACE + notes.reduce((h, n) => h + n.flow.height, 0) : 0;
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
  /** Note a piè di pagina richiamate nella pagina, in ordine di lettura. */
  notes: PageNote[];
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
  readonly footnotes: ReadonlyMap<Id, number>;
  /** Richiami di nota nei paragrafi del corpo: offset e nota. */
  private readonly noteRefs = new Map<Id, { offset: number; noteId: Id }[]>();
  private readonly placedNotes = new Set<Id>();

  constructor(
    readonly doc: DocumentModel,
    readonly measurer: TextMeasurer,
    private readonly numPages: number,
    private readonly now: Date,
    private readonly cache: (p: Paragraph, width: Points, fields: FieldContext, marker?: MarkerInfo) => ParaBox,
  ) {
    this.markers = computeMarkers(doc);
    this.footnotes = footnoteNumbers(doc);
    if (this.footnotes.size) {
      for (const { paragraph: p } of paragraphs(doc)) {
        let offset = 0;
        for (const c of p.children) {
          if (c.type === 'footnoteRef' && this.footnotes.has(c.noteId)) {
            const refs = this.noteRefs.get(p.id) ?? [];
            refs.push({ offset, noteId: c.noteId });
            this.noteRefs.set(p.id, refs);
          }
          offset += c.type === 'text' ? c.text.length : 1;
        }
      }
    }
  }

  get fields(): FieldContext {
    return { page: this.page?.number ?? 1, numPages: this.numPages, now: this.now, footnotes: this.footnotes };
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
      notes: [],
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
      this.frames = columnFrames(s, prev.maxY, prev.contentRect.y + prev.contentRect.height - notesHeight(prev.notes));
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

  // -- Note a piè di pagina ----------------------------------------------------

  /** Note richiamate dalle righe e non ancora collocate, già impaginate. */
  private newNotes(lines: readonly { paragraphId: Id; start: number; end: number }[]): PageNote[] {
    if (!this.noteRefs.size) return [];
    const out: PageNote[] = [];
    for (const l of lines) {
      for (const r of this.noteRefs.get(l.paragraphId) ?? []) {
        if (r.offset < l.start || r.offset >= l.end || this.placedNotes.has(r.noteId) || out.some((n) => n.noteId === r.noteId)) continue;
        const note = this.doc.footnotes?.[r.noteId];
        if (note) out.push({ noteId: r.noteId, flow: layoutFlow(this, note.blocks, this.page.contentRect.width) });
      }
    }
    return out;
  }

  /** Spazio in più che occuperebbero le note sulla pagina corrente. */
  private notesDelta(notes: readonly PageNote[]): Points {
    return notes.length ? notesHeight([...this.page.notes, ...notes]) - notesHeight(this.page.notes) : 0;
  }

  /** Il fondo del corpo si alza dell'altezza delle note (in tutte le colonne). */
  private updateNoteArea(): void {
    const bottom = this.page.contentRect.y + this.page.contentRect.height - notesHeight(this.page.notes);
    this.frames = this.frames.map((f) => ({ ...f, bottom }));
  }

  private addNotes(notes: readonly PageNote[], line?: LineBox): void {
    if (!notes.length) return;
    for (const n of notes) {
      this.page.notes.push({ ...n, line });
      this.placedNotes.add(n.noteId);
    }
    this.updateNoteArea();
  }

  private removeNotes(line: LineBox): void {
    const gone = this.page.notes.filter((n) => n.line === line);
    if (!gone.length) return;
    this.page.notes = this.page.notes.filter((n) => n.line !== line);
    for (const n of gone) this.placedNotes.delete(n.noteId);
    this.updateNoteArea();
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
      this.removeNotes(l);
    };

    let i = 0;
    while (i < n) {
      const l = lines[i]!;
      const notes = this.newNotes([{ paragraphId: p.id, start: l.start, end: l.end }]);
      if (this.y + l.height + this.notesDelta(notes) > this.frame.bottom + 0.01 && !this.atTop()) {
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
      this.addNotes(notes, lb);
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
    const lines = row.lines.map((l) => translateLine(l, dx, dy));
    this.page.lines.push(...lines);
    this.addNotes(this.newNotes(lines), lines[0]);
    this.page.hasBody = true;
    this.y += row.height;
    this.page.maxY = Math.max(this.page.maxY, this.y);
  }

  private placeTable(t: Table): void {
    const box = layoutTable(this, t, this.frame.width);
    let headerCount = 0;
    while (headerCount < box.rows.length && box.rows[headerCount]!.isHeader) headerCount++;
    box.rows.forEach((row, r) => {
      if (this.y + row.height + this.notesDelta(this.newNotes(row.lines)) > this.frame.bottom + 0.01 && !this.atTop()) {
        this.nextFrame();
        // Le righe di intestazione si ripetono in cima a ogni pagina.
        if (r >= headerCount) for (let h = 0; h < headerCount; h++) this.placeRow(box.rows[h]!);
      }
      this.placeRow(row);
    });
  }
}

/** Riga separatrice e testo delle note, impilati in fondo all'area del corpo. */
function placeNotes(p: MutablePage): { lines: LineBox[]; decorations: Decoration[] } {
  const lines: LineBox[] = [];
  const decorations: Decoration[] = [];
  if (!p.notes.length) return { lines, decorations };
  const x = p.contentRect.x;
  let y = p.contentRect.y + p.contentRect.height - notesHeight(p.notes);
  decorations.push({
    kind: 'border',
    color: '#000000',
    width: 0.5,
    rect: { x, y: y + FOOTNOTE_SEPARATOR_SPACE / 2 - 0.25, width: Math.min(FOOTNOTE_SEPARATOR_WIDTH, p.contentRect.width), height: 0.5 },
  });
  y += FOOTNOTE_SEPARATOR_SPACE;
  for (const n of p.notes) {
    lines.push(...n.flow.lines.map((l) => translateLine(l, x, y)));
    decorations.push(...n.flow.decorations.map((d) => translateDecoration(d, x, y)));
    y += n.flow.height;
  }
  return { lines, decorations };
}

function toPageLayout(p: MutablePage): PageLayout {
  const notes = placeNotes(p);
  return {
    index: p.index,
    number: p.number,
    sectionId: p.sectionId,
    width: p.width,
    height: p.height,
    contentRect: p.contentRect,
    lines: p.lines,
    decorations: notes.decorations.length ? [...p.decorations, ...notes.decorations] : p.decorations,
    header: p.header,
    footer: p.footer,
    ...(notes.lines.length ? { footnotes: notes.lines } : {}),
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
