/**
 * Vista del documento nel DOM.
 *
 * - Con un `LayoutEngine` in `editor.services.layout` disegna le pagine
 *   impaginate: righe e run in posizione assoluta, cursore e selezione dai
 *   rettangoli del motore, clic tramite `hitTest`.
 * - Senza impaginazione (modulo non ancora presente) usa una vista a
 *   scorrimento: una "pagina" per sezione, paragrafi nel flusso HTML.
 *
 * In entrambi i casi la vista è di sola lettura: ogni modifica passa da un
 * comando (`edit.setSelection` per i clic).
 */

import {
  footnoteNumbers,
  inlineLength,
  isCollapsed,
  orderedRange,
  paragraphLength,
  paragraphs,
  ptToPx,
  pxToPt,
  resolveCharFormat,
  resolveParaFormat,
  type Block,
  type CharFormat,
  type DocumentLayout,
  type DocumentModel,
  type Editor,
  type GlyphRun,
  type Inline,
  type LineBox,
  type Paragraph,
  type Position,
  type Section,
  type Selection,
} from '../core';
import { collectMarks, onExtensionsChange, viewExtensions, type TextMark } from './extensions';
import { runFirstExecutable } from './input';
import { TabLeaders } from './tab-leaders';
import type { ViewState, ViewStore } from './view-state';

const FIELD_PLACEHOLDER: Record<string, string> = { page: '1', numPages: '1', date: 'gg/mm/aaaa', time: 'hh:mm' };

/** Stile CSS di un formato di carattere, allo zoom dato. */
export function charStyle(f: CharFormat, zoom: number): Partial<CSSStyleDeclaration> {
  const deco: string[] = [];
  if (f.underline) deco.push('underline');
  if (f.strike) deco.push('line-through');
  const style: Partial<CSSStyleDeclaration> = {
    fontFamily: f.fontFamily ? `"${f.fontFamily}", sans-serif` : '',
    fontSize: f.fontSize ? `${ptToPx(f.fontSize, zoom)}px` : '',
    fontWeight: f.bold ? 'bold' : '',
    fontStyle: f.italic ? 'italic' : '',
    textDecorationLine: deco.join(' '),
    textDecorationStyle: f.underline && f.underline !== 'single' ? (f.underline === 'dashed' ? 'dashed' : f.underline) : '',
    color: f.color ?? '',
    backgroundColor: f.highlight ?? '',
    verticalAlign: f.verticalAlign === 'superscript' ? 'super' : f.verticalAlign === 'subscript' ? 'sub' : '',
    fontVariant: f.smallCaps ? 'small-caps' : '',
    textTransform: f.allCaps ? 'uppercase' : '',
  };
  if (f.verticalAlign && f.verticalAlign !== 'baseline') style.fontSize = `${ptToPx(scriptFontSize(f), zoom)}px`;
  return style;
}

// Stessi valori di src/layout/measure.ts (apice e pedice come in Word).
const DEFAULT_SIZE = 11;
const SCRIPT_SCALE = 0.65;

function scriptFontSize(f: CharFormat): number {
  return (f.fontSize ?? DEFAULT_SIZE) * SCRIPT_SCALE;
}

/** Spostamento della linea di base in punti (positivo = verso l'alto), come nell'impaginazione. */
export function runBaselineShift(f: CharFormat): number {
  const size = f.fontSize ?? DEFAULT_SIZE;
  if (f.verticalAlign === 'superscript') return size * 0.33;
  if (f.verticalAlign === 'subscript') return -size * 0.14;
  return 0;
}

let metricsCtx: CanvasRenderingContext2D | null | undefined;
const metricsCache = new Map<string, { ascent: number; descent: number }>();

/**
 * Ascendente e discendente (in px) del carattere con cui il browser disegna
 * uno span con lo stile dato. Senza canvas (test) si usano proporzioni tipiche.
 */
function fontMetrics(style: Partial<CSSStyleDeclaration>, sizePx: number): { ascent: number; descent: number } {
  const font = `${style.fontStyle || 'normal'} ${style.fontWeight || 'normal'} ${sizePx}px ${style.fontFamily || 'sans-serif'}`;
  let m = metricsCache.get(font);
  if (m) return m;
  if (metricsCtx === undefined) {
    try {
      metricsCtx = document.createElement('canvas').getContext('2d');
    } catch {
      metricsCtx = null;
    }
  }
  let ascent = sizePx * 0.8;
  let descent = sizePx * 0.2;
  if (metricsCtx) {
    metricsCtx.font = font;
    const tm = metricsCtx.measureText('Hg');
    if (tm.fontBoundingBoxAscent) ascent = tm.fontBoundingBoxAscent;
    if (tm.fontBoundingBoxDescent) descent = tm.fontBoundingBoxDescent;
  }
  m = { ascent, descent };
  metricsCache.set(font, m);
  return m;
}

/**
 * Posizione verticale di un run di testo nella riga: la sua linea di base
 * cade su quella della riga, alzata o abbassata per apice e pedice. Senza
 * questo ogni run starebbe in cima alla riga e un pedice (più piccolo)
 * sembrerebbe un apice.
 */
export function runVerticalStyle(f: CharFormat, line: LineBox, zoom: number): { top: string; lineHeight: string } {
  const style = charStyle(f, zoom);
  const sizePx = parseFloat(style.fontSize || '') || ptToPx(f.fontSize ?? DEFAULT_SIZE, zoom);
  const { ascent, descent } = fontMetrics(style, sizePx);
  const baseline = ptToPx(line.baseline - runBaselineShift(f), zoom);
  return { top: `${baseline - ascent}px`, lineHeight: `${ascent + descent}px` };
}

/** Disegna un segno delle estensioni su uno span di testo (vista a scorrimento). */
function applyMark(el: HTMLElement, mark: TextMark): void {
  el.classList.add('text-mark', `mark-${mark.kind}`);
  if (mark.className) el.classList.add(mark.className);
  if (mark.kind === 'highlight') {
    el.style.backgroundColor = mark.color;
    return;
  }
  const lines = new Set(el.style.textDecorationLine.split(' ').filter(Boolean));
  lines.add('underline');
  el.style.textDecorationLine = [...lines].join(' ');
  el.style.textDecorationStyle = mark.kind === 'squiggle' ? 'wavy' : 'solid';
  el.style.textDecorationColor = mark.color;
}

/** Testo commentato: classe ed elenco degli id, per evidenziarlo e trovarlo. */
function markComments(el: HTMLElement, f: CharFormat): void {
  if (!f.commentIds) return;
  el.classList.add('run-comment');
  el.dataset.comments = f.commentIds;
}

function applyStyle(el: HTMLElement, style: Partial<CSSStyleDeclaration>): void {
  for (const [k, v] of Object.entries(style)) if (v) (el.style as unknown as Record<string, string>)[k] = v as string;
}

/** Intervallo selezionato in ciascun paragrafo (offset), per la vista a scorrimento. */
export function selectionByParagraph(doc: DocumentModel, sel: Selection): Map<string, [number, number]> {
  const out = new Map<string, [number, number]>();
  if (isCollapsed(sel)) return out;
  const [from, to] = orderedRange(doc, sel);
  let inside = false;
  for (const { paragraph: p } of paragraphs(doc)) {
    const isFrom = p.id === from.paragraphId;
    const isTo = p.id === to.paragraphId;
    if (isFrom) inside = true;
    if (inside) out.set(p.id, [isFrom ? from.offset : 0, isTo ? to.offset : paragraphLength(p) + 1]);
    if (isTo) break;
  }
  return out;
}

export class DocumentView {
  readonly element: HTMLElement;
  private readonly pages: HTMLElement;
  private caretEl: HTMLElement | null = null;
  private cleanup: (() => void)[] = [];
  private layoutCache: { doc: DocumentModel; layout: DocumentLayout } | null = null;
  private dragging = false;
  private marks = new Map<string, TextMark[]>();

  constructor(
    private readonly editor: Editor,
    private readonly view: ViewStore,
  ) {
    this.element = document.createElement('div');
    this.element.className = 'doc-scroll';
    this.pages = document.createElement('div');
    this.pages.className = 'doc-pages';
    this.element.append(this.pages);

    const onDown = (e: MouseEvent) => this.onMouseDown(e);
    const onMove = (e: MouseEvent) => this.dragging && this.select(e, true);
    const onUp = () => {
      this.dragging = false;
    };
    const onDbl = (e: MouseEvent) => this.onDoubleClick(e);
    this.pages.addEventListener('mousedown', onDown);
    this.pages.addEventListener('dblclick', onDbl);
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
    this.cleanup.push(
      () => this.pages.removeEventListener('mousedown', onDown),
      () => this.pages.removeEventListener('dblclick', onDbl),
      () => document.removeEventListener('mousemove', onMove),
      () => document.removeEventListener('mouseup', onUp),
      editor.events.on('change', () => this.render()),
      view.events.on('change', () => this.render()),
      onExtensionsChange(editor, () => this.render()),
    );
    this.render();
  }

  destroy(): void {
    for (const c of this.cleanup.splice(0)) c();
    this.element.remove();
  }

  /** Rettangolo del cursore sullo schermo (per posizionare l'input e l'IME). */
  caretClientRect(): DOMRect | null {
    return this.caretEl?.getBoundingClientRect() ?? null;
  }

  render(): void {
    const { state } = this.editor;
    const vs = this.view.state;
    this.pages.replaceChildren();
    this.caretEl = null;
    this.pages.classList.toggle('show-marks', vs.formattingMarks);
    this.pages.dataset.headerFooter = vs.headerFooter ?? '';
    this.marks = new Map();
    for (const m of collectMarks(this.editor, state)) {
      if (m.end <= m.start) continue;
      const list = this.marks.get(m.paragraphId);
      if (list) list.push(m);
      else this.marks.set(m.paragraphId, [m]);
    }
    const engine = this.editor.services.layout;
    if (engine) {
      if (this.layoutCache?.doc !== state.doc) this.layoutCache = { doc: state.doc, layout: engine.layout(state.doc) };
      this.renderLayout(this.layoutCache.layout, vs);
    } else {
      this.layoutCache = null;
      state.doc.sections.forEach((s) => this.pages.append(this.renderFlowSection(state.doc, s, vs)));
    }
  }

  // -------------------------------------------------------------------------
  // Vista impaginata

  private renderLayout(layout: DocumentLayout, vs: ViewState): void {
    const engine = this.editor.services.layout!;
    const { doc, selection } = this.editor.state;
    const px = (pt: number) => `${ptToPx(pt, vs.zoom)}px`;
    const pageEls: HTMLElement[] = [];
    const leaders = new TabLeaders(doc);
    for (const page of layout.pages) {
      const el = document.createElement('div');
      el.className = 'page page-laid-out';
      el.dataset.page = String(page.index);
      el.style.width = px(page.width);
      el.style.height = px(page.height);
      for (const d of page.decorations) {
        const box = document.createElement('div');
        box.className = `decoration decoration-${d.kind}`;
        Object.assign(box.style, { left: px(d.rect.x), top: px(d.rect.y), width: px(d.rect.width), height: px(d.rect.height) });
        // Per i bordi il rettangolo è già il tratto (spesso quanto il bordo): si riempie come un fondo.
        box.style.background = d.color;
        el.append(box);
      }
      const area = (cls: string, lines: readonly LineBox[]) => {
        const div = document.createElement('div');
        div.className = cls;
        for (const line of lines) div.append(this.renderLine(line, vs.zoom, leaders, page.contentRect.x));
        el.append(div);
      };
      area('page-area page-header-area', page.header);
      area('page-area page-body-area', page.lines);
      if (page.footnotes?.length) area('page-area page-footnote-area', page.footnotes);
      area('page-area page-footer-area', page.footer);
      this.pages.append(el);
      pageEls.push(el);
    }
    if (!isCollapsed(selection)) {
      const [from, to] = orderedRange(doc, selection);
      for (const { page, rect } of engine.selectionRects(layout, from, to)) {
        const r = document.createElement('div');
        r.className = 'selection-rect';
        Object.assign(r.style, { left: px(rect.x), top: px(rect.y), width: px(rect.width), height: px(rect.height) });
        pageEls[page]?.append(r);
      }
    }
    for (const mark of [...this.marks.values()].flat()) {
      const from = { paragraphId: mark.paragraphId, offset: mark.start };
      const to = { paragraphId: mark.paragraphId, offset: mark.end };
      for (const { page, rect } of engine.selectionRects(layout, from, to)) {
        const r = document.createElement('div');
        r.className = `text-mark mark-${mark.kind}`;
        if (mark.className) r.classList.add(mark.className);
        r.style.setProperty('--mark-color', mark.color);
        const h = mark.kind === 'highlight' ? rect.height : Math.min(rect.height, 3 / ptToPx(1, vs.zoom));
        Object.assign(r.style, { left: px(rect.x), top: px(rect.y + rect.height - h), width: px(rect.width), height: px(h) });
        pageEls[page]?.append(r);
      }
    }
    for (const ext of viewExtensions(this.editor)) ext.decoratePages?.(pageEls, layout, vs.zoom);
    const c = engine.caretRect(layout, selection.focus);
    if (c && pageEls[c.page]) {
      const caret = document.createElement('div');
      caret.className = 'caret';
      Object.assign(caret.style, { left: px(c.rect.x), top: px(c.rect.y), height: px(c.rect.height) });
      pageEls[c.page]!.append(caret);
      this.caretEl = caret;
    }
  }

  private renderLine(line: LineBox, zoom: number, leaders: TabLeaders, columnLeft: number): HTMLElement {
    const px = (pt: number) => `${ptToPx(pt, zoom)}px`;
    const el = document.createElement('div');
    el.className = 'line';
    Object.assign(el.style, { left: px(line.x), top: px(line.y), width: px(line.width), height: px(line.height) });
    for (const run of line.runs) {
      const runEl = this.renderRun(run, line, zoom);
      const leader = leaders.leader(run, line.x - columnLeft + run.x);
      if (leader) {
        // Più caratteri del necessario: l'eccesso è tagliato a sinistra.
        runEl.textContent = leader.repeat(Math.ceil(run.width / 2) + 1);
        runEl.classList.add('run-tab-leader');
      }
      el.append(runEl);
    }
    return el;
  }

  private renderRun(run: GlyphRun, line: LineBox, zoom: number): HTMLElement {
    const px = (pt: number) => `${ptToPx(pt, zoom)}px`;
    let el: HTMLElement;
    if (run.kind === 'image' && run.src) {
      const img = document.createElement('img');
      img.src = this.editor.state.doc.media[run.src] ?? run.src;
      img.alt = '';
      img.style.height = px(run.height ?? line.height);
      el = img;
    } else {
      el = document.createElement('span');
      el.textContent = run.kind === 'tab' ? '\t' : run.text;
      applyStyle(el, charStyle(run.format, zoom));
      Object.assign(el.style, runVerticalStyle(run.format, line, zoom));
      markComments(el, run.format);
      if (run.format.link) el.dataset.link = run.format.link;
    }
    el.classList.add('run', `run-${run.kind}`);
    el.style.left = px(run.x);
    el.style.width = px(run.width);
    return el;
  }

  // -------------------------------------------------------------------------
  // Vista a scorrimento (senza impaginazione)

  private renderFlowSection(doc: DocumentModel, section: Section, vs: ViewState): HTMLElement {
    const px = (pt: number) => `${ptToPx(pt, vs.zoom)}px`;
    const { margins } = section.page;
    const page = document.createElement('div');
    page.className = 'page page-flow';
    page.dataset.section = section.id;
    page.style.width = px(section.page.width);
    page.style.minHeight = px(section.page.height);
    page.style.padding = `${px(margins.top)} ${px(margins.right)} ${px(margins.bottom)} ${px(margins.left)}`;

    const sel = selectionByParagraph(doc, this.editor.state.selection);
    const part = (cls: string, blocks: readonly Block[] | undefined, top: boolean) => {
      const div = document.createElement('div');
      div.className = `page-area ${cls}`;
      div.style.left = px(margins.left);
      div.style.right = px(margins.right);
      if (top) div.style.top = px(margins.header);
      else div.style.bottom = px(margins.footer);
      for (const b of blocks ?? []) div.append(this.renderBlock(doc, b, vs, sel));
      return div;
    };
    const body = document.createElement('div');
    body.className = 'page-area page-body-area';
    for (const b of section.blocks) body.append(this.renderBlock(doc, b, vs, sel));
    page.append(part('page-header-area', section.header?.default, true), body, part('page-footer-area', section.footer?.default, false));
    return page;
  }

  private renderBlock(doc: DocumentModel, block: Block, vs: ViewState, sel: Map<string, [number, number]>): HTMLElement {
    if (block.type === 'paragraph') return this.renderParagraph(doc, block, vs, sel);
    const table = document.createElement('table');
    table.className = 'doc-table';
    table.dataset.id = block.id;
    for (const row of block.rows) {
      const tr = document.createElement('tr');
      for (const cell of row.cells) {
        if (cell.vMerge === 'continue') continue;
        const td = document.createElement('td');
        if (cell.colSpan && cell.colSpan > 1) td.colSpan = cell.colSpan;
        if (cell.format.shading) td.style.background = cell.format.shading;
        for (const b of cell.blocks) td.append(this.renderBlock(doc, b, vs, sel));
        tr.append(td);
      }
      table.append(tr);
    }
    return table;
  }

  private renderParagraph(doc: DocumentModel, p: Paragraph, vs: ViewState, sel: Map<string, [number, number]>): HTMLElement {
    const px = (pt: number) => `${ptToPx(pt, vs.zoom)}px`;
    const sheet = doc.styles;
    const pf = resolveParaFormat(sheet, p);
    const el = document.createElement('p');
    el.className = 'para';
    el.dataset.pid = p.id;
    if (pf.align) el.style.textAlign = pf.align;
    el.style.marginTop = px(pf.spaceBefore ?? 0);
    el.style.marginBottom = px(pf.spaceAfter ?? 0);
    if (pf.indentLeft) el.style.paddingLeft = px(pf.indentLeft);
    if (pf.indentRight) el.style.paddingRight = px(pf.indentRight);
    if (pf.indentFirstLine) el.style.textIndent = px(pf.indentFirstLine);
    if (pf.shading) el.style.background = pf.shading;
    for (const side of ['top', 'right', 'bottom', 'left'] as const) {
      const b = pf.borders?.[side];
      if (!b || b.style === 'none' || b.width <= 0) continue;
      el.style.setProperty(`border-${side}`, `${px(b.width)} ${b.style === 'single' ? 'solid' : b.style} ${b.color}`);
      if (side === 'top' || side === 'bottom') el.style.setProperty(`padding-${side}`, px(1));
    }
    if (pf.lineSpacing?.rule === 'multiple') el.style.lineHeight = String(pf.lineSpacing.value * 1.15);
    else if (pf.lineSpacing) el.style.lineHeight = px(pf.lineSpacing.value);
    applyStyle(el, charStyle(resolveCharFormat(sheet, p), vs.zoom));

    const { focus } = this.editor.state.selection;
    const caretAt = focus.paragraphId === p.id && vs.headerFooter === null ? focus.offset : -1;
    const [selStart, selEnd] = sel.get(p.id) ?? [-1, -1];
    const marks = this.marks.get(p.id) ?? [];
    const markCuts = marks.flatMap((m) => [m.start, m.end]);
    const cuts = (start: number, end: number) =>
      [...new Set([start, caretAt, selStart, selEnd, ...markCuts, end])].filter((n) => n >= start && n <= end).sort((a, b) => a - b);
    const selected = (a: number, b: number) => selStart >= 0 && a >= selStart && b <= selEnd;

    let pos = 0;
    const pushCaret = (at: number) => {
      if (at === caretAt && !this.caretEl) {
        const c = document.createElement('span');
        c.className = 'caret caret-inline';
        el.append(c);
        this.caretEl = c;
      }
    };
    for (const inline of p.children) {
      const len = inlineLength(inline);
      const points = cuts(pos, pos + len);
      for (let i = 0; i < points.length - 1; i++) {
        pushCaret(points[i]!);
        const a = points[i]!;
        const b = points[i + 1]!;
        const span = this.renderInline(sheet, p, inline, a - pos, b - pos, vs);
        span.dataset.pid = p.id;
        span.dataset.start = String(a);
        if (selected(a, b)) span.classList.add('selected');
        if (inline.type === 'text') for (const m of marks) if (m.start <= a && b <= m.end) applyMark(span, m);
        el.append(span);
      }
      pos += len;
    }
    pushCaret(pos);
    const mark = document.createElement('span');
    mark.className = 'para-mark';
    mark.dataset.pid = p.id;
    mark.dataset.start = String(pos);
    mark.textContent = vs.formattingMarks ? '¶' : '​';
    if (selected(pos, pos + 1)) mark.classList.add('selected');
    el.append(mark);
    return el;
  }

  private renderInline(sheet: DocumentModel['styles'], p: Paragraph, inline: Inline, from: number, to: number, vs: ViewState): HTMLElement {
    switch (inline.type) {
      case 'text': {
        const span = document.createElement('span');
        span.className = 'run run-text';
        const text = inline.text.slice(from, to);
        span.textContent = vs.formattingMarks ? text.replace(/ /g, '·') : text;
        applyStyle(span, charStyle(resolveCharFormat(sheet, p, inline), vs.zoom));
        if (inline.format.link) {
          span.classList.add('run-link');
          span.dataset.link = inline.format.link;
        }
        markComments(span, inline.format);
        return span;
      }
      case 'image': {
        const img = document.createElement('img');
        img.className = 'run run-image';
        img.src = this.editor.state.doc.media[inline.src] ?? inline.src;
        img.alt = inline.alt ?? '';
        img.style.width = `${ptToPx(inline.width, vs.zoom)}px`;
        img.style.height = `${ptToPx(inline.height, vs.zoom)}px`;
        return img;
      }
      case 'tab': {
        const span = document.createElement('span');
        span.className = 'run run-tab';
        span.textContent = vs.formattingMarks ? '→\t' : '\t';
        return span;
      }
      case 'break': {
        const span = document.createElement('span');
        span.className = `run run-break run-break-${inline.kind}`;
        if (vs.formattingMarks) span.textContent = inline.kind === 'line' ? '↵' : '···Interruzione di pagina···';
        span.append(document.createElement('br'));
        return span;
      }
      case 'field': {
        const span = document.createElement('span');
        span.className = 'run run-field';
        span.textContent = FIELD_PLACEHOLDER[inline.kind] ?? '';
        applyStyle(span, charStyle(resolveCharFormat(sheet, p, inline.format), vs.zoom));
        return span;
      }
      case 'footnoteRef': {
        // Vista senza impaginazione: solo il numero in apice.
        const span = document.createElement('span');
        span.className = 'run run-footnote-ref';
        span.textContent = String(footnoteNumbers(this.editor.state.doc).get(inline.noteId) ?? '');
        applyStyle(span, charStyle({ ...resolveCharFormat(sheet, p, inline.format), verticalAlign: 'superscript' }, vs.zoom));
        return span;
      }
    }
  }

  // -------------------------------------------------------------------------
  // Mouse

  private onMouseDown(e: MouseEvent): void {
    if (e.button !== 0) return;
    // In modalità intestazione/piè di pagina il corpo non è modificabile.
    if (this.view.state.headerFooter !== null) return;
    e.preventDefault();
    this.dragging = true;
    this.select(e, e.shiftKey);
  }

  private onDoubleClick(e: MouseEvent): void {
    const target = e.target as HTMLElement | null;
    const area = target?.closest('.page-header-area, .page-footer-area');
    if (area) {
      const part = area.classList.contains('page-header-area') ? 'header' : 'footer';
      runFirstExecutable(this.editor, [{ commandId: 'page.headerFooter', args: { part, edit: true } }]);
    } else if (this.view.state.headerFooter && target?.closest('.page-body-area')) {
      runFirstExecutable(this.editor, [{ commandId: 'page.headerFooter', args: { part: this.view.state.headerFooter, edit: false } }]);
    }
  }

  private select(e: MouseEvent, extend: boolean): void {
    const pos = this.positionAt(e.clientX, e.clientY, e.target as Node | null);
    if (!pos) return;
    const { selection } = this.editor.state;
    const next: Selection = extend ? { anchor: selection.anchor, focus: pos } : { anchor: pos, focus: pos };
    if (next.anchor.paragraphId === selection.anchor.paragraphId && next.anchor.offset === selection.anchor.offset && next.focus.paragraphId === selection.focus.paragraphId && next.focus.offset === selection.focus.offset) return;
    runFirstExecutable(this.editor, [{ commandId: 'edit.setSelection', args: { selection: next } }]);
  }

  /** Posizione del modello sotto un punto dello schermo. */
  positionAt(clientX: number, clientY: number, target: Node | null = null): Position | null {
    const engine = this.editor.services.layout;
    if (engine && this.layoutCache) {
      const pageEl = (target instanceof Element ? target : target?.parentElement)?.closest<HTMLElement>('.page-laid-out')
        ?? this.pageElementAt(clientY);
      if (!pageEl) return null;
      const r = pageEl.getBoundingClientRect();
      const zoom = this.view.state.zoom;
      return engine.hitTest(this.layoutCache.layout, Number(pageEl.dataset.page), pxToPt(clientX - r.left, zoom), pxToPt(clientY - r.top, zoom));
    }
    return this.flowPositionAt(clientX, clientY, target);
  }

  private pageElementAt(clientY: number): HTMLElement | null {
    for (const el of Array.from(this.pages.querySelectorAll<HTMLElement>('.page-laid-out'))) {
      const r = el.getBoundingClientRect();
      if (clientY >= r.top && clientY <= r.bottom) return el;
    }
    return null;
  }

  private flowPositionAt(clientX: number, clientY: number, target: Node | null): Position | null {
    const doc = document as Document & {
      caretPositionFromPoint?: (x: number, y: number) => { offsetNode: Node; offset: number } | null;
      caretRangeFromPoint?: (x: number, y: number) => Range | null;
    };
    let node: Node | null = null;
    let offset = 0;
    const cp = doc.caretPositionFromPoint?.(clientX, clientY);
    if (cp) {
      node = cp.offsetNode;
      offset = cp.offset;
    } else {
      const range = doc.caretRangeFromPoint?.(clientX, clientY);
      if (range) {
        node = range.startContainer;
        offset = range.startOffset;
      }
    }
    node ??= target;
    const el = node instanceof Element ? node : node?.parentElement;
    const span = el?.closest<HTMLElement>('[data-start]');
    if (span?.dataset.pid) {
      const start = Number(span.dataset.start);
      if (span.classList.contains('para-mark')) return { paragraphId: span.dataset.pid, offset: start };
      if (span.classList.contains('run-text') && node?.nodeType === Node.TEXT_NODE) return { paragraphId: span.dataset.pid, offset: start + offset };
      // Inline non testuale: prima o dopo a seconda della metà cliccata.
      const r = span.getBoundingClientRect();
      return { paragraphId: span.dataset.pid, offset: start + (clientX > r.left + r.width / 2 ? 1 : 0) };
    }
    // Clic nel margine o fuori dal testo: fine del paragrafo cliccato.
    const para = el?.closest<HTMLElement>('.para');
    if (para?.dataset.pid) {
      const p = [...paragraphs(this.editor.state.doc)].find((l) => l.paragraph.id === para.dataset.pid);
      if (p) return { paragraphId: p.paragraph.id, offset: paragraphLength(p.paragraph) };
    }
    return null;
  }
}
