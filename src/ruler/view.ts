/**
 * Il righello orizzontale sopra le pagine, come in Word.
 *
 * - Zona bianca fra i margini, grigia fuori; numeri in centimetri contati
 *   dal margine sinistro.
 * - Marcatori del paragrafo con il cursore: rientro prima riga (triangolo in
 *   alto), rientro sporgente (triangolo in basso), rientro sinistro
 *   (rettangolino, sposta entrambi) e rientro destro.
 * - Tabulazioni: clic sul righello per aggiungerne una del tipo scelto nel
 *   quadratino a sinistra; trascinandola si sposta, trascinandola fuori dal
 *   righello si toglie. I segnetti grigi sotto il righello sono le
 *   tabulazioni predefinite.
 * - I bordi fra grigio e bianco si trascinano per cambiare i margini.
 *
 * Durante il trascinamento una linea tratteggiata attraversa il documento;
 * la modifica entra nel documento (e in Annulla) solo al rilascio.
 */

import {
  cmToPt,
  findParagraph,
  ptToPx,
  pxToPt,
  type DocumentLayout,
  type Editor,
  type PageLayout,
  type Points,
  type TabAlignment,
  type TabStop,
} from '../core';
import { effectiveIndents, effectiveTabs, snap, TAB_LABELS, TAB_TYPES, type Indents } from './model';

/** Intervallo delle tabulazioni predefinite usato dall'impaginazione (½ pollice). */
const DEFAULT_TAB: Points = 36;
/** Spazio minimo fra rientro sinistro e destro. */
const MIN_TEXT: Points = cmToPt(0.5);
/** Oltre questa distanza verticale (px) una tabulazione trascinata viene tolta. */
const REMOVE_DISTANCE = 24;

const SVG_NS = 'http://www.w3.org/2000/svg';

function svg(width: number, height: number, path: string): SVGSVGElement {
  const el = document.createElementNS(SVG_NS, 'svg');
  el.setAttribute('width', String(width));
  el.setAttribute('height', String(height));
  el.setAttribute('viewBox', `0 0 ${width} ${height}`);
  const p = document.createElementNS(SVG_NS, 'path');
  p.setAttribute('d', path);
  el.append(p);
  return el;
}

/** Disegni dei marcatori (larghezza 9 o 11 px, punta sulla posizione). */
const SHAPES = {
  first: () => svg(9, 7, 'M0.5 0.5H8.5L4.5 6.5Z'),
  hanging: () => svg(9, 7, 'M4.5 0.5L8.5 6.5H0.5Z'),
  left: () => svg(9, 5, 'M0.5 0.5H8.5V4.5H0.5Z'),
  right: () => svg(9, 7, 'M4.5 0.5L8.5 6.5H0.5Z'),
};

const TAB_SHAPES: Record<TabAlignment, string> = {
  left: 'M1.5 0.5V6.5H7.5',
  center: 'M4.5 0.5V6.5M0.5 6.5H8.5',
  right: 'M7.5 0.5V6.5H1.5',
  decimal: 'M4.5 0.5V6.5M0.5 6.5H8.5M7 3.5h1',
};

function tabGlyph(align: TabAlignment): SVGSVGElement {
  const el = svg(9, 8, TAB_SHAPES[align]);
  el.classList.add('ruler-tab-glyph');
  return el;
}

type DragKind = 'first' | 'hanging' | 'left' | 'right' | 'tab' | 'newTab' | 'marginLeft' | 'marginRight';

interface Drag {
  readonly kind: DragKind;
  readonly startIndents: Indents;
  /** Tabulazione trascinata (posizione originale; per una nuova, il tipo). */
  readonly tab?: TabStop;
  /** Posizione attuale in punti dal margine sinistro (dal bordo pagina per i margini). */
  x: Points;
  removing: boolean;
  moved: boolean;
}

export interface RulerHost {
  /** Tipo di tabulazione scelto nel quadratino. */
  tabType: TabAlignment;
}

export class RulerView {
  readonly element: HTMLElement;
  private readonly track: HTMLElement;
  private readonly strip: HTMLElement;
  private readonly body: HTMLElement;
  private readonly ticks: HTMLElement;
  private readonly marks: HTMLElement;
  private readonly typeButton: HTMLButtonElement;
  private guide: HTMLElement | null = null;
  private scroller: HTMLElement | null = null;
  private pageEl: HTMLElement | null = null;
  private page: PageLayout | null = null;
  private zoom = 1;
  private ticksKey = '';
  private drag: Drag | null = null;
  private readonly cleanup: (() => void)[] = [];

  constructor(
    private readonly editor: Editor,
    private readonly host: RulerHost,
  ) {
    this.element = document.createElement('div');
    this.element.className = 'ruler-bar';
    this.element.setAttribute('aria-label', 'Righello');

    this.typeButton = document.createElement('button');
    this.typeButton.type = 'button';
    this.typeButton.className = 'ruler-tab-type';
    this.typeButton.addEventListener('mousedown', (e) => e.preventDefault());
    this.typeButton.addEventListener('click', () => {
      const i = TAB_TYPES.indexOf(this.host.tabType);
      this.host.tabType = TAB_TYPES[(i + 1) % TAB_TYPES.length]!;
      this.renderTypeButton();
    });

    this.track = document.createElement('div');
    this.track.className = 'ruler-track';
    this.strip = document.createElement('div');
    this.strip.className = 'ruler-strip';
    this.body = document.createElement('div');
    this.body.className = 'ruler-body';
    this.ticks = document.createElement('div');
    this.ticks.className = 'ruler-ticks';
    this.marks = document.createElement('div');
    this.marks.className = 'ruler-marks';
    this.strip.append(this.body, this.ticks, this.marks);
    this.track.append(this.strip);
    this.element.append(this.typeButton, this.track);
    this.renderTypeButton();

    this.strip.addEventListener('mousedown', (e) => this.onStripDown(e));
    const onMove = (e: MouseEvent) => this.onMove(e);
    const onUp = (e: MouseEvent) => this.onUp(e);
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
    const onResize = () => this.place();
    globalThis.addEventListener?.('resize', onResize);
    this.cleanup.push(
      () => document.removeEventListener('mousemove', onMove),
      () => document.removeEventListener('mouseup', onUp),
      () => globalThis.removeEventListener?.('resize', onResize),
      editor.events.on('change', () => this.renderMarks()),
    );
  }

  /** Mette il righello sopra l'area del documento `scroller`. */
  attach(scroller: HTMLElement): void {
    this.scroller = scroller;
    // Il righello va nella colonna dell'app, non dentro eventuali righe (es. riquadro commenti).
    let anchor: HTMLElement = scroller;
    while (anchor.parentElement && !anchor.parentElement.classList.contains('app')) anchor = anchor.parentElement;
    anchor.before(this.element);
    const onScroll = () => this.place();
    scroller.addEventListener('scroll', onScroll);
    this.cleanup.push(() => scroller.removeEventListener('scroll', onScroll));
  }

  destroy(): void {
    for (const c of this.cleanup.splice(0)) c();
    this.guide?.remove();
    this.element.remove();
  }

  setVisible(visible: boolean): void {
    this.element.hidden = !visible;
    if (visible) this.place();
  }

  /** Chiamato a ogni disegno delle pagine. */
  update(pages: readonly HTMLElement[], layout: DocumentLayout, zoom: number): void {
    const engine = this.editor.services.layout;
    const caret = engine?.caretRect(layout, this.editor.state.selection.focus);
    const index = caret && layout.pages[caret.page] ? caret.page : 0;
    this.page = layout.pages[index] ?? null;
    this.pageEl = pages[index] ?? null;
    this.zoom = zoom;
    this.renderTicks();
    this.renderMarks();
    this.place();
  }

  // -------------------------------------------------------------------------
  // Geometria

  private px(pt: Points): number {
    return ptToPx(pt, this.zoom);
  }

  /** Allinea il righello alla pagina con il cursore. */
  private place(): void {
    if (!this.pageEl || this.element.hidden) return;
    const left = this.pageEl.getBoundingClientRect().left - this.track.getBoundingClientRect().left;
    this.strip.style.transform = `translateX(${Math.round(left)}px)`;
  }

  /** Punti dal bordo sinistro della pagina per una coordinata dello schermo. */
  private pagePt(clientX: number): Points {
    return pxToPt(clientX - this.strip.getBoundingClientRect().left, this.zoom);
  }

  private renderTypeButton(): void {
    const t = this.host.tabType;
    this.typeButton.replaceChildren(tabGlyph(t));
    this.typeButton.title = `${TAB_LABELS[t]}: clic sul righello per inserirla, clic qui per cambiare tipo`;
    this.typeButton.dataset.type = t;
  }

  private renderTicks(): void {
    const page = this.page;
    if (!page) return;
    const c = page.contentRect;
    const key = `${page.width}|${c.x}|${c.width}|${this.zoom}`;
    if (key === this.ticksKey) return;
    this.ticksKey = key;
    this.strip.style.width = `${this.px(page.width)}px`;
    this.body.replaceChildren();
    const white = document.createElement('div');
    white.className = 'ruler-text-area';
    Object.assign(white.style, { left: `${this.px(c.x)}px`, width: `${this.px(c.width)}px` });
    this.body.append(white);

    this.ticks.replaceChildren();
    const quarter = cmToPt(0.25);
    const from = Math.ceil(-c.x / quarter);
    const to = Math.floor((page.width - c.x) / quarter);
    for (let q = from; q <= to; q++) {
      if (q === 0) continue;
      const x = this.px(c.x + q * quarter);
      if (q % 4 === 0) {
        const label = document.createElement('span');
        label.className = 'ruler-number';
        label.textContent = String(Math.abs(q / 4));
        label.style.left = `${x}px`;
        this.ticks.append(label);
      } else {
        const tick = document.createElement('span');
        tick.className = q % 2 === 0 ? 'ruler-tick ruler-tick-half' : 'ruler-tick';
        tick.style.left = `${x}px`;
        this.ticks.append(tick);
      }
    }
  }

  /** Paragrafo con il cursore, se il righello può modificarlo (non in tabella né intestazione). */
  private target() {
    const loc = findParagraph(this.editor.state.doc, this.editor.state.selection.focus.paragraphId);
    return loc && loc.path.length === 1 ? loc.paragraph : undefined;
  }

  private renderMarks(): void {
    const page = this.page;
    this.marks.replaceChildren();
    if (!page) return;
    const c = page.contentRect;
    const doc = this.editor.state.doc;
    const at = (el: HTMLElement, pt: Points) => {
      el.style.left = `${this.px(pt)}px`;
      this.marks.append(el);
      return el;
    };

    // Margini trascinabili (sotto i marcatori).
    for (const side of ['Left', 'Right'] as const) {
      const h = document.createElement('div');
      h.className = `ruler-margin-handle ruler-margin-${side.toLowerCase()}`;
      h.title = side === 'Left' ? 'Margine sinistro' : 'Margine destro';
      const x = this.drag?.kind === `margin${side}` ? this.drag.x : side === 'Left' ? c.x : c.x + c.width;
      at(h, x).addEventListener('mousedown', (e) => this.start(e, `margin${side}`));
    }

    const p = this.target();
    if (!p) return;
    let ind = effectiveIndents(doc, p);
    let tabs = effectiveTabs(doc, p);
    const d = this.drag;
    if (d) {
      ind = this.previewIndents(d, c.width);
      if (d.kind === 'tab' || d.kind === 'newTab') {
        tabs = tabs.filter((t) => d.kind === 'newTab' || Math.abs(t.position - d.tab!.position) > 0.5);
      }
    }

    // Tabulazioni predefinite: dopo l'ultima personalizzata e dopo il rientro sinistro.
    const lastCustom = Math.max(ind.left, tabs.length ? tabs[tabs.length - 1]!.position : 0);
    for (let x = (Math.floor(lastCustom / DEFAULT_TAB + 0.01) + 1) * DEFAULT_TAB; x < c.width - ind.right; x += DEFAULT_TAB) {
      const t = document.createElement('div');
      t.className = 'ruler-default-tab';
      at(t, c.x + x);
    }

    for (const tab of tabs) {
      const el = document.createElement('div');
      el.className = `ruler-tab ruler-tab-${tab.align}`;
      el.dataset.position = String(tab.position);
      el.title = TAB_LABELS[tab.align]!;
      el.append(tabGlyph(tab.align));
      at(el, c.x + tab.position).addEventListener('mousedown', (e) => this.start(e, 'tab', tab));
    }
    if (d && (d.kind === 'tab' || d.kind === 'newTab')) {
      const el = document.createElement('div');
      el.className = `ruler-tab ruler-tab-${d.tab!.align} dragging`;
      el.classList.toggle('removing', d.removing);
      el.append(tabGlyph(d.tab!.align));
      at(el, c.x + d.x);
    }

    const marker = (kind: 'first' | 'hanging' | 'left' | 'right', pt: Points, title: string) => {
      const el = document.createElement('div');
      el.className = `ruler-marker ruler-${kind}`;
      el.title = title;
      el.append(SHAPES[kind]());
      at(el, c.x + pt).addEventListener('mousedown', (e) => this.start(e, kind));
    };
    marker('first', ind.left + ind.firstLine, 'Rientro prima riga');
    marker('hanging', ind.left, 'Rientro sporgente');
    marker('left', ind.left, 'Rientro sinistro');
    marker('right', c.width - ind.right, 'Rientro destro');
  }

  // -------------------------------------------------------------------------
  // Trascinamento

  private start(e: MouseEvent, kind: DragKind, tab?: TabStop): void {
    if (e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    const page = this.page;
    if (!page) return;
    const p = this.target();
    const startIndents = p ? effectiveIndents(this.editor.state.doc, p) : { left: 0, right: 0, firstLine: 0 };
    const c = page.contentRect;
    const x =
      kind === 'marginLeft' ? c.x
      : kind === 'marginRight' ? c.x + c.width
      : kind === 'first' ? startIndents.left + startIndents.firstLine
      : kind === 'hanging' || kind === 'left' ? startIndents.left
      : kind === 'right' ? c.width - startIndents.right
      : (tab?.position ?? 0);
    this.drag = { kind, startIndents, ...(tab ? { tab } : {}), x, removing: false, moved: false };
    this.showGuide(e.clientX);
    this.renderMarks();
  }

  private onStripDown(e: MouseEvent): void {
    if (e.button !== 0 || this.drag) return;
    e.preventDefault();
    const page = this.page;
    if (!page || !this.target()) return;
    const x = this.pagePt(e.clientX) - page.contentRect.x;
    // Solo nella zona bianca, come in Word.
    if (x <= 0 || x >= page.contentRect.width) return;
    this.start(e, 'newTab', { position: x, align: this.host.tabType });
  }

  private onMove(e: MouseEvent): void {
    const d = this.drag;
    const page = this.page;
    if (!d || !page) return;
    const c = page.contentRect;
    const raw = this.pagePt(e.clientX);
    const prev = d.x;
    // Alt: posizione libera, senza aggancio alla griglia.
    const free = (pt: Points) => (e.altKey ? Math.round(pt * 100) / 100 : snap(pt));
    if (d.kind === 'marginLeft' || d.kind === 'marginRight') {
      // Margini agganciati alla griglia dal bordo pagina.
      d.x = free(raw);
    } else {
      d.x = free(raw - c.x);
    }
    if (d.kind === 'tab' || d.kind === 'newTab') {
      const r = this.element.getBoundingClientRect();
      d.removing = e.clientY > r.bottom + REMOVE_DISTANCE || e.clientY < r.top - REMOVE_DISTANCE;
    }
    if (Math.abs(d.x - prev) > 0.01) d.moved = true;
    this.showGuide(this.strip.getBoundingClientRect().left + this.px((d.kind.startsWith('margin') ? 0 : c.x) + d.x));
    this.renderMarks();
  }

  /** Rientri durante il trascinamento, con i limiti di Word. */
  private previewIndents(d: Drag, width: Points): Indents {
    const s = d.startIndents;
    const page = this.page!;
    const minX = -page.contentRect.x;
    const maxX = page.width - page.contentRect.x;
    const clamp = (v: Points, lo: Points, hi: Points) => Math.min(hi, Math.max(lo, v));
    const textEnd = width - s.right - MIN_TEXT;
    switch (d.kind) {
      case 'first': {
        const x = clamp(d.x, minX, textEnd);
        return { ...s, firstLine: x - s.left };
      }
      case 'hanging': {
        const x = clamp(d.x, minX, textEnd);
        return { ...s, left: x, firstLine: s.left + s.firstLine - x };
      }
      case 'left': {
        const lo = Math.max(minX, minX - s.firstLine);
        const hi = Math.min(textEnd, textEnd - s.firstLine);
        return { ...s, left: clamp(d.x, lo, hi) };
      }
      case 'right': {
        const x = clamp(d.x, Math.max(s.left, s.left + s.firstLine) + MIN_TEXT, maxX);
        return { ...s, right: width - x };
      }
      default:
        return s;
    }
  }

  private onUp(e: MouseEvent): void {
    const d = this.drag;
    if (!d) return;
    this.onMove(e);
    this.drag = null;
    this.hideGuide();
    const page = this.page;
    if (!page) return;
    const c = page.contentRect;
    switch (d.kind) {
      case 'first':
      case 'hanging':
      case 'left':
      case 'right': {
        if (!d.moved) break;
        const ind = this.previewIndents(d, c.width);
        const s = d.startIndents;
        const change = {
          ...(ind.left !== s.left ? { left: ind.left } : {}),
          ...(ind.right !== s.right ? { right: ind.right } : {}),
          ...(ind.firstLine !== s.firstLine ? { firstLine: ind.firstLine } : {}),
        };
        if (Object.keys(change).length) this.editor.execute('ruler.indent', change);
        break;
      }
      case 'newTab': {
        const x = Math.min(c.width, Math.max(0, d.x));
        if (!d.removing) this.editor.execute('ruler.tabs', { add: { ...d.tab!, position: x } });
        break;
      }
      case 'tab': {
        if (d.removing) this.editor.execute('ruler.tabs', { remove: d.tab!.position });
        else if (d.moved) {
          const x = Math.min(c.width, Math.max(0, d.x));
          if (Math.abs(x - d.tab!.position) > 0.01) this.editor.execute('ruler.tabs', { move: { from: d.tab!.position, to: x } });
        }
        break;
      }
      case 'marginLeft':
      case 'marginRight': {
        if (!d.moved) break;
        this.setMargin(d.kind === 'marginLeft' ? 'left' : 'right', d.x);
        break;
      }
    }
    this.renderMarks();
  }

  private setMargin(side: 'left' | 'right', x: Points): void {
    const page = this.page!;
    const section = this.editor.state.doc.sections.find((s) => s.id === page.sectionId);
    if (!section) return;
    const m = section.page.margins;
    const value = side === 'left' ? x : page.width - x;
    const other = side === 'left' ? m.right : m.left;
    const clamped = Math.round(Math.max(0, Math.min(page.width - other - cmToPt(1), value)) * 100) / 100;
    if (Math.abs(clamped - m[side]) < 0.05) return;
    this.editor.execute('page.setup', { page: { margins: { ...m, [side]: clamped } } });
  }

  private showGuide(clientX: number): void {
    const scroller = this.scroller;
    if (!scroller) return;
    if (!this.guide) {
      this.guide = document.createElement('div');
      this.guide.className = 'ruler-guide';
      document.body.append(this.guide);
    }
    const r = scroller.getBoundingClientRect();
    Object.assign(this.guide.style, { left: `${Math.round(clientX)}px`, top: `${r.top}px`, height: `${r.height}px` });
  }

  private hideGuide(): void {
    this.guide?.remove();
    this.guide = null;
  }
}
