/**
 * Forme sulle pagine: disegno, selezione con le maniglie, spostamento e
 * ridimensionamento col mouse, modifica del testo con doppio clic e menu
 * contestuale. Si aggancia alla vista con `decoratePages` (estensione della
 * vista) e, quando l'app è pronta, con `mount` per la tastiera.
 */

import { pxToPt, ptToPx, type DocumentLayout, type Drawing, type Editor, type Id } from '../core';
import { ContextMenu } from '../ui/context-menu';
import type { ContextMenuItem, ViewExtension } from '../ui/extensions';
import { openChartEditor } from './chart-editor';
import type { DrawingChanges } from './commands';
import { drawingList, drawingPage, fromLineEnds, getDrawing, isLine, lineEnds } from './model';
import { applyTextStyle, drawingElement } from './render';
import type { DrawingSelection } from './selection';
import './styles.css';

type Handle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w' | 'start' | 'end';
type Geometry = Pick<Drawing, 'x' | 'y' | 'width' | 'height' | 'flipH' | 'flipV'>;

const BOX_HANDLES: readonly Handle[] = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];
const MIN_SIZE = 4;

interface Drag {
  readonly id: Id;
  readonly handle: Handle | 'move';
  readonly origin: Drawing;
  readonly page: number;
  readonly startX: number;
  readonly startY: number;
  preview: Geometry;
  moved: boolean;
}

/** Nuova geometria trascinando una maniglia di (dx, dy) punti. */
export function resizeGeometry(o: Drawing, handle: Handle, dx: number, dy: number, keepRatio = false): Geometry {
  if (handle === 'start' || handle === 'end') {
    const [a, b] = lineEnds(o);
    return handle === 'start' ? fromLineEnds([a[0] + dx, a[1] + dy], b) : fromLineEnds(a, [b[0] + dx, b[1] + dy]);
  }
  let { x, y, width, height } = o;
  if (handle.includes('w')) {
    x += dx;
    width -= dx;
  }
  if (handle.includes('e')) width += dx;
  if (handle.includes('n')) {
    y += dy;
    height -= dy;
  }
  if (handle.includes('s')) height += dy;
  if (keepRatio && handle.length === 2 && o.width > 0 && o.height > 0) {
    const scale = Math.max(width / o.width, height / o.height);
    const w = o.width * scale;
    const h = o.height * scale;
    if (handle.includes('w')) x = o.x + o.width - w;
    if (handle.includes('n')) y = o.y + o.height - h;
    width = w;
    height = h;
  }
  if (width < MIN_SIZE) {
    if (handle.includes('w')) x -= MIN_SIZE - width;
    width = MIN_SIZE;
  }
  if (height < MIN_SIZE) {
    if (handle.includes('n')) y -= MIN_SIZE - height;
    height = MIN_SIZE;
  }
  return { x, y, width, height, flipH: o.flipH, flipV: o.flipV };
}

export class DrawingView implements ViewExtension {
  readonly id = 'drawings';
  private layout: DocumentLayout | null = null;
  private pages: readonly HTMLElement[] = [];
  private zoom = 1;
  private drag: Drag | null = null;
  private editing: { readonly id: Id; readonly el: HTMLTextAreaElement } | null = null;
  private closeDialog: (() => void) | null = null;
  private readonly listeners = new Set<() => void>();
  private readonly cleanup: (() => void)[] = [];
  /** Riporta il focus al documento (fornito da `mount`). */
  private focusDocument: () => void = () => {};

  constructor(
    private readonly editor: Editor,
    readonly selection: DrawingSelection,
  ) {
    this.cleanup.push(
      selection.onChange(() => this.notify()),
      editor.events.on('docChange', ({ state }) => {
        if (selection.id && !getDrawing(state.doc, selection.id)) selection.select(null);
        if (this.editing && !getDrawing(state.doc, this.editing.id)) this.editing = null;
      }),
    );
  }

  onChange(listener: () => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  private notify(): void {
    for (const l of [...this.listeners]) l();
  }

  get isEditing(): boolean {
    return this.editing !== null;
  }

  destroy(): void {
    this.closeDialog?.();
    for (const c of this.cleanup.splice(0)) c();
    document.removeEventListener('mousemove', this.onMouseMove);
    document.removeEventListener('mouseup', this.onMouseUp);
  }

  // -------------------------------------------------------------------------
  // Disegno

  decoratePages(pages: readonly HTMLElement[], layout: DocumentLayout, zoom: number): void {
    this.pages = pages;
    this.layout = layout;
    this.zoom = zoom;
    for (const d of drawingList(layout.doc)) {
      const page = pages[drawingPage(layout, d)];
      if (page) page.append(this.element(d));
    }
  }

  private element(d: Drawing): HTMLElement {
    const shown = this.drag?.id === d.id ? { ...d, ...this.drag.preview } : d;
    const el = drawingElement(shown, this.zoom);
    if (this.selection.id === d.id) {
      el.classList.add('selected');
      this.addHandles(el, shown);
    }
    if (this.editing?.id === d.id) {
      el.classList.add('editing');
      const text = el.querySelector('.drawing-text');
      applyTextStyle(this.editing.el, shown, this.zoom);
      if (text) text.replaceWith(this.editing.el);
      else el.append(this.editing.el);
      const editor = this.editing.el;
      queueMicrotask(() => editor.isConnected && editor.focus({ preventScroll: true }));
    }
    el.addEventListener('mousedown', (e) => this.onMouseDown(e, d.id));
    el.addEventListener('dblclick', (e) => {
      e.stopPropagation();
      this.open(d.id);
    });
    el.addEventListener('contextmenu', (e) => this.onContextMenu(e, d.id));
    return el;
  }

  private addHandles(el: HTMLElement, d: Geometry & Pick<Drawing, 'kind' | 'shape'>): void {
    const px = (pt: number) => `${ptToPx(pt, this.zoom)}px`;
    if (isLine(d as Drawing)) {
      const [a, b] = lineEnds(d);
      for (const [handle, [x, y]] of [['start', a], ['end', b]] as const) {
        const h = document.createElement('div');
        h.className = `drawing-handle drawing-handle-${handle}`;
        h.dataset.handle = handle;
        h.style.left = px(x - d.x);
        h.style.top = px(y - d.y);
        el.append(h);
      }
      return;
    }
    for (const handle of BOX_HANDLES) {
      const h = document.createElement('div');
      h.className = `drawing-handle drawing-handle-${handle}`;
      h.dataset.handle = handle;
      el.append(h);
    }
  }

  // -------------------------------------------------------------------------
  // Mouse

  private onMouseDown(e: MouseEvent, id: Id): void {
    const target = e.target as HTMLElement | null;
    if (this.editing?.id === id && target === this.editing.el) {
      // Clic dentro il testo in modifica: lo gestisce il campo di testo.
      e.stopPropagation();
      return;
    }
    e.stopPropagation();
    if (e.button !== 0) return;
    e.preventDefault();
    const d = getDrawing(this.editor.state.doc, id);
    if (!d || !this.layout) return;
    if (this.editing && this.editing.id !== id) this.commitEditing();
    const handle = (target?.dataset.handle as Handle | undefined) ?? 'move';
    this.drag = {
      id,
      handle,
      origin: d,
      page: drawingPage(this.layout, d),
      startX: e.clientX,
      startY: e.clientY,
      preview: d,
      moved: false,
    };
    document.addEventListener('mousemove', this.onMouseMove);
    document.addEventListener('mouseup', this.onMouseUp);
    this.selection.select(id);
  }

  private readonly onMouseMove = (e: MouseEvent): void => {
    const drag = this.drag;
    if (!drag) return;
    const dxPx = e.clientX - drag.startX;
    const dyPx = e.clientY - drag.startY;
    if (!drag.moved && Math.hypot(dxPx, dyPx) < 3) return;
    drag.moved = true;
    const dx = pxToPt(dxPx, this.zoom);
    const dy = pxToPt(dyPx, this.zoom);
    const o = drag.origin;
    drag.preview = drag.handle === 'move' ? { ...o, x: o.x + dx, y: o.y + dy } : resizeGeometry(o, drag.handle, dx, dy, e.shiftKey);
    // Solo l'elemento trascinato si ridisegna, non tutta la vista.
    const old = Array.from(this.pages[drag.page]?.querySelectorAll<HTMLElement>('.drawing') ?? []).find((el) => el.dataset.drawing === drag.id);
    if (old) old.replaceWith(this.element(o));
  };

  private readonly onMouseUp = (e: MouseEvent): void => {
    document.removeEventListener('mousemove', this.onMouseMove);
    document.removeEventListener('mouseup', this.onMouseUp);
    const drag = this.drag;
    this.drag = null;
    if (!drag?.moved) return;
    const changes: DrawingChanges = { ...drag.preview };
    if (drag.handle === 'move' && this.layout) {
      // Trascinata su un'altra pagina: si ancora a un paragrafo di quella pagina.
      const target = this.pageAt(e.clientY);
      if (target !== null && target !== drag.page) {
        const rect = this.pages[target]!.getBoundingClientRect();
        const origin = this.pages[drag.page]!.getBoundingClientRect();
        const grabX = pxToPt(drag.startX - origin.left, this.zoom) - drag.origin.x;
        const grabY = pxToPt(drag.startY - origin.top, this.zoom) - drag.origin.y;
        const x = pxToPt(e.clientX - rect.left, this.zoom) - grabX;
        const y = pxToPt(e.clientY - rect.top, this.zoom) - grabY;
        const anchor = this.paragraphOnPage(target, x, y);
        Object.assign(changes, { x, y }, anchor ? { anchor } : {});
      }
    }
    this.editor.execute('drawing.update', { id: drag.id, changes });
  };

  /** Indice della pagina sotto la coordinata verticale dello schermo (la più vicina). */
  private pageAt(clientY: number): number | null {
    let best: number | null = null;
    let bestDist = Infinity;
    this.pages.forEach((el, i) => {
      const r = el.getBoundingClientRect();
      const dist = clientY < r.top ? r.top - clientY : clientY > r.bottom ? clientY - r.bottom : 0;
      if (dist < bestDist) {
        bestDist = dist;
        best = i;
      }
    });
    return best;
  }

  private paragraphOnPage(page: number, x: number, y: number): Id | undefined {
    const layout = this.layout;
    if (!layout) return undefined;
    const hit = this.editor.services.layout?.hitTest(layout, page, x, y);
    return hit?.paragraphId ?? layout.pages[page]?.lines[0]?.paragraphId;
  }

  private onContextMenu(e: MouseEvent, id: Id): void {
    e.preventDefault();
    e.stopPropagation();
    const d = getDrawing(this.editor.state.doc, id);
    if (!d) return;
    this.selection.select(id);
    const run = (commandId: string, args: unknown) => () => this.editor.execute(commandId, args);
    const items: ContextMenuItem[] = [];
    if (d.kind === 'chart') items.push({ label: 'Modifica dati…', run: () => this.open(id), strong: true });
    else if (!isLine(d)) items.push({ label: 'Modifica testo', run: () => this.open(id), strong: true });
    items.push(
      { label: 'Porta in primo piano', run: run('drawing.order', { id, to: 'front' }), separator: true },
      { label: 'Porta in secondo piano', run: run('drawing.order', { id, to: 'back' }) },
      { label: 'Elimina', run: run('drawing.delete', { id }), separator: true },
    );
    new ContextMenu(items, e.clientX, e.clientY, () => this.focusDocument());
  }

  // -------------------------------------------------------------------------
  // Testo e dati

  /** Doppio clic: modifica il testo, o i dati se è un grafico. */
  open(id: Id, initialText?: string): void {
    const d = getDrawing(this.editor.state.doc, id);
    if (!d) return;
    this.selection.select(id);
    if (d.kind === 'chart') {
      if (!d.chart || this.closeDialog) return;
      this.closeDialog = openChartEditor(
        d.chart,
        (chart) => this.editor.execute('drawing.update', { id, changes: { chart } }),
        () => {
          this.closeDialog = null;
          this.focusDocument();
        },
      );
      return;
    }
    if (isLine(d)) return;
    this.startEditing(d, initialText);
  }

  private startEditing(d: Drawing, initialText?: string): void {
    if (this.editing?.id === d.id) return;
    if (this.editing) this.commitEditing();
    const el = document.createElement('textarea');
    el.className = 'drawing-text drawing-editor';
    el.value = (d.text?.text ?? '') + (initialText ?? '');
    el.setAttribute('aria-label', 'Testo della forma');
    el.addEventListener('mousedown', (e) => e.stopPropagation());
    el.addEventListener('dblclick', (e) => e.stopPropagation());
    el.addEventListener('keydown', (e) => {
      e.stopPropagation();
      if (e.key === 'Escape') {
        e.preventDefault();
        this.commitEditing();
      }
    });
    // Il campo esce dal DOM a ogni ridisegno: quella non è una vera uscita.
    el.addEventListener('blur', () => {
      if (el.isConnected && this.editing?.el === el) queueMicrotask(() => !el.matches(':focus') && el.isConnected && this.commitEditing());
    });
    this.editing = { id: d.id, el };
    this.notify();
    el.setSelectionRange(el.value.length, el.value.length);
  }

  /** Salva il testo in modifica (un passo di Annulla) ed esce dalla modifica. */
  commitEditing(): void {
    const editing = this.editing;
    if (!editing) return;
    this.editing = null;
    const d = getDrawing(this.editor.state.doc, editing.id);
    const text = editing.el.value.replace(/\r\n?/g, '\n');
    if (d && text !== (d.text?.text ?? '')) {
      const base = d.text ?? { text: '', format: {} };
      this.editor.execute('drawing.update', { id: d.id, changes: { text: { ...base, text } } });
    } else this.notify();
    this.focusDocument();
  }

  // -------------------------------------------------------------------------
  // Tastiera e app

  /** Tasti con una forma selezionata. Restituisce true se gestito. */
  handleKey(e: Pick<KeyboardEvent, 'key' | 'ctrlKey' | 'metaKey' | 'altKey' | 'shiftKey'>): boolean {
    const id = this.selection.id;
    if (!id || this.editing) return false;
    const d = getDrawing(this.editor.state.doc, id);
    if (!d) return false;
    const step = e.ctrlKey || e.altKey ? 1 : 6;
    const nudge: Record<string, [number, number]> = { ArrowLeft: [-step, 0], ArrowRight: [step, 0], ArrowUp: [0, -step], ArrowDown: [0, step] };
    if (e.key in nudge) {
      const [dx, dy] = nudge[e.key]!;
      this.editor.execute('drawing.update', { id, changes: { x: d.x + dx, y: d.y + dy } });
      return true;
    }
    switch (e.key) {
      case 'Delete':
      case 'Backspace':
        this.editor.execute('drawing.delete', { id });
        return true;
      case 'Escape':
        this.selection.select(null);
        return true;
      case 'Enter':
      case 'F2':
        this.open(id);
        return true;
    }
    // Scrivere con una forma selezionata scrive dentro la forma, come in Word.
    if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && d.kind !== 'chart' && !isLine(d)) {
      this.open(id, e.key);
      return true;
    }
    return false;
  }

  /** Collega tastiera e clic fuori dalle forme all'app (chiamato da `src/ui/main.ts`). */
  mount(app: { readonly view: { readonly element: HTMLElement }; readonly input: { readonly element: HTMLElement; focus(): void } }): () => void {
    this.focusDocument = () => app.input.focus();
    const onKey = (e: KeyboardEvent) => {
      if (e.isComposing || !this.handleKey(e)) return;
      e.preventDefault();
      e.stopImmediatePropagation();
    };
    const onDown = () => {
      // Clic sul testo o sul margine: si esce dalla forma.
      if (this.editing) this.commitEditing();
      this.selection.select(null);
    };
    app.input.element.addEventListener('keydown', onKey, true);
    app.view.element.addEventListener('mousedown', onDown);
    return () => {
      app.input.element.removeEventListener('keydown', onKey, true);
      app.view.element.removeEventListener('mousedown', onDown);
      this.focusDocument = () => {};
    };
  }
}
