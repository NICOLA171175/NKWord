/**
 * Schermata della Modalità lettura: copre la finestra con il documento, una
 * barra in alto (menu Visualizza, titolo, chiudi) e una in basso (pagina e
 * zoom). Con le pagine affiancate si sfoglia con le frecce, la rotellina o
 * i pulsanti ai lati; con lo scorrimento le pagine stanno una sotto l'altra.
 * Esc torna al documento. I colori seguono il tema dell'interfaccia.
 */

import { ptToPx, type DocumentLayout, type Editor } from '../core';
import { renderPage } from '../print/render';
import { drawingList, drawingPage } from '../drawings/model';
import { drawingElement } from '../drawings/render';
import '../drawings/styles.css';
import { TabLeaders } from '../ui/tab-leaders';
import {
  MAX_READ_ZOOM,
  MIN_READ_ZOOM,
  PAGE_COLOR_LABELS,
  PAGE_COLORS,
  READ_LAYOUT_LABELS,
  READ_LAYOUTS,
  READ_ZOOM_STEP,
  type ReadModeModule,
} from './module';
import './styles.css';

export interface ReadModeHost {
  readonly editor: Editor;
  /** Riporta il focus al documento. */
  readonly input?: { focus(): void };
}

/** Le pagine affiancate non sono mai più di tante, anche su schermi larghissimi. */
export const MAX_PAGES_PER_SCREEN = 3;
/** Spazio fra le pagine e ai bordi, in pixel. */
const GAP = 24;
const MARGIN = 24;
/** Spazio ai lati per le frecce, con le pagine affiancate. */
const ARROW_SPACE = 56;
/** Misure di ripiego se la finestra non ha dimensioni (es. nei test). */
const FALLBACK_SIZE = { width: 1280, height: 800 };

function el<K extends keyof HTMLElementTagNameMap>(d: Document, tag: K, className: string, text = ''): HTMLElementTagNameMap[K] {
  const e = d.createElement(tag);
  e.className = className;
  e.textContent = text;
  return e;
}

function button(d: Document, className: string, text: string, title: string): HTMLButtonElement {
  const b = el(d, 'button', className, text);
  b.type = 'button';
  b.title = title;
  b.setAttribute('aria-label', title);
  return b;
}

/** Testo della barra in basso: "Pagina 2 di 5" o "Pagine 3-4 di 5". */
export function pageLabel(first: number, count: number, total: number): string {
  if (total === 0) return 'Nessuna pagina';
  const last = Math.min(total, first + count);
  return last > first + 1 ? `Pagine ${first + 1}-${last} di ${total}` : `Pagina ${first + 1} di ${total}`;
}

/** Quante pagine stanno affiancate in una larghezza, a una data scala. */
export function pagesPerScreen(availableWidth: number, pageWidthPx: number): number {
  if (pageWidthPx <= 0) return 1;
  const n = Math.floor((availableWidth + GAP) / (pageWidthPx + GAP));
  return Math.max(1, Math.min(MAX_PAGES_PER_SCREEN, n));
}

export function mountReadMode(host: ReadModeHost, module: ReadModeModule, d: Document = document): () => void {
  const { editor } = host;

  const root = el(d, 'div', 'readmode');
  root.setAttribute('role', 'dialog');
  root.setAttribute('aria-label', 'Modalità lettura');
  root.tabIndex = -1;
  root.hidden = true;

  // Barra in alto --------------------------------------------------------
  const top = el(d, 'div', 'readmode-top');
  const menuButton = button(d, 'readmode-menu-button', 'Visualizza ▾', 'Opzioni di visualizzazione');
  menuButton.setAttribute('aria-haspopup', 'menu');
  const title = el(d, 'div', 'readmode-title');
  const close = button(d, 'readmode-close', '✕', 'Modifica documento (Esc)');
  top.append(menuButton, title, close);

  const menu = el(d, 'div', 'readmode-menu');
  menu.setAttribute('role', 'menu');
  menu.hidden = true;
  const menuItem = (label: string, action: () => void, checked?: () => boolean, hint = label) => {
    const item = button(d, 'readmode-menu-item', '', hint);
    item.setAttribute('role', checked ? 'menuitemradio' : 'menuitem');
    item.append(el(d, 'span', 'readmode-check'), el(d, 'span', '', label));
    item.addEventListener('click', () => {
      menu.hidden = true;
      action();
    });
    return { item, checked };
  };
  const items = [
    menuItem('Modifica documento', () => void editor.executeAsync('readmode.close')),
    ...READ_LAYOUTS.map((layout) => menuItem(READ_LAYOUT_LABELS[layout], () => void editor.executeAsync('readmode.layout', { layout }), () => module.layout === layout)),
    ...PAGE_COLORS.map((color) => menuItem(PAGE_COLOR_LABELS[color], () => void editor.executeAsync('readmode.pageColor', { color }), () => module.pageColor === color, `Colore pagina: ${PAGE_COLOR_LABELS[color]}`)),
    menuItem('Zoom 100%', () => void editor.executeAsync('readmode.zoom', { zoom: 100 })),
  ];
  const heading = (text: string) => el(d, 'div', 'readmode-menu-heading', text);
  menu.append(
    items[0]!.item,
    heading('Layout'),
    ...items.slice(1, 1 + READ_LAYOUTS.length).map((i) => i.item),
    heading('Colore pagina'),
    ...items.slice(1 + READ_LAYOUTS.length, -1).map((i) => i.item),
    heading('Zoom'),
    items[items.length - 1]!.item,
  );

  // Pagine ---------------------------------------------------------------
  const stage = el(d, 'div', 'readmode-stage');
  const pages = el(d, 'div', 'readmode-pages');
  const prev = button(d, 'readmode-arrow readmode-arrow-prev', '‹', 'Pagina precedente');
  const next = button(d, 'readmode-arrow readmode-arrow-next', '›', 'Pagina successiva');
  stage.append(pages, prev, next);

  // Barra in basso -------------------------------------------------------
  const bottom = el(d, 'div', 'readmode-bottom');
  const status = el(d, 'span', 'readmode-status');
  const zoomOut = button(d, 'readmode-zoom-button', '−', 'Riduci');
  const zoomIn = button(d, 'readmode-zoom-button', '+', 'Ingrandisci');
  const zoomInput = el(d, 'input', 'readmode-zoom');
  zoomInput.type = 'range';
  zoomInput.min = String(MIN_READ_ZOOM);
  zoomInput.max = String(MAX_READ_ZOOM);
  zoomInput.step = String(READ_ZOOM_STEP);
  zoomInput.setAttribute('aria-label', 'Zoom');
  const zoomLabel = el(d, 'span', 'readmode-zoom-label');
  const zoomBox = el(d, 'div', 'readmode-zoom-box');
  zoomBox.append(zoomOut, zoomInput, zoomIn, zoomLabel);
  bottom.append(status, zoomBox);

  root.append(top, menu, stage, bottom);

  // Disegno --------------------------------------------------------------
  let rendered: { layout: DocumentLayout; key: string } | undefined;
  let total = 0;

  const size = () => {
    const width = stage.clientWidth || FALLBACK_SIZE.width;
    const height = stage.clientHeight || FALLBACK_SIZE.height;
    return { width, height };
  };

  /** Scala dei pixel delle pagine: la pagina riempie l'altezza (affiancate) o al massimo la larghezza (scorrimento). */
  const scaleFor = (layout: DocumentLayout) => {
    const first = layout.pages[0];
    if (!first) return 1;
    const { width, height } = size();
    const w = ptToPx(first.width, 1);
    const h = ptToPx(first.height, 1);
    const fit = module.layout === 'columns' ? (height - 2 * MARGIN) / h : Math.min(1, (width - 2 * MARGIN) / w);
    return Math.max(0.1, fit) * (module.zoom / 100);
  };

  const drawPages = (layout: DocumentLayout, indexes: readonly number[], scale: number) => {
    const leaders = new TabLeaders(layout.doc);
    const els = indexes.map((i) => renderPage(layout.pages[i]!, { zoom: scale, doc: layout.doc, leaders, document: d }));
    // Forme e grafici, come nella vista normale (ma senza maniglie: qui non si modificano).
    for (const drawing of drawingList(layout.doc)) {
      const k = indexes.indexOf(drawingPage(layout, drawing));
      if (k >= 0) els[k]!.append(drawingElement(drawing, scale));
    }
    return els;
  };

  const updateStatus = (first: number) => {
    const count = module.layout === 'columns' ? module.perScreen : 1;
    status.textContent = pageLabel(first, count, total);
  };

  const render = () => {
    const layout = module.documentLayout();
    title.textContent = editor.state.doc.metadata.title || 'Documento1';
    root.dataset.layout = module.layout;
    root.dataset.pageColor = module.pageColor;
    zoomInput.value = String(module.zoom);
    zoomLabel.textContent = `${module.zoom}%`;
    zoomOut.disabled = module.zoom <= MIN_READ_ZOOM;
    zoomIn.disabled = module.zoom >= MAX_READ_ZOOM;
    for (const { item, checked } of items) {
      if (!checked) continue;
      item.setAttribute('aria-checked', String(checked()));
      item.firstElementChild!.textContent = checked() ? '✓' : '';
    }
    if (!layout) {
      total = 0;
      pages.replaceChildren(el(d, 'p', 'readmode-empty', "L'impaginazione non è disponibile."));
      updateStatus(0);
      return;
    }
    total = layout.pages.length;
    const scale = scaleFor(layout);
    const first = layout.pages[0];

    if (module.layout === 'columns') {
      const perScreen = first ? pagesPerScreen(size().width - 2 * ARROW_SPACE, ptToPx(first.width, scale)) : 1;
      if (perScreen !== module.perScreen) {
        module.setPerScreen(perScreen); // avvisa e ridisegna
        return;
      }
      const indexes: number[] = [];
      for (let i = module.page; i < Math.min(total, module.page + perScreen); i++) indexes.push(i);
      const key = `c|${module.page}|${perScreen}|${scale}`;
      if (rendered?.layout !== layout || rendered.key !== key) {
        pages.replaceChildren(...drawPages(layout, indexes, scale));
        rendered = { layout, key };
      }
      prev.disabled = module.page <= 0;
      next.disabled = module.page + perScreen >= total;
      updateStatus(module.page);
    } else {
      const key = `s|${scale}`;
      if (rendered?.layout !== layout || rendered.key !== key) {
        const all = layout.pages.map((_, i) => i);
        pages.replaceChildren(...drawPages(layout, all, scale));
        rendered = { layout, key };
        scrollToPage(module.page);
      }
      updateStatus(visiblePage());
    }
  };

  const scrollToPage = (page: number) => {
    const target = pages.children[page] as HTMLElement | undefined;
    if (target) stage.scrollTop = target.offsetTop - MARGIN;
  };

  /** Pagina in cima alla zona visibile, con lo scorrimento. */
  const visiblePage = () => {
    const children = Array.from(pages.children) as HTMLElement[];
    const top = stage.scrollTop + MARGIN;
    const i = children.findIndex((c) => c.offsetTop + c.offsetHeight > top);
    return Math.max(0, i);
  };

  let wasOpen = false;
  const show = () => {
    root.hidden = !module.isOpen;
    if (!module.isOpen) {
      menu.hidden = true;
      rendered = undefined;
      pages.replaceChildren();
      if (wasOpen) host.input?.focus();
      wasOpen = false;
      return;
    }
    if (!wasOpen) {
      stage.scrollTop = 0;
      // Il focus esce dal documento: quello che si scrive non lo modifica.
      root.focus();
    }
    wasOpen = true;
    render();
  };

  // Eventi ---------------------------------------------------------------
  const zoomBy = (delta: number) => void editor.executeAsync('readmode.zoom', { zoom: module.zoom + delta });
  close.addEventListener('click', () => void editor.executeAsync('readmode.close'));
  prev.addEventListener('click', () => void editor.executeAsync('readmode.previous'));
  next.addEventListener('click', () => void editor.executeAsync('readmode.next'));
  zoomOut.addEventListener('click', () => zoomBy(-READ_ZOOM_STEP));
  zoomIn.addEventListener('click', () => zoomBy(READ_ZOOM_STEP));
  zoomInput.addEventListener('input', () => void editor.executeAsync('readmode.zoom', { zoom: Number(zoomInput.value) }));
  menuButton.addEventListener('click', () => {
    menu.hidden = !menu.hidden;
    if (!menu.hidden) (menu.querySelector('button') as HTMLButtonElement | null)?.focus();
  });
  root.addEventListener('mousedown', (e) => {
    if (!menu.hidden && !menu.contains(e.target as Node) && e.target !== menuButton) menu.hidden = true;
  });
  stage.addEventListener('scroll', () => {
    if (module.layout === 'scroll') updateStatus(visiblePage());
  });

  // Rotellina: con le pagine affiancate sfoglia (se la pagina sta tutta nello schermo); con Ctrl cambia lo zoom.
  let lastWheel = 0;
  stage.addEventListener(
    'wheel',
    (e) => {
      if (!module.isOpen) return;
      if (e.ctrlKey || e.metaKey) {
        e.preventDefault();
        zoomBy(e.deltaY < 0 ? READ_ZOOM_STEP : -READ_ZOOM_STEP);
        return;
      }
      if (module.layout !== 'columns' || stage.scrollHeight > stage.clientHeight + 1) return;
      const delta = Math.abs(e.deltaY) >= Math.abs(e.deltaX) ? e.deltaY : e.deltaX;
      if (delta === 0) return;
      e.preventDefault();
      const now = Date.now();
      if (now - lastWheel < 300) return;
      lastWheel = now;
      void editor.executeAsync(delta > 0 ? 'readmode.next' : 'readmode.previous');
    },
    { passive: false },
  );

  const onKey = (e: KeyboardEvent) => {
    if (!module.isOpen) return;
    const handled = () => {
      e.preventDefault();
      e.stopPropagation();
    };
    if (e.key === 'Escape') {
      handled();
      if (!menu.hidden) {
        menu.hidden = true;
        menuButton.focus();
      } else void editor.executeAsync('readmode.close');
      return;
    }
    if (!menu.hidden || e.altKey) return;
    const mod = e.ctrlKey || e.metaKey;
    if (mod && (e.key === '+' || e.key === '=')) return (handled(), zoomBy(READ_ZOOM_STEP));
    if (mod && e.key === '-') return (handled(), zoomBy(-READ_ZOOM_STEP));
    if (mod && e.key === '0') return (handled(), void editor.executeAsync('readmode.zoom', { zoom: 100 }));
    if (mod) return;
    const t = e.target as HTMLElement | null;
    if (t?.tagName === 'INPUT') return;
    if (module.layout === 'columns') {
      if (['ArrowRight', 'ArrowDown', 'PageDown', ' '].includes(e.key)) return (handled(), void editor.executeAsync('readmode.next'));
      if (['ArrowLeft', 'ArrowUp', 'PageUp'].includes(e.key)) return (handled(), void editor.executeAsync('readmode.previous'));
      if (e.key === 'Home') return (handled(), module.goTo(0));
      if (e.key === 'End') return (handled(), module.goTo(total - 1));
    } else {
      if (e.key === 'Home') return (handled(), (stage.scrollTop = 0));
      if (e.key === 'End') return (handled(), (stage.scrollTop = stage.scrollHeight));
    }
    // Nessun altro tasto deve arrivare al documento: in lettura non si scrive.
    if (e.key.length === 1 || e.key === 'Backspace' || e.key === 'Delete' || e.key === 'Enter' || e.key === 'Tab') {
      if (e.key === 'Tab' || t?.tagName === 'BUTTON') return;
      handled();
    }
  };

  const view = d.defaultView;
  const onResize = () => module.isOpen && render();
  const offChange = module.onChange(show);
  d.addEventListener('keydown', onKey, true);
  view?.addEventListener('resize', onResize);
  d.body.append(root);
  show();

  return () => {
    offChange();
    d.removeEventListener('keydown', onKey, true);
    view?.removeEventListener('resize', onResize);
    root.remove();
  };
}
