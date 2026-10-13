/** Operazioni pure sui disegni (`DocumentModel.drawings`). */

import {
  newId,
  paragraphs,
  type ChartData,
  type DocumentLayout,
  type DocumentModel,
  type Drawing,
  type Id,
  type Points,
  type ShapeKind,
} from '../core';

/** Colori del tema di Office usati dalle forme e dai grafici nuovi. */
export const ACCENT = '#4472c4';
export const ACCENT_DARK = '#2f528f';
export const CHART_COLORS = ['#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47', '#264478', '#9e480e'];

export const LINE_SHAPES: ReadonlySet<ShapeKind> = new Set(['line', 'arrow', 'doubleArrow']);

export const SHAPE_LABELS: Record<ShapeKind, string> = {
  rect: 'Rettangolo',
  roundRect: 'Rettangolo arrotondato',
  ellipse: 'Ovale',
  triangle: 'Triangolo isoscele',
  rtTriangle: 'Triangolo rettangolo',
  diamond: 'Rombo',
  pentagon: 'Pentagono',
  hexagon: 'Esagono',
  star5: 'Stella a 5 punte',
  heart: 'Cuore',
  rightArrow: 'Freccia a destra',
  leftArrow: 'Freccia a sinistra',
  upArrow: 'Freccia in su',
  downArrow: 'Freccia in giù',
  wedgeRectCallout: 'Fumetto',
  line: 'Linea',
  arrow: 'Freccia',
  doubleArrow: 'Freccia doppia',
};

export const CHART_LABELS: Record<ChartData['type'], string> = {
  column: 'Istogramma',
  bar: 'Barre',
  line: 'Linee',
  pie: 'Torta',
};

export function isLine(d: Drawing): boolean {
  return d.kind === 'shape' && !!d.shape && LINE_SHAPES.has(d.shape);
}

/** Disegni del documento nell'ordine di sovrapposizione (il primo è il più in basso). */
export function drawingList(doc: DocumentModel): Drawing[] {
  return Object.values(doc.drawings ?? {});
}

export function getDrawing(doc: DocumentModel, id: Id): Drawing | undefined {
  return doc.drawings?.[id];
}

/** Id con una lettera davanti: le chiavi numeriche cambierebbero l'ordine dell'oggetto. */
export function newDrawingId(): Id {
  return `d${newId()}`;
}

export function setDrawing(doc: DocumentModel, d: Drawing): DocumentModel {
  return { ...doc, drawings: { ...doc.drawings, [d.id]: d } };
}

export function removeDrawing(doc: DocumentModel, id: Id): DocumentModel {
  if (!doc.drawings?.[id]) return doc;
  const { [id]: _removed, ...rest } = doc.drawings;
  const { drawings: _d, ...base } = doc;
  return Object.keys(rest).length ? { ...base, drawings: rest } : base;
}

/** Porta il disegno in cima (davanti a tutti) o in fondo alla pila. */
export function reorderDrawing(doc: DocumentModel, id: Id, to: 'front' | 'back'): DocumentModel {
  const d = doc.drawings?.[id];
  if (!d) return doc;
  const others = drawingList(doc).filter((o) => o.id !== id);
  const list = to === 'front' ? [...others, d] : [d, ...others];
  return { ...doc, drawings: Object.fromEntries(list.map((o) => [o.id, o])) };
}

/** Dati di esempio di un grafico nuovo, come quelli che propone Word. */
export function sampleChart(type: ChartData['type']): ChartData {
  const categories = ['Categoria 1', 'Categoria 2', 'Categoria 3', 'Categoria 4'];
  if (type === 'pie') {
    return { type, title: 'Vendite', categories: ['1° trim.', '2° trim.', '3° trim.', '4° trim.'], series: [{ name: 'Vendite', values: [8.2, 3.2, 1.4, 1.2] }] };
  }
  return {
    type,
    title: 'Titolo del grafico',
    categories,
    series: [
      { name: 'Serie 1', values: [4.3, 2.5, 3.5, 4.5] },
      { name: 'Serie 2', values: [2.4, 4.4, 1.8, 2.8] },
      { name: 'Serie 3', values: [2, 2, 3, 5] },
    ],
  };
}

export type NewDrawing =
  | { readonly kind: 'shape'; readonly shape: ShapeKind }
  | { readonly kind: 'textbox' }
  | { readonly kind: 'chart'; readonly type: ChartData['type'] };

/** Disegno nuovo con dimensioni e colori predefiniti di Word, in (x, y) della pagina. */
export function createDrawing(spec: NewDrawing, anchor: Id, x: Points, y: Points): Drawing {
  const id = newDrawingId();
  switch (spec.kind) {
    case 'textbox':
      return {
        id, kind: 'textbox', anchor, x, y, width: 144, height: 72,
        fill: '#ffffff', stroke: '#000000', strokeWidth: 0.75,
        text: { text: '', format: { fontSize: 11, color: '#000000' }, align: 'left', verticalAlign: 'top' },
      };
    case 'chart':
      return { id, kind: 'chart', anchor, x, y, width: 360, height: 216, fill: '#ffffff', stroke: '#d9d9d9', strokeWidth: 0.75, chart: sampleChart(spec.type) };
    case 'shape': {
      if (LINE_SHAPES.has(spec.shape)) {
        return { id, kind: 'shape', shape: spec.shape, anchor, x, y, width: 108, height: 0, fill: null, stroke: ACCENT, strokeWidth: 1.5 };
      }
      const square = ['ellipse', 'star5', 'heart', 'diamond', 'pentagon', 'hexagon', 'triangle', 'rtTriangle'].includes(spec.shape);
      return {
        id, kind: 'shape', shape: spec.shape, anchor, x, y, width: square ? 72 : 108, height: 72,
        fill: ACCENT, stroke: ACCENT_DARK, strokeWidth: 1,
        text: { text: '', format: { fontSize: 11, color: '#ffffff' }, align: 'center', verticalAlign: 'middle' },
      };
    }
  }
}

/** Id dei paragrafi del corpo, delle tabelle e delle note. */
function paragraphIds(doc: DocumentModel): Set<Id> {
  const ids = new Set<Id>();
  for (const { paragraph } of paragraphs(doc)) ids.add(paragraph.id);
  return ids;
}

/**
 * Riancora i disegni il cui paragrafo è sparito (es. unito al precedente)
 * al paragrafo indicato: come in Word, la forma resta nel documento.
 */
export function reanchorOrphans(doc: DocumentModel, fallback: Id): DocumentModel {
  const list = drawingList(doc);
  if (!list.length) return doc;
  const ids = paragraphIds(doc);
  if (list.every((d) => ids.has(d.anchor))) return doc;
  const target = ids.has(fallback) ? fallback : [...ids][0];
  if (!target) return doc;
  return { ...doc, drawings: Object.fromEntries(list.map((d) => [d.id, ids.has(d.anchor) ? d : { ...d, anchor: target }])) };
}

/** Pagina su cui compare il disegno: quella del suo paragrafo (la prima se non si trova). */
export function drawingPage(layout: DocumentLayout, d: Drawing): number {
  const index = pageIndex(layout);
  return index.get(d.anchor) ?? 0;
}

const pageIndexCache = new WeakMap<DocumentLayout, Map<Id, number>>();

/** Prima pagina di ciascun paragrafo (memorizzata per impaginazione). */
export function pageIndex(layout: DocumentLayout): Map<Id, number> {
  let map = pageIndexCache.get(layout);
  if (!map) {
    map = new Map();
    for (const page of layout.pages) {
      for (const lines of [page.lines, page.footnotes ?? []]) {
        for (const l of lines) if (!map.has(l.paragraphId)) map.set(l.paragraphId, page.index);
      }
    }
    pageIndexCache.set(layout, map);
  }
  return map;
}

/** Disegni di una pagina, nell'ordine di sovrapposizione. */
export function drawingsOnPage(layout: DocumentLayout, page: number): Drawing[] {
  return drawingList(layout.doc).filter((d) => drawingPage(layout, d) === page);
}

/** Estremi di una linea, dal rettangolo e dai ribaltamenti. */
export function lineEnds(d: Pick<Drawing, 'x' | 'y' | 'width' | 'height' | 'flipH' | 'flipV'>): [[number, number], [number, number]] {
  const x1 = d.flipH ? d.x + d.width : d.x;
  const x2 = d.flipH ? d.x : d.x + d.width;
  const y1 = d.flipV ? d.y + d.height : d.y;
  const y2 = d.flipV ? d.y : d.y + d.height;
  return [[x1, y1], [x2, y2]];
}

/** Rettangolo e ribaltamenti di una linea dai suoi estremi. */
export function fromLineEnds(a: readonly [number, number], b: readonly [number, number]): Pick<Drawing, 'x' | 'y' | 'width' | 'height' | 'flipH' | 'flipV'> {
  return {
    x: Math.min(a[0], b[0]),
    y: Math.min(a[1], b[1]),
    width: Math.abs(b[0] - a[0]),
    height: Math.abs(b[1] - a[1]),
    flipH: b[0] < a[0],
    flipV: b[1] < a[1],
  };
}
