/**
 * Geometria delle forme in coordinate locali (origine nell'angolo in alto a
 * sinistra del disegno, y verso il basso, in punti). La stessa descrizione
 * serve allo schermo (SVG) e al PDF.
 */

import type { Drawing, ShapeKind } from '../core';

export type Segment =
  | { readonly op: 'M' | 'L'; readonly x: number; readonly y: number }
  | { readonly op: 'C'; readonly x1: number; readonly y1: number; readonly x2: number; readonly y2: number; readonly x: number; readonly y: number }
  | { readonly op: 'Z' };

export interface ShapePath {
  readonly segments: readonly Segment[];
  /** Colore di riempimento (null = vuoto). */
  readonly fill: string | null;
  /** Colore del contorno (null = nessuno). */
  readonly stroke: string | null;
  readonly strokeWidth: number;
}

/** Costante per approssimare un quarto di cerchio con una curva di Bézier. */
const K = 0.5522847498;

const M = (x: number, y: number): Segment => ({ op: 'M', x, y });
const L = (x: number, y: number): Segment => ({ op: 'L', x, y });
const C = (x1: number, y1: number, x2: number, y2: number, x: number, y: number): Segment => ({ op: 'C', x1, y1, x2, y2, x, y });
const Z: Segment = { op: 'Z' };

function polygon(points: readonly (readonly [number, number])[]): Segment[] {
  return [...points.map(([x, y], i) => (i ? L(x, y) : M(x, y))), Z];
}

export function ellipseSegments(cx: number, cy: number, rx: number, ry: number): Segment[] {
  const kx = rx * K;
  const ky = ry * K;
  return [
    M(cx + rx, cy),
    C(cx + rx, cy + ky, cx + kx, cy + ry, cx, cy + ry),
    C(cx - kx, cy + ry, cx - rx, cy + ky, cx - rx, cy),
    C(cx - rx, cy - ky, cx - kx, cy - ry, cx, cy - ry),
    C(cx + kx, cy - ry, cx + rx, cy - ky, cx + rx, cy),
    Z,
  ];
}

function roundRect(w: number, h: number, r: number): Segment[] {
  const k = r * (1 - K);
  return [
    M(r, 0), L(w - r, 0), C(w - k, 0, w, k, w, r),
    L(w, h - r), C(w, h - k, w - k, h, w - r, h),
    L(r, h), C(k, h, 0, h - k, 0, h - r),
    L(0, r), C(0, k, k, 0, r, 0), Z,
  ];
}

/** Arco di cerchio (in senso orario, angoli in radianti da ore 3) come curve di Bézier, senza il punto iniziale. */
export function arcSegments(cx: number, cy: number, r: number, from: number, to: number): Segment[] {
  const out: Segment[] = [];
  const steps = Math.max(1, Math.ceil(Math.abs(to - from) / (Math.PI / 2)));
  const step = (to - from) / steps;
  const k = (4 / 3) * Math.tan(step / 4) * r;
  for (let i = 0; i < steps; i++) {
    const a = from + i * step;
    const b = a + step;
    out.push(
      C(
        cx + r * Math.cos(a) - k * Math.sin(a), cy + r * Math.sin(a) + k * Math.cos(a),
        cx + r * Math.cos(b) + k * Math.sin(b), cy + r * Math.sin(b) - k * Math.cos(b),
        cx + r * Math.cos(b), cy + r * Math.sin(b),
      ),
    );
  }
  return out;
}

/** Freccia piena orientata verso destra in coordinate (lungo, trasversale), poi ruotata. */
function blockArrow(w: number, h: number, dir: 'right' | 'left' | 'up' | 'down'): Segment[] {
  const horizontal = dir === 'right' || dir === 'left';
  const len = horizontal ? w : h;
  const thick = horizontal ? h : w;
  const head = Math.min(len, thick * 0.5);
  const pts: [number, number][] = [
    [0, thick * 0.25], [len - head, thick * 0.25], [len - head, 0], [len, thick / 2],
    [len - head, thick], [len - head, thick * 0.75], [0, thick * 0.75],
  ];
  const map = ([a, c]: [number, number]): [number, number] => {
    switch (dir) {
      case 'right': return [a, c];
      case 'left': return [len - a, c];
      case 'down': return [c, a];
      case 'up': return [c, len - a];
    }
  };
  return polygon(pts.map(map));
}

function star(w: number, h: number): Segment[] {
  const pts: [number, number][] = [];
  for (let i = 0; i < 10; i++) {
    const r = i % 2 ? 0.382 : 1;
    const a = -Math.PI / 2 + (i * Math.PI) / 5;
    pts.push([w / 2 + (w / 2) * r * Math.cos(a), h * 0.525 + (h * 0.525) * r * Math.sin(a)]);
  }
  return polygon(pts);
}

function heart(w: number, h: number): Segment[] {
  return [
    M(w / 2, h * 0.25),
    C(w / 2, 0, 0, 0, 0, h * 0.3),
    C(0, h * 0.6, w / 2, h * 0.8, w / 2, h),
    C(w / 2, h * 0.8, w, h * 0.6, w, h * 0.3),
    C(w, 0, w / 2, 0, w / 2, h * 0.25),
    Z,
  ];
}

/** Contorno di una forma chiusa di dimensioni w × h. */
export function shapeOutline(shape: ShapeKind, w: number, h: number): Segment[] {
  switch (shape) {
    case 'roundRect':
      return roundRect(w, h, Math.min(w, h) * 0.1667);
    case 'ellipse':
      return ellipseSegments(w / 2, h / 2, w / 2, h / 2);
    case 'triangle':
      return polygon([[w / 2, 0], [w, h], [0, h]]);
    case 'rtTriangle':
      return polygon([[0, 0], [w, h], [0, h]]);
    case 'diamond':
      return polygon([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]);
    case 'pentagon':
      return polygon([[w / 2, 0], [w, h * 0.38], [w * 0.81, h], [w * 0.19, h], [0, h * 0.38]]);
    case 'hexagon':
      return polygon([[w * 0.25, 0], [w * 0.75, 0], [w, h / 2], [w * 0.75, h], [w * 0.25, h], [0, h / 2]]);
    case 'star5':
      return star(w, h);
    case 'heart':
      return heart(w, h);
    case 'rightArrow':
      return blockArrow(w, h, 'right');
    case 'leftArrow':
      return blockArrow(w, h, 'left');
    case 'upArrow':
      return blockArrow(w, h, 'up');
    case 'downArrow':
      return blockArrow(w, h, 'down');
    case 'wedgeRectCallout':
      return polygon([[0, 0], [w, 0], [w, h * 0.8], [w * 0.42, h * 0.8], [w * 0.15, h], [w * 0.25, h * 0.8], [0, h * 0.8]]);
    default:
      return polygon([[0, 0], [w, 0], [w, h], [0, h]]);
  }
}

/** Punta di freccia piena in `tip`, che arriva da `from`. */
function arrowHead(tip: readonly [number, number], from: readonly [number, number], width: number): Segment[] {
  const dx = tip[0] - from[0];
  const dy = tip[1] - from[1];
  const len = Math.hypot(dx, dy) || 1;
  const ux = dx / len;
  const uy = dy / len;
  const size = 3 * width + 4;
  const half = 1.5 * width + 2.5;
  const bx = tip[0] - ux * size;
  const by = tip[1] - uy * size;
  return polygon([[tip[0], tip[1]], [bx - uy * half, by + ux * half], [bx + uy * half, by - ux * half]]);
}

/** Percorsi da disegnare per una forma o una casella di testo (non per i grafici). */
export function drawingPaths(d: Drawing): ShapePath[] {
  const stroke = d.stroke === undefined ? null : d.stroke;
  const strokeWidth = d.strokeWidth ?? 0.75;
  const shape = d.kind === 'textbox' ? 'rect' : d.shape ?? 'rect';
  if (shape === 'line' || shape === 'arrow' || shape === 'doubleArrow') {
    const a: [number, number] = [d.flipH ? d.width : 0, d.flipV ? d.height : 0];
    const b: [number, number] = [d.flipH ? 0 : d.width, d.flipV ? 0 : d.height];
    const color = stroke ?? '#000000';
    const out: ShapePath[] = [{ segments: [M(a[0], a[1]), L(b[0], b[1])], fill: null, stroke: color, strokeWidth }];
    if (shape !== 'line') out.push({ segments: arrowHead(b, a, strokeWidth), fill: color, stroke: null, strokeWidth: 0 });
    if (shape === 'doubleArrow') out.push({ segments: arrowHead(a, b, strokeWidth), fill: color, stroke: null, strokeWidth: 0 });
    return out;
  }
  return [{ segments: shapeOutline(shape, d.width, d.height), fill: d.fill ?? null, stroke, strokeWidth }];
}

const n = (v: number) => String(Math.round(v * 100) / 100);

export function segmentsToSvg(segments: readonly Segment[]): string {
  return segments
    .map((s) => (s.op === 'Z' ? 'Z' : s.op === 'C' ? `C${n(s.x1)} ${n(s.y1)} ${n(s.x2)} ${n(s.y2)} ${n(s.x)} ${n(s.y)}` : `${s.op}${n(s.x)} ${n(s.y)}`))
    .join(' ');
}
