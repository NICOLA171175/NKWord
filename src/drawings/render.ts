/** Disegno di forme, caselle di testo e grafici nel DOM (SVG più testo HTML). */

import { ptToPx, type Drawing } from '../core';
import { chartPrimitives, type ChartPrimitive } from './chart';
import { drawingPaths, segmentsToSvg } from './geometry';
import { isLine } from './model';

const SVG_NS = 'http://www.w3.org/2000/svg';

/** Margini interni del testo, come in Word (0,1" ai lati, 0,05" sopra e sotto). */
export const TEXT_INSET = { x: 7.2, y: 3.6 } as const;

function svgEl<K extends keyof SVGElementTagNameMap>(name: K, attrs: Record<string, string | number>): SVGElementTagNameMap[K] {
  const el = document.createElementNS(SVG_NS, name);
  for (const [k, v] of Object.entries(attrs)) el.setAttribute(k, String(v));
  return el;
}

function primitiveSvg(p: ChartPrimitive): SVGElement {
  switch (p.t) {
    case 'rect':
      return svgEl('rect', { x: p.x, y: p.y, width: p.w, height: p.h, fill: p.fill });
    case 'path':
      return svgEl('path', { d: segmentsToSvg(p.segments), fill: p.fill ?? 'none', stroke: p.stroke ?? 'none', 'stroke-width': p.width, 'stroke-linejoin': 'round' });
    case 'text': {
      const t = svgEl('text', { x: p.x, y: p.y, 'font-size': p.size, fill: p.color, 'text-anchor': p.anchor, 'font-family': 'Calibri, Carlito, sans-serif' });
      t.textContent = p.text;
      return t;
    }
  }
}

/** SVG della forma o del grafico, con viewBox in punti. */
export function drawingSvg(d: Drawing, zoom: number): SVGSVGElement {
  // Le linee orizzontali o verticali hanno un lato nullo: si allarga la tela per il tratto e le punte.
  const pad = isLine(d) ? 3 * (d.strokeWidth ?? 1) + 6 : (d.strokeWidth ?? 0) / 2 + 1;
  const svg = svgEl('svg', {
    class: 'drawing-svg',
    viewBox: `${-pad} ${-pad} ${d.width + 2 * pad} ${d.height + 2 * pad}`,
    width: ptToPx(d.width + 2 * pad, zoom),
    height: ptToPx(d.height + 2 * pad, zoom),
  });
  svg.style.left = `${-ptToPx(pad, zoom)}px`;
  svg.style.top = `${-ptToPx(pad, zoom)}px`;
  if (d.kind === 'chart') {
    if (d.fill || d.stroke) {
      svg.append(svgEl('rect', { x: 0, y: 0, width: d.width, height: d.height, fill: d.fill ?? 'none', stroke: d.stroke ?? 'none', 'stroke-width': d.strokeWidth ?? 0.75 }));
    }
    if (d.chart) for (const p of chartPrimitives(d.chart, d.width, d.height)) svg.append(primitiveSvg(p));
    return svg;
  }
  for (const p of drawingPaths(d)) {
    svg.append(
      svgEl('path', {
        d: segmentsToSvg(p.segments),
        fill: p.fill ?? 'none',
        stroke: p.stroke ?? 'none',
        'stroke-width': p.strokeWidth,
        'stroke-linejoin': 'round',
        'stroke-linecap': 'round',
      }),
    );
  }
  return svg;
}

/** Stile CSS del testo di una forma. */
export function applyTextStyle(el: HTMLElement, d: Drawing, zoom: number): void {
  const t = d.text;
  const f = t?.format ?? {};
  el.style.fontFamily = f.fontFamily ? `"${f.fontFamily}", Calibri, sans-serif` : 'Calibri, Carlito, sans-serif';
  el.style.fontSize = `${ptToPx(f.fontSize ?? 11, zoom)}px`;
  el.style.color = f.color ?? '#000000';
  if (f.bold) el.style.fontWeight = 'bold';
  if (f.italic) el.style.fontStyle = 'italic';
  el.style.textAlign = t?.align ?? 'left';
  el.style.padding = `${ptToPx(TEXT_INSET.y, zoom)}px ${ptToPx(TEXT_INSET.x, zoom)}px`;
  el.style.justifyContent = t?.verticalAlign === 'middle' ? 'center' : t?.verticalAlign === 'bottom' ? 'flex-end' : 'flex-start';
}

/** Elemento posizionato sulla pagina che contiene il disegno. */
export function drawingElement(d: Drawing, zoom: number): HTMLElement {
  const el = document.createElement('div');
  el.className = `drawing drawing-${d.kind}`;
  if (d.shape) el.classList.add(`drawing-shape-${d.shape}`);
  el.dataset.drawing = d.id;
  placeElement(el, d, zoom);
  el.append(drawingSvg(d, zoom));
  if (isLine(d)) {
    const hit = document.createElement('div');
    hit.className = 'drawing-hit';
    el.append(hit);
  }
  if (d.kind !== 'chart' && !isLine(d)) {
    const text = document.createElement('div');
    text.className = 'drawing-text';
    applyTextStyle(text, d, zoom);
    for (const line of (d.text?.text ?? '').split('\n')) {
      const p = document.createElement('div');
      p.textContent = line || '​';
      text.append(p);
    }
    el.append(text);
  }
  return el;
}

export function placeElement(el: HTMLElement, r: Pick<Drawing, 'x' | 'y' | 'width' | 'height'>, zoom: number): void {
  el.style.left = `${ptToPx(r.x, zoom)}px`;
  el.style.top = `${ptToPx(r.y, zoom)}px`;
  el.style.width = `${ptToPx(r.width, zoom)}px`;
  el.style.height = `${ptToPx(r.height, zoom)}px`;
}
