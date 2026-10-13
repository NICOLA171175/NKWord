/**
 * Disegno dei grafici: dai dati (`ChartData`) a primitive semplici
 * (rettangoli, percorsi, testi) in coordinate locali. Lo schermo le
 * trasforma in SVG, il PDF in operatori di disegno.
 */

import type { ChartData } from '../core';
import { arcSegments, type Segment } from './geometry';
import { CHART_COLORS } from './model';

export type ChartPrimitive =
  | { readonly t: 'rect'; readonly x: number; readonly y: number; readonly w: number; readonly h: number; readonly fill: string }
  | { readonly t: 'path'; readonly segments: readonly Segment[]; readonly fill: string | null; readonly stroke: string | null; readonly width: number }
  | {
      readonly t: 'text';
      readonly x: number;
      /** Linea di base. */
      readonly y: number;
      readonly text: string;
      readonly size: number;
      readonly color: string;
      readonly anchor: 'start' | 'middle' | 'end';
    };

const TEXT = '#595959';
const GRID = '#d9d9d9';

/** Larghezza approssimata di un testo (Calibri ≈ mezzo em per carattere). */
export const approxWidth = (text: string, size: number): number => text.length * size * 0.5;

export function seriesColor(chart: ChartData, index: number): string {
  return chart.series[index]?.color ?? CHART_COLORS[index % CHART_COLORS.length]!;
}

export function pointColor(index: number): string {
  return CHART_COLORS[index % CHART_COLORS.length]!;
}

/** Passo "rotondo" (1, 2, 5 × 10^n) per circa `count` intervalli. */
export function niceStep(range: number, count = 5): number {
  if (!(range > 0)) return 1;
  const raw = range / count;
  const mag = 10 ** Math.floor(Math.log10(raw));
  const f = raw / mag;
  return (f <= 1 ? 1 : f <= 2 ? 2 : f <= 2.5 ? 2.5 : f <= 5 ? 5 : 10) * mag;
}

export function axisRange(chart: ChartData): { min: number; max: number; step: number } {
  const values = chart.series.flatMap((s) => s.values).filter((v) => Number.isFinite(v));
  let min = Math.min(0, ...values);
  let max = Math.max(0, ...values);
  if (min === max) max = min + 1;
  const step = niceStep(max - min);
  min = Math.floor(min / step) * step;
  max = Math.ceil(max / step) * step;
  return { min, max, step };
}

export function formatNumber(v: number): string {
  const r = Math.round(v * 1000) / 1000;
  return String(r).replace('.', ',');
}

interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
}

/** Legenda in basso, centrata. Restituisce l'altezza usata. */
function legend(out: ChartPrimitive[], entries: readonly { label: string; color: string }[], box: Box): number {
  if (!entries.length) return 0;
  const size = 9;
  const items = entries.map((e) => ({ ...e, w: 10 + approxWidth(e.label, size) + 12 }));
  const total = items.reduce((s, i) => s + i.w, 0);
  let x = box.x + Math.max(0, (box.w - total) / 2);
  const y = box.y + box.h - 4;
  for (const i of items) {
    out.push({ t: 'rect', x, y: y - 7, w: 7, h: 7, fill: i.color });
    out.push({ t: 'text', x: x + 10, y, text: i.label, size, color: TEXT, anchor: 'start' });
    x += i.w;
  }
  return 16;
}

/** Primitive di un grafico di dimensioni w × h. */
export function chartPrimitives(chart: ChartData, w: number, h: number): ChartPrimitive[] {
  const out: ChartPrimitive[] = [];
  const pad = 9;
  const box: Box = { x: pad, y: pad, w: Math.max(0, w - 2 * pad), h: Math.max(0, h - 2 * pad) };
  if (chart.title) {
    out.push({ t: 'text', x: w / 2, y: box.y + 13, text: chart.title, size: 14, color: TEXT, anchor: 'middle' });
    box.y += 22;
    box.h -= 22;
  }
  const n = chart.categories.length;
  if (chart.type === 'pie') {
    const series = chart.series[0];
    const lh = legend(out, chart.categories.map((c, i) => ({ label: c, color: pointColor(i) })), box);
    box.h -= lh + 4;
    const values = (series?.values ?? []).slice(0, n).map((v) => (Number.isFinite(v) && v > 0 ? v : 0));
    const sum = values.reduce((a, b) => a + b, 0);
    const r = Math.max(0, Math.min(box.w, box.h) / 2);
    const cx = box.x + box.w / 2;
    const cy = box.y + box.h / 2;
    if (!(sum > 0) || r <= 0) return out;
    let a = -Math.PI / 2;
    values.forEach((v, i) => {
      if (!v) return;
      const b = a + (v / sum) * Math.PI * 2;
      const segments: Segment[] =
        v === sum
          ? [{ op: 'M', x: cx + r, y: cy }, ...arcSegments(cx, cy, r, 0, Math.PI * 2), { op: 'Z' }]
          : [{ op: 'M', x: cx, y: cy }, { op: 'L', x: cx + r * Math.cos(a), y: cy + r * Math.sin(a) }, ...arcSegments(cx, cy, r, a, b), { op: 'Z' }];
      out.push({ t: 'path', segments, fill: pointColor(i), stroke: '#ffffff', width: 1 });
      a = b;
    });
    return out;
  }

  const lh = legend(out, chart.series.map((s, i) => ({ label: s.name, color: seriesColor(chart, i) })), box);
  box.h -= lh + 4;
  const { min, max, step } = axisRange(chart);
  const ticks: number[] = [];
  for (let v = min; v <= max + step / 1000; v += step) ticks.push(Math.round(v / step) * step);
  const size = 9;
  const horizontal = chart.type === 'bar';
  // Etichette dei valori a sinistra (o in basso per le barre orizzontali), categorie sull'altro asse.
  const valueLabelW = Math.max(...ticks.map((t) => approxWidth(formatNumber(t), size))) + 6;
  const catLabelW = horizontal ? Math.max(0, ...chart.categories.map((c) => approxWidth(c, size))) + 6 : 0;
  const plot: Box = horizontal
    ? { x: box.x + catLabelW, y: box.y, w: Math.max(1, box.w - catLabelW - 6), h: Math.max(1, box.h - 14) }
    : { x: box.x + valueLabelW, y: box.y, w: Math.max(1, box.w - valueLabelW), h: Math.max(1, box.h - 14) };
  const span = max - min || 1;

  for (const t of ticks) {
    if (horizontal) {
      const x = plot.x + ((t - min) / span) * plot.w;
      out.push({ t: 'path', segments: [{ op: 'M', x, y: plot.y }, { op: 'L', x, y: plot.y + plot.h }], fill: null, stroke: GRID, width: 0.75 });
      out.push({ t: 'text', x, y: plot.y + plot.h + 11, text: formatNumber(t), size, color: TEXT, anchor: 'middle' });
    } else {
      const y = plot.y + plot.h - ((t - min) / span) * plot.h;
      out.push({ t: 'path', segments: [{ op: 'M', x: plot.x, y }, { op: 'L', x: plot.x + plot.w, y }], fill: null, stroke: GRID, width: 0.75 });
      out.push({ t: 'text', x: plot.x - 4, y: y + 3, text: formatNumber(t), size, color: TEXT, anchor: 'end' });
    }
  }
  if (!n) return out;

  const slot = (horizontal ? plot.h : plot.w) / n;
  chart.categories.forEach((c, i) => {
    if (horizontal) {
      // Come in Excel la prima categoria è in basso.
      const y = plot.y + plot.h - (i + 0.5) * slot;
      out.push({ t: 'text', x: plot.x - 4, y: y + 3, text: c, size, color: TEXT, anchor: 'end' });
    } else {
      out.push({ t: 'text', x: plot.x + (i + 0.5) * slot, y: plot.y + plot.h + 11, text: c, size, color: TEXT, anchor: 'middle' });
    }
  });

  const zero = (v: number) => (horizontal ? plot.x + ((v - min) / span) * plot.w : plot.y + plot.h - ((v - min) / span) * plot.h);
  if (chart.type === 'line') {
    chart.series.forEach((s, si) => {
      const segments: Segment[] = [];
      s.values.slice(0, n).forEach((v, i) => {
        if (!Number.isFinite(v)) return;
        const x = plot.x + (i + 0.5) * slot;
        segments.push({ op: segments.length ? 'L' : 'M', x, y: zero(v) });
      });
      if (segments.length) out.push({ t: 'path', segments, fill: null, stroke: seriesColor(chart, si), width: 2.25 });
    });
    return out;
  }

  const count = Math.max(1, chart.series.length);
  const group = slot * 0.65;
  const bar = group / count;
  chart.series.forEach((s, si) => {
    const fill = seriesColor(chart, si);
    s.values.slice(0, n).forEach((v, i) => {
      if (!Number.isFinite(v) || v === 0) return;
      const a = zero(0);
      const b = zero(v);
      const lo = Math.min(a, b);
      const len = Math.abs(b - a);
      if (horizontal) {
        const y = plot.y + plot.h - (i + 1) * slot + (slot - group) / 2 + (count - 1 - si) * bar;
        out.push({ t: 'rect', x: lo, y, w: len, h: bar, fill });
      } else {
        const x = plot.x + i * slot + (slot - group) / 2 + si * bar;
        out.push({ t: 'rect', x, y: lo, w: bar, h: len, fill });
      }
    });
  });
  return out;
}
