/**
 * Forme, caselle di testo e grafici nel PDF: operatori di disegno per una
 * pagina. Usato da `src/pdf/render.ts` dopo il testo della pagina, così i
 * disegni restano davanti al testo come a schermo.
 */

import type { CharFormat, DocumentLayout, Drawing } from '../core';
import { encodedWidth, encodeWinAnsi, standardFont, type StandardFont } from '../pdf/fonts';
import { hexString, num } from '../pdf/writer';
import { chartPrimitives } from './chart';
import { drawingPaths, type Segment } from './geometry';
import { drawingsOnPage, isLine } from './model';
import { TEXT_INSET } from './render';

/** Nome di risorsa (/F1...) di un carattere standard, condiviso con il resto della pagina. */
export type PdfFontResolver = (font: StandardFont) => string;

function rgb(hex: string): string {
  const v = (i: number) => num(parseInt(hex.slice(i, i + 2), 16) / 255);
  return /^#[0-9a-f]{6}$/i.test(hex) ? `${v(1)} ${v(3)} ${v(5)}` : '0 0 0';
}

function pathOps(segments: readonly Segment[]): string {
  return segments
    .map((s) =>
      s.op === 'Z' ? 'h' : s.op === 'C' ? `${num(s.x1)} ${num(s.y1)} ${num(s.x2)} ${num(s.y2)} ${num(s.x)} ${num(s.y)} c` : `${num(s.x)} ${num(s.y)} ${s.op === 'M' ? 'm' : 'l'}`,
    )
    .join(' ');
}

function paint(segments: readonly Segment[], fill: string | null, stroke: string | null, width: number): string {
  if (!fill && !stroke) return '';
  const op = fill && stroke ? 'B' : fill ? 'f' : 'S';
  const colors = `${fill ? `${rgb(fill)} rg ` : ''}${stroke ? `${rgb(stroke)} RG ${num(width)} w ` : ''}`;
  return `${colors}1 j ${pathOps(segments)} ${op}`;
}

/**
 * Testo in coordinate locali con l'asse y verso il basso: la matrice del
 * testo ribalta di nuovo i glifi, che altrimenti uscirebbero capovolti.
 */
function textOp(fonts: PdfFontResolver, font: StandardFont, size: number, color: string, x: number, baseline: number, bytes: Uint8Array): string {
  return `BT ${fonts(font)} ${num(size)} Tf ${rgb(color)} rg 1 0 0 -1 ${num(x)} ${num(baseline)} Tm ${hexString(bytes)} Tj ET`;
}

/** Va a capo per parole dentro `width` (le parole troppo lunghe si spezzano). */
export function wrapText(text: string, width: number, measure: (s: string) => number): string[] {
  const out: string[] = [];
  for (const para of text.split('\n')) {
    let line = '';
    for (const word of para.split(/(?<=\s)/)) {
      const candidate = line + word;
      if (!line || measure(candidate.trimEnd()) <= width) {
        line = candidate;
        continue;
      }
      out.push(line.trimEnd());
      line = word;
    }
    // Parole più larghe della riga: si spezzano carattere per carattere.
    while (line.length > 1 && measure(line.trimEnd()) > width) {
      let cut = line.length - 1;
      while (cut > 1 && measure(line.slice(0, cut)) > width) cut--;
      out.push(line.slice(0, cut));
      line = line.slice(cut);
    }
    out.push(line.trimEnd());
  }
  return out;
}

function shapeText(d: Drawing, fonts: PdfFontResolver): string[] {
  const t = d.text;
  if (!t?.text.trim()) return [];
  const f: CharFormat = { fontFamily: 'Calibri', fontSize: 11, color: '#000000', ...t.format };
  const size = f.fontSize ?? 11;
  const font = standardFont(f);
  const measure = (s: string) => encodedWidth(font, encodeWinAnsi(s), size);
  const inner = Math.max(1, d.width - 2 * TEXT_INSET.x);
  const lines = wrapText(t.text, inner, measure);
  const lineHeight = size * 1.15;
  const total = lines.length * lineHeight;
  const avail = d.height - 2 * TEXT_INSET.y;
  const top = TEXT_INSET.y + (t.verticalAlign === 'middle' ? (avail - total) / 2 : t.verticalAlign === 'bottom' ? avail - total : 0);
  const ops: string[] = [];
  // Come a schermo, il testo che non sta nella forma non si vede.
  ops.push(`0 0 ${num(d.width)} ${num(d.height)} re W n`);
  lines.forEach((line, i) => {
    if (!line) return;
    const bytes = encodeWinAnsi(line);
    const w = encodedWidth(font, bytes, size);
    const x = TEXT_INSET.x + (t.align === 'center' ? (inner - w) / 2 : t.align === 'right' ? inner - w : 0);
    const baseline = top + i * lineHeight + size * 0.9;
    ops.push(textOp(fonts, font, size, f.color ?? '#000000', x, baseline, bytes));
  });
  return ops;
}

function chartOps(d: Drawing, fonts: PdfFontResolver): string[] {
  const ops: string[] = [];
  const frame = [{ op: 'M', x: 0, y: 0 }, { op: 'L', x: d.width, y: 0 }, { op: 'L', x: d.width, y: d.height }, { op: 'L', x: 0, y: d.height }, { op: 'Z' }] as const;
  ops.push(paint(frame, d.fill ?? null, d.stroke ?? null, d.strokeWidth ?? 0.75));
  if (!d.chart) return ops;
  for (const p of chartPrimitives(d.chart, d.width, d.height)) {
    if (p.t === 'rect') ops.push(`${rgb(p.fill)} rg ${num(p.x)} ${num(p.y)} ${num(p.w)} ${num(p.h)} re f`);
    else if (p.t === 'path') ops.push(paint(p.segments, p.fill, p.stroke, p.width));
    else {
      const font = standardFont({ fontFamily: 'Calibri' });
      const bytes = encodeWinAnsi(p.text);
      const w = encodedWidth(font, bytes, p.size);
      const x = p.anchor === 'middle' ? p.x - w / 2 : p.anchor === 'end' ? p.x - w : p.x;
      ops.push(textOp(fonts, font, p.size, p.color, x, p.y, bytes));
    }
  }
  return ops;
}

/** Operatori PDF dei disegni della pagina `index`, in coordinate PDF (origine in basso). */
export function drawingPdfOps(layout: DocumentLayout, index: number, fonts: PdfFontResolver): string[] {
  const page = layout.pages[index];
  if (!page) return [];
  const out: string[] = [];
  for (const d of drawingsOnPage(layout, index)) {
    const body: string[] = [];
    if (d.kind === 'chart') body.push(...chartOps(d, fonts));
    else {
      for (const p of drawingPaths(d)) body.push(paint(p.segments, p.fill, p.stroke, p.strokeWidth));
      if (!isLine(d)) body.push(...shapeText(d, fonts));
    }
    // Origine nell'angolo in alto a sinistra del disegno, y verso il basso come nel layout.
    out.push(`q 1 0 0 -1 ${num(d.x)} ${num(page.height - d.y)} cm ${body.filter(Boolean).join(' ')} Q`);
  }
  return out;
}
