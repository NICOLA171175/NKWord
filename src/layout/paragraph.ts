/**
 * Divisione di un paragrafo in righe. Il risultato è indipendente dalla
 * pagina: coordinate relative al bordo sinistro del blocco e all'inizio del
 * paragrafo (spazio prima escluso). L'impaginatore poi distribuisce le righe
 * sulle pagine.
 */

import type {
  CharFormat,
  DocumentModel,
  FieldKind,
  GlyphRun,
  Id,
  ParaFormat,
  Paragraph,
  Points,
  TabStop,
  TextMeasurer,
} from '../core';
import { resolveCharFormat, resolveParaFormat } from '../core';
import { baselineShift, displayText } from './measure';

/** Distanza fra le tabulazioni predefinite (1,27 cm). */
export const DEFAULT_TAB_INTERVAL: Points = 36;

/** Dati per calcolare i campi (numero di pagina, data...). */
export interface FieldContext {
  readonly page: number;
  readonly numPages: number;
  readonly now: Date;
  /** Numeri delle note a piè di pagina (vedi `footnoteNumbers`). */
  readonly footnotes?: ReadonlyMap<Id, number>;
}

export interface ParaLine {
  readonly start: number;
  readonly end: number;
  /** Relativo al bordo sinistro del blocco (rientri inclusi). */
  readonly x: Points;
  /** Relativo all'inizio del testo del paragrafo. */
  readonly y: Points;
  readonly width: Points;
  readonly height: Points;
  readonly baseline: Points;
  readonly runs: readonly GlyphRun[];
  /** La riga termina con un'interruzione di pagina o di colonna. */
  readonly breakAfter?: 'page' | 'column';
}

export interface ParaBox {
  readonly format: ParaFormat;
  readonly lines: readonly ParaLine[];
  /** Altezza delle righe, senza spazio prima/dopo. */
  readonly height: Points;
  readonly spaceBefore: Points;
  readonly spaceAfter: Points;
}

interface Atom {
  kind: 'text' | 'image' | 'tab' | 'field' | 'break';
  start: number;
  end: number;
  text: string;
  format: CharFormat;
  width: number;
  /** Larghezza degli spazi finali, che possono sporgere oltre il margine. */
  trailing: number;
  ascent: number;
  descent: number;
  canBreakAfter: boolean;
  breakKind?: 'line' | 'page' | 'column';
  src?: string;
  height?: number;
}

function fieldText(kind: FieldKind, ctx: FieldContext): string {
  switch (kind) {
    case 'page':
      return String(ctx.page);
    case 'numPages':
      return String(ctx.numPages);
    case 'date':
      return ctx.now.toLocaleDateString('it-IT');
    case 'time':
      return ctx.now.toLocaleTimeString('it-IT', { hour: '2-digit', minute: '2-digit' });
  }
}

export function hasFields(p: Paragraph): boolean {
  return p.children.some((c) => c.type === 'field' || c.type === 'footnoteRef');
}

function textAtom(
  measurer: TextMeasurer,
  text: string,
  start: number,
  format: CharFormat,
): Atom {
  const shown = displayText(text, format);
  const m = measurer.measure(shown, format);
  const body = shown.replace(/\s+$/, '');
  const trailing = body.length === shown.length ? 0 : m.width - measurer.measure(body, format).width;
  const shift = baselineShift(format);
  return {
    kind: 'text',
    start,
    end: start + text.length,
    text: shown,
    format,
    width: m.width,
    trailing,
    ascent: m.ascent + Math.max(0, shift),
    descent: m.descent + Math.max(0, -shift),
    canBreakAfter: trailing > 0,
  };
}

/** Divide un atomo di testo in due all'offset relativo `k`. */
function splitAtom(measurer: TextMeasurer, a: Atom, k: number): [Atom, Atom] {
  return [textAtom(measurer, a.text.slice(0, k), a.start, a.format), textAtom(measurer, a.text.slice(k), a.start + k, a.format)];
}

function buildAtoms(doc: DocumentModel, p: Paragraph, measurer: TextMeasurer, fields: FieldContext): Atom[] {
  const atoms: Atom[] = [];
  const mark = resolveCharFormat(doc.styles, p);
  let offset = 0;
  let lastFormat = mark;
  const metrics = (f: CharFormat) => measurer.measure('', f);
  for (const inline of p.children) {
    switch (inline.type) {
      case 'text': {
        const format = resolveCharFormat(doc.styles, p, inline);
        lastFormat = format;
        // Parole con i loro spazi finali: si può andare a capo solo dopo uno spazio.
        const re = /\S+\s*|\s+/g;
        let m: RegExpExecArray | null;
        while ((m = re.exec(inline.text))) atoms.push(textAtom(measurer, m[0], offset + m.index, format));
        offset += inline.text.length;
        break;
      }
      case 'image':
        atoms.push({
          kind: 'image',
          start: offset,
          end: offset + 1,
          text: '',
          format: lastFormat,
          width: inline.width,
          trailing: 0,
          ascent: inline.height,
          descent: 0,
          canBreakAfter: true,
          src: inline.src,
          height: inline.height,
        });
        offset += 1;
        break;
      case 'tab': {
        const m = metrics(lastFormat);
        atoms.push({
          kind: 'tab',
          start: offset,
          end: offset + 1,
          text: '\t',
          format: lastFormat,
          width: 0,
          trailing: 0,
          ascent: m.ascent,
          descent: m.descent,
          canBreakAfter: true,
        });
        offset += 1;
        break;
      }
      case 'break': {
        const m = metrics(lastFormat);
        atoms.push({
          kind: 'break',
          start: offset,
          end: offset + 1,
          text: '',
          format: lastFormat,
          width: 0,
          trailing: 0,
          ascent: m.ascent,
          descent: m.descent,
          canBreakAfter: true,
          breakKind: inline.kind,
        });
        offset += 1;
        break;
      }
      case 'field': {
        const format = resolveCharFormat(doc.styles, p, inline.format);
        const text = fieldText(inline.kind, fields);
        const m = measurer.measure(text, format);
        atoms.push({
          kind: 'field',
          start: offset,
          end: offset + 1,
          text,
          format,
          width: m.width,
          trailing: 0,
          ascent: m.ascent,
          descent: m.descent,
          canBreakAfter: false,
        });
        offset += 1;
        break;
      }
      case 'footnoteRef': {
        // Il richiamo è sempre in apice; il numero dipende dall'ordine nel documento.
        const format: CharFormat = { ...resolveCharFormat(doc.styles, p, inline.format), verticalAlign: 'superscript' };
        const text = String(fields.footnotes?.get(inline.noteId) ?? '');
        const m = measurer.measure(text, format);
        const shift = baselineShift(format);
        atoms.push({
          kind: 'field',
          start: offset,
          end: offset + 1,
          text,
          format,
          width: m.width,
          trailing: 0,
          ascent: m.ascent + shift,
          descent: m.descent,
          canBreakAfter: false,
        });
        offset += 1;
        break;
      }
    }
  }
  return atoms;
}

function nextTabStop(x: Points, tabs: readonly TabStop[] | undefined): TabStop {
  const custom = (tabs ?? []).filter((t) => t.position > x + 0.01).sort((a, b) => a.position - b.position)[0];
  if (custom) return custom;
  return { position: (Math.floor(x / DEFAULT_TAB_INTERVAL + 0.01) + 1) * DEFAULT_TAB_INTERVAL, align: 'left' };
}

function lineHeightFor(spacing: ParaFormat['lineSpacing'], natural: Points): Points {
  if (!spacing) return natural;
  switch (spacing.rule) {
    case 'multiple':
      return natural * spacing.value;
    case 'exact':
      return spacing.value;
    case 'atLeast':
      return Math.max(natural, spacing.value);
  }
}

export interface MarkerInfo {
  readonly text: string;
  readonly format: CharFormat;
  readonly indentLeft: Points;
  readonly hanging: Points;
}

/**
 * Divide il paragrafo in righe larghe al massimo `width`.
 * `marker` è il marcatore dell'elenco già calcolato (vedi `lists.ts`).
 */
export function layoutParagraph(
  doc: DocumentModel,
  p: Paragraph,
  width: Points,
  measurer: TextMeasurer,
  fields: FieldContext,
  marker?: MarkerInfo,
): ParaBox {
  const format = resolveParaFormat(doc.styles, p);
  const indentLeft = p.format.indentLeft ?? marker?.indentLeft ?? format.indentLeft ?? 0;
  const indentRight = format.indentRight ?? 0;
  const firstLine = p.format.indentFirstLine ?? (marker ? -marker.hanging : format.indentFirstLine ?? 0);
  const align = format.align ?? 'left';
  const atoms = buildAtoms(doc, p, measurer, fields);
  const mark = resolveCharFormat(doc.styles, p);
  const markMetrics = measurer.measure('', mark);
  const total = atoms.length ? atoms[atoms.length - 1]!.end : 0;

  const lines: ParaLine[] = [];
  let y = 0;
  let i = 0;
  let lineIndex = 0;
  do {
    const first = lineIndex === 0;
    const lineX = indentLeft + (first ? firstLine : 0);
    const lineWidth = Math.max(1, width - lineX - indentRight);
    const placed: { atom: Atom; x: number }[] = [];
    let x = 0;
    let markerRun: GlyphRun | undefined;

    if (first && marker) {
      const m = measurer.measure(marker.text, marker.format);
      markerRun = {
        kind: 'listMarker',
        paragraphId: p.id,
        start: 0,
        end: 0,
        text: marker.text,
        format: marker.format,
        x: 0,
        width: m.width,
      };
      // Dopo il marcatore il testo parte dal rientro sinistro (come una tabulazione).
      const after = lineX + m.width;
      const target = after < indentLeft - 0.01 ? indentLeft : nextTabStop(after, format.tabs).position;
      x = target - lineX;
    }

    let breakAfter: ParaLine['breakAfter'];
    while (i < atoms.length) {
      const a = atoms[i]!;
      if (a.kind === 'break') {
        placed.push({ atom: a, x });
        i++;
        if (a.breakKind !== 'line') breakAfter = a.breakKind;
        break;
      }
      if (a.kind === 'tab') {
        const abs = lineX + x;
        const stop = nextTabStop(abs, format.tabs);
        let w = stop.position - abs;
        if (stop.align !== 'left') {
          let follow = 0;
          for (let j = i + 1; j < atoms.length && atoms[j]!.kind !== 'tab' && atoms[j]!.kind !== 'break'; j++) {
            follow += atoms[j]!.width;
          }
          if (stop.align === 'right' || stop.align === 'decimal') w -= follow;
          else w -= follow / 2;
        }
        a.width = Math.max(0, w);
      }
      const fits = x + a.width - a.trailing <= lineWidth + 0.01;
      if (fits || (placed.length === 0 && a.kind !== 'text')) {
        placed.push({ atom: a, x });
        x += a.width;
        i++;
        continue;
      }
      // Non ci sta: si torna all'ultima opportunità di andare a capo.
      let bp = -1;
      for (let j = placed.length - 1; j >= 0; j--) {
        if (placed[j]!.atom.canBreakAfter) {
          bp = j;
          break;
        }
      }
      if (bp >= 0) {
        const removed = placed.length - 1 - bp;
        placed.length = bp + 1;
        i -= removed;
      } else if (placed.length === 0 || placed.every((e) => e.atom.kind === 'text')) {
        // Parola più lunga della riga: si spezza per caratteri.
        if (a.kind === 'text') {
          let k = 1;
          while (k < a.text.length && x + measurer.measure(a.text.slice(0, k + 1), a.format).width <= lineWidth) k++;
          if (placed.length > 0 && x + measurer.measure(a.text.slice(0, 1), a.format).width > lineWidth) break;
          if (k >= a.text.length) {
            placed.push({ atom: a, x });
          } else {
            const [head, tail] = splitAtom(measurer, a, k);
            atoms.splice(i, 1, head, tail);
            placed.push({ atom: head, x });
          }
          i++;
        }
      }
      break;
    }

    // Metriche verticali.
    let ascent = 0;
    let descent = 0;
    for (const { atom } of placed) {
      ascent = Math.max(ascent, atom.ascent);
      descent = Math.max(descent, atom.descent);
    }
    if (markerRun) {
      const m = measurer.measure(marker!.text, marker!.format);
      ascent = Math.max(ascent, m.ascent);
      descent = Math.max(descent, m.descent);
    }
    if (ascent === 0 && descent === 0) {
      ascent = markMetrics.ascent;
      descent = markMetrics.descent;
    }
    const natural = ascent + descent;
    const height = lineHeightFor(format.lineSpacing, natural);
    const baseline = natural > 0 ? (height * ascent) / natural : 0;

    // Allineamento orizzontale.
    const last = placed[placed.length - 1];
    const used = last ? last.x + last.atom.width - last.atom.trailing : x;
    const isLastLine = i >= atoms.length;
    const endsWithBreak = last?.atom.kind === 'break';
    let shift = 0;
    let extraPerGap = 0;
    const free = Math.max(0, lineWidth - used);
    if (align === 'center') shift = free / 2;
    else if (align === 'right') shift = free;
    else if (align === 'justify' && !isLastLine && !endsWithBreak) {
      // Lo spazio extra va dopo le parole, a partire dall'ultima tabulazione.
      let from = 0;
      placed.forEach((e, j) => {
        if (e.atom.kind === 'tab') from = j + 1;
      });
      const gaps = placed.slice(from, -1).filter((e) => e.atom.trailing > 0).length;
      if (gaps > 0) extraPerGap = free / gaps;
    }

    const runs: GlyphRun[] = markerRun ? [markerRun] : [];
    let extra = 0;
    let firstGap = 0;
    placed.forEach((e, j) => {
      if (e.atom.kind === 'tab') firstGap = j + 1;
    });
    placed.forEach(({ atom, x: ax }, j) => {
      const isGap = extraPerGap > 0 && j >= firstGap && j < placed.length - 1 && atom.trailing > 0;
      const run: GlyphRun = {
        kind: atom.kind === 'break' ? 'text' : atom.kind,
        paragraphId: p.id,
        start: atom.start,
        end: atom.end,
        text: atom.kind === 'break' ? '' : atom.text,
        format: atom.format,
        x: ax + shift + extra,
        width: atom.width + (isGap ? extraPerGap : 0),
        ...(atom.kind === 'image' ? { src: atom.src!, height: atom.height! } : {}),
      };
      if (isGap) extra += extraPerGap;
      const prev = runs[runs.length - 1];
      // Unisce i pezzi di testo contigui con la stessa formattazione.
      if (
        prev &&
        extraPerGap === 0 &&
        prev.kind === 'text' &&
        run.kind === 'text' &&
        atom.kind === 'text' &&
        prev.format === run.format &&
        prev.end === run.start &&
        prev.text !== ''
      ) {
        runs[runs.length - 1] = { ...prev, end: run.end, text: prev.text + run.text, width: prev.width + run.width };
      } else runs.push(run);
    });

    const start = placed[0]?.atom.start ?? (lines.length ? lines[lines.length - 1]!.end : 0);
    const end = last ? last.atom.end : start;
    lines.push({
      start,
      end: isLastLine ? Math.max(end, total) : end,
      x: lineX,
      y,
      width: lineWidth,
      height,
      baseline,
      runs,
      ...(breakAfter ? { breakAfter } : {}),
    });
    y += height;
    lineIndex++;
    // Un'interruzione alla fine del paragrafo lascia una riga vuota dopo,
    // dove sta il segno di paragrafo.
    if (isLastLine && endsWithBreak) {
      lines.push({ start: total, end: total, x: indentLeft, y, width: Math.max(1, width - indentLeft - indentRight), height, baseline, runs: [] });
      y += height;
    }
  } while (i < atoms.length);

  return {
    format,
    lines,
    height: y,
    spaceBefore: format.spaceBefore ?? 0,
    spaceAfter: format.spaceAfter ?? 0,
  };
}
