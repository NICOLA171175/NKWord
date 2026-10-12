/**
 * Misurazione del testo. Il motore di impaginazione riceve un `TextMeasurer`
 * iniettato: nel browser si usa quello basato su canvas, nei test quello
 * approssimato (deterministico, senza DOM).
 *
 * Tutte le misure sono in punti: la dimensione del carattere in punti viene
 * passata al canvas come se fossero pixel, quindi le larghezze restituite
 * sono già nella stessa unità.
 */

import type { CharFormat, Points, TextMeasurer } from '../core';

/** Riduzione della dimensione per apice e pedice, come in Word. */
export const SCRIPT_SCALE = 0.65;
/** Riduzione per il maiuscoletto. */
export const SMALL_CAPS_SCALE = 0.8;

export const DEFAULT_FONT_FAMILY = 'Calibri';
export const DEFAULT_FONT_SIZE = 11;

/** Dimensione effettiva del carattere, tenendo conto di apice/pedice. */
export function effectiveFontSize(format: CharFormat): Points {
  const size = format.fontSize ?? DEFAULT_FONT_SIZE;
  return format.verticalAlign && format.verticalAlign !== 'baseline' ? size * SCRIPT_SCALE : size;
}

/** Spostamento verticale della linea di base (positivo = verso l'alto). */
export function baselineShift(format: CharFormat): Points {
  const size = format.fontSize ?? DEFAULT_FONT_SIZE;
  if (format.verticalAlign === 'superscript') return size * 0.33;
  if (format.verticalAlign === 'subscript') return -size * 0.14;
  return 0;
}

/** Testo come viene mostrato (maiuscolo se `allCaps`). Stessa lunghezza dell'originale. */
export function displayText(text: string, format: CharFormat): string {
  if (!format.allCaps) return text;
  const upper = text.toUpperCase();
  // Alcuni caratteri (es. 'ß') cambiano lunghezza: in quel caso si lascia
  // l'originale per non rompere la corrispondenza con gli offset.
  return upper.length === text.length ? upper : text;
}

/** Stringa CSS `font` con la dimensione in "px" che in realtà sono punti. */
export function cssFont(format: CharFormat, size = effectiveFontSize(format)): string {
  const family = format.fontFamily ?? DEFAULT_FONT_FAMILY;
  return `${format.italic ? 'italic ' : ''}${format.bold ? 'bold ' : ''}${size}px ${quoteFamily(family)}`;
}

export function quoteFamily(family: string): string {
  const generic = /^(serif|sans-serif|monospace|cursive|fantasy|system-ui)$/;
  const quoted = generic.test(family) ? family : `"${family.replace(/"/g, '')}"`;
  return `${quoted}, sans-serif`;
}

// ---------------------------------------------------------------------------
// Misuratore approssimato

/**
 * Misuratore senza DOM: ogni carattere vale una frazione fissa della
 * dimensione del carattere. Usato nei test e come ripiego fuori dal browser.
 */
export function createApproxMeasurer(opts: { charWidth?: number; spaceWidth?: number } = {}): TextMeasurer {
  const charWidth = opts.charWidth ?? 0.5;
  const spaceWidth = opts.spaceWidth ?? 0.25;
  return {
    measure(text, format) {
      let size = effectiveFontSize(format);
      if (format.smallCaps) size *= SMALL_CAPS_SCALE;
      let units = 0;
      for (const ch of text) units += ch === ' ' ? spaceWidth : format.bold ? charWidth * 1.1 : charWidth;
      const full = format.fontSize ?? DEFAULT_FONT_SIZE;
      return { width: units * size, ascent: full * 0.8, descent: full * 0.2 };
    },
  };
}

// ---------------------------------------------------------------------------
// Misuratore canvas

interface CanvasLike {
  getContext(type: '2d'): CanvasRenderingContext2D | null;
}

/**
 * Misuratore basato su `CanvasRenderingContext2D.measureText`, con cache.
 * `canvas` si può passare per usare un `OffscreenCanvas`.
 */
export function createCanvasMeasurer(canvas?: CanvasLike): TextMeasurer {
  const c: CanvasLike =
    canvas ??
    (typeof OffscreenCanvas !== 'undefined'
      ? (new OffscreenCanvas(1, 1) as unknown as CanvasLike)
      : document.createElement('canvas'));
  const ctx = c.getContext('2d');
  if (!ctx) throw new Error('Canvas 2D non disponibile');
  const widths = new Map<string, number>();
  const metrics = new Map<string, { ascent: number; descent: number }>();

  const vertical = (font: string, size: number) => {
    let m = metrics.get(font);
    if (!m) {
      ctx.font = font;
      const tm = ctx.measureText('Hg');
      m = {
        ascent: tm.fontBoundingBoxAscent || size * 0.8,
        descent: tm.fontBoundingBoxDescent || size * 0.2,
      };
      metrics.set(font, m);
    }
    return m;
  };

  return {
    measure(text, format) {
      let size = effectiveFontSize(format);
      if (format.smallCaps) size *= SMALL_CAPS_SCALE;
      const font = cssFont(format, size);
      const key = `${font}\u0000${text}`;
      let width = widths.get(key);
      if (width === undefined) {
        ctx.font = font;
        width = ctx.measureText(format.smallCaps ? text.toUpperCase() : text).width;
        if (widths.size > 20000) widths.clear();
        widths.set(key, width);
      }
      // L'altezza della riga dipende dalla dimensione piena, non ridotta.
      const full = format.fontSize ?? DEFAULT_FONT_SIZE;
      const v = vertical(cssFont({ ...format, verticalAlign: 'baseline', smallCaps: false }, full), full);
      return { width, ascent: v.ascent, descent: v.descent };
    },
  };
}
