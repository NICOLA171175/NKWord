/**
 * I 14 caratteri standard del PDF (non incorporati: ogni lettore li ha) con
 * la codifica WinAnsi e le larghezze dei glifi dai file AFM di Adobe.
 *
 * Le larghezze servono solo a ripartire lo spazio dentro un pezzo di testo:
 * la larghezza totale viene poi riportata a quella misurata dal layout con
 * la scala orizzontale (`Tz`), così il PDF ricalca le pagine a schermo anche
 * se il carattere non è identico. Per questo le varianti corsive usano le
 * larghezze del tondo.
 */

import type { CharFormat } from '../core';

export type FontFamily = 'Helvetica' | 'Times' | 'Courier';

export interface StandardFont {
  /** Nome PostScript, es. 'Helvetica-BoldOblique'. */
  readonly name: string;
  /** Larghezze in millesimi di em per i codici WinAnsi 32…255. */
  readonly widths: readonly number[];
}

/** Larghezze dei caratteri ASCII 32…126 (spazio … tilde). */
const ASCII_WIDTHS: Record<'helvetica' | 'helveticaBold' | 'times' | 'timesBold', string> = {
  helvetica:
    '278 278 355 556 556 889 667 191 333 333 389 584 278 333 278 278 556 556 556 556 556 556 556 556 556 556 278 278 584 584 584 556 ' +
    '1015 667 667 722 722 667 611 778 722 278 500 667 556 833 722 778 667 778 722 667 611 722 667 944 667 667 611 278 278 278 469 556 ' +
    '333 556 556 500 556 556 278 556 556 222 222 500 222 833 556 556 556 556 333 500 278 556 500 722 500 500 500 334 260 334 584',
  helveticaBold:
    '278 333 474 556 556 889 722 238 333 333 389 584 278 333 278 278 556 556 556 556 556 556 556 556 556 556 333 333 584 584 584 611 ' +
    '975 722 722 722 722 667 611 778 722 278 556 722 611 833 722 778 667 778 722 667 611 722 667 944 667 667 611 333 278 333 584 556 ' +
    '333 556 611 556 611 556 333 611 611 278 278 556 278 889 611 611 611 611 389 556 333 611 556 778 556 556 500 389 280 389 584',
  times:
    '250 333 408 500 500 833 778 180 333 333 500 564 250 333 250 278 500 500 500 500 500 500 500 500 500 500 278 278 564 564 564 444 ' +
    '921 722 667 667 722 611 556 722 722 333 389 722 611 889 722 722 556 722 667 556 611 722 722 944 722 722 611 333 278 333 469 500 ' +
    '333 444 500 444 500 444 333 500 500 278 278 500 278 778 500 500 500 500 333 389 278 500 500 722 500 500 444 480 200 480 541',
  timesBold:
    '250 333 555 500 500 1000 833 278 333 333 500 570 250 333 250 278 500 500 500 500 500 500 500 500 500 500 333 333 570 570 570 500 ' +
    '930 722 667 722 722 667 611 778 778 389 500 778 667 944 722 778 611 778 722 556 667 722 722 1000 722 722 667 333 278 333 581 500 ' +
    '333 500 556 444 556 444 333 500 556 278 333 556 278 833 556 500 556 556 444 389 333 556 500 722 500 500 444 394 220 394 520',
};

/** Codici WinAnsi 0x80…0x9F diversi da Latin-1. */
const WIN_ANSI_EXTRA: Record<string, number> = {
  '€': 0x80, '‚': 0x82, 'ƒ': 0x83, '„': 0x84, '…': 0x85, '†': 0x86, '‡': 0x87, 'ˆ': 0x88, '‰': 0x89, 'Š': 0x8a,
  '‹': 0x8b, 'Œ': 0x8c, 'Ž': 0x8e, '‘': 0x91, '’': 0x92, '“': 0x93, '”': 0x94, '•': 0x95, '–': 0x96, '—': 0x97,
  '˜': 0x98, '™': 0x99, 'š': 0x9a, '›': 0x9b, 'œ': 0x9c, 'ž': 0x9e, 'Ÿ': 0x9f,
};
const WIN_ANSI_CHAR = new Map(Object.entries(WIN_ANSI_EXTRA).map(([ch, code]) => [code, ch]));

/** Larghezze approssimate dei simboli 0x80…0x9F, in proporzione al corpo. */
const EXTRA_WIDTH_EM: Record<number, number> = { 0x85: 1, 0x89: 1, 0x8c: 1, 0x95: 0.35, 0x96: 0.556, 0x97: 1, 0x99: 1, 0x9c: 0.9 };

/** Codice WinAnsi di un carattere; `?` per quelli che la codifica non ha. */
export function winAnsiCode(ch: string): number {
  const cp = ch.codePointAt(0) ?? 63;
  if ((cp >= 0x20 && cp <= 0x7e) || (cp >= 0xa0 && cp <= 0xff)) return cp;
  if (cp === 0x09) return 0x20;
  return WIN_ANSI_EXTRA[ch] ?? 0x3f;
}

/** Testo → byte WinAnsi. */
export function encodeWinAnsi(text: string): Uint8Array {
  const out: number[] = [];
  for (const ch of text) out.push(winAnsiCode(ch));
  return Uint8Array.from(out);
}

function buildWidths(ascii: string): number[] {
  const base = ascii.split(' ').map(Number);
  const widths: number[] = [];
  const of = (code: number) => base[code - 32] ?? base[78]!; // 'n' come ripiego
  for (let code = 32; code <= 255; code++) {
    if (code <= 126) widths.push(of(code));
    else if (code >= 0xa0) {
      // Lettere accentate: la larghezza della lettera base (à → a).
      const base0 = String.fromCharCode(code).normalize('NFD')[0]!;
      const b = base0.charCodeAt(0);
      widths.push(b >= 32 && b <= 126 && b !== code ? of(b) : code === 0xa0 ? of(32) : of(78));
    } else {
      const em = EXTRA_WIDTH_EM[code];
      widths.push(em ? Math.round(em * 1000) : WIN_ANSI_CHAR.has(code) ? of(78) : of(63));
    }
  }
  return widths;
}

const HELVETICA = buildWidths(ASCII_WIDTHS.helvetica);
const HELVETICA_BOLD = buildWidths(ASCII_WIDTHS.helveticaBold);
const TIMES = buildWidths(ASCII_WIDTHS.times);
const TIMES_BOLD = buildWidths(ASCII_WIDTHS.timesBold);
const COURIER = new Array<number>(224).fill(600);

/** Famiglia standard più vicina al nome del carattere del documento. */
export function familyFor(fontFamily: string | undefined): FontFamily {
  const f = (fontFamily ?? '').toLowerCase();
  if (/courier|consolas|mono|menlo|lucida console/.test(f)) return 'Courier';
  if (/times|georgia|cambria|garamond|book antiqua|palatino|serif$/.test(f) && !/sans/.test(f)) return 'Times';
  return 'Helvetica';
}

export function standardFont(format: CharFormat): StandardFont {
  const family = familyFor(format.fontFamily);
  const bold = !!format.bold;
  const italic = !!format.italic;
  switch (family) {
    case 'Times': {
      const style = bold && italic ? 'BoldItalic' : bold ? 'Bold' : italic ? 'Italic' : 'Roman';
      return { name: `Times-${style}`, widths: bold ? TIMES_BOLD : TIMES };
    }
    case 'Courier': {
      const style = bold && italic ? '-BoldOblique' : bold ? '-Bold' : italic ? '-Oblique' : '';
      return { name: `Courier${style}`, widths: COURIER };
    }
    default: {
      const style = bold && italic ? '-BoldOblique' : bold ? '-Bold' : italic ? '-Oblique' : '';
      return { name: `Helvetica${style}`, widths: bold ? HELVETICA_BOLD : HELVETICA };
    }
  }
}

/** Larghezza in punti di byte WinAnsi già codificati. */
export function encodedWidth(font: StandardFont, bytes: Uint8Array, size: number): number {
  let units = 0;
  for (const b of bytes) units += font.widths[b - 32] ?? 0;
  return (units * size) / 1000;
}
