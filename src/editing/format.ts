/**
 * Formattazione del carattere sulla selezione.
 *
 * Come in Word: con selezione vuota dentro una parola la modifica si applica
 * all'intera parola; altrimenti diventa "formattazione in attesa" per il
 * prossimo testo digitato.
 */

import {
  inlineLength,
  isCollapsed,
  paragraphLength,
  paragraphText,
  replaceBlock,
  resolveCharFormat,
  type CharFormat,
  type DocumentModel,
  type EditorState,
  type Inline,
  type Paragraph,
  type Position,
  type StyleSheet,
  type Transaction,
} from '../core';
import { formatAt, mapFormatInRange, textToInlines, insertInParagraph } from './inlines';
import { locate, paragraphsBetween } from './range';
import { selectionRange } from './selection';
import { wordAround } from './words';

type Mutable<T> = { -readonly [K in keyof T]: T[K] };

/** Intervallo su cui agisce un comando di formato: la selezione o la parola al cursore. */
export function targetRange(state: EditorState): [Position, Position] | null {
  const [start, end] = selectionRange(state);
  if (!isCollapsed(state.selection)) return [start, end];
  const p = locate(state.doc, start.paragraphId).paragraph;
  const word = wordAround(paragraphText(p), start.offset);
  return word && [{ paragraphId: p.id, offset: word[0] }, { paragraphId: p.id, offset: word[1] }];
}

function resolved(sheet: StyleSheet, p: Paragraph, format: CharFormat, styleId?: string): CharFormat {
  return resolveCharFormat(sheet, p, styleId ? { type: 'text', text: '', format, styleId } : { type: 'text', text: '', format });
}

/** Formattazione effettiva al cursore, inclusa quella in attesa. */
export function caretFormat(state: EditorState): CharFormat {
  const pos = state.selection.focus;
  const p = locate(state.doc, pos.paragraphId).paragraph;
  const base = formatAt(p, pos.offset);
  return { ...resolved(state.doc.styles, p, base.format, base.styleId), ...state.pendingFormat };
}

/** Formattazioni effettive di tutti i run di testo nella selezione (o al cursore). */
export function selectionFormats(state: EditorState): CharFormat[] {
  const range = targetRange(state);
  if (!range || isCollapsed({ anchor: range[0], focus: range[1] })) return [caretFormat(state)];
  const [start, end] = range;
  const out: CharFormat[] = [];
  for (const { paragraph: p } of paragraphsBetween(state.doc, start, end)) {
    const from = p.id === start.paragraphId ? start.offset : 0;
    const to = p.id === end.paragraphId ? end.offset : paragraphLength(p);
    let pos = 0;
    for (const inline of p.children) {
      const len = inlineLength(inline);
      if (pos < to && pos + len > from && (inline.type === 'text' || inline.type === 'field')) {
        out.push(resolved(state.doc.styles, p, inline.format, inline.type === 'text' ? inline.styleId : undefined));
      }
      pos += len;
    }
  }
  return out.length ? out : [caretFormat(state)];
}

export type CharChange = (effective: CharFormat) => CharFormat;

/** Toglie le chiavi cambiate che coincidono con quanto già dato dallo stile. */
function prune(sheet: StyleSheet, p: Paragraph, format: CharFormat, changed: CharFormat, styleId?: string): CharFormat {
  const base = resolved(sheet, p, {}, styleId);
  const out: Mutable<CharFormat> = { ...format };
  for (const key of Object.keys(changed) as (keyof CharFormat)[]) {
    const v = out[key];
    const b = base[key];
    const neutral = v === null || v === false || v === 'baseline';
    if (v === b || (b === undefined && neutral)) delete out[key];
  }
  return out;
}

function changeParagraph(
  sheet: StyleSheet,
  p: Paragraph,
  from: number,
  to: number,
  change: CharChange,
  includeMark: boolean,
): Paragraph {
  let next = mapFormatInRange(p, from, to, (format, inline) => {
    const styleId = inline.type === 'text' ? inline.styleId : undefined;
    const delta = change(resolved(sheet, p, format, styleId));
    return { format: prune(sheet, p, { ...format, ...delta }, delta, styleId), styleId };
  });
  if (includeMark) {
    const mark = p.markFormat ?? {};
    const delta = change(resolved(sheet, p, mark));
    next = { ...next, markFormat: prune(sheet, p, { ...mark, ...delta }, delta) };
  }
  return next;
}

/** Applica una modifica di formato all'intervallo [start, end]. */
export function applyCharChangeToRange(
  doc: DocumentModel,
  start: Position,
  end: Position,
  change: CharChange,
): DocumentModel {
  let out = doc;
  for (const { paragraph: p } of paragraphsBetween(doc, start, end)) {
    const from = p.id === start.paragraphId ? start.offset : 0;
    const to = p.id === end.paragraphId ? end.offset : paragraphLength(p);
    // Il segno di paragrafo è selezionato se la selezione prosegue oltre.
    const includeMark = p.id !== end.paragraphId || paragraphLength(p) === 0;
    out = replaceBlock(out, p.id, [changeParagraph(doc.styles, p, from, to, change, includeMark)]);
  }
  return out;
}

/** Comando di formato carattere generico. */
export function applyCharChange(state: EditorState, change: CharChange, label: string): Transaction {
  const range = targetRange(state);
  if (!range) {
    const pendingFormat = { ...state.pendingFormat, ...change(caretFormat(state)) };
    return { state: { ...state, pendingFormat }, meta: { label, addToHistory: false } };
  }
  const doc = applyCharChangeToRange(state.doc, range[0], range[1], change);
  return { state: { doc, selection: state.selection }, meta: { label } };
}

export interface Toggle {
  readonly isOn: (f: CharFormat) => boolean;
  readonly on: CharFormat;
  readonly off: CharFormat;
}

export const TOGGLES = {
  bold: { isOn: (f) => f.bold === true, on: { bold: true }, off: { bold: false } },
  italic: { isOn: (f) => f.italic === true, on: { italic: true }, off: { italic: false } },
  underline: { isOn: (f) => !!f.underline, on: { underline: 'single' }, off: { underline: null } },
  strike: { isOn: (f) => f.strike === true, on: { strike: true }, off: { strike: false } },
  superscript: {
    isOn: (f) => f.verticalAlign === 'superscript',
    on: { verticalAlign: 'superscript' },
    off: { verticalAlign: 'baseline' },
  },
  subscript: {
    isOn: (f) => f.verticalAlign === 'subscript',
    on: { verticalAlign: 'subscript' },
    off: { verticalAlign: 'baseline' },
  },
} satisfies Record<string, Toggle>;

export function isToggleActive(state: EditorState, toggle: Toggle): boolean {
  return selectionFormats(state).every(toggle.isOn);
}

export function applyToggle(state: EditorState, toggle: Toggle, label: string): Transaction {
  const value = isToggleActive(state, toggle) ? toggle.off : toggle.on;
  return applyCharChange(state, () => value, label);
}

/** Dimensioni dei pulsanti "Aumenta/Riduci carattere" di Word. */
const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
export const MIN_FONT_SIZE = 1;
export const MAX_FONT_SIZE = 1638;

export function growSize(size: number): number {
  const next = FONT_SIZES.find((s) => s > size);
  if (next !== undefined) return next;
  return Math.min(MAX_FONT_SIZE, (Math.floor(size / 10) + 1) * 10);
}

export function shrinkSize(size: number): number {
  if (size <= FONT_SIZES[0]!) return Math.max(MIN_FONT_SIZE, Math.ceil(size) - 1);
  if (size > 72) return Math.max(72, (Math.ceil(size / 10) - 1) * 10);
  return [...FONT_SIZES].reverse().find((s) => s < size)!;
}

/** Cancella formattazione: toglie formato diretto e stile dai caratteri e riporta i paragrafi allo stile Normale. */
export function clearFormatting(state: EditorState): Transaction {
  let [start, end] = selectionRange(state);
  if (isCollapsed(state.selection)) {
    const p = locate(state.doc, start.paragraphId).paragraph;
    start = { paragraphId: p.id, offset: 0 };
    end = { paragraphId: p.id, offset: paragraphLength(p) };
  }
  let doc = state.doc;
  for (const { paragraph: p } of paragraphsBetween(doc, start, end)) {
    const from = p.id === start.paragraphId ? start.offset : 0;
    const to = p.id === end.paragraphId ? end.offset : paragraphLength(p);
    let next = mapFormatInRange(p, from, to, (format) => ({
      format: { ...(format.link ? { link: format.link } : {}), ...(format.commentIds ? { commentIds: format.commentIds } : {}) },
    }));
    const { styleId: _s, markFormat: _m, ...rest } = next;
    next = { ...rest, format: p.format.list ? { list: p.format.list } : {} };
    doc = replaceBlock(doc, p.id, [next]);
  }
  return { state: { doc, selection: state.selection }, meta: { label: 'Cancella formattazione' } };
}

/** Intervallo del collegamento che contiene il cursore. */
function linkRangeAt(p: Paragraph, offset: number): [number, number] | null {
  const spans: { from: number; to: number; inline: Inline }[] = [];
  let pos = 0;
  for (const inline of p.children) {
    spans.push({ from: pos, to: pos + inlineLength(inline), inline });
    pos += inlineLength(inline);
  }
  const linkOf = (i: Inline) => (i.type === 'text' ? i.format.link : undefined);
  const hit = spans.findIndex((s) => linkOf(s.inline) && s.from <= offset && offset <= s.to);
  if (hit < 0) return null;
  const url = linkOf(spans[hit]!.inline);
  let a = hit;
  let b = hit;
  while (a > 0 && linkOf(spans[a - 1]!.inline) === url) a--;
  while (b + 1 < spans.length && linkOf(spans[b + 1]!.inline) === url) b++;
  return [spans[a]!.from, spans[b]!.to];
}

/** Inserisce, cambia o rimuove (`url: null`) un collegamento ipertestuale. */
export function setLink(state: EditorState, url: string | null): Transaction | null {
  const change: CharChange = () => (url ? { link: url } : {});
  const strip = (doc: DocumentModel, start: Position, end: Position) => {
    let out = doc;
    for (const { paragraph: p } of paragraphsBetween(doc, start, end)) {
      const from = p.id === start.paragraphId ? start.offset : 0;
      const to = p.id === end.paragraphId ? end.offset : paragraphLength(p);
      out = replaceBlock(out, p.id, [
        mapFormatInRange(p, from, to, (format, inline) => {
          const { link: _l, ...rest } = format;
          return { format: rest, styleId: inline.type === 'text' ? inline.styleId : undefined };
        }),
      ]);
    }
    return out;
  };

  if (!isCollapsed(state.selection)) {
    const [start, end] = selectionRange(state);
    const doc = url ? applyCharChangeToRange(state.doc, start, end, change) : strip(state.doc, start, end);
    return { state: { doc, selection: state.selection }, meta: { label: 'Collegamento' } };
  }
  const pos = state.selection.focus;
  const p = locate(state.doc, pos.paragraphId).paragraph;
  const range = linkRangeAt(p, pos.offset) ?? wordAround(paragraphText(p), pos.offset);
  if (range) {
    const start = { paragraphId: p.id, offset: range[0] };
    const end = { paragraphId: p.id, offset: range[1] };
    const doc = url ? applyCharChangeToRange(state.doc, start, end, change) : strip(state.doc, start, end);
    return { state: { doc, selection: state.selection }, meta: { label: 'Collegamento' } };
  }
  if (!url) return null;
  // Nessun testo: come Word, si inserisce l'indirizzo stesso come collegamento.
  const base = formatAt(p, pos.offset);
  const inlines = textToInlines(url, { ...base.format, ...state.pendingFormat, link: url }, base.styleId);
  const doc = replaceBlock(state.doc, p.id, [insertInParagraph(p, pos.offset, inlines)]);
  const offset = pos.offset + url.length;
  const at = { paragraphId: p.id, offset };
  return { state: { doc, selection: { anchor: at, focus: at } }, meta: { label: 'Collegamento' } };
}
