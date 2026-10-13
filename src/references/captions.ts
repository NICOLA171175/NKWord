/**
 * Didascalie numerate (Riferimenti > Inserisci didascalia), come in Word.
 *
 * Una didascalia è un paragrafo con lo stile "Didascalia" (nome di Word
 * "caption") il cui testo inizia con etichetta e numero: "Figura 1: Grafico
 * delle vendite". Il numero non è salvato a parte: si ricalcola contando le
 * didascalie con la stessa etichetta in ordine di lettura, così quelle
 * importate da Word (campo SEQ) vengono riconosciute e rinumerate.
 */

import {
  caret,
  findParagraph,
  paragraph,
  paragraphLength,
  paragraphs,
  replaceBlock,
  textRun,
  type DocumentModel,
  type EditorState,
  type Id,
  type Paragraph,
  type Style,
  type StyleSheet,
  type Transaction,
  STYLE_NORMAL,
} from '../core';
import { formatAt } from '../editing/inlines';
import { ensureStyles, findStyleId, oneLine, replaceText } from './common';

export const CAPTION_STYLE = 'Caption';
/** Etichette proposte, come in Word in italiano. */
export const CAPTION_LABELS: readonly string[] = ['Figura', 'Tabella', 'Equazione'];

const CAPTION_NAME = /^(caption|didascalia)$/i;

const captionStyle: Style = {
  id: CAPTION_STYLE,
  name: 'caption',
  type: 'paragraph',
  basedOn: STYLE_NORMAL,
  next: STYLE_NORMAL,
  paraFormat: { spaceBefore: 2, spaceAfter: 10 },
  charFormat: { italic: true, fontSize: 9, color: '#44546a' },
};

export function captionStyleId(sheet: StyleSheet): Id | undefined {
  return findStyleId(sheet, CAPTION_STYLE, CAPTION_NAME);
}

export function isCaptionStyle(sheet: StyleSheet, styleId: Id | undefined): boolean {
  if (!styleId) return false;
  return styleId === CAPTION_STYLE || CAPTION_NAME.test(sheet.styles[styleId]?.name ?? '');
}

export function ensureCaptionStyle(sheet: StyleSheet): StyleSheet {
  return captionStyleId(sheet) ? sheet : ensureStyles(sheet, [captionStyle]);
}

// ---------------------------------------------------------------------------
// Lettura

export interface Caption {
  readonly paragraphId: Id;
  readonly label: string;
  readonly number: number;
  /** Posizione del numero nel paragrafo: [start, end). */
  readonly numberStart: number;
  readonly numberEnd: number;
  /** Testo intero su una riga ("Figura 1: Grafico"). */
  readonly text: string;
  /** Solo il testo dopo etichetta e numero ("Grafico"). */
  readonly caption: string;
}

const HEAD = /^([\p{L}][\p{L}\p{M}'’]*)[  ](\d+)/u;

/** Etichetta e numero all'inizio di un paragrafo di didascalia. */
export function parseCaption(p: Paragraph): Caption | null {
  // Il testo per offset del modello: ogni oggetto vale un carattere.
  const raw = p.children.map((c) => (c.type === 'text' ? c.text : '\u0000')).join('');
  const lead = raw.length - raw.trimStart().length;
  const m = HEAD.exec(raw.slice(lead));
  if (!m) return null;
  const numberStart = lead + m[1]!.length + 1;
  const text = oneLine(p);
  return {
    paragraphId: p.id,
    label: m[1]!,
    number: Number(m[2]),
    numberStart,
    numberEnd: numberStart + m[2]!.length,
    text,
    caption: text.slice(text.indexOf(m[2]!) + m[2]!.length).replace(/^[\s:.\-–—]+/, ''),
  };
}

/** Didascalie del documento in ordine di lettura (anche dentro le tabelle). */
export function collectCaptions(doc: DocumentModel): Caption[] {
  const out: Caption[] = [];
  for (const { paragraph: p } of paragraphs(doc)) {
    if (!isCaptionStyle(doc.styles, p.styleId)) continue;
    const c = parseCaption(p);
    if (c) out.push(c);
  }
  return out;
}

/** Etichette proposte più quelle già usate nel documento. */
export function captionLabels(doc: DocumentModel): string[] {
  const labels = [...CAPTION_LABELS];
  for (const c of collectCaptions(doc)) if (!labels.includes(c.label)) labels.push(c.label);
  return labels;
}

/** Prossimo numero libero per un'etichetta (per l'anteprima della finestra). */
export function nextCaptionNumber(doc: DocumentModel, label: string): number {
  return collectCaptions(doc).filter((c) => c.label === label).length + 1;
}

// ---------------------------------------------------------------------------
// Numerazione

/** Rinumera le didascalie: 1, 2, 3... per ogni etichetta, in ordine di lettura. */
export function renumberCaptions(doc: DocumentModel): DocumentModel {
  const counters = new Map<string, number>();
  let out = doc;
  for (const c of collectCaptions(doc)) {
    const n = (counters.get(c.label) ?? 0) + 1;
    counters.set(c.label, n);
    if (n === c.number && String(n).length === c.numberEnd - c.numberStart) continue;
    const p = findParagraph(out, c.paragraphId)!.paragraph;
    const { format, styleId } = formatAt(p, c.numberStart + 1);
    out = replaceBlock(out, p.id, [replaceText(p, c.numberStart, c.numberEnd, String(n), format, styleId)]);
  }
  return out;
}

// ---------------------------------------------------------------------------
// Inserimento

export type CaptionPosition = 'above' | 'below';

export interface CaptionArgs {
  readonly label: string;
  /** Testo dopo il numero ("Grafico delle vendite"); facoltativo. */
  readonly text?: string;
  /** Predefinito come in Word: sopra per le tabelle, sotto per il resto. */
  readonly position?: CaptionPosition;
}

export function defaultPosition(label: string): CaptionPosition {
  return label === 'Tabella' ? 'above' : 'below';
}

/** Dove va la didascalia: accanto al blocco di primo livello del cursore (la tabella intera, se è in una tabella). */
function captionSlot(state: EditorState): { sectionIndex: number; index: number } | undefined {
  const loc = findParagraph(state.doc, state.selection.focus.paragraphId);
  return loc && { sectionIndex: loc.sectionIndex, index: loc.path[0]! };
}

export function canInsertCaption(state: EditorState): boolean {
  return !!captionSlot(state);
}

/**
 * Inserisce una didascalia sopra o sotto l'elemento del cursore e rinumera
 * tutte le didascalie. Il cursore va alla fine della didascalia.
 */
export function insertCaption(state: EditorState, args: CaptionArgs): Transaction | null {
  const label = args.label.trim().replace(/\s+/g, ' ');
  const slot = captionSlot(state);
  if (!label || !/^[\p{L}][\p{L}\p{M}'’]*$/u.test(label) || !slot) return null;
  const styles = ensureCaptionStyle(state.doc.styles);
  const text = args.text?.trim();
  const p = paragraph([textRun(`${label} 0${text ? `: ${text}` : ''}`)], {}, captionStyleId(styles));
  const position = args.position ?? defaultPosition(label);
  const at = slot.index + (position === 'below' ? 1 : 0);
  const sections = state.doc.sections.map((s, i) =>
    i === slot.sectionIndex ? { ...s, blocks: [...s.blocks.slice(0, at), p, ...s.blocks.slice(at)] } : s,
  );
  const doc = renumberCaptions({ ...state.doc, styles, sections });
  const inserted = findParagraph(doc, p.id)!.paragraph;
  return { state: { doc, selection: caret(p.id, paragraphLength(inserted)) }, meta: { label: 'Inserisci didascalia' } };
}

/** Testo che la finestra mostra come anteprima ("Figura 3"). */
export function captionPreview(doc: DocumentModel, label: string): string {
  return `${label} ${nextCaptionNumber(doc, label)}`;
}
