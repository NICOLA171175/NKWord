/** Formattazione di paragrafo e applicazione degli stili. */

import {
  cmToPt,
  paragraphLength,
  replaceBlock,
  resolveParaFormat,
  type EditorState,
  type ParaFormat,
  type Paragraph,
  type DocumentModel,
  type Position,
  type Transaction,
} from '../core';
import { targetRange } from './format';
import { mapFormatInRange } from './inlines';
import { paragraphsBetween } from './range';
import { selectionRange } from './selection';

/** Passo di "Aumenta/Riduci rientro" (1,25 cm come in Word italiano). */
export const INDENT_STEP = cmToPt(1.25);

export function selectedParagraphs(state: EditorState): Paragraph[] {
  const [start, end] = selectionRange(state);
  return paragraphsBetween(state.doc, start, end).map((l) => l.paragraph);
}

export function mapParagraphs(
  state: EditorState,
  fn: (p: Paragraph) => Paragraph,
  label: string,
): Transaction | null {
  let doc = state.doc;
  for (const p of selectedParagraphs(state)) {
    const next = fn(p);
    if (next !== p) doc = replaceBlock(doc, p.id, [next]);
  }
  if (doc === state.doc) return null;
  return { state: { ...state, doc }, meta: { label } };
}

export function setParaFormat(state: EditorState, format: ParaFormat, label: string): Transaction | null {
  return mapParagraphs(state, (p) => ({ ...p, format: { ...p.format, ...format } }), label);
}

/** Rientro sinistro al multiplo successivo (o precedente) del passo. */
export function shiftIndent(state: EditorState, dir: 1 | -1): Transaction | null {
  return mapParagraphs(
    state,
    (p) => {
      const current = resolveParaFormat(state.doc.styles, p).indentLeft ?? 0;
      const steps = current / INDENT_STEP;
      // Tolleranza: rientri arrivati da file (twip, mm) non sono multipli esatti.
      const eps = 1e-3;
      const target = dir > 0 ? (Math.floor(steps + eps) + 1) * INDENT_STEP : (Math.ceil(steps - eps) - 1) * INDENT_STEP;
      const indentLeft = Math.max(0, target);
      return Math.abs(indentLeft - current) < 1e-9 ? p : { ...p, format: { ...p.format, indentLeft } };
    },
    dir > 0 ? 'Aumenta rientro' : 'Riduci rientro',
  );
}

export function isAlignActive(state: EditorState, align: ParaFormat['align']): boolean {
  return selectedParagraphs(state).every((p) => resolveParaFormat(state.doc.styles, p).align === align);
}

/**
 * Applica uno stile. Stile di paragrafo: ai paragrafi selezionati (la
 * formattazione diretta resta, come in Word). Stile di carattere: al testo
 * selezionato o alla parola al cursore.
 */
export function applyStyle(state: EditorState, styleId: string): Transaction | null {
  const style = state.doc.styles.styles[styleId];
  if (!style) return null;
  const label = `Stile ${style.name}`;
  if (style.type === 'paragraph') {
    const isDefault = styleId === state.doc.styles.defaultParagraphStyle;
    return mapParagraphs(
      state,
      (p) => {
        if ((p.styleId ?? state.doc.styles.defaultParagraphStyle) === styleId) return p;
        const { styleId: _s, ...rest } = p;
        // Il livello di struttura viene dallo stile (es. i titoli).
        const { outlineLevel: _o, ...format } = p.format;
        return isDefault ? { ...rest, format } : { ...rest, format, styleId };
      },
      label,
    );
  }
  if (style.type === 'character') {
    const range = targetRange(state);
    if (!range) return null;
    const doc = setRunStyle(state.doc, range, styleId);
    return { state: { doc, selection: state.selection }, meta: { label } };
  }
  return null;
}

function setRunStyle(
  doc: DocumentModel,
  [start, end]: [Position, Position],
  styleId: string,
): DocumentModel {
  let out = doc;
  for (const { paragraph: p } of paragraphsBetween(doc, start, end)) {
    const from = p.id === start.paragraphId ? start.offset : 0;
    const to = p.id === end.paragraphId ? end.offset : paragraphLength(p);
    out = replaceBlock(out, p.id, [mapFormatInRange(p, from, to, (format) => ({ format, styleId }))]);
  }
  return out;
}
