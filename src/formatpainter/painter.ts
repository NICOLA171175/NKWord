/**
 * Copia formato: le funzioni pure che leggono la formattazione dalla
 * selezione e la applicano a un'altra parte del testo, come il pennello di
 * Word.
 *
 * - Si copia sempre la formattazione del carattere (quella effettiva del
 *   primo testo selezionato, o della parola al cursore).
 * - Se la selezione è vuota o comprende un segno di paragrafo (va oltre la
 *   fine di un paragrafo), si copia anche lo stile e il formato del
 *   paragrafo, che poi si applica ai paragrafi toccati dalla destinazione.
 * - Collegamenti e commenti non sono formattazione: restano dove sono.
 */

import { isCollapsed, replaceBlock, type CharFormat, type EditorState, type Id, type ParaFormat, type Paragraph, type Transaction } from '../core';
import { applyCharChangeToRange, caretFormat, selectionFormats, targetRange } from '../editing/format';
import { locate, paragraphsBetween } from '../editing/range';
import { selectionRange } from '../editing/selection';

type Mutable<T> = { -readonly [K in keyof T]: T[K] };

/** Formattazione copiata dal pennello. */
export interface CopiedFormat {
  readonly char: CharFormat;
  /** Presente solo se è stato copiato anche il paragrafo. */
  readonly paragraph?: { readonly styleId?: Id; readonly format: ParaFormat };
}

/**
 * Valori "spenti" delle proprietà a interruttore: se il testo di partenza
 * non è in grassetto, quello di arrivo deve perdere il grassetto.
 */
const NEUTRAL: CharFormat = {
  bold: false,
  italic: false,
  underline: null,
  strike: false,
  highlight: null,
  verticalAlign: 'baseline',
  smallCaps: false,
  allCaps: false,
};

/** Proprietà del carattere che il pennello copia (non collegamenti né commenti). */
const CHAR_KEYS = ['fontFamily', 'fontSize', 'bold', 'italic', 'underline', 'strike', 'color', 'highlight', 'verticalAlign', 'smallCaps', 'allCaps'] as const;

export const LABEL = 'Copia formato';

/** Legge la formattazione da copiare dalla selezione corrente. */
export function copyFormat(state: EditorState): CopiedFormat {
  const source = selectionFormats(state)[0] ?? caretFormat(state);
  const char: Mutable<CharFormat> = {};
  for (const key of CHAR_KEYS) (char as Record<string, unknown>)[key] = source[key] ?? NEUTRAL[key];
  const [start, end] = selectionRange(state);
  if (!isCollapsed(state.selection) && start.paragraphId === end.paragraphId) return { char };
  const p = locate(state.doc, start.paragraphId).paragraph;
  return { char, paragraph: { ...(p.styleId ? { styleId: p.styleId } : {}), format: p.format } };
}

function withParagraph(p: Paragraph, copied: NonNullable<CopiedFormat['paragraph']>): Paragraph {
  const { styleId: _s, ...rest } = p;
  return { ...rest, ...(copied.styleId ? { styleId: copied.styleId } : {}), format: copied.format };
}

/**
 * Applica la formattazione copiata alla selezione o, con selezione vuota,
 * alla parola al cursore. Fuori da una parola diventa la formattazione del
 * prossimo testo digitato.
 */
export function pasteFormat(state: EditorState, copied: CopiedFormat): Transaction | null {
  const range = targetRange(state);
  let doc = state.doc;
  if (copied.paragraph) {
    const [start, end] = range ?? selectionRange(state);
    for (const { paragraph: p } of paragraphsBetween(doc, start, end)) doc = replaceBlock(doc, p.id, [withParagraph(p, copied.paragraph)]);
  }
  // `undefined` toglie il formato diretto: torna quello dello stile.
  const change = () => copied.char;
  if (range) {
    doc = applyCharChangeToRange(doc, range[0], range[1], change);
    return { state: { doc, selection: state.selection }, meta: { label: LABEL } };
  }
  const pendingFormat: Mutable<CharFormat> = { ...copied.char };
  for (const key of Object.keys(pendingFormat) as (keyof CharFormat)[]) if (pendingFormat[key] === undefined) delete pendingFormat[key];
  if (doc === state.doc) return { state: { ...state, pendingFormat }, meta: { label: LABEL, addToHistory: false } };
  return { state: { doc, selection: state.selection, pendingFormat }, meta: { label: LABEL } };
}
