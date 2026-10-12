/**
 * Letture sullo stato usate dalla UI per mostrare la formattazione al
 * cursore (carattere e dimensione nella barra, pulsanti attivi).
 */

import {
  findParagraph,
  inlineLength,
  resolveCharFormat,
  resolveParaFormat,
  type CharFormat,
  type EditorState,
  type ParaFormat,
  type Paragraph,
} from '../core';

export function focusParagraph(state: EditorState): Paragraph | undefined {
  return findParagraph(state.doc, state.selection.focus.paragraphId)?.paragraph;
}

/** Formattazione del carattere effettiva al cursore (quella che avrebbe il testo digitato). */
export function charFormatAtCaret(state: EditorState): CharFormat {
  const p = focusParagraph(state);
  if (!p) return { ...state.doc.styles.defaults.char, ...state.pendingFormat };
  const offset = state.selection.focus.offset;
  let pos = 0;
  let base: CharFormat | undefined;
  // Come in Word, il testo digitato eredita dal carattere che precede il cursore.
  for (const inline of p.children) {
    const len = inlineLength(inline);
    if (inline.type === 'text' && (offset > pos || (offset === 0 && pos === 0)) && offset <= pos + len) {
      base = resolveCharFormat(state.doc.styles, p, inline);
    }
    pos += len;
    if (pos >= offset && base) break;
  }
  return { ...(base ?? resolveCharFormat(state.doc.styles, p)), ...state.pendingFormat };
}

export function paraFormatAtCaret(state: EditorState): ParaFormat {
  const p = focusParagraph(state);
  return p ? resolveParaFormat(state.doc.styles, p) : { ...state.doc.styles.defaults.para };
}

export function styleIdAtCaret(state: EditorState): string {
  return focusParagraph(state)?.styleId ?? state.doc.styles.defaultParagraphStyle;
}
