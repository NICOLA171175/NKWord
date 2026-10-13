/**
 * Letture sullo stato usate dalla UI per mostrare la formattazione al
 * cursore (carattere e dimensione nella barra, pulsanti attivi).
 */

import {
  findParagraph,
  inlineLength,
  isCollapsed,
  orderedRange,
  paragraphs,
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

/**
 * Formattazioni effettive dei caratteri selezionati (o quella al cursore).
 * Servono per lasciare vuoto il carattere o la dimensione quando la
 * selezione ne contiene più d'uno, come fa Word.
 */
export function selectionCharFormats(state: EditorState): CharFormat[] {
  if (isCollapsed(state.selection)) return [charFormatAtCaret(state)];
  const [from, to] = orderedRange(state.doc, state.selection);
  const out: CharFormat[] = [];
  let inside = false;
  for (const { paragraph: p } of paragraphs(state.doc)) {
    if (p.id === from.paragraphId) inside = true;
    if (!inside) continue;
    const a = p.id === from.paragraphId ? from.offset : 0;
    const b = p.id === to.paragraphId ? to.offset : Infinity;
    let pos = 0;
    for (const inline of p.children) {
      const len = inlineLength(inline);
      if (pos < b && pos + len > a) {
        if (inline.type === 'text') out.push(resolveCharFormat(state.doc.styles, p, inline));
        else if (inline.type === 'field') out.push(resolveCharFormat(state.doc.styles, p, inline.format));
      }
      pos += len;
    }
    if (p.id === to.paragraphId) break;
  }
  return out.length ? out : [charFormatAtCaret(state)];
}

/** Valore comune a tutta la selezione, `undefined` se è misto. */
export function commonCharValue<K extends keyof CharFormat>(state: EditorState, key: K): CharFormat[K] | undefined {
  const [first, ...rest] = selectionCharFormats(state);
  const value = first![key];
  return rest.every((f) => f[key] === value) ? value : undefined;
}

export function paraFormatAtCaret(state: EditorState): ParaFormat {
  const p = focusParagraph(state);
  return p ? resolveParaFormat(state.doc.styles, p) : { ...state.doc.styles.defaults.para };
}

export function styleIdAtCaret(state: EditorState): string {
  return focusParagraph(state)?.styleId ?? state.doc.styles.defaultParagraphStyle;
}
