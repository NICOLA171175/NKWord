/**
 * Operazioni sugli inline di un paragrafo per offset del modello.
 * Gli offset seguono la convenzione del core: un run di testo occupa la
 * lunghezza del suo testo, ogni altro inline occupa 1.
 */

import { inlineLength, normalizeInlines, type CharFormat, type Inline, type Paragraph } from '../core';

/** Divide gli inline all'offset dato, spezzando il run di testo che lo contiene. */
export function splitInlines(inlines: readonly Inline[], offset: number): [Inline[], Inline[]] {
  const before: Inline[] = [];
  const after: Inline[] = [];
  let pos = 0;
  for (const inline of inlines) {
    const len = inlineLength(inline);
    if (pos + len <= offset) before.push(inline);
    else if (pos >= offset) after.push(inline);
    else if (inline.type === 'text') {
      const cut = offset - pos;
      before.push({ ...inline, text: inline.text.slice(0, cut) });
      after.push({ ...inline, text: inline.text.slice(cut) });
    } else after.push(inline);
    pos += len;
  }
  return [before, after];
}

/** Divide gli inline in [prima di from, from..to, da to in poi]. */
export function splitRange(
  inlines: readonly Inline[],
  from: number,
  to: number,
): [Inline[], Inline[], Inline[]] {
  const [head, rest] = splitInlines(inlines, from);
  const [middle, tail] = splitInlines(rest, to - from);
  return [head, middle, tail];
}

export function inlinesLength(inlines: readonly Inline[]): number {
  let n = 0;
  for (const i of inlines) n += inlineLength(i);
  return n;
}

/** Toglie l'intervallo [from, to) dal paragrafo. */
export function deleteInParagraph(p: Paragraph, from: number, to: number): Paragraph {
  if (from >= to) return p;
  const [head, , tail] = splitRange(p.children, from, to);
  return { ...p, children: normalizeInlines([...head, ...tail]) };
}

/** Inserisce inline all'offset dato. */
export function insertInParagraph(p: Paragraph, offset: number, inlines: readonly Inline[]): Paragraph {
  if (!inlines.length) return p;
  const [head, tail] = splitInlines(p.children, offset);
  return { ...p, children: normalizeInlines([...head, ...inlines, ...tail]) };
}

/**
 * Inline "di riferimento" per il testo digitato all'offset: quello che
 * precede il cursore, oppure il primo del paragrafo se il cursore è
 * all'inizio. Restituisce formattazione diretta e stile carattere.
 */
export function formatAt(p: Paragraph, offset: number): { format: CharFormat; styleId?: string } {
  let pos = 0;
  let candidate: Inline | undefined;
  for (const inline of p.children) {
    const len = inlineLength(inline);
    if (offset > 0 ? pos < offset && offset <= pos + len : true) {
      candidate = inline;
      break;
    }
    pos += len;
  }
  if (!candidate) return { format: p.markFormat ?? {} };
  if (candidate.type === 'text') {
    return candidate.styleId ? { format: candidate.format, styleId: candidate.styleId } : { format: candidate.format };
  }
  if (candidate.type === 'field') return { format: candidate.format };
  // Immagini, tab e interruzioni non hanno formato: si guarda il testo vicino.
  for (const inline of p.children) if (inline.type === 'text') return { format: inline.format };
  return { format: p.markFormat ?? {} };
}

/** Applica `fn` alla formattazione di run di testo e campi in [from, to). */
export function mapFormatInRange(
  p: Paragraph,
  from: number,
  to: number,
  fn: (format: CharFormat, inline: Inline) => { format: CharFormat; styleId?: string | undefined },
): Paragraph {
  if (from >= to) return p;
  const [head, middle, tail] = splitRange(p.children, from, to);
  const mapped = middle.map((inline): Inline => {
    if (inline.type !== 'text' && inline.type !== 'field') return inline;
    const { format, styleId } = fn(inline.format, inline);
    if (inline.type === 'field') return { ...inline, format };
    const { styleId: _old, ...rest } = inline;
    return styleId ? { ...rest, format, styleId } : { ...rest, format };
  });
  return { ...p, children: normalizeInlines([...head, ...mapped, ...tail]) };
}

/** Converte testo (con '\t') in inline: i tab diventano `TabInline`. */
export function textToInlines(text: string, format: CharFormat, styleId?: string): Inline[] {
  const out: Inline[] = [];
  text.split('\t').forEach((part, i) => {
    if (i > 0) out.push({ type: 'tab' });
    if (part) out.push(styleId ? { type: 'text', text: part, format, styleId } : { type: 'text', text: part, format });
  });
  return out;
}
