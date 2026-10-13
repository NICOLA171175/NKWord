/**
 * Flusso di "segni": una sequenza di paragrafi vista come in Word, cioè
 * caratteri, oggetti in linea e segni di paragrafo (¶) uno dopo l'altro. Il
 * segno di paragrafo porta le proprietà del paragrafo che chiude; togliere
 * un segno unisce il paragrafo al successivo.
 *
 * Ogni segno occupa una unità: l'indice nel flusso del punto (paragrafo,
 * offset) è l'inizio del paragrafo più l'offset, come per `Position`.
 */

import {
  newId,
  normalizeInlines,
  paragraphLength,
  type Id,
  type Inline,
  type Paragraph,
  type Position,
  type Revision,
  type TextRun,
} from '../core';

export type Token =
  /** Una unità di codice UTF-16 di un run di testo. */
  | { readonly kind: 'char'; readonly text: string; readonly run: TextRun }
  | { readonly kind: 'inline'; readonly inline: Exclude<Inline, TextRun> }
  | { readonly kind: 'mark'; readonly paragraph: Paragraph };

export function tokenize(paras: readonly Paragraph[]): Token[] {
  const out: Token[] = [];
  for (const p of paras) {
    for (const c of p.children) {
      if (c.type === 'text') {
        for (let i = 0; i < c.text.length; i++) out.push({ kind: 'char', text: c.text[i]!, run: c });
      } else out.push({ kind: 'inline', inline: c });
    }
    out.push({ kind: 'mark', paragraph: p });
  }
  return out;
}

export function revisionOf(t: Token): Revision | undefined {
  switch (t.kind) {
    case 'char':
      return t.run.revision;
    case 'inline':
      return t.inline.revision;
    case 'mark':
      return t.paragraph.revision;
  }
}

function setRevision<T extends { readonly revision?: Revision }>(x: T, revision: Revision | undefined): T {
  if (x.revision === revision) return x;
  const { revision: _old, ...rest } = x;
  return (revision ? { ...rest, revision } : rest) as T;
}

export function withRevision(t: Token, revision: Revision | undefined): Token {
  switch (t.kind) {
    case 'char':
      return { ...t, run: setRevision(t.run, revision) };
    case 'inline':
      return { ...t, inline: setRevision(t.inline, revision) };
    case 'mark':
      return { ...t, paragraph: setRevision(t.paragraph, revision) };
  }
}

/** Stesso contenuto, ignorando formattazione e revisione. */
export function sameContent(a: Token, b: Token): boolean {
  if (a.kind === 'char') return b.kind === 'char' && a.text === b.text;
  if (a.kind === 'mark' || b.kind !== 'inline') return a.kind === b.kind;
  const x = a.inline;
  const y = b.inline;
  if (x.type !== y.type) return false;
  switch (x.type) {
    case 'break':
      return x.kind === (y as typeof x).kind;
    case 'field':
      return x.kind === (y as typeof x).kind;
    case 'image':
      return x.src === (y as typeof x).src;
    default:
      return true;
  }
}

/**
 * Ricostruisce i paragrafi. Il flusso deve finire con un segno di paragrafo.
 * Un id già in `usedIds` viene sostituito da uno nuovo (succede quando un
 * paragrafo diviso o unito ricompare due volte).
 */
export function buildParagraphs(tokens: readonly Token[], usedIds: Set<Id> = new Set()): Paragraph[] {
  const out: Paragraph[] = [];
  let inlines: Inline[] = [];
  for (const t of tokens) {
    if (t.kind === 'char') inlines.push({ ...t.run, text: t.text });
    else if (t.kind === 'inline') inlines.push(t.inline);
    else {
      const id = usedIds.has(t.paragraph.id) ? newId() : t.paragraph.id;
      usedIds.add(id);
      out.push({ ...t.paragraph, id, children: normalizeInlines(inlines) });
      inlines = [];
    }
  }
  if (inlines.length) throw new Error('Flusso senza segno di paragrafo finale');
  return out;
}

/** Indice nel flusso di una posizione, o undefined se il paragrafo non c'è. */
export function indexOf(paras: readonly Paragraph[], pos: Position): number | undefined {
  let start = 0;
  for (const p of paras) {
    const len = paragraphLength(p);
    if (p.id === pos.paragraphId) return start + Math.min(pos.offset, len);
    start += len + 1;
  }
  return undefined;
}

/** Posizione corrispondente a un indice del flusso (prima del segno di indice `index`). */
export function positionAt(paras: readonly Paragraph[], index: number): Position {
  let rest = Math.max(0, index);
  for (const p of paras) {
    const len = paragraphLength(p);
    if (rest <= len) return { paragraphId: p.id, offset: rest };
    rest -= len + 1;
  }
  const last = paras[paras.length - 1]!;
  return { paragraphId: last.id, offset: paragraphLength(last) };
}
