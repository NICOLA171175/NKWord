/** Trova e sostituisci. Le corrispondenze non attraversano i paragrafi. */

import {
  isCollapsed,
  paragraphText,
  paragraphs,
  replaceBlock,
  type DocumentModel,
  type EditorState,
  type Inline,
  type Paragraph,
  type Position,
  type Transaction,
} from '../core';
import { deleteInParagraph, formatAt, insertInParagraph, textToInlines } from './inlines';
import { clampPosition, selectionRange, stateOf } from './selection';
import { isWordChar } from './words';

export interface FindOptions {
  readonly matchCase?: boolean;
  readonly wholeWord?: boolean;
}

export interface Match {
  readonly paragraphId: string;
  readonly from: number;
  readonly to: number;
}

const escapeRegExp = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/** Corrispondenze non sovrapposte di `query` nel testo. */
export function matchesInText(text: string, query: string, opts: FindOptions = {}): [number, number][] {
  if (!query) return [];
  const re = new RegExp(escapeRegExp(query), opts.matchCase ? 'gu' : 'giu');
  const out: [number, number][] = [];
  for (const m of text.matchAll(re)) {
    const from = m.index;
    const to = from + m[0].length;
    if (opts.wholeWord && (isWordChar(text[from - 1]) || isWordChar(text[to]))) continue;
    out.push([from, to]);
  }
  return out;
}

export function findAll(doc: DocumentModel, query: string, opts: FindOptions = {}): Match[] {
  const out: Match[] = [];
  for (const { paragraph: p } of paragraphs(doc)) {
    for (const [from, to] of matchesInText(paragraphText(p), query, opts)) out.push({ paragraphId: p.id, from, to });
  }
  return out;
}

/** Prima corrispondenza dopo `from`, ricominciando dall'inizio se serve. */
export function findNext(doc: DocumentModel, query: string, opts: FindOptions, from: Position): Match | null {
  const all = findAll(doc, query, opts);
  if (!all.length) return null;
  const order = [...paragraphs(doc)].map((l) => l.paragraph.id);
  const fromIndex = order.indexOf(from.paragraphId);
  const after = all.find((m) => {
    const i = order.indexOf(m.paragraphId);
    return i > fromIndex || (i === fromIndex && m.from >= from.offset);
  });
  return after ?? all[0]!;
}

function selectMatch(doc: DocumentModel, m: Match): EditorState {
  return stateOf(doc, {
    anchor: { paragraphId: m.paragraphId, offset: m.from },
    focus: { paragraphId: m.paragraphId, offset: m.to },
  });
}

export function find(state: EditorState, query: string, opts: FindOptions, from?: Position): Transaction | null {
  const start = from ?? selectionRange(state)[1];
  const m = findNext(state.doc, query, opts, start);
  if (!m) return null;
  return { state: selectMatch(state.doc, m), meta: { addToHistory: false } };
}

function replacementInlines(text: string, p: Paragraph, at: number): Inline[] {
  const base = formatAt(p, at);
  const out: Inline[] = [];
  text.replace(/\r\n?/g, '\n').split('\n').forEach((line, i) => {
    if (i > 0) out.push({ type: 'break', kind: 'line' });
    out.push(...textToInlines(line, base.format, base.styleId));
  });
  return out;
}

/** Sostituisce [from, to) mantenendo la formattazione del primo carattere sostituito. */
function replaceInParagraph(p: Paragraph, from: number, to: number, text: string): [Paragraph, number] {
  const inlines = replacementInlines(text, p, from + 1);
  const next = insertInParagraph(deleteInParagraph(p, from, to), from, inlines);
  let len = 0;
  for (const i of inlines) len += i.type === 'text' ? i.text.length : 1;
  return [next, len];
}

export function replace(
  state: EditorState,
  query: string,
  replacement: string,
  opts: FindOptions & { all?: boolean },
): Transaction | null {
  if (!query) return null;
  if (opts.all) {
    let doc = state.doc;
    let count = 0;
    for (const { paragraph: p } of paragraphs(state.doc)) {
      const matches = matchesInText(paragraphText(p), query, opts);
      if (!matches.length) continue;
      let next = p;
      for (const [from, to] of [...matches].reverse()) next = replaceInParagraph(next, from, to, replacement)[0];
      doc = replaceBlock(doc, p.id, [next]);
      count += matches.length;
    }
    if (!count) return null;
    const anchor = clampPosition(doc, state.selection.anchor)!;
    const focus = clampPosition(doc, state.selection.focus)!;
    return { state: stateOf(doc, { anchor, focus }), meta: { label: 'Sostituisci tutto' } };
  }

  // Sostituisci: se la selezione è una corrispondenza la sostituisce e passa
  // alla successiva; altrimenti cerca la prossima, come in Word.
  const [start, end] = selectionRange(state);
  const selected =
    !isCollapsed(state.selection) && start.paragraphId === end.paragraphId
      ? matchesInText(
          paragraphText([...paragraphs(state.doc)].find((l) => l.paragraph.id === start.paragraphId)!.paragraph),
          query,
          opts,
        ).some(([a, b]) => a === start.offset && b === end.offset)
      : false;
  if (!selected) return find(state, query, opts);

  const p = [...paragraphs(state.doc)].find((l) => l.paragraph.id === start.paragraphId)!.paragraph;
  const [next, len] = replaceInParagraph(p, start.offset, end.offset, replacement);
  const doc = replaceBlock(state.doc, p.id, [next]);
  const after = { paragraphId: p.id, offset: start.offset + len };
  const m = findNext(doc, query, opts, after);
  const nextState = m ? selectMatch(doc, m) : stateOf(doc, { anchor: after, focus: after });
  return { state: nextState, meta: { label: 'Sostituisci' } };
}
