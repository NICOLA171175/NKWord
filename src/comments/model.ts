/**
 * Commenti nel modello: funzioni pure. Un commento vive in
 * `DocumentModel.comments`; il testo a cui si riferisce porta il suo id in
 * `CharFormat.commentIds` (id separati da spazio). Le risposte hanno
 * `parentId` e si ancorano al testo del commento principale.
 *
 * Un commento principale il cui testo è stato cancellato non ha più
 * ancoraggio: resta nel modello (così Annulla lo fa ricomparire) ma non si
 * mostra e non si esporta.
 */

import {
  normalizeInlines,
  paragraphLength,
  paragraphs,
  paragraphText,
  updateParagraph,
  type CharFormat,
  type Comment,
  type DocumentModel,
  type Id,
  type Inline,
  type Paragraph,
  type Position,
} from '../core';

// ---------------------------------------------------------------------------
// Id nel formato del carattere

export function commentIdsOf(format: CharFormat | undefined): string[] {
  return format?.commentIds ? format.commentIds.split(' ').filter(Boolean) : [];
}

/** Restituisce il formato con la lista di id data (senza il campo se vuota). */
export function withCommentIds(format: CharFormat, ids: readonly string[]): CharFormat {
  const { commentIds: _old, ...rest } = format;
  const unique = [...new Set(ids)].sort();
  return unique.length ? { ...rest, commentIds: unique.join(' ') } : rest;
}

function inlineIds(inline: Inline): string[] {
  return inline.type === 'text' || inline.type === 'field' ? commentIdsOf(inline.format) : [];
}

/** Applica `change` agli id dei run di testo fra gli offset `from` e `to` del paragrafo. */
function mapIdsInRange(p: Paragraph, from: number, to: number, change: (ids: string[]) => string[]): Paragraph {
  const out: Inline[] = [];
  let pos = 0;
  let changed = false;
  for (const inline of p.children) {
    const len = inline.type === 'text' ? inline.text.length : 1;
    const start = pos;
    const end = pos + len;
    pos = end;
    if (end <= from || start >= to || len === 0) {
      out.push(inline);
      continue;
    }
    if (inline.type === 'text') {
      const a = Math.max(from, start) - start;
      const b = Math.min(to, end) - start;
      const format = withCommentIds(inline.format, change(commentIdsOf(inline.format)));
      if (a > 0) out.push({ ...inline, text: inline.text.slice(0, a) });
      out.push({ ...inline, text: inline.text.slice(a, b), format });
      if (b < len) out.push({ ...inline, text: inline.text.slice(b) });
      changed = true;
    } else if (inline.type === 'field') {
      out.push({ ...inline, format: withCommentIds(inline.format, change(commentIdsOf(inline.format))) });
      changed = true;
    } else out.push(inline);
  }
  return changed ? { ...p, children: normalizeInlines(out) } : p;
}

// ---------------------------------------------------------------------------
// Ancoraggi

export interface CommentRange {
  readonly from: Position;
  readonly to: Position;
}

/** Intervallo di testo di ogni commento ancorato, in ordine di documento. */
export function commentRanges(doc: DocumentModel): Map<Id, CommentRange> {
  const ranges = new Map<Id, { from: Position; to: Position }>();
  for (const { paragraph: p } of paragraphs(doc)) {
    let pos = 0;
    for (const inline of p.children) {
      const len = inline.type === 'text' ? inline.text.length : 1;
      for (const id of inlineIds(inline)) {
        const r = ranges.get(id);
        const end = { paragraphId: p.id, offset: pos + len };
        if (r) r.to = end;
        else ranges.set(id, { from: { paragraphId: p.id, offset: pos }, to: end });
      }
      pos += len;
    }
  }
  return ranges;
}

/** Id dei commenti ancorati alla posizione (cursore dentro o ai bordi del testo). */
export function commentsAt(doc: DocumentModel, pos: Position): Id[] {
  for (const { paragraph: p } of paragraphs(doc)) {
    if (p.id !== pos.paragraphId) continue;
    const ids = new Set<Id>();
    let at = 0;
    for (const inline of p.children) {
      const len = inline.type === 'text' ? inline.text.length : 1;
      if (at <= pos.offset && pos.offset <= at + len) for (const id of inlineIds(inline)) ids.add(id);
      at += len;
    }
    return [...ids].filter((id) => doc.comments?.[id] && !doc.comments[id]!.parentId);
  }
  return [];
}

export interface CommentThread {
  readonly comment: Comment;
  readonly replies: readonly Comment[];
  readonly range: CommentRange;
}

/** Commenti principali ancorati, in ordine di documento, con le risposte in ordine di data. */
export function commentThreads(doc: DocumentModel): CommentThread[] {
  const all = Object.values(doc.comments ?? {});
  if (!all.length) return [];
  const out: CommentThread[] = [];
  for (const [id, range] of commentRanges(doc)) {
    const comment = doc.comments?.[id];
    if (!comment || comment.parentId) continue;
    const replies = all.filter((c) => c.parentId === id).sort((a, b) => a.date.localeCompare(b.date));
    out.push({ comment, replies, range });
  }
  return out;
}

// ---------------------------------------------------------------------------
// Modifiche

function setComments(doc: DocumentModel, comments: Record<Id, Comment>): DocumentModel {
  const { comments: _old, ...rest } = doc;
  return Object.keys(comments).length ? { ...rest, comments } : rest;
}

export function initialsOf(name: string): string {
  return name
    .split(/\s+/)
    .filter(Boolean)
    .map((w) => w[0]!.toUpperCase())
    .join('')
    .slice(0, 3);
}

/**
 * Ancora un nuovo commento al testo fra `from` e `to` (ordinati). Restituisce
 * null se l'intervallo non contiene testo.
 */
export function addComment(doc: DocumentModel, from: Position, to: Position, comment: Comment): DocumentModel | null {
  let next = doc;
  let inside = false;
  let anchored = false;
  for (const { paragraph: p } of paragraphs(doc)) {
    const isFrom = p.id === from.paragraphId;
    const isTo = p.id === to.paragraphId;
    if (isFrom) inside = true;
    if (inside) {
      const a = isFrom ? from.offset : 0;
      const b = isTo ? to.offset : paragraphLength(p);
      if (b > a) {
        const updated = mapIdsInRange(p, a, b, (ids) => [...ids, comment.id]);
        if (updated !== p) {
          anchored = true;
          next = updateParagraph(next, updated);
        }
      }
    }
    if (isTo) break;
  }
  if (!anchored) return null;
  return setComments(next, { ...doc.comments, [comment.id]: comment });
}

export function addReply(doc: DocumentModel, reply: Comment & { parentId: Id }): DocumentModel | null {
  if (!doc.comments?.[reply.parentId]) return null;
  return setComments(doc, { ...doc.comments, [reply.id]: reply });
}

export function updateComment(doc: DocumentModel, id: Id, patch: Partial<Omit<Comment, 'id'>>): DocumentModel | null {
  const c = doc.comments?.[id];
  if (!c) return null;
  return setComments(doc, { ...doc.comments, [id]: { ...c, ...patch } });
}

/** Elimina un commento: se è principale, anche le sue risposte e il suo ancoraggio. */
export function deleteComment(doc: DocumentModel, id: Id): DocumentModel | null {
  const comments = { ...doc.comments };
  if (!comments[id]) return null;
  const removed = new Set([id]);
  for (const c of Object.values(comments)) if (c.parentId === id) removed.add(c.id);
  for (const r of removed) delete comments[r];
  let next = doc;
  for (const { paragraph: p } of paragraphs(doc)) {
    if (!p.children.some((i) => inlineIds(i).some((x) => removed.has(x)))) continue;
    next = updateParagraph(next, mapIdsInRange(p, 0, paragraphLength(p), (ids) => ids.filter((x) => !removed.has(x))));
  }
  return setComments(next, comments);
}

// ---------------------------------------------------------------------------
// Parola al cursore

const WORD = /[\p{L}\p{N}_'’]/u;

/** Intervallo della parola che contiene o tocca l'offset, oppure null. */
export function wordAt(p: Paragraph, offset: number): [number, number] | null {
  const text = paragraphText(p);
  let a = offset;
  let b = offset;
  while (a > 0 && WORD.test(text[a - 1]!)) a--;
  while (b < text.length && WORD.test(text[b]!)) b++;
  return b > a ? [a, b] : null;
}
