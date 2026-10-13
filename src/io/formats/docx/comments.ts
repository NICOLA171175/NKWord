/**
 * Commenti in DOCX: `word/comments.xml` (testo, autore, data) e
 * `word/commentsExtended.xml` (risposte e stato "risolto", Word 2013+).
 * Nel corpo il testo commentato sta fra `w:commentRangeStart` e
 * `w:commentRangeEnd`, seguiti da un run con `w:commentReference`.
 */

import { newId, paragraphs, type Comment, type DocumentModel, type Id } from '../../../core';
import { elements, escapeAttr, escapeText, textContent, type MarkupElement } from '../../util/markup';
import { NS } from './common';

export const COMMENTS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
export const COMMENTS_EXT_REL = 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended';
export const COMMENTS_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml';
export const COMMENTS_EXT_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const idsOf = (s: string | undefined) => (s ? s.split(' ').filter(Boolean) : []);

// ---------------------------------------------------------------------------
// Scrittura

/** Inizio e fine dei commenti per paragrafo: indice del figlio → id DOCX numerici. */
export interface CommentMarks {
  readonly starts: Map<string, Map<number, number[]>>;
  readonly ends: Map<string, Map<number, number[]>>;
  /** Commenti da scrivere, nell'ordine dei loro id DOCX (0, 1, 2...). */
  readonly comments: readonly Comment[];
}

export function commentMarks(doc: DocumentModel): CommentMarks | null {
  const all = doc.comments;
  if (!all || !Object.keys(all).length) return null;
  const first = new Map<Id, [string, number]>();
  const last = new Map<Id, [string, number]>();
  for (const { paragraph: p } of paragraphs(doc)) {
    p.children.forEach((c, i) => {
      if (c.type !== 'text' && c.type !== 'field') return;
      for (const id of idsOf(c.format.commentIds)) {
        if (!all[id] || all[id]!.parentId) continue;
        if (!first.has(id)) first.set(id, [p.id, i]);
        last.set(id, [p.id, i]);
      }
    });
  }
  if (!first.size) return null;
  const comments: Comment[] = [];
  const starts = new Map<string, Map<number, number[]>>();
  const ends = new Map<string, Map<number, number[]>>();
  const mark = (map: Map<string, Map<number, number[]>>, [pid, i]: [string, number], n: number) => {
    let byIndex = map.get(pid);
    if (!byIndex) map.set(pid, (byIndex = new Map()));
    byIndex.set(i, [...(byIndex.get(i) ?? []), n]);
  };
  const replies = Object.values(all).filter((c) => c.parentId).sort((a, b) => a.date.localeCompare(b.date));
  for (const [id, start] of first) {
    // Come fa Word, le risposte condividono l'intervallo del commento principale.
    for (const c of [all[id]!, ...replies.filter((r) => r.parentId === id)]) {
      const n = comments.length;
      comments.push(c);
      mark(starts, start, n);
      mark(ends, last.get(id)!, n);
    }
  }
  return { starts, ends, comments };
}

export function commentRangeStartXml(ids: readonly number[] | undefined): string {
  return (ids ?? []).map((n) => `<w:commentRangeStart w:id="${n}"/>`).join('');
}

export function commentRangeEndXml(ids: readonly number[] | undefined): string {
  return (ids ?? []).map((n) => `<w:commentRangeEnd w:id="${n}"/><w:r><w:commentReference w:id="${n}"/></w:r>`).join('');
}

/** paraId di Word: 8 cifre esadecimali, minori di 0x80000000. */
const paraId = (n: number) => (n + 1).toString(16).toUpperCase().padStart(8, '0');

export function commentsXml(comments: readonly Comment[]): string {
  const items = comments.map((c, n) => {
    const lines = c.text.split('\n');
    const paras = lines.map((line, i) => {
      const id = i === lines.length - 1 ? ` w14:paraId="${paraId(n)}"` : '';
      const ref = i === 0 ? '<w:r><w:annotationRef/></w:r>' : '';
      const run = line ? `<w:r><w:t xml:space="preserve">${escapeText(line)}</w:t></w:r>` : '';
      return `<w:p${id}>${ref}${run}</w:p>`;
    });
    const initials = c.initials ? ` w:initials="${escapeAttr(c.initials)}"` : '';
    return `<w:comment w:id="${n}" w:author="${escapeAttr(c.author)}" w:date="${escapeAttr(c.date)}"${initials}>${paras.join('')}</w:comment>`;
  });
  return `${XML_DECL}<w:comments xmlns:w="${NS.w}" xmlns:w14="${W14}" xmlns:mc="${MC}" mc:Ignorable="w14">${items.join('')}</w:comments>`;
}

export function commentsExtendedXml(comments: readonly Comment[]): string {
  const index = new Map(comments.map((c, n) => [c.id, n]));
  const items = comments.map((c, n) => {
    const parent = c.parentId !== undefined ? index.get(c.parentId) : undefined;
    return `<w15:commentEx w15:paraId="${paraId(n)}"${parent !== undefined ? ` w15:paraIdParent="${paraId(parent)}"` : ''} w15:done="${c.resolved ? 1 : 0}"/>`;
  });
  return `${XML_DECL}<w15:commentsEx xmlns:w15="${W15}" xmlns:mc="${MC}" mc:Ignorable="w15">${items.join('')}</w15:commentsEx>`;
}

// ---------------------------------------------------------------------------
// Lettura

const attrBySuffix = (el: MarkupElement, suffix: string) => {
  for (const [k, v] of Object.entries(el.attrs)) if (k === suffix || k.endsWith(`:${suffix}`)) return v;
  return undefined;
};

export interface ImportedComments {
  readonly comments: Record<Id, Comment>;
  /** id DOCX → id del modello, solo per i commenti principali (quelli che si ancorano al testo). */
  readonly anchors: Map<string, Id>;
}

export function readComments(root: MarkupElement | undefined, extended: MarkupElement | undefined): ImportedComments {
  const comments: Record<Id, Comment> = {};
  const anchors = new Map<string, Id>();
  if (!root) return { comments, anchors };

  // commentsExtended: paraId dell'ultimo paragrafo → genitore e stato.
  const ext = new Map<string, { parent?: string; done: boolean }>();
  for (const el of extended ? elements(extended) : []) {
    if (!el.name.endsWith('commentEx')) continue;
    const id = attrBySuffix(el, 'paraId');
    if (id) ext.set(id, { parent: attrBySuffix(el, 'paraIdParent'), done: attrBySuffix(el, 'done') === '1' });
  }

  const byParaId = new Map<string, Id>();
  const parsed: { docxId: string; comment: Comment; parentPara?: string }[] = [];
  for (const el of elements(root, 'w:comment')) {
    const docxId = el.attrs['w:id'];
    if (docxId === undefined) continue;
    const paras = elements(el, 'w:p');
    const text = paras.map((p) => textContent(p).replace(/[\r\n\t]+/g, ' ')).join('\n').trim();
    const lastPara = paras[paras.length - 1];
    const pid = lastPara && attrBySuffix(lastPara, 'paraId');
    const info = pid ? ext.get(pid) : undefined;
    const id = newId();
    if (pid) byParaId.set(pid, id);
    const author = el.attrs['w:author'] ?? '';
    const initials = el.attrs['w:initials'];
    parsed.push({
      docxId,
      parentPara: info?.parent,
      comment: {
        id,
        author,
        date: el.attrs['w:date'] ?? '',
        text,
        ...(initials ? { initials } : {}),
        ...(info?.done ? { resolved: true } : {}),
      },
    });
  }
  const parentOf = new Map(parsed.map((x) => [x.comment.id, x.parentPara ? byParaId.get(x.parentPara) : undefined]));
  // Nel modello si risponde sempre al commento principale della discussione.
  const rootOf = (id: Id | undefined) => {
    const seen = new Set<Id>();
    while (id && parentOf.get(id) && !seen.has(id)) {
      seen.add(id);
      id = parentOf.get(id);
    }
    return id;
  };
  for (const { docxId, comment, parentPara } of parsed) {
    const parentId = parentPara ? rootOf(byParaId.get(parentPara)) : undefined;
    if (parentId && parentId !== comment.id) comments[comment.id] = { ...comment, parentId };
    else {
      comments[comment.id] = comment;
      anchors.set(docxId, comment.id);
    }
  }
  return { comments, anchors };
}
