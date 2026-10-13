/**
 * Rilevamento modifiche. Il modulo di editing modifica il documento come
 * sempre; qui si confronta il documento prima e dopo la transazione e si
 * riscrive il risultato come farebbe Word:
 *
 * - il contenuto nuovo riceve una revisione di inserimento;
 * - il contenuto tolto torna al suo posto con una revisione di eliminazione
 *   (barrato), tranne quello che era a sua volta un inserimento non ancora
 *   accettato, che sparisce davvero;
 * - il contenuto già eliminato resta com'è (Backspace lo "salta").
 *
 * Il confronto si fa solo sui blocchi cambiati (il modello è immutabile:
 * i blocchi intatti sono gli stessi oggetti) e, dentro questi, sul flusso
 * di segni (vedi `stream.ts`): prefisso e suffisso comuni restano, la parte
 * centrale è la modifica. Inserire o togliere tabelle intere o sezioni non
 * viene rilevato.
 */

import {
  paragraphs,
  type Block,
  type DocumentModel,
  type EditorState,
  type Id,
  type Paragraph,
  type Position,
  type Revision,
  type Selection,
  type Table,
} from '../core';
import { buildParagraphs, indexOf, positionAt, revisionOf, sameContent, tokenize, withRevision, type Token } from './stream';

export interface TrackOptions {
  readonly author: string;
  /** Data ISO della modifica. */
  readonly date: string;
  /** La modifica è un Backspace: il cursore resta prima del testo barrato. */
  readonly backward?: boolean;
}

interface Env {
  readonly opts: TrackOptions;
  readonly usedIds: Set<Id>;
  /** Cursore del nuovo stato: limita il suffisso comune (vedi `trackRegion`). */
  readonly caret: Position;
  /** Traduttori di posizioni per le regioni riscritte. */
  readonly mappers: ((pos: Position) => Position | undefined)[];
}

/** Revisione da usare: quella del vicino se è dello stesso tipo e autore, così i run si uniscono. */
function pick(kind: Revision['kind'], neighbours: readonly (Revision | undefined)[], opts: TrackOptions): Revision {
  for (const r of neighbours) if (r && r.kind === kind && r.author === opts.author) return r;
  return { kind, author: opts.author, date: opts.date };
}

function trackRegion(oldParas: readonly Paragraph[], newParas: readonly Paragraph[], env: Env): Paragraph[] {
  const o = tokenize(oldParas);
  const n = tokenize(newParas);
  // Il cursore segna la fine della modifica: il suffisso comune non lo
  // supera. Così in "ab" → "abb" la "b" nuova è quella prima del cursore.
  const caretIndex = indexOf(newParas, env.caret);
  const maxSuffix = Math.min(o.length, n.length - (caretIndex ?? 0));
  let suf = 0;
  while (suf < maxSuffix && sameContent(o[o.length - 1 - suf]!, n[n.length - 1 - suf]!)) suf++;
  let pre = 0;
  while (pre < o.length - suf && pre < n.length - suf && sameContent(o[pre]!, n[pre]!)) pre++;

  const oldMid = o.slice(pre, o.length - suf);
  const newMid = n.slice(pre, n.length - suf);
  const prefix = n.slice(0, pre).map((t, i) => withRevision(t, revisionOf(o[i]!)));
  const suffix = n.slice(n.length - suf).map((t, i) => withRevision(t, revisionOf(o[o.length - suf + i]!)));

  const before = prefix.length ? revisionOf(prefix[prefix.length - 1]!) : undefined;
  const after = suffix.length ? revisionOf(suffix[0]!) : undefined;

  const deleteRev = pick('delete', [before, after], env.opts);
  const deleted: Token[] = [];
  for (const t of oldMid) {
    const rev = revisionOf(t);
    if (rev?.kind === 'insert') continue;
    deleted.push(rev?.kind === 'delete' ? t : withRevision(t, deleteRev));
  }
  const insertRev = pick('insert', [before, after], env.opts);
  const inserted = newMid.map((t) => withRevision(t, insertRev));

  for (const p of newParas) env.usedIds.delete(p.id);
  const result = buildParagraphs([...prefix, ...deleted, ...inserted, ...suffix], env.usedIds);

  const kept = deleted.length;
  const backwardDelete = newMid.length === 0 && !!env.opts.backward;
  env.mappers.push((pos) => {
    const i = indexOf(newParas, pos);
    if (i === undefined) return undefined;
    const r = i < pre || (i === pre && backwardDelete) ? i : i + kept;
    return positionAt(result, r);
  });
  return result;
}

function trackTable(o: Table, n: Table, env: Env): Table {
  const sameShape =
    o.rows.length === n.rows.length &&
    o.rows.every((r, i) => {
      const nr = n.rows[i]!;
      return r.id === nr.id && r.cells.length === nr.cells.length && r.cells.every((c, j) => c.id === nr.cells[j]!.id);
    });
  if (!sameShape) return n;
  let changed = false;
  const rows = n.rows.map((row, i) => {
    const cells = row.cells.map((cell, j) => {
      const blocks = trackBlocks(o.rows[i]!.cells[j]!.blocks, cell.blocks, env);
      if (blocks === cell.blocks) return cell;
      changed = true;
      return { ...cell, blocks };
    });
    return cells.some((c, j) => c !== row.cells[j]) ? { ...row, cells } : row;
  });
  return changed ? { ...n, rows } : n;
}

const isParagraph = (b: Block): b is Paragraph => b.type === 'paragraph';

function trackBlocks(old: readonly Block[], neu: readonly Block[], env: Env): readonly Block[] {
  if (old === neu) return neu;
  let start = 0;
  while (start < old.length && start < neu.length && old[start] === neu[start]) start++;
  let endO = old.length;
  let endN = neu.length;
  while (endO > start && endN > start && old[endO - 1] === neu[endN - 1]) {
    endO--;
    endN--;
  }
  const oMid = old.slice(start, endO);
  const nMid = neu.slice(start, endN);
  let mid: readonly Block[];
  if (nMid.length && oMid.every(isParagraph) && nMid.every(isParagraph)) {
    mid = trackRegion(oMid, nMid, env);
  } else if (oMid.length === 1 && nMid.length === 1 && oMid[0]!.type === 'table' && nMid[0]!.type === 'table' && oMid[0]!.id === nMid[0]!.id) {
    mid = [trackTable(oMid[0]!, nMid[0]!, env)];
  } else return neu;
  const out = [...neu.slice(0, start), ...mid, ...neu.slice(endN)];
  // Nulla di nuovo (es. Backspace su testo già barrato): si torna al vecchio array.
  return out.length === old.length && out.every((b, i) => sameBlock(b, old[i]!)) ? old : out;
}

function sameBlock(a: Block, b: Block): boolean {
  if (a === b) return true;
  if (a.type !== 'paragraph' || b.type !== 'paragraph') return false;
  if (a.id !== b.id || a.children.length !== b.children.length) return false;
  const { children: ac, ...ar } = a;
  const { children: bc, ...br } = b;
  return JSON.stringify(ar) === JSON.stringify(br) && JSON.stringify(ac) === JSON.stringify(bc);
}

/**
 * Riscrive `next` (risultato di una modifica a partire da `prev`) con le
 * revisioni. Restituisce `next` se non c'è nulla da rilevare.
 */
export function trackChanges(prev: EditorState, next: EditorState, opts: TrackOptions): EditorState {
  const a = prev.doc;
  const b = next.doc;
  if (a === b || a.sections.length !== b.sections.length) return next;
  const usedIds = new Set<Id>();
  for (const loc of paragraphs(b)) usedIds.add(loc.paragraph.id);
  const env: Env = { opts, usedIds, caret: next.selection.focus, mappers: [] };

  let changed = false;
  const sections = b.sections.map((s, i) => {
    const o = a.sections[i]!;
    if (o.id !== s.id || o.blocks === s.blocks) return s;
    const blocks = trackBlocks(o.blocks, s.blocks, env);
    if (blocks === s.blocks) return s;
    changed = true;
    // Se il risultato coincide col vecchio contenuto, si riusa la vecchia sezione.
    return blocks === o.blocks && sectionPropsEqual(o, s) ? o : { ...s, blocks };
  });
  if (!changed) return next;

  const unchanged =
    sections.every((s, i) => s === a.sections[i]) &&
    b.styles === a.styles &&
    b.lists === a.lists &&
    b.media === a.media;
  const doc: DocumentModel = unchanged ? a : { ...b, sections };
  const map = (pos: Position): Position => {
    for (const m of env.mappers) {
      const r = m(pos);
      if (r) return r;
    }
    return pos;
  };
  const selection: Selection = { anchor: map(next.selection.anchor), focus: map(next.selection.focus) };
  return { ...next, doc, selection };
}

function sectionPropsEqual(a: { readonly blocks: unknown }, b: { readonly blocks: unknown }): boolean {
  const { blocks: _a, ...ra } = a;
  const { blocks: _b, ...rb } = b;
  const ka = Object.keys(ra) as (keyof typeof ra)[];
  return ka.length === Object.keys(rb).length && ka.every((k) => ra[k] === rb[k as keyof typeof rb]);
}
