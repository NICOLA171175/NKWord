/**
 * Revisione: elenco delle revisioni, accetta/rifiuta, spostamento fra
 * revisioni. Funzioni pure sul modello.
 *
 * Accettare un inserimento toglie solo il segno di revisione; accettare
 * un'eliminazione toglie il contenuto. Rifiutare fa il contrario. Togliere
 * un segno di paragrafo unisce il paragrafo al successivo; l'ultimo segno
 * di una cella o di una sezione non si può togliere e perde solo la
 * revisione.
 */

import {
  paragraphs,
  type Block,
  type DocumentModel,
  type Id,
  type Paragraph,
  type Position,
  type Revision,
  type Selection,
} from '../core';
import { buildParagraphs, indexOf, positionAt, revisionOf, tokenize, withRevision } from './stream';

/** Una revisione nel documento: segni contigui con stesso tipo e autore. */
export interface RevisionSpan {
  readonly revision: Revision;
  readonly from: Position;
  readonly to: Position;
  /** Chiavi `paragrafo:offset` dei segni coperti (il segno di paragrafo ha offset = lunghezza). */
  readonly keys: ReadonlySet<string>;
}

const key = (paragraphId: Id, offset: number) => `${paragraphId}:${offset}`;

/** Sequenze di paragrafi consecutivi (corpo e celle), dove i segni di paragrafo si possono unire. */
function* paragraphRuns(blocks: readonly Block[]): Generator<Paragraph[]> {
  let run: Paragraph[] = [];
  for (const b of blocks) {
    if (b.type === 'paragraph') run.push(b);
    else {
      if (run.length) yield run;
      run = [];
      for (const row of b.rows) for (const cell of row.cells) yield* paragraphRuns(cell.blocks);
    }
  }
  if (run.length) yield run;
}

function* allRuns(doc: DocumentModel): Generator<Paragraph[]> {
  for (const s of doc.sections) yield* paragraphRuns(s.blocks);
}

/** Posizione di ciascun segno del flusso di una sequenza. */
function tokenPositions(run: readonly Paragraph[]): Position[] {
  const out: Position[] = [];
  for (const p of run) {
    const len = tokenize([p]).length;
    for (let i = 0; i < len; i++) out.push({ paragraphId: p.id, offset: i });
  }
  return out;
}

const cache = new WeakMap<DocumentModel, RevisionSpan[]>();

/** Tutte le revisioni del corpo del documento, in ordine di lettura. */
export function revisionSpans(doc: DocumentModel): readonly RevisionSpan[] {
  const cached = cache.get(doc);
  if (cached) return cached;
  const spans: RevisionSpan[] = [];
  for (const run of allRuns(doc)) {
    const tokens = tokenize(run);
    const pos = tokenPositions(run);
    let i = 0;
    while (i < tokens.length) {
      const rev = revisionOf(tokens[i]!);
      if (!rev) {
        i++;
        continue;
      }
      const keys = new Set<string>();
      let j = i;
      while (j < tokens.length) {
        const r = revisionOf(tokens[j]!);
        if (!r || r.kind !== rev.kind || r.author !== rev.author) break;
        keys.add(key(pos[j]!.paragraphId, pos[j]!.offset));
        j++;
      }
      spans.push({ revision: rev, from: pos[i]!, to: positionAt(run, j), keys });
      i = j;
    }
  }
  cache.set(doc, spans);
  return spans;
}

export function hasRevisions(doc: DocumentModel): boolean {
  return revisionSpans(doc).length > 0;
}

/** Ordine di lettura dei paragrafi, per confrontare posizioni rapidamente. */
function paragraphOrder(doc: DocumentModel): Map<Id, number> {
  const order = new Map<Id, number>();
  let i = 0;
  for (const loc of paragraphs(doc)) order.set(loc.paragraph.id, i++);
  return order;
}

function compare(order: Map<Id, number>, a: Position, b: Position): number {
  const pa = order.get(a.paragraphId) ?? -1;
  const pb = order.get(b.paragraphId) ?? -1;
  return pa !== pb ? pa - pb : a.offset - b.offset;
}

/** Segni da trattare: quelli di un intervallo [from, to) oppure un insieme di chiavi. */
export type Scope =
  | { readonly kind: 'all' }
  | { readonly kind: 'range'; readonly from: Position; readonly to: Position }
  | { readonly kind: 'keys'; readonly keys: ReadonlySet<string> };

/**
 * Revisioni interessate dalla selezione: quelle che la intersecano o, con
 * il solo cursore, quella che lo tocca.
 */
export function scopeForSelection(doc: DocumentModel, sel: Selection): Scope | null {
  const order = paragraphOrder(doc);
  const [from, to] = compare(order, sel.anchor, sel.focus) <= 0 ? [sel.anchor, sel.focus] : [sel.focus, sel.anchor];
  const spans = revisionSpans(doc);
  if (compare(order, from, to) === 0) {
    const at = key(from.paragraphId, from.offset);
    const before = key(from.paragraphId, from.offset - 1);
    const span = spans.find((s) => s.keys.has(at)) ?? spans.find((s) => s.keys.has(before));
    return span ? { kind: 'keys', keys: span.keys } : null;
  }
  const touched = spans.some((s) => compare(order, s.from, to) < 0 && compare(order, s.to, from) > 0);
  return touched ? { kind: 'range', from, to } : null;
}

/** Accetta o rifiuta le revisioni nell'ambito; la selezione segue il contenuto. */
export function resolveRevisions(
  doc: DocumentModel,
  selection: Selection,
  action: 'accept' | 'reject',
  scope: Scope,
): { doc: DocumentModel; selection: Selection } | null {
  const order = scope.kind === 'range' ? paragraphOrder(doc) : undefined;
  const inScope = (pos: Position): boolean => {
    switch (scope.kind) {
      case 'all':
        return true;
      case 'keys':
        return scope.keys.has(key(pos.paragraphId, pos.offset));
      case 'range':
        return compare(order!, scope.from, pos) <= 0 && compare(order!, pos, scope.to) < 0;
    }
  };

  const replaced = new Map<Id, Paragraph[]>();
  const mappers: ((pos: Position) => Position | undefined)[] = [];
  const usedIds = new Set<Id>();
  for (const loc of paragraphs(doc)) usedIds.add(loc.paragraph.id);

  for (const run of allRuns(doc)) {
    const tokens = tokenize(run);
    const pos = tokenPositions(run);
    const out: typeof tokens = [];
    const keptBefore: number[] = [];
    let changed = false;
    tokens.forEach((t, i) => {
      keptBefore.push(out.length);
      const rev = revisionOf(t);
      if (!rev || !inScope(pos[i]!)) {
        out.push(t);
        return;
      }
      changed = true;
      const remove = (action === 'accept') === (rev.kind === 'delete');
      const lastMark = i === tokens.length - 1;
      if (!remove || lastMark) out.push(withRevision(t, undefined));
    });
    keptBefore.push(out.length);
    if (!changed) continue;
    for (const p of run) usedIds.delete(p.id);
    const built = buildParagraphs(out, usedIds);
    replaced.set(run[0]!.id, built);
    for (const p of run.slice(1)) replaced.set(p.id, []);
    mappers.push((p) => {
      const i = indexOf(run, p);
      return i === undefined ? undefined : positionAt(built, keptBefore[i]!);
    });
  }
  if (!replaced.size) return null;

  const mapBlocks = (blocks: readonly Block[]): readonly Block[] => {
    let changed = false;
    const out: Block[] = [];
    for (const b of blocks) {
      if (b.type === 'paragraph') {
        const r = replaced.get(b.id);
        if (r) {
          changed = true;
          out.push(...r);
        } else out.push(b);
        continue;
      }
      const rows = b.rows.map((row) => {
        const cells = row.cells.map((cell) => {
          const inner = mapBlocks(cell.blocks);
          return inner === cell.blocks ? cell : { ...cell, blocks: inner };
        });
        return cells.some((c, i) => c !== row.cells[i]) ? { ...row, cells } : row;
      });
      if (rows.some((r, i) => r !== b.rows[i])) {
        changed = true;
        out.push({ ...b, rows });
      } else out.push(b);
    }
    return changed ? out : blocks;
  };
  const sections = doc.sections.map((s) => {
    const blocks = mapBlocks(s.blocks);
    return blocks === s.blocks ? s : { ...s, blocks };
  });
  const map = (p: Position): Position => {
    for (const m of mappers) {
      const r = m(p);
      if (r) return r;
    }
    return p;
  };
  return { doc: { ...doc, sections }, selection: { anchor: map(selection.anchor), focus: map(selection.focus) } };
}

/** Revisione successiva (o precedente) rispetto al cursore, ricominciando dall'inizio (o dalla fine). */
export function adjacentRevision(doc: DocumentModel, sel: Selection, direction: 'next' | 'previous'): RevisionSpan | undefined {
  const spans = revisionSpans(doc);
  if (!spans.length) return undefined;
  const order = paragraphOrder(doc);
  const from = compare(order, sel.anchor, sel.focus) <= 0 ? sel.anchor : sel.focus;
  if (direction === 'next') return spans.find((s) => compare(order, s.from, from) > 0) ?? spans[0];
  for (let i = spans.length - 1; i >= 0; i--) if (compare(order, spans[i]!.from, from) < 0) return spans[i];
  return spans[spans.length - 1];
}
