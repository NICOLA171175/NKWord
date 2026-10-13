/**
 * Elenchi a più livelli: stili pronti (1. 1.1. 1.1.1., I. A. 1. ...),
 * riavvio e continuazione della numerazione, valore iniziale e definizione
 * del formato di ogni livello. Le funzioni sono pure: ricevono lo stato e
 * restituiscono il documento nuovo (o `null` se non c'è nulla da fare).
 *
 * Il modello usa già `ListDefinition` a 9 livelli con testo come '%1.%2.'
 * (vedi `src/core/types.ts`), letto e scritto anche da DOCX: qui si
 * costruiscono e si assegnano le definizioni giuste.
 */

import {
  newId,
  paragraphs,
  replaceBlock,
  resolveParaFormat,
  type DocumentModel,
  type EditorState,
  type Id,
  type ListDefinition,
  type ListLevel,
  type ListRef,
  type NumberFormat,
  type Paragraph,
  type ParaFormat,
} from '../core';
import { formatNumber } from '../layout/lists';
import { selectedParagraphs } from '../structures/util';

export const LEVELS = 9;

export type PresetId = 'outline' | 'parens' | 'roman' | 'letters' | 'bullets';

export interface MultilevelPreset {
  readonly id: PresetId;
  readonly label: string;
  levels(): ListLevel[];
}

/** Rientri di Word per lo stile 1. 1.1. 1.1.1. (il numero si allunga a ogni livello). */
const OUTLINE_INDENTS: readonly [number, number][] = [
  [18, 18], [39.6, 21.6], [61.2, 28.8], [86.4, 36], [111.6, 43.2], [136.8, 50.4], [162, 57.6], [187.2, 64.8], [216, 79.2],
];

function lvl(format: NumberFormat, text: string, i: number, step = 18, hanging = 18): ListLevel {
  return { format, text, start: 1, indentLeft: step * (i + 1), hanging };
}

/** Testo '%1.%2.…%n.' del livello `i` (da 0). */
function nested(i: number, sep = '.'): string {
  return Array.from({ length: i + 1 }, (_, k) => `%${k + 1}${sep}`).join('');
}

export const PRESETS: readonly MultilevelPreset[] = [
  {
    id: 'outline',
    label: '1. 1.1. 1.1.1.',
    levels: () =>
      OUTLINE_INDENTS.map(([left, hanging], i) => ({ format: 'decimal' as const, text: nested(i), start: 1, indentLeft: left, hanging })),
  },
  {
    id: 'parens',
    label: '1) a) i)',
    levels: () => {
      const f: NumberFormat[] = ['decimal', 'lowerLetter', 'lowerRoman'];
      return Array.from({ length: LEVELS }, (_, i) => {
        const text = i < 3 ? `%${i + 1})` : i < 6 ? `(%${i + 1})` : `%${i + 1}.`;
        return lvl(f[i % 3]!, text, i);
      });
    },
  },
  {
    id: 'roman',
    label: 'I. A. 1. a) i)',
    levels: () => {
      const f: NumberFormat[] = ['upperRoman', 'upperLetter', 'decimal', 'lowerLetter', 'lowerRoman', 'decimal', 'lowerLetter', 'lowerRoman', 'decimal'];
      const t = ['%1.', '%2.', '%3.', '%4)', '%5)', '(%6)', '(%7)', '(%8)', '(%9)'];
      return f.map((format, i) => lvl(format, t[i]!, i, 27, i === 0 ? 27 : 18));
    },
  },
  {
    id: 'letters',
    label: 'A. 1. a. i.',
    levels: () => {
      const f: NumberFormat[] = ['upperLetter', 'decimal', 'lowerLetter', 'lowerRoman'];
      return Array.from({ length: LEVELS }, (_, i) => lvl(f[i % 4]!, `%${i + 1}.`, i));
    },
  },
  {
    id: 'bullets',
    label: '• ◦ ▪ (punti)',
    levels: () => {
      const b = ['•', '◦', '▪', '–', '➢', '✓'];
      return Array.from({ length: LEVELS }, (_, i) => lvl('bullet', b[i % b.length]!, i));
    },
  },
];

export function presetById(id: string): MultilevelPreset | undefined {
  return PRESETS.find((p) => p.id === id);
}

// ---------------------------------------------------------------------------
// Anteprima del marcatore

/** Testo del marcatore di un livello quando ogni contatore vale `counters[k]` (default 1). */
export function markerText(levels: readonly ListLevel[], i: number, counters: readonly number[] = []): string {
  const level = levels[i];
  if (!level) return '';
  if (level.format === 'bullet') return level.text;
  return level.text.replace(/%(\d)/g, (_, d: string) => {
    const k = Number(d) - 1;
    const lv = levels[k];
    if (!lv || lv.format === 'none') return '';
    return formatNumber(counters[k] ?? lv.start, lv.format === 'bullet' ? 'decimal' : lv.format);
  });
}

// ---------------------------------------------------------------------------
// Testo del livello: prima, numeri dei livelli superiori, numero, dopo

export interface LevelPattern {
  readonly before: string;
  /** Mostra i numeri dei livelli superiori (1.1.1). */
  readonly parents: boolean;
  readonly after: string;
}

/** Scompone '%1.%2.%3)' del livello 2 in { before: '', parents: true, after: ')' }. */
export function parsePattern(text: string, i: number): LevelPattern {
  const own = `%${i + 1}`;
  const at = text.lastIndexOf(own);
  if (at < 0) return { before: text.replace(/%\d/g, ''), parents: false, after: '' };
  const head = text.slice(0, at);
  const after = text.slice(at + own.length).replace(/%\d/g, '');
  const parentsText = i > 0 ? nested(i - 1) : '';
  if (i > 0 && head.endsWith(parentsText)) return { before: head.slice(0, -parentsText.length).replace(/%\d/g, ''), parents: true, after };
  return { before: head.replace(/%\d/g, ''), parents: false, after };
}

export function buildPattern(p: LevelPattern, i: number): string {
  const clean = (s: string) => s.replace(/%/g, '');
  return `${clean(p.before)}${p.parents && i > 0 ? nested(i - 1) : ''}%${i + 1}${clean(p.after)}`;
}

// ---------------------------------------------------------------------------
// Paragrafi degli elenchi

interface ListItem {
  readonly paragraph: Paragraph;
  readonly ref: ListRef;
}

/** Paragrafi di elenco in ordine di lettura, con il riferimento effettivo (anche da stile). */
export function listItems(doc: DocumentModel): ListItem[] {
  const out: ListItem[] = [];
  for (const { paragraph } of paragraphs(doc)) {
    const ref = resolveParaFormat(doc.styles, paragraph).list;
    if (ref && doc.lists[ref.listId]) out.push({ paragraph, ref });
  }
  return out;
}

function withList(p: Paragraph, list: ListRef | null): Paragraph {
  const { indentLeft: _l, indentFirstLine: _f, ...rest } = p.format;
  const format: ParaFormat = list ? { ...rest, list } : { ...p.format, list: null };
  return { ...p, format };
}

function isNumbered(def: ListDefinition | undefined): boolean {
  return !!def && def.levels[0]?.format !== 'bullet';
}

/** Elenco del paragrafo col cursore, se c'è. */
export function focusItem(state: EditorState): ListItem | undefined {
  const id = state.selection.focus.paragraphId;
  return listItems(state.doc).find((i) => i.paragraph.id === id);
}

// ---------------------------------------------------------------------------
// Applica uno stile

/**
 * Applica i livelli dati ai paragrafi selezionati. Se sono tutti dello
 * stesso elenco cambia l'aspetto di tutto l'elenco (come Word); altrimenti
 * crea un elenco nuovo, conservando il livello di chi era già in elenco.
 */
export function applyLevels(state: EditorState, levels: readonly ListLevel[]): DocumentModel | null {
  const locs = selectedParagraphs(state);
  if (!locs.length) return null;
  let doc = state.doc;
  const refs = locs.map((l) => resolveParaFormat(doc.styles, l.paragraph).list);
  const first = refs[0];
  if (first && doc.lists[first.listId] && refs.every((r) => r?.listId === first.listId)) {
    const listId = first.listId;
    return { ...doc, lists: { ...doc.lists, [listId]: { id: listId, levels: [...levels] } } };
  }
  const def: ListDefinition = { id: newId(), levels: [...levels] };
  doc = { ...doc, lists: { ...doc.lists, [def.id]: def } };
  locs.forEach((l, k) => {
    doc = replaceBlock(doc, l.paragraph.id, [withList(l.paragraph, { listId: def.id, level: refs[k]?.level ?? 0 })]);
  });
  return doc;
}

export function applyPreset(state: EditorState, preset: PresetId): DocumentModel | null {
  const p = presetById(preset);
  return p ? applyLevels(state, p.levels()) : null;
}

/** Toglie l'elenco dai paragrafi selezionati. */
export function removeList(state: EditorState): DocumentModel | null {
  let doc = state.doc;
  let changed = false;
  for (const { paragraph: p } of selectedParagraphs(state)) {
    if (!resolveParaFormat(doc.styles, p).list) continue;
    changed = true;
    doc = replaceBlock(doc, p.id, [withList(p, null)]);
  }
  return changed ? doc : null;
}

// ---------------------------------------------------------------------------
// Riavvio, continuazione e valore iniziale

/** Sposta in `target` il paragrafo `from` e quelli che lo seguono nel suo elenco. */
function moveFrom(doc: DocumentModel, from: Id, listId: Id, target: Id): DocumentModel {
  let started = false;
  for (const { paragraph, ref } of listItems(doc)) {
    if (paragraph.id === from) started = true;
    if (!started || ref.listId !== listId) continue;
    doc = replaceBlock(doc, paragraph.id, [{ ...paragraph, format: { ...paragraph.format, list: { listId: target, level: ref.level } } }]);
  }
  return doc;
}

/** Il paragrafo col cursore è la prima voce del suo elenco? */
function isFirstOfList(doc: DocumentModel, item: ListItem): boolean {
  return listItems(doc).find((i) => i.ref.listId === item.ref.listId)?.paragraph.id === item.paragraph.id;
}

/** Elenco numerato precedente, diverso da quello della voce, a cui agganciarsi. */
function previousNumberedList(doc: DocumentModel, item: ListItem): Id | undefined {
  let found: Id | undefined;
  for (const i of listItems(doc)) {
    if (i.paragraph.id === item.paragraph.id) break;
    if (i.ref.listId !== item.ref.listId && isNumbered(doc.lists[i.ref.listId])) found = i.ref.listId;
  }
  return found;
}

export function canRestart(state: EditorState): boolean {
  const item = focusItem(state);
  return !!item && isNumbered(state.doc.lists[item.ref.listId]) && !isFirstOfList(state.doc, item);
}

export function canContinue(state: EditorState): boolean {
  const item = focusItem(state);
  return !!item && isNumbered(state.doc.lists[item.ref.listId]) && isFirstOfList(state.doc, item) && !!previousNumberedList(state.doc, item);
}

/** Ricomincia da 1: da questa voce in poi l'elenco diventa un elenco nuovo con lo stesso aspetto. */
export function restartNumbering(state: EditorState, startAt?: number): DocumentModel | null {
  const item = focusItem(state);
  const def = item && state.doc.lists[item.ref.listId];
  if (!item || !def || !isNumbered(def)) return null;
  const levels = def.levels.map((l, i) => (startAt !== undefined && i === item.ref.level ? { ...l, start: startAt } : l));
  if (isFirstOfList(state.doc, item)) {
    // Prima voce: basta cambiare il valore iniziale dell'elenco.
    if (startAt === undefined) return null;
    return { ...state.doc, lists: { ...state.doc.lists, [def.id]: { ...def, levels } } };
  }
  const copy: ListDefinition = { id: newId(), levels };
  const doc = { ...state.doc, lists: { ...state.doc.lists, [copy.id]: copy } };
  return moveFrom(doc, item.paragraph.id, def.id, copy.id);
}

/** Continua numerazione: l'elenco prosegue i numeri dell'elenco numerato precedente. */
export function continueNumbering(state: EditorState): DocumentModel | null {
  const item = focusItem(state);
  if (!item || !canContinue(state)) return null;
  const target = previousNumberedList(state.doc, item)!;
  return moveFrom(state.doc, item.paragraph.id, item.ref.listId, target);
}

/** Numero che la voce col cursore mostra ora al suo livello (per la finestra "Imposta valore"). */
export function currentValue(state: EditorState): number {
  const item = focusItem(state);
  if (!item) return 1;
  const counters: number[] = [];
  for (const i of listItems(state.doc)) {
    if (i.ref.listId !== item.ref.listId) continue;
    const start = state.doc.lists[i.ref.listId]!.levels[i.ref.level]?.start ?? 1;
    counters[i.ref.level] = (counters[i.ref.level] ?? start - 1) + 1;
    counters.length = i.ref.level + 1;
    if (i.paragraph.id === item.paragraph.id) return counters[i.ref.level]!;
  }
  return 1;
}

/** Livelli dell'elenco col cursore (o `undefined` se il cursore non è in un elenco). */
export function focusLevels(state: EditorState): { levels: readonly ListLevel[]; level: number } | undefined {
  const item = focusItem(state);
  const def = item && state.doc.lists[item.ref.listId];
  return def && item ? { levels: def.levels, level: item.ref.level } : undefined;
}

/** I 9 livelli, completati con quelli dello stile 1. 1.1. se ne mancano. */
export function fullLevels(levels: readonly ListLevel[] | undefined): ListLevel[] {
  const base = presetById('outline')!.levels();
  return base.map((b, i) => levels?.[i] ?? b);
}
