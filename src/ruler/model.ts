/**
 * Calcoli del righello, senza DOM: rientri effettivi di un paragrafo (come
 * li usa l'impaginazione), modifica dei rientri e delle tabulazioni dei
 * paragrafi selezionati, griglia di aggancio in centimetri.
 */

import {
  cmToPt,
  orderedRange,
  paragraphs,
  replaceBlock,
  resolveParaFormat,
  type DocumentModel,
  type EditorState,
  type Paragraph,
  type ParaFormat,
  type ParagraphLocation,
  type Points,
  type TabAlignment,
  type TabStop,
  type Transaction,
} from '../core';

/** Tipi di tabulazione che si scelgono nel quadratino a sinistra del righello, nell'ordine di Word. */
export const TAB_TYPES: readonly TabAlignment[] = ['left', 'center', 'right', 'decimal'];

export const TAB_LABELS: Record<string, string> = {
  left: 'Tabulazione sinistra',
  center: 'Tabulazione centrata',
  right: 'Tabulazione destra',
  decimal: 'Tabulazione decimale',
};

/** Passo di aggancio dei marcatori: 0,25 cm, come Word con unità metriche. */
export const SNAP_STEP: Points = cmToPt(0.25);

export function snap(pt: Points, step: Points = SNAP_STEP): Points {
  return Math.round(Math.round(pt / step) * step * 100) / 100;
}

/** Rientri effettivi di un paragrafo, misurati dal margine sinistro (o destro). */
export interface Indents {
  readonly left: Points;
  readonly right: Points;
  /** Rispetto a `left`: positivo = prima riga, negativo = sporgente. */
  readonly firstLine: Points;
}

/** Stessa regola dell'impaginazione: la formattazione diretta vince sull'elenco, l'elenco sullo stile. */
export function effectiveIndents(doc: DocumentModel, p: Paragraph): Indents {
  const format = resolveParaFormat(doc.styles, p);
  const ref = format.list;
  const level = ref ? doc.lists[ref.listId]?.levels[ref.level] : undefined;
  return {
    left: p.format.indentLeft ?? level?.indentLeft ?? format.indentLeft ?? 0,
    right: format.indentRight ?? 0,
    firstLine: p.format.indentFirstLine ?? (level ? -level.hanging : (format.indentFirstLine ?? 0)),
  };
}

/** Tabulazioni del paragrafo (stile compreso), ordinate per posizione. */
export function effectiveTabs(doc: DocumentModel, p: Paragraph): TabStop[] {
  return [...(resolveParaFormat(doc.styles, p).tabs ?? [])].sort((a, b) => a.position - b.position);
}

/** Paragrafi toccati dalla selezione, nell'ordine del documento. */
export function selectedLocations(state: EditorState): ParagraphLocation[] {
  const [start, end] = orderedRange(state.doc, state.selection);
  const out: ParagraphLocation[] = [];
  let inside = false;
  for (const loc of paragraphs(state.doc)) {
    if (loc.paragraph.id === start.paragraphId) inside = true;
    if (inside) out.push(loc);
    if (inside && loc.paragraph.id === end.paragraphId) break;
  }
  return out;
}

function mapSelected(state: EditorState, fn: (p: Paragraph) => Paragraph, label: string): Transaction | null {
  let doc = state.doc;
  for (const { paragraph: p } of selectedLocations(state)) {
    const next = fn(p);
    if (next !== p) doc = replaceBlock(doc, p.id, [next]);
  }
  if (doc === state.doc) return null;
  return { state: { ...state, doc }, meta: { label } };
}

const round = (pt: Points) => Math.round(pt * 100) / 100;
const same = (a: Points, b: Points) => Math.abs(a - b) < 0.05;

export interface IndentChange {
  readonly left?: Points;
  readonly right?: Points;
  readonly firstLine?: Points;
}

/** Imposta i rientri dati su tutti i paragrafi selezionati. */
export function setIndents(state: EditorState, change: IndentChange): Transaction | null {
  return mapSelected(
    state,
    (p) => {
      const now = effectiveIndents(state.doc, p);
      const format: { -readonly [K in keyof ParaFormat]: ParaFormat[K] } = { ...p.format };
      let changed = false;
      if (change.left !== undefined && !same(change.left, now.left)) {
        format.indentLeft = round(change.left);
        changed = true;
      }
      if (change.right !== undefined && !same(change.right, now.right)) {
        format.indentRight = round(change.right);
        changed = true;
      }
      if (change.firstLine !== undefined && !same(change.firstLine, now.firstLine)) {
        format.indentFirstLine = round(change.firstLine);
        changed = true;
      }
      // Su una voce di elenco i rientri vengono dal livello: se ne cambia uno si fissa anche l'altro.
      if (changed && format.list !== undefined) {
        format.indentLeft ??= round(now.left);
        format.indentFirstLine ??= round(now.firstLine);
      }
      return changed ? { ...p, format } : p;
    },
    'Rientro',
  );
}

export interface TabChange {
  /** Aggiunge (o sostituisce) una tabulazione. */
  readonly add?: TabStop;
  /** Toglie la tabulazione in questa posizione. */
  readonly remove?: Points;
  /** Sposta la tabulazione da `from` a `to`, mantenendo tipo e riempitivo. */
  readonly move?: { readonly from: Points; readonly to: Points };
}

/** Tolleranza per riconoscere "la stessa" tabulazione (mezzo punto). */
const near = (a: Points, b: Points) => Math.abs(a - b) < 0.5;

export function changeTabs(tabs: readonly TabStop[], change: TabChange): TabStop[] {
  let out = [...tabs];
  if (change.remove !== undefined) out = out.filter((t) => !near(t.position, change.remove!));
  if (change.move) {
    const { from, to } = change.move;
    const moving = out.find((t) => near(t.position, from));
    if (moving) {
      out = out.filter((t) => t !== moving && !near(t.position, to));
      out.push({ ...moving, position: round(to) });
    }
  }
  if (change.add) {
    const add = { ...change.add, position: round(change.add.position) };
    out = out.filter((t) => !near(t.position, add.position));
    out.push(add);
  }
  return out.sort((a, b) => a.position - b.position);
}

const sameTabs = (a: readonly TabStop[], b: readonly TabStop[]) =>
  a.length === b.length && a.every((t, i) => t.position === b[i]!.position && t.align === b[i]!.align && t.leader === b[i]!.leader);

/** Aggiunge, sposta o toglie una tabulazione in tutti i paragrafi selezionati. */
export function setTabs(state: EditorState, change: TabChange): Transaction | null {
  return mapSelected(
    state,
    (p) => {
      const before = effectiveTabs(state.doc, p);
      const tabs = changeTabs(before, change);
      return sameTabs(before, tabs) ? p : { ...p, format: { ...p.format, tabs } };
    },
    'Tabulazioni',
  );
}
