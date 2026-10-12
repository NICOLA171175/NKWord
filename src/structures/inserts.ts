/**
 * Comandi `insert.*` (immagini, interruzioni, campi) e `page.setup`.
 */

import {
  caret,
  newId,
  replaceBlock,
  type EditorState,
  type PageSetup,
  type Section,
} from '../core';
import {
  charFormatAt,
  defineCommand,
  deleteSelection,
  innermostCell,
  insertInline,
  replaceSection,
  requireParagraph,
  selectedSections,
  splitParagraph,
} from './util';

function inTableCell(state: EditorState): boolean {
  const loc = requireParagraph(state.doc, state.selection.focus.paragraphId);
  return !!innermostCell(state.doc, loc);
}

/**
 * Interruzione di pagina: l'inline `break` chiude il paragrafo corrente e il
 * testo dopo il cursore passa in un paragrafo nuovo, come in Word.
 */
function insertPageBreak(state: EditorState): EditorState {
  const { doc, pos } = deleteSelection(state);
  const p = requireParagraph(doc, pos.paragraphId).paragraph;
  const [left, right] = splitParagraph(p, pos.offset);
  const withBreak = { ...left, children: [...left.children, { type: 'break' as const, kind: 'page' as const }] };
  return {
    ...state,
    doc: replaceBlock(doc, p.id, [withBreak, right]),
    selection: caret(right.id, 0),
    pendingFormat: undefined,
  };
}

/**
 * Interruzione di sezione: la sezione corrente si divide al cursore. La nuova
 * sezione eredita impostazioni di pagina, intestazioni e piè di pagina.
 */
function insertSectionBreak(state: EditorState, type: 'nextPage' | 'continuous'): EditorState {
  const { doc, pos } = deleteSelection(state);
  const loc = requireParagraph(doc, pos.paragraphId);
  const sec = doc.sections[loc.sectionIndex]!;
  const index = loc.path[0]!;
  const [left, right] = splitParagraph(loc.paragraph, pos.offset);
  const first: Section = { ...sec, blocks: [...sec.blocks.slice(0, index), left] };
  const { pageNumberStart: _n, ...inherited } = sec;
  const second: Section = { ...inherited, id: newId(), breakType: type, blocks: [right, ...sec.blocks.slice(index + 1)] };
  const sections = [...doc.sections.slice(0, loc.sectionIndex), first, second, ...doc.sections.slice(loc.sectionIndex + 1)];
  return { ...state, doc: { ...doc, sections }, selection: caret(right.id, 0), pendingFormat: undefined };
}

/** Applica una modifica parziale alle impostazioni di pagina. */
export function applyPageSetup(page: PageSetup, change: Partial<PageSetup>): PageSetup {
  let next: PageSetup = { ...page, ...change, margins: { ...page.margins, ...change.margins } };
  // Cambiando orientamento senza dimensioni esplicite si scambiano larghezza e altezza.
  if (change.orientation && change.orientation !== page.orientation && change.width === undefined && change.height === undefined) {
    next = { ...next, width: page.height, height: page.width };
  }
  return next;
}

export function isValidPageSetup(p: PageSetup): boolean {
  const m = p.margins;
  const nonNeg = [m.top, m.right, m.bottom, m.left, m.header, m.footer].every((v) => Number.isFinite(v) && v >= 0);
  const cols = !p.columns || (Number.isInteger(p.columns.count) && p.columns.count >= 1 && p.columns.gap >= 0);
  return p.width > 0 && p.height > 0 && nonNeg && m.left + m.right < p.width && m.top + m.bottom < p.height && cols;
}

function setupSections(state: EditorState, change: Partial<PageSetup>): EditorState | null {
  let doc = state.doc;
  for (const i of selectedSections(state)) {
    const sec = doc.sections[i]!;
    const page = applyPageSetup(sec.page, change);
    if (!isValidPageSetup(page)) return null;
    doc = replaceSection(doc, i, { ...sec, page });
  }
  return { ...state, doc };
}

export const insertCommands = [
  defineCommand('insert.image', {
    label: 'Immagine',
    isEnabled: (_ctx, args) => !!args && !!args.src && args.width > 0 && args.height > 0,
    run: ({ state }, { src, width, height, alt }) => {
      const image = { type: 'image' as const, src, width, height, ...(alt ? { alt } : {}) };
      return { state: { ...insertInline(state, image), pendingFormat: state.pendingFormat }, meta: { label: 'Inserisci immagine' } };
    },
  }),
  defineCommand('insert.pageBreak', {
    label: 'Interruzione di pagina',
    run: ({ state }) => ({ state: insertPageBreak(state), meta: { label: 'Interruzione di pagina' } }),
  }),
  defineCommand('insert.sectionBreak', {
    label: 'Interruzione di sezione',
    isEnabled: ({ state }) => !inTableCell(state),
    run: ({ state }, { type }) => {
      if (inTableCell(state)) return null;
      return { state: insertSectionBreak(state, type), meta: { label: 'Interruzione di sezione' } };
    },
  }),
  defineCommand('insert.field', {
    label: 'Campo',
    run: ({ state }, { kind }) => {
      const format = charFormatAt(state, state.selection.focus);
      return { state: { ...insertInline(state, { type: 'field', kind, format }), pendingFormat: state.pendingFormat }, meta: { label: 'Inserisci campo' } };
    },
  }),
  defineCommand('page.setup', {
    label: 'Imposta pagina',
    isEnabled: ({ state }, args) => !!args && setupSections(state, args.page) !== null,
    run: ({ state }, { page }) => {
      const next = setupSections(state, page);
      return next && { state: next, meta: { label: 'Imposta pagina' } };
    },
  }),
];
