/**
 * Operazioni pure sulle note a piè di pagina: inserire, modificare ed
 * eliminare una nota. Il richiamo nel testo è un inline `footnoteRef`; il
 * testo della nota sta in `doc.footnotes` e il suo primo paragrafo comincia
 * con lo stesso richiamo, come in Word. Il numero non è salvato da nessuna
 * parte: lo calcola `footnoteNumbers` dall'ordine dei richiami.
 */

import {
  caret,
  findParagraph,
  isCollapsed,
  newId,
  normalizeInlines,
  orderedRange,
  paragraphs,
  paragraphText,
  STYLE_NORMAL,
  textRun,
  updateParagraph,
  type Block,
  type CharFormat,
  type DocumentModel,
  type EditorState,
  type Footnote,
  type FootnoteRefInline,
  type Id,
  type Inline,
  type Paragraph,
  type Position,
  type Style,
  type StyleSheet,
} from '../core';

/** Stile di paragrafo del testo delle note (nome interno di Word: "footnote text"). */
export const FOOTNOTE_TEXT_STYLE = 'FootnoteText';

export function footnoteTextStyle(): Style {
  return {
    id: FOOTNOTE_TEXT_STYLE,
    name: 'Testo nota a piè di pagina',
    type: 'paragraph',
    basedOn: STYLE_NORMAL,
    paraFormat: { spaceBefore: 0, spaceAfter: 0, lineSpacing: { rule: 'multiple', value: 1 } },
    charFormat: { fontSize: 10 },
  };
}

/** Aggiunge lo stile del testo delle note se manca. */
export function ensureFootnoteStyles(sheet: StyleSheet): StyleSheet {
  if (sheet.styles[FOOTNOTE_TEXT_STYLE]) return sheet;
  return { ...sheet, styles: { ...sheet.styles, [FOOTNOTE_TEXT_STYLE]: footnoteTextStyle() } };
}

export function footnoteRef(noteId: Id, format: CharFormat = {}): FootnoteRefInline {
  return { type: 'footnoteRef', noteId, format };
}

/** Paragrafi della nota: il primo comincia con il richiamo e uno spazio. Ogni riga di `text` è un paragrafo. */
export function footnoteBlocks(noteId: Id, text: string, styleId: Id = FOOTNOTE_TEXT_STYLE): Paragraph[] {
  const lines = text.replace(/\t/g, ' ').split(/\r\n|\r|\n/);
  return lines.map((line, i) => {
    const children: Inline[] = i === 0 ? [footnoteRef(noteId), textRun(` ${line}`)] : normalizeInlines([textRun(line)]);
    return { type: 'paragraph', id: newId(), styleId, format: {}, children };
  });
}

/** Testo semplice della nota, senza il richiamo iniziale; i paragrafi sono separati da '\n'. */
export function footnoteText(doc: DocumentModel, noteId: Id): string | undefined {
  const note = doc.footnotes?.[noteId];
  if (!note) return undefined;
  const out: string[] = [];
  const visit = (blocks: readonly Block[]) => {
    for (const b of blocks) {
      if (b.type === 'paragraph') out.push(paragraphText({ ...b, children: b.children.filter((c) => c.type !== 'footnoteRef') }));
      else for (const row of b.rows) for (const cell of row.cells) visit(cell.blocks);
    }
  };
  visit(note.blocks);
  if (out.length) out[0] = out[0]!.replace(/^ /, '');
  return out.join('\n');
}

/** Richiami presenti nel corpo, in ordine di lettura. */
export function footnoteRefs(doc: DocumentModel): { noteId: Id; position: Position }[] {
  const out: { noteId: Id; position: Position }[] = [];
  for (const { paragraph: p } of paragraphs(doc)) {
    let offset = 0;
    for (const c of p.children) {
      if (c.type === 'footnoteRef') out.push({ noteId: c.noteId, position: { paragraphId: p.id, offset } });
      offset += c.type === 'text' ? c.text.length : 1;
    }
  }
  return out;
}

// ---------------------------------------------------------------------------
// Inserimento

/** Formato del testo subito prima dell'offset, senza collegamento e apice/pedice. */
function formatBefore(p: Paragraph, offset: number): CharFormat {
  let pos = 0;
  let found: CharFormat | undefined;
  for (const c of p.children) {
    if (pos >= offset) break;
    if (c.type === 'text' || c.type === 'field' || c.type === 'footnoteRef') found = c.format;
    pos += c.type === 'text' ? c.text.length : 1;
  }
  const { link: _link, verticalAlign: _va, ...rest } = found ?? p.markFormat ?? {};
  return rest;
}

/** Inserisce un inline all'offset, dividendo il run di testo se serve. */
export function insertInlineAt(p: Paragraph, offset: number, inline: Inline): Paragraph {
  const out: Inline[] = [];
  let pos = 0;
  let done = false;
  for (const c of p.children) {
    const len = c.type === 'text' ? c.text.length : 1;
    if (!done && offset <= pos) {
      out.push(inline);
      done = true;
    }
    if (!done && c.type === 'text' && offset < pos + len) {
      const k = offset - pos;
      out.push({ ...c, text: c.text.slice(0, k) }, inline, { ...c, text: c.text.slice(k) });
      done = true;
    } else out.push(c);
    pos += len;
  }
  if (!done) out.push(inline);
  return { ...p, children: normalizeInlines(out) };
}

/** Il cursore è nel corpo del documento (non in intestazioni o note)? */
export function canInsertFootnote(state: EditorState): boolean {
  const [, to] = orderedRange(state.doc, state.selection);
  return !!findParagraph(state.doc, to.paragraphId);
}

/**
 * Inserisce una nota alla fine della selezione (come Word, il testo
 * selezionato resta). Restituisce null se il cursore non è nel corpo.
 */
export function insertFootnote(state: EditorState, text = ''): { state: EditorState; noteId: Id } | null {
  const { doc } = state;
  const [, to] = orderedRange(doc, state.selection);
  const loc = findParagraph(doc, to.paragraphId);
  if (!loc) return null;
  const p = loc.paragraph;
  const noteId = newId();
  const updated = insertInlineAt(p, to.offset, footnoteRef(noteId, formatBefore(p, to.offset)));
  const note: Footnote = { id: noteId, blocks: footnoteBlocks(noteId, text) };
  const next: DocumentModel = {
    ...updateParagraph(doc, updated),
    styles: ensureFootnoteStyles(doc.styles),
    footnotes: { ...doc.footnotes, [noteId]: note },
  };
  return { state: { doc: next, selection: caret(p.id, to.offset + 1) }, noteId };
}

// ---------------------------------------------------------------------------
// Nota "corrente"

/**
 * Nota a cui si riferisce il cursore: il primo richiamo nella selezione,
 * oppure quello subito prima o subito dopo il cursore.
 */
export function footnoteAtSelection(state: EditorState): Id | undefined {
  const { doc, selection } = state;
  const notes = doc.footnotes ?? {};
  const [from, to] = orderedRange(doc, selection);
  let inside = false;
  for (const { paragraph: p } of paragraphs(doc)) {
    if (p.id === from.paragraphId) inside = true;
    if (!inside) continue;
    const start = p.id === from.paragraphId ? from.offset : 0;
    const end = p.id === to.paragraphId ? to.offset : Infinity;
    let pos = 0;
    for (const c of p.children) {
      if (c.type === 'footnoteRef' && notes[c.noteId]) {
        const hit = isCollapsed(selection) ? pos === start - 1 || pos === start : pos >= start && pos < end;
        if (hit) return c.noteId;
      }
      pos += c.type === 'text' ? c.text.length : 1;
    }
    if (p.id === to.paragraphId) break;
  }
  return undefined;
}

// ---------------------------------------------------------------------------
// Modifica ed eliminazione

export function editFootnote(state: EditorState, noteId: Id, text: string): EditorState | null {
  const note = state.doc.footnotes?.[noteId];
  if (!note) return null;
  const first = note.blocks[0];
  const styleId = (first?.type === 'paragraph' ? first.styleId : undefined) ?? FOOTNOTE_TEXT_STYLE;
  const blocks = footnoteBlocks(noteId, text, styleId);
  const doc: DocumentModel = {
    ...state.doc,
    styles: styleId === FOOTNOTE_TEXT_STYLE ? ensureFootnoteStyles(state.doc.styles) : state.doc.styles,
    footnotes: { ...state.doc.footnotes, [noteId]: { ...note, blocks } },
  };
  return { ...state, doc };
}

/** Toglie i richiami della nota dal corpo e la nota dal documento; il cursore resta sullo stesso testo. */
export function deleteFootnote(state: EditorState, noteId: Id): EditorState | null {
  if (!state.doc.footnotes?.[noteId]) return null;
  let doc = state.doc;
  const removed = new Map<Id, number[]>();
  for (const { paragraph: p } of paragraphs(state.doc)) {
    if (!p.children.some((c) => c.type === 'footnoteRef' && c.noteId === noteId)) continue;
    const offsets: number[] = [];
    let pos = 0;
    const children: Inline[] = [];
    for (const c of p.children) {
      if (c.type === 'footnoteRef' && c.noteId === noteId) offsets.push(pos);
      else children.push(c);
      pos += c.type === 'text' ? c.text.length : 1;
    }
    removed.set(p.id, offsets);
    doc = updateParagraph(doc, { ...p, children: normalizeInlines(children) });
  }
  const { [noteId]: _gone, ...rest } = doc.footnotes ?? {};
  doc = { ...doc, footnotes: rest };
  const shift = (pos: Position): Position => {
    const offsets = removed.get(pos.paragraphId);
    if (!offsets) return pos;
    return { ...pos, offset: pos.offset - offsets.filter((o) => o < pos.offset).length };
  };
  return { ...state, doc, selection: { anchor: shift(state.selection.anchor), focus: shift(state.selection.focus) } };
}
