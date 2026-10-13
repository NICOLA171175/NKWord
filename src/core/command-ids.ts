/**
 * Catalogo dei comandi: id, argomenti e modulo che li registra.
 *
 * È il contratto fra chi invoca i comandi (UI, tastiera) e chi li
 * implementa. La UI può collegare un pulsante a un id prima che il modulo
 * proprietario lo abbia implementato: `editor.canExecute` restituisce false
 * finché il comando non è registrato. Aggiungere un id qui richiede
 * coordinamento, come ogni modifica a `src/core`.
 */

import type { FileData } from './interfaces';
import type {
  Alignment,
  CharFormat,
  DocumentModel,
  FieldKind,
  Id,
  LineSpacing,
  NumberFormat,
  PageSetup,
  ParaFormat,
  Position,
  Selection,
} from './types';
import type { Points } from './units';

export type CommandOwner = 'core' | 'editing' | 'structures' | 'layout' | 'io' | 'ui' | 'comments' | 'footnotes';

/** Argomenti di ogni comando (`void` = nessun argomento). */
export interface CommandArgs {
  // core ---------------------------------------------------------------
  'edit.undo': void;
  'edit.redo': void;

  // editing: testo e selezione ------------------------------------------
  'edit.insertText': { text: string };
  'edit.deleteBackward': { unit?: 'char' | 'word' };
  'edit.deleteForward': { unit?: 'char' | 'word' };
  'edit.deleteSelection': void;
  /** Invio: divide il paragrafo. */
  'edit.splitParagraph': void;
  /** Maiusc+Invio: interruzione di riga. */
  'edit.insertLineBreak': void;
  'edit.insertTab': void;
  'edit.selectAll': void;
  'edit.setSelection': { selection: Selection };
  'edit.moveCaret': {
    direction: 'left' | 'right' | 'wordLeft' | 'wordRight' | 'lineStart' | 'lineEnd' | 'docStart' | 'docEnd';
    extend?: boolean;
  };
  /** Incolla un frammento già convertito (vedi `ClipboardService.parse`). */
  'edit.insertFragment': { fragment: DocumentModel };
  'edit.find': { query: string; matchCase?: boolean; wholeWord?: boolean; from?: Position };
  'edit.replace': { query: string; replacement: string; all?: boolean; matchCase?: boolean; wholeWord?: boolean };

  // editing: formattazione carattere -----------------------------------
  'format.bold': void;
  'format.italic': void;
  'format.underline': void;
  'format.strike': void;
  'format.superscript': void;
  'format.subscript': void;
  'format.fontFamily': { fontFamily: string };
  'format.fontSize': { fontSize: Points };
  'format.growFont': void;
  'format.shrinkFont': void;
  'format.color': { color: string };
  'format.highlight': { color: string | null };
  'format.setChar': { format: CharFormat };
  'format.clear': void;
  'format.link': { url: string | null };

  // editing: formattazione paragrafo e stili ---------------------------
  'paragraph.align': { align: Alignment };
  'paragraph.indent': void;
  'paragraph.outdent': void;
  'paragraph.lineSpacing': { lineSpacing: LineSpacing };
  'paragraph.spacing': { before?: Points; after?: Points };
  'paragraph.setFormat': { format: ParaFormat };
  'style.apply': { styleId: Id };

  // structures: elenchi -------------------------------------------------
  'list.toggleBullet': void;
  'list.toggleNumbered': { format?: NumberFormat };
  'list.increaseLevel': void;
  'list.decreaseLevel': void;

  // structures: tabelle -------------------------------------------------
  'table.insert': { rows: number; cols: number };
  'table.insertRowAbove': void;
  'table.insertRowBelow': void;
  'table.insertColumnLeft': void;
  'table.insertColumnRight': void;
  'table.deleteRow': void;
  'table.deleteColumn': void;
  'table.delete': void;
  'table.mergeCells': void;
  'table.splitCell': { rows: number; cols: number };

  // structures: oggetti e sezioni ---------------------------------------
  'insert.image': { src: string; width: Points; height: Points; alt?: string };
  'insert.pageBreak': void;
  'insert.sectionBreak': { type: 'nextPage' | 'continuous' };
  'insert.field': { kind: FieldKind };
  'page.setup': { page: Partial<PageSetup> };
  'page.headerFooter': { part: 'header' | 'footer'; edit: boolean };

  // io: file ------------------------------------------------------------
  'file.new': void;
  'file.open': { data: FileData; formatId: string };
  'file.export': { formatId: string };

  // comments: commenti a margine ---------------------------------------
  /** Commenta la selezione (o la parola al cursore). */
  'comment.insert': { text?: string; author?: string };
  'comment.reply': { parentId: Id; text: string; author?: string };
  'comment.edit': { id: Id; text: string };
  /** Senza `id` elimina il commento al cursore. */
  'comment.delete': { id?: Id };
  'comment.resolve': { id: Id; resolved: boolean };
  'comment.next': void;
  'comment.previous': void;

  // ui: vista -----------------------------------------------------------
  'view.zoom': { zoom: number };
  'view.toggleFormattingMarks': void;
  'view.print': void;

  // footnotes: note a piè di pagina ------------------------------------
  /** Inserisce il richiamo alla fine della selezione e crea la nota. */
  'footnote.insert': { text?: string };
  /** Sostituisce il testo della nota (default: quella vicina al cursore). */
  'footnote.edit': { text: string; noteId?: Id };
  /** Elimina richiamo e nota (default: quella vicina al cursore). */
  'footnote.delete': { noteId?: Id };
}

export type CommandId = keyof CommandArgs;

/** Modulo responsabile di registrare ciascun comando. */
export const COMMAND_OWNERS: { readonly [K in CommandId]: CommandOwner } = {
  'edit.undo': 'core',
  'edit.redo': 'core',
  'edit.insertText': 'editing',
  'edit.deleteBackward': 'editing',
  'edit.deleteForward': 'editing',
  'edit.deleteSelection': 'editing',
  'edit.splitParagraph': 'editing',
  'edit.insertLineBreak': 'editing',
  'edit.insertTab': 'editing',
  'edit.selectAll': 'editing',
  'edit.setSelection': 'editing',
  'edit.moveCaret': 'editing',
  'edit.insertFragment': 'editing',
  'edit.find': 'editing',
  'edit.replace': 'editing',
  'format.bold': 'editing',
  'format.italic': 'editing',
  'format.underline': 'editing',
  'format.strike': 'editing',
  'format.superscript': 'editing',
  'format.subscript': 'editing',
  'format.fontFamily': 'editing',
  'format.fontSize': 'editing',
  'format.growFont': 'editing',
  'format.shrinkFont': 'editing',
  'format.color': 'editing',
  'format.highlight': 'editing',
  'format.setChar': 'editing',
  'format.clear': 'editing',
  'format.link': 'editing',
  'paragraph.align': 'editing',
  'paragraph.indent': 'editing',
  'paragraph.outdent': 'editing',
  'paragraph.lineSpacing': 'editing',
  'paragraph.spacing': 'editing',
  'paragraph.setFormat': 'editing',
  'style.apply': 'editing',
  'list.toggleBullet': 'structures',
  'list.toggleNumbered': 'structures',
  'list.increaseLevel': 'structures',
  'list.decreaseLevel': 'structures',
  'table.insert': 'structures',
  'table.insertRowAbove': 'structures',
  'table.insertRowBelow': 'structures',
  'table.insertColumnLeft': 'structures',
  'table.insertColumnRight': 'structures',
  'table.deleteRow': 'structures',
  'table.deleteColumn': 'structures',
  'table.delete': 'structures',
  'table.mergeCells': 'structures',
  'table.splitCell': 'structures',
  'insert.image': 'structures',
  'insert.pageBreak': 'structures',
  'insert.sectionBreak': 'structures',
  'insert.field': 'structures',
  'page.setup': 'structures',
  'page.headerFooter': 'ui',
  'file.new': 'io',
  'file.open': 'io',
  'file.export': 'io',
  'comment.insert': 'comments',
  'comment.reply': 'comments',
  'comment.edit': 'comments',
  'comment.delete': 'comments',
  'comment.resolve': 'comments',
  'comment.next': 'comments',
  'comment.previous': 'comments',
  'view.zoom': 'ui',
  'view.toggleFormattingMarks': 'ui',
  'view.print': 'ui',
  'footnote.insert': 'footnotes',
  'footnote.edit': 'footnotes',
  'footnote.delete': 'footnotes',
};

export const COMMAND_IDS = Object.keys(COMMAND_OWNERS) as CommandId[];
