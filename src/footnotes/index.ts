/** Note a piè di pagina: comandi `footnote.*` e operazioni pure sul modello. */

export { footnotesModule, footnoteCommands } from './module';
export {
  FOOTNOTE_TEXT_STYLE,
  canInsertFootnote,
  deleteFootnote,
  editFootnote,
  ensureFootnoteStyles,
  footnoteAtSelection,
  footnoteBlocks,
  footnoteRef,
  footnoteRefs,
  footnoteText,
  footnoteTextStyle,
  insertFootnote,
  insertInlineAt,
} from './model';
