import type { CharFormat, DocumentModel, Selection } from './types';
import { caret, createDocument, firstParagraph } from './model';

/**
 * Stato completo dell'editor in un istante. Immutabile: ogni modifica
 * produce un nuovo `EditorState`.
 */
export interface EditorState {
  readonly doc: DocumentModel;
  readonly selection: Selection;
  /**
   * Formattazione "in attesa" applicata al prossimo testo digitato
   * (es. Ctrl+B con selezione vuota). Si azzera quando il cursore si sposta.
   */
  readonly pendingFormat?: CharFormat;
}

export function createState(doc: DocumentModel = createDocument()): EditorState {
  return { doc, selection: caret(firstParagraph(doc).id, 0) };
}

/** Metadati di una transazione, letti dalla cronologia e dagli ascoltatori. */
export interface TransactionMeta {
  /** Etichetta per "Annulla ..." nella UI. */
  readonly label?: string;
  /** false: la modifica non entra nella cronologia (es. solo spostamento cursore). */
  readonly addToHistory?: boolean;
  /**
   * Transazioni consecutive con la stessa chiave, ravvicinate nel tempo, si
   * fondono in un solo passo di annullamento (es. 'typing').
   */
  readonly mergeKey?: string;
}

/** Risultato di un comando: il nuovo stato più i metadati. */
export interface Transaction {
  readonly state: EditorState;
  readonly meta?: TransactionMeta;
}
