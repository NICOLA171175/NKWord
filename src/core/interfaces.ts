/**
 * Contratti fra i moduli. Ogni modulo parallelo implementa una o più di
 * queste interfacce nella propria cartella e non modifica `src/core` senza
 * coordinarsi. Un modulo si aggancia all'editor con `EditorModule` e
 * pubblica i suoi servizi in `editor.services`.
 */

import type { Editor } from './editor';
import type { CharFormat, DocumentModel, Id, Position } from './types';
import type { Points } from './units';

// ---------------------------------------------------------------------------
// Moduli

export interface EditorModule {
  readonly id: string;
  /** Registra comandi/servizi; può restituire una funzione di pulizia. */
  activate(editor: Editor): void | (() => void);
}

/** Servizi pubblicati dai moduli. Tutti opzionali: un modulo può mancare. */
export interface EditorServices {
  layout?: LayoutEngine;
  formats?: FormatRegistry;
  clipboard?: ClipboardService;
}

// ---------------------------------------------------------------------------
// Impaginazione (modulo layout)

export interface Rect {
  readonly x: Points;
  readonly y: Points;
  readonly width: Points;
  readonly height: Points;
}

/** Misurazione del testo, iniettata per poter testare il layout senza DOM. */
export interface TextMeasurer {
  measure(text: string, format: CharFormat): { width: Points; ascent: Points; descent: Points };
}

/** Pezzo di testo posizionato su una riga, con formattazione già risolta. */
export interface GlyphRun {
  readonly kind: 'text' | 'image' | 'tab' | 'field' | 'listMarker';
  /** Paragrafo e intervallo di offset del modello coperti (vuoto per i marcatori). */
  readonly paragraphId: Id;
  readonly start: number;
  readonly end: number;
  readonly text: string;
  readonly format: CharFormat;
  /** Relativo alla riga. */
  readonly x: Points;
  readonly width: Points;
  /** Solo per le immagini. */
  readonly src?: string;
  readonly height?: Points;
}

export interface LineBox {
  readonly paragraphId: Id;
  readonly start: number;
  readonly end: number;
  /** Relativo alla pagina. */
  readonly x: Points;
  readonly y: Points;
  readonly width: Points;
  readonly height: Points;
  readonly baseline: Points;
  readonly runs: readonly GlyphRun[];
}

/** Elemento disegnabile non testuale: bordi, sfondi, celle di tabella. */
export interface Decoration {
  readonly kind: 'fill' | 'border';
  readonly rect: Rect;
  readonly color: string;
  readonly width?: Points;
}

export interface PageLayout {
  readonly index: number;
  /** Numero mostrato (tiene conto di `pageNumberStart`). */
  readonly number: number;
  readonly sectionId: Id;
  readonly width: Points;
  readonly height: Points;
  readonly contentRect: Rect;
  readonly lines: readonly LineBox[];
  readonly decorations: readonly Decoration[];
  readonly header: readonly LineBox[];
  readonly footer: readonly LineBox[];
}

export interface DocumentLayout {
  readonly doc: DocumentModel;
  readonly pages: readonly PageLayout[];
}

export interface LayoutEngine {
  layout(doc: DocumentModel): DocumentLayout;
  /** Rettangolo del cursore per una posizione (pagina + coordinate in punti). */
  caretRect(layout: DocumentLayout, pos: Position): { page: number; rect: Rect } | null;
  /** Posizione del modello più vicina al punto (coordinate della pagina). */
  hitTest(layout: DocumentLayout, page: number, x: Points, y: Points): Position | null;
  /** Rettangoli da evidenziare per una selezione. */
  selectionRects(layout: DocumentLayout, from: Position, to: Position): { page: number; rect: Rect }[];
}

// ---------------------------------------------------------------------------
// Formati di file (modulo io)

export type FileData = string | Uint8Array;

export interface DocumentFormat {
  /** Es. 'docx', 'html', 'txt', 'md', 'pdf'. */
  readonly id: string;
  readonly label: string;
  readonly extensions: readonly string[];
  readonly mimeType: string;
  /** Assente se il formato è solo di esportazione (es. PDF). */
  import?(data: FileData): Promise<DocumentModel>;
  /** Assente se il formato è solo di importazione. */
  export?(doc: DocumentModel): Promise<FileData>;
}

export interface FormatRegistry {
  register(format: DocumentFormat): void;
  get(id: string): DocumentFormat | undefined;
  byExtension(ext: string): DocumentFormat | undefined;
  list(): readonly DocumentFormat[];
}

/** Appunti: frammenti di modello convertiti da/a HTML e testo semplice. */
export interface ClipboardService {
  serialize(doc: DocumentModel, from: Position, to: Position): { html: string; text: string };
  /** Restituisce un documento-frammento da incollare. */
  parse(data: { html?: string; text?: string }): DocumentModel;
}

// ---------------------------------------------------------------------------
// Interfaccia utente (modulo ui)

/** Voce della barra multifunzione, collegata a un comando per id. */
export interface RibbonItem {
  readonly commandId: string;
  readonly tab: 'home' | 'insert' | 'layout' | 'references' | 'review' | 'view' | 'file';
  readonly group: string;
  readonly icon?: string;
  readonly kind?: 'button' | 'toggle' | 'dropdown';
  /** Argomenti fissi passati al comando (es. { size: 12 }). */
  readonly args?: unknown;
}

/** Scorciatoia: 'Mod' = Ctrl su Windows/Linux, Cmd su macOS. Es. 'Mod-b', 'Shift-Enter'. */
export interface Keybinding {
  readonly key: string;
  readonly commandId: string;
  readonly args?: unknown;
}
