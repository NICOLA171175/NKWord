/**
 * Modello del documento.
 *
 * Regole valide per tutti i moduli:
 * - Il modello è immutabile: nessun modulo modifica un oggetto esistente,
 *   si crea sempre una copia (structural sharing). Per questo i tipi sono
 *   `readonly`.
 * - Le lunghezze sono in punti (vedi `units.ts`).
 * - Ogni blocco (paragrafo, tabella, riga, cella) ha un `id` univoco
 *   nell'intero documento; le posizioni del cursore si riferiscono agli id
 *   dei paragrafi, quindi restano valide anche se cambiano gli indici.
 * - I colori sono stringhe esadecimali `#rrggbb`.
 */

import type { Points } from './units';

export type Id = string;
export type Color = string;

// ---------------------------------------------------------------------------
// Formattazione

export type UnderlineStyle = 'single' | 'double' | 'dotted' | 'dashed' | 'wave';
export type VerticalAlign = 'baseline' | 'superscript' | 'subscript';

/** Formattazione del carattere. Ogni campo assente eredita dallo stile. */
export interface CharFormat {
  readonly fontFamily?: string;
  readonly fontSize?: Points;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: UnderlineStyle | null;
  readonly strike?: boolean;
  readonly color?: Color;
  readonly highlight?: Color | null;
  readonly verticalAlign?: VerticalAlign;
  readonly smallCaps?: boolean;
  readonly allCaps?: boolean;
  /** URL del collegamento ipertestuale. */
  readonly link?: string;
}

export type Alignment = 'left' | 'center' | 'right' | 'justify';

export type LineSpacing =
  /** Multiplo dell'interlinea singola (1, 1.15, 1.5, 2...). */
  | { readonly rule: 'multiple'; readonly value: number }
  | { readonly rule: 'exact'; readonly value: Points }
  | { readonly rule: 'atLeast'; readonly value: Points };

export type TabAlignment = 'left' | 'center' | 'right' | 'decimal';

export interface TabStop {
  readonly position: Points;
  readonly align: TabAlignment;
  readonly leader?: 'none' | 'dot' | 'hyphen' | 'underscore';
}

export interface BorderLine {
  readonly style: 'none' | 'single' | 'double' | 'dotted' | 'dashed';
  readonly width: Points;
  readonly color: Color;
}

export interface Borders {
  readonly top?: BorderLine;
  readonly right?: BorderLine;
  readonly bottom?: BorderLine;
  readonly left?: BorderLine;
}

/** Riferimento a un elenco puntato/numerato (vedi `ListDefinition`). */
export interface ListRef {
  readonly listId: Id;
  /** Livello di rientro, 0..8. */
  readonly level: number;
}

/** Formattazione del paragrafo. Ogni campo assente eredita dallo stile. */
export interface ParaFormat {
  readonly align?: Alignment;
  readonly indentLeft?: Points;
  readonly indentRight?: Points;
  /** Positivo: rientro prima riga; negativo: rientro sporgente. */
  readonly indentFirstLine?: Points;
  readonly spaceBefore?: Points;
  readonly spaceAfter?: Points;
  readonly lineSpacing?: LineSpacing;
  readonly keepWithNext?: boolean;
  readonly keepLinesTogether?: boolean;
  readonly pageBreakBefore?: boolean;
  readonly widowControl?: boolean;
  readonly tabs?: readonly TabStop[];
  readonly borders?: Borders;
  readonly shading?: Color | null;
  readonly list?: ListRef | null;
  /** 1..9 per i titoli; usato da sommario e struttura. */
  readonly outlineLevel?: number;
}

// ---------------------------------------------------------------------------
// Revisioni

/**
 * Revisione del rilevamento modifiche: il contenuto che la porta è stato
 * inserito o eliminato da `author` ed è in attesa di essere accettato o
 * rifiutato. Il testo eliminato resta nel modello finché non si accetta.
 */
export interface Revision {
  readonly kind: 'insert' | 'delete';
  readonly author: string;
  /** Data ISO 8601. */
  readonly date?: string;
}

// ---------------------------------------------------------------------------
// Contenuto in linea

export interface TextRun {
  readonly type: 'text';
  /** Non contiene mai '\n' né '\t': si usano `BreakInline` e `TabInline`. */
  readonly text: string;
  readonly format: CharFormat;
  readonly styleId?: Id;
  /** Revisione in attesa (rilevamento modifiche). */
  readonly revision?: Revision;
}

export interface ImageInline {
  readonly type: 'image';
  /** URL, data URI o chiave di `DocumentModel.media`. */
  readonly src: string;
  readonly width: Points;
  readonly height: Points;
  readonly alt?: string;
  /** Revisione in attesa (rilevamento modifiche). */
  readonly revision?: Revision;
}

export interface BreakInline {
  readonly type: 'break';
  readonly kind: 'line' | 'page' | 'column';
  /** Revisione in attesa (rilevamento modifiche). */
  readonly revision?: Revision;
}

export interface TabInline {
  readonly type: 'tab';
  /** Revisione in attesa (rilevamento modifiche). */
  readonly revision?: Revision;
}

export type FieldKind = 'page' | 'numPages' | 'date' | 'time';

/** Campo calcolato al momento dell'impaginazione (es. numero di pagina). */
export interface FieldInline {
  readonly type: 'field';
  readonly kind: FieldKind;
  readonly format: CharFormat;
  /** Revisione in attesa (rilevamento modifiche). */
  readonly revision?: Revision;
}

export type Inline = TextRun | ImageInline | BreakInline | TabInline | FieldInline;

// ---------------------------------------------------------------------------
// Blocchi

export interface Paragraph {
  readonly type: 'paragraph';
  readonly id: Id;
  readonly styleId?: Id;
  readonly format: ParaFormat;
  /** Un paragrafo vuoto ha `children: []`. Vedi `normalizeInlines`. */
  readonly children: readonly Inline[];
  /** Formattazione del segno di paragrafo (usata per il cursore su riga vuota). */
  readonly markFormat?: CharFormat;
  /** Revisione del segno di paragrafo: eliminarlo unisce il paragrafo al successivo. */
  readonly revision?: Revision;
}

export interface TableCellFormat {
  readonly verticalAlign?: 'top' | 'center' | 'bottom';
  readonly shading?: Color | null;
  readonly borders?: Borders;
  readonly padding?: { readonly top: Points; readonly right: Points; readonly bottom: Points; readonly left: Points };
}

export interface TableCell {
  readonly id: Id;
  /** Contiene almeno un paragrafo. */
  readonly blocks: readonly Block[];
  readonly colSpan?: number;
  /** 'restart' apre un'unione verticale, 'continue' la prosegue. */
  readonly vMerge?: 'restart' | 'continue';
  readonly format: TableCellFormat;
}

export interface TableRow {
  readonly id: Id;
  readonly cells: readonly TableCell[];
  readonly height?: { readonly rule: 'auto' | 'atLeast' | 'exact'; readonly value: Points };
  readonly isHeader?: boolean;
}

export interface TableFormat {
  readonly align?: 'left' | 'center' | 'right';
  readonly indentLeft?: Points;
  readonly borders?: Borders & { readonly insideH?: BorderLine; readonly insideV?: BorderLine };
}

export interface Table {
  readonly type: 'table';
  readonly id: Id;
  readonly styleId?: Id;
  /** Larghezza di ciascuna colonna della griglia. */
  readonly columnWidths: readonly Points[];
  readonly rows: readonly TableRow[];
  readonly format: TableFormat;
}

export type Block = Paragraph | Table;

// ---------------------------------------------------------------------------
// Sezioni e pagina

export interface PageMargins {
  readonly top: Points;
  readonly right: Points;
  readonly bottom: Points;
  readonly left: Points;
  /** Distanza dell'intestazione dal bordo superiore. */
  readonly header: Points;
  /** Distanza del piè di pagina dal bordo inferiore. */
  readonly footer: Points;
}

export interface PageSetup {
  /** Dimensioni effettive (già scambiate se orizzontale). */
  readonly width: Points;
  readonly height: Points;
  readonly orientation: 'portrait' | 'landscape';
  readonly margins: PageMargins;
  readonly columns?: { readonly count: number; readonly gap: Points };
}

export interface HeaderFooter {
  readonly default?: readonly Block[];
  readonly first?: readonly Block[];
  readonly even?: readonly Block[];
}

export interface Section {
  readonly id: Id;
  readonly page: PageSetup;
  readonly blocks: readonly Block[];
  readonly header?: HeaderFooter;
  readonly footer?: HeaderFooter;
  readonly titlePage?: boolean;
  /** Come inizia la sezione rispetto alla precedente. */
  readonly breakType?: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage';
  readonly pageNumberStart?: number;
}

// ---------------------------------------------------------------------------
// Stili ed elenchi

export interface Style {
  readonly id: Id;
  readonly name: string;
  readonly type: 'paragraph' | 'character' | 'table';
  readonly basedOn?: Id;
  /** Stile del paragrafo successivo quando si preme Invio. */
  readonly next?: Id;
  readonly paraFormat?: ParaFormat;
  readonly charFormat?: CharFormat;
  /** Visibile nella galleria stili della barra multifunzione. */
  readonly quickStyle?: boolean;
}

export interface StyleSheet {
  readonly styles: Readonly<Record<Id, Style>>;
  /** Formattazione di base applicata sotto ogni stile. */
  readonly defaults: { readonly char: CharFormat; readonly para: ParaFormat };
  readonly defaultParagraphStyle: Id;
}

export type NumberFormat =
  | 'bullet'
  | 'decimal'
  | 'lowerLetter'
  | 'upperLetter'
  | 'lowerRoman'
  | 'upperRoman'
  | 'none';

export interface ListLevel {
  readonly format: NumberFormat;
  /** Es. '%1.' oppure '%1.%2.'; per i punti elenco è il carattere stesso. */
  readonly text: string;
  readonly start: number;
  readonly indentLeft: Points;
  readonly hanging: Points;
  readonly charFormat?: CharFormat;
}

export interface ListDefinition {
  readonly id: Id;
  readonly levels: readonly ListLevel[];
}

// ---------------------------------------------------------------------------
// Documento

export interface DocumentMetadata {
  readonly title?: string;
  readonly author?: string;
  readonly subject?: string;
  readonly keywords?: readonly string[];
  /** Date ISO 8601. */
  readonly created?: string;
  readonly modified?: string;
}

export interface DocumentModel {
  readonly sections: readonly Section[];
  readonly styles: StyleSheet;
  readonly lists: Readonly<Record<Id, ListDefinition>>;
  readonly metadata: DocumentMetadata;
  /** Risorse binarie (immagini) indicizzate per chiave, come data URI. */
  readonly media: Readonly<Record<string, string>>;
}

// ---------------------------------------------------------------------------
// Posizioni e selezione

/**
 * Punto nel testo: paragrafo + offset. Ogni inline non testuale (immagine,
 * tab, interruzione, campo) occupa 1 unità di offset; l'offset va da 0 a
 * `paragraphLength(p)` incluso.
 */
export interface Position {
  readonly paragraphId: Id;
  readonly offset: number;
}

/** `anchor` è dove è iniziata la selezione, `focus` dove si trova il cursore. */
export interface Selection {
  readonly anchor: Position;
  readonly focus: Position;
}
