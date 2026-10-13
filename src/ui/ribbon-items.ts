/**
 * Contenuto della barra multifunzione. Ogni voce è un `RibbonItem` del core
 * collegato a un comando per id: se il modulo proprietario non ha ancora
 * registrato il comando, il pulsante resta disabilitato (`canExecute` false).
 */

import { cmToPt, PAGE_SIZES, type CommandArgs, type CommandId, type PageMargins, type RibbonItem } from '../core';
import type { PaletteKind } from './color-picker';

export type RibbonTab = RibbonItem['tab'];

export interface RibbonOption {
  readonly label: string;
  readonly args: unknown;
}

export interface UiRibbonItem extends RibbonItem {
  readonly commandId: CommandId;
  readonly label: string;
  /** Voci di un menu a tendina: ognuna esegue il comando con i propri argomenti. */
  readonly options?: readonly RibbonOption[];
  /** Pulsante colore con tavolozza (stile Word): il comando riceve `{ color }`. */
  readonly palette?: PaletteKind;
  /**
   * Con `prompt`: il comando agisce su qualcosa al cursore (es. la nota da
   * modificare) e sa dire senza argomenti se c'è; altrimenti il pulsante è spento.
   */
  readonly needsTarget?: boolean;
  /** Chiede gli argomenti all'utente; `undefined` annulla. */
  readonly prompt?: (ask: (message: string, initial?: string) => string | null) => unknown | undefined;
}

export const RIBBON_TABS: readonly { readonly id: RibbonTab; readonly label: string }[] = [
  { id: 'file', label: 'File' },
  { id: 'home', label: 'Home' },
  { id: 'insert', label: 'Inserisci' },
  { id: 'layout', label: 'Layout' },
  { id: 'references', label: 'Riferimenti' },
  { id: 'review', label: 'Revisione' },
  { id: 'view', label: 'Visualizza' },
];

function item<K extends CommandId>(
  tab: RibbonTab,
  group: string,
  commandId: K,
  label: string,
  extra: Partial<Omit<UiRibbonItem, 'commandId' | 'tab' | 'group' | 'label' | 'args'>> & { args?: CommandArgs[K] } = {},
): UiRibbonItem {
  return { tab, group, commandId, label, kind: 'button', ...extra };
}

const FONTS = ['Calibri', 'Calibri Light', 'Arial', 'Times New Roman', 'Georgia', 'Verdana', 'Courier New'];
const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72];
const LINE_SPACINGS = [1, 1.15, 1.5, 2, 2.5, 3];
const ZOOMS = [0.5, 0.75, 1, 1.25, 1.5, 2];

const percent = (z: number) => `${Math.round(z * 100)}%`;

/** Margini predefiniti di Word (Layout > Margini), in cm: superiore, destro, inferiore, sinistro. */
const MARGINS: readonly [string, number, number, number, number][] = [
  ['Normali', 2.5, 2, 2, 2],
  ['Stretti', 1.27, 1.27, 1.27, 1.27],
  ['Moderati', 2.54, 1.91, 2.54, 1.91],
  ['Larghi', 2.54, 5.08, 2.54, 5.08],
];
const margins = ([, top, right, bottom, left]: (typeof MARGINS)[number]): PageMargins => ({
  top: cmToPt(top),
  right: cmToPt(right),
  bottom: cmToPt(bottom),
  left: cmToPt(left),
  header: cmToPt(1.25),
  footer: cmToPt(1.25),
});
/** Formati carta (Layout > Dimensioni), in verticale: con la pagina orizzontale si scambiano. */
const PAPER: readonly [string, keyof typeof PAGE_SIZES][] = [
  ['A4 (21 × 29,7 cm)', 'A4'],
  ['A5 (14,8 × 21 cm)', 'A5'],
  ['A3 (29,7 × 42 cm)', 'A3'],
  ['Lettera (21,6 × 27,9 cm)', 'Letter'],
  ['Legale (21,6 × 35,6 cm)', 'Legal'],
];

export const RIBBON_ITEMS: readonly UiRibbonItem[] = [
  // File ----------------------------------------------------------------
  item('file', 'File', 'file.new', 'Nuovo', { icon: '🗋' }),
  // `file.open` richiede i dati del file: la shell apre il selettore.
  item('file', 'File', 'file.open', 'Apri…', { icon: '📂' }),
  item('file', 'File', 'file.export', 'Salva', { icon: '⤓', args: { formatId: 'docx' } }),
  item('file', 'Esporta', 'file.export', 'Word (.docx)', { icon: '⤓', args: { formatId: 'docx' } }),
  item('file', 'Esporta', 'file.export', 'HTML', { icon: '⤓', args: { formatId: 'html' } }),
  item('file', 'Esporta', 'file.export', 'Testo', { icon: '⤓', args: { formatId: 'txt' } }),
  item('file', 'Esporta', 'file.export', 'PDF', { icon: '⤓', args: { formatId: 'pdf' } }),
  item('file', 'Stampa', 'view.print', 'Stampa', { icon: '🖶' }),

  // Home ----------------------------------------------------------------
  item('home', 'Annulla', 'edit.undo', 'Annulla', { icon: '↶' }),
  item('home', 'Annulla', 'edit.redo', 'Ripeti', { icon: '↷' }),

  item('home', 'Carattere', 'format.fontFamily', 'Tipo di carattere', {
    kind: 'dropdown',
    options: FONTS.map((f) => ({ label: f, args: { fontFamily: f } })),
  }),
  item('home', 'Carattere', 'format.fontSize', 'Dimensione carattere', {
    kind: 'dropdown',
    options: SIZES.map((s) => ({ label: String(s), args: { fontSize: s } })),
  }),
  item('home', 'Carattere', 'format.growFont', 'Aumenta dimensione', { icon: 'A↑' }),
  item('home', 'Carattere', 'format.shrinkFont', 'Riduci dimensione', { icon: 'A↓' }),
  item('home', 'Carattere', 'format.clear', 'Cancella formattazione', { icon: 'A⌫' }),
  item('home', 'Carattere', 'format.bold', 'Grassetto', { icon: 'G', kind: 'toggle' }),
  item('home', 'Carattere', 'format.italic', 'Corsivo', { icon: 'C', kind: 'toggle' }),
  item('home', 'Carattere', 'format.underline', 'Sottolineato', { icon: 'S', kind: 'toggle' }),
  item('home', 'Carattere', 'format.strike', 'Barrato', { icon: 'ab', kind: 'toggle' }),
  item('home', 'Carattere', 'format.subscript', 'Pedice', { icon: 'x₂', kind: 'toggle' }),
  item('home', 'Carattere', 'format.superscript', 'Apice', { icon: 'x²', kind: 'toggle' }),
  item('home', 'Carattere', 'format.highlight', 'Colore evidenziazione testo', { palette: 'highlight', args: { color: null } }),
  item('home', 'Carattere', 'format.color', 'Colore carattere', { palette: 'text', args: { color: null } }),

  item('home', 'Paragrafo', 'list.toggleBullet', 'Elenco puntato', { icon: '•≡', kind: 'toggle' }),
  item('home', 'Paragrafo', 'list.toggleNumbered', 'Elenco numerato', { icon: '1≡', kind: 'toggle', args: {} }),
  item('home', 'Paragrafo', 'paragraph.outdent', 'Riduci rientro', { icon: '⇤' }),
  item('home', 'Paragrafo', 'paragraph.indent', 'Aumenta rientro', { icon: '⇥' }),
  item('home', 'Paragrafo', 'paragraph.align', 'Allinea a sinistra', { icon: '⫷', kind: 'toggle', args: { align: 'left' } }),
  item('home', 'Paragrafo', 'paragraph.align', 'Centra', { icon: '≡', kind: 'toggle', args: { align: 'center' } }),
  item('home', 'Paragrafo', 'paragraph.align', 'Allinea a destra', { icon: '⫸', kind: 'toggle', args: { align: 'right' } }),
  item('home', 'Paragrafo', 'paragraph.align', 'Giustifica', { icon: '☰', kind: 'toggle', args: { align: 'justify' } }),
  item('home', 'Paragrafo', 'paragraph.lineSpacing', 'Interlinea', {
    icon: '↕',
    kind: 'dropdown',
    options: LINE_SPACINGS.map((v) => ({ label: v.toFixed(v % 1 ? 2 : 1).replace(/0$/, ''), args: { lineSpacing: { rule: 'multiple', value: v } } })),
  }),
  item('home', 'Paragrafo', 'view.toggleFormattingMarks', 'Mostra tutto', { icon: '¶', kind: 'toggle' }),
  item('home', 'Paragrafo', 'paragraph.shading', 'Sfondo', { palette: 'shading', args: { color: null } }),
  item('home', 'Paragrafo', 'paragraph.borders', 'Bordi', {
    icon: '▢',
    kind: 'dropdown',
    options: [
      { label: 'Bordo inferiore', args: { sides: 'bottom' } },
      { label: 'Bordo superiore', args: { sides: 'top' } },
      { label: 'Bordo sinistro', args: { sides: 'left' } },
      { label: 'Bordo destro', args: { sides: 'right' } },
      { label: 'Nessun bordo', args: { sides: 'none' } },
      { label: 'Bordi esterni', args: { sides: 'outside' } },
    ],
  }),

  item('home', 'Stili', 'style.apply', 'Normale', { kind: 'toggle', args: { styleId: 'Normal' } }),
  item('home', 'Stili', 'style.apply', 'Titolo 1', { kind: 'toggle', args: { styleId: 'Heading1' } }),
  item('home', 'Stili', 'style.apply', 'Titolo 2', { kind: 'toggle', args: { styleId: 'Heading2' } }),
  item('home', 'Stili', 'style.apply', 'Titolo 3', { kind: 'toggle', args: { styleId: 'Heading3' } }),
  item('home', 'Stili', 'style.apply', 'Titolo', { kind: 'toggle', args: { styleId: 'Title' } }),

  item('home', 'Modifica', 'edit.find', 'Trova', {
    icon: '🔍',
    prompt: (ask) => {
      const query = ask('Trova:');
      return query ? { query } : undefined;
    },
  }),
  item('home', 'Modifica', 'edit.replace', 'Sostituisci', {
    icon: '⇄',
    prompt: (ask) => {
      const query = ask('Trova:');
      if (!query) return undefined;
      const replacement = ask('Sostituisci con:');
      return replacement === null ? undefined : { query, replacement, all: true };
    },
  }),
  item('home', 'Modifica', 'edit.selectAll', 'Seleziona tutto', { icon: '⬚' }),

  // Inserisci -----------------------------------------------------------
  item('insert', 'Pagine', 'insert.pageBreak', 'Interruzione di pagina', { icon: '⤓' }),
  item('insert', 'Tabelle', 'table.insert', 'Tabella', {
    icon: '▦',
    kind: 'dropdown',
    options: [
      { label: '2 × 2', args: { rows: 2, cols: 2 } },
      { label: '3 × 3', args: { rows: 3, cols: 3 } },
      { label: '4 × 4', args: { rows: 4, cols: 4 } },
      { label: '5 × 3', args: { rows: 5, cols: 3 } },
    ],
  }),
  item('insert', 'Illustrazioni', 'insert.image', 'Immagine…', {
    icon: '🖼',
    prompt: (ask) => {
      const src = ask('Indirizzo dell\'immagine:');
      return src ? { src, width: 200, height: 150 } : undefined;
    },
  }),
  item('insert', 'Intestazione e piè di pagina', 'page.headerFooter', 'Intestazione', {
    icon: '▔',
    kind: 'toggle',
    args: { part: 'header', edit: true },
  }),
  item('insert', 'Intestazione e piè di pagina', 'page.headerFooter', 'Piè di pagina', {
    icon: '▁',
    kind: 'toggle',
    args: { part: 'footer', edit: true },
  }),
  item('insert', 'Testo', 'insert.field', 'Data', { icon: '📅', args: { kind: 'date' } }),
  item('insert', 'Testo', 'insert.field', 'Ora', { icon: '🕑', args: { kind: 'time' } }),

  // Layout --------------------------------------------------------------
  item('layout', 'Imposta pagina', 'page.setup', 'Margini', {
    icon: '▭',
    kind: 'dropdown',
    options: MARGINS.map((m) => ({ label: m[0], args: { page: { margins: margins(m) } } })),
  }),
  item('layout', 'Imposta pagina', 'page.setup', 'Orientamento', {
    icon: '⤾',
    kind: 'dropdown',
    options: [
      { label: 'Verticale', args: { page: { orientation: 'portrait' } } },
      { label: 'Orizzontale', args: { page: { orientation: 'landscape' } } },
    ],
  }),
  item('layout', 'Imposta pagina', 'page.setup', 'Dimensioni', {
    icon: '▯',
    kind: 'dropdown',
    options: PAPER.map(([label, size]) => ({ label, args: { page: { ...PAGE_SIZES[size] } } })),
  }),
  item('layout', 'Imposta pagina', 'insert.sectionBreak', 'Interruzioni', {
    icon: '⤶',
    kind: 'dropdown',
    options: [
      { label: 'Sezione: pagina successiva', args: { type: 'nextPage' } },
      { label: 'Sezione: continua', args: { type: 'continuous' } },
    ],
  }),
  item('layout', 'Paragrafo', 'paragraph.spacing', 'Spaziatura', {
    icon: '⇳',
    kind: 'dropdown',
    options: [
      { label: 'Nessuna', args: { before: 0, after: 0 } },
      { label: 'Dopo: 8 pt', args: { after: 8 } },
      { label: 'Prima e dopo: 12 pt', args: { before: 12, after: 12 } },
    ],
  }),

  // Riferimenti ---------------------------------------------------------
  item('references', 'Note a piè di pagina', 'footnote.insert', 'Inserisci nota…', {
    icon: 'ab¹',
    prompt: (ask) => {
      const text = ask('Testo della nota a piè di pagina:');
      return text === null ? undefined : { text };
    },
  }),
  item('references', 'Note a piè di pagina', 'footnote.edit', 'Modifica nota…', {
    icon: '✎¹',
    needsTarget: true,
    prompt: (ask) => {
      const text = ask('Nuovo testo della nota (cursore accanto al numero della nota):');
      return text === null ? undefined : { text };
    },
  }),
  item('references', 'Note a piè di pagina', 'footnote.delete', 'Elimina nota', { icon: '✕¹' }),
  item('references', 'Campi', 'insert.field', 'Numero di pagine', { icon: '#', args: { kind: 'numPages' } }),

  // Revisione -----------------------------------------------------------
  item('review', 'Modifiche', 'edit.undo', 'Annulla', { icon: '↶' }),
  item('review', 'Modifiche', 'edit.replace', 'Sostituisci', {
    icon: '⇄',
    prompt: (ask) => {
      const query = ask('Trova:');
      if (!query) return undefined;
      const replacement = ask('Sostituisci con:');
      return replacement === null ? undefined : { query, replacement, all: true };
    },
  }),

  item('review', 'Commenti', 'comment.insert', 'Nuovo commento', { icon: '💬' }),
  item('review', 'Commenti', 'comment.delete', 'Elimina', { icon: '🗑' }),
  item('review', 'Commenti', 'comment.previous', 'Precedente', { icon: '◂' }),
  item('review', 'Commenti', 'comment.next', 'Successivo', { icon: '▸' }),

  // Visualizza ----------------------------------------------------------
  item('view', 'Mostra', 'view.toggleFormattingMarks', 'Segni di formattazione', { icon: '¶', kind: 'toggle' }),
  item('view', 'Zoom', 'view.zoom', 'Zoom', {
    icon: '🔎',
    kind: 'dropdown',
    options: ZOOMS.map((z) => ({ label: percent(z), args: { zoom: z } })),
  }),
  item('view', 'Zoom', 'view.zoom', '100%', { icon: '1:1', kind: 'toggle', args: { zoom: 1 } }),
  item('view', 'Stampa', 'view.print', 'Stampa', { icon: '🖶' }),
];
