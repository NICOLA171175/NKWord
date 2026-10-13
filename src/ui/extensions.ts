/**
 * Estensioni della vista: un modulo può aggiungere alla UI segni sul testo
 * (es. sottolineatura ondulata rossa del controllo ortografico), voci di
 * menu contestuale, scorciatoie e lo stato dei suoi pulsanti, senza che la
 * UI lo conosca. Il modulo pubblica le estensioni nel suo `activate` con
 * `addViewExtension(editor, ...)`; la UI le legge da `editor.services`.
 * I pulsanti si dichiarano invece con l'esportazione `ribbonItems`
 * dell'index.ts del modulo (vedi `collectRibbonItems`).
 */

import type { DocumentLayout, Editor, EditorState, Id, Keybinding, Position } from '../core';

/** Segno disegnato sopra un intervallo di testo di un paragrafo. */
export interface TextMark {
  readonly paragraphId: Id;
  readonly start: number;
  readonly end: number;
  /** squiggle: ondulata (errori); underline: linea continua; highlight: sfondo. */
  readonly kind: 'squiggle' | 'underline' | 'highlight';
  readonly color: string;
  /** Classe CSS aggiuntiva, utile anche ai test. */
  readonly className?: string;
}

export interface ContextMenuItem {
  readonly label: string;
  readonly run?: () => void;
  readonly disabled?: boolean;
  /** In grassetto, come i suggerimenti di Word. */
  readonly strong?: boolean;
  /** Linea di separazione prima della voce. */
  readonly separator?: boolean;
}

export interface ViewExtension {
  readonly id: string;
  /** Segni da disegnare sul testo per lo stato dato. */
  marks?(state: EditorState): readonly TextMark[];
  /** Voci del menu contestuale per un clic destro nella posizione data. */
  contextMenu?(pos: Position, state: EditorState): readonly ContextMenuItem[];
  readonly keybindings?: readonly Keybinding[];
  /** Stato attivo di un pulsante a interruttore o di una voce di menu a tendina. */
  isActive?(commandId: string, args?: unknown): boolean;
  /**
   * Aggiunge elementi sopra le pagine impaginate (es. forme e grafici), dopo
   * testo e selezione. `pages[i]` è l'elemento della pagina `i` del layout.
   */
  decoratePages?(pages: readonly HTMLElement[], layout: DocumentLayout, zoom: number): void;
  /** Avvisa quando segni o stato cambiano senza un cambiamento dell'editor (es. dizionario caricato). */
  onChange?(listener: () => void): () => void;
}

declare module '../core/interfaces' {
  interface EditorServices {
    viewExtensions?: ViewExtension[];
  }
}

export function viewExtensions(editor: Editor): readonly ViewExtension[] {
  return editor.services.viewExtensions ?? [];
}

/** Pubblica un'estensione; restituisce la funzione per toglierla. */
export function addViewExtension(editor: Editor, extension: ViewExtension): () => void {
  const list = (editor.services.viewExtensions ??= []);
  list.push(extension);
  return () => {
    const i = list.indexOf(extension);
    if (i >= 0) list.splice(i, 1);
  };
}

/** Tutti i segni delle estensioni per lo stato dato. */
export function collectMarks(editor: Editor, state: EditorState = editor.state): TextMark[] {
  return viewExtensions(editor).flatMap((e) => e.marks?.(state) ?? []);
}

export function extensionIsActive(editor: Editor, commandId: string, args?: unknown): boolean {
  return viewExtensions(editor).some((e) => e.isActive?.(commandId, args));
}

/** Si iscrive a `onChange` di tutte le estensioni presenti. */
export function onExtensionsChange(editor: Editor, listener: () => void): () => void {
  const offs = viewExtensions(editor).map((e) => e.onChange?.(listener) ?? (() => {}));
  return () => offs.forEach((off) => off());
}
