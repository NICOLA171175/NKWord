/**
 * Modulo Modalità lettura, come Visualizza > Modalità lettura in Word: il
 * documento a schermo intero, senza barra multifunzione, da sfogliare a
 * pagine affiancate o a scorrimento. È solo una vista: non modifica il
 * documento e non entra in Annulla.
 *
 * - `readmode.open` / `readmode.close` aprono e chiudono la vista (Esc chiude).
 * - `readmode.next` / `readmode.previous` sfogliano le pagine.
 * - `readmode.zoom`, `readmode.layout`, `readmode.pageColor` cambiano zoom,
 *   disposizione delle pagine e colore della pagina.
 */

import type { AsyncCommand, DocumentLayout, Editor, EditorModule } from '../core';

export const READ_MODE_MODULE_ID = 'readmode';

/** Pagine affiancate (si sfoglia) o una sotto l'altra (si scorre). */
export type ReadLayout = 'columns' | 'scroll';
/** Colore della pagina, come in Word: normale, seppia o inverso (testo chiaro su nero). */
export type PageColor = 'none' | 'sepia' | 'inverse';

export const READ_LAYOUTS: readonly ReadLayout[] = ['columns', 'scroll'];
export const PAGE_COLORS: readonly PageColor[] = ['none', 'sepia', 'inverse'];
export const READ_LAYOUT_LABELS: Record<ReadLayout, string> = { columns: 'Pagine affiancate', scroll: 'Scorrimento verticale' };
export const PAGE_COLOR_LABELS: Record<PageColor, string> = { none: 'Nessuno', sepia: 'Seppia', inverse: 'Inverso' };

/** Zoom in percentuale, rispetto alla pagina che riempie lo schermo. */
export const MIN_READ_ZOOM = 50;
export const MAX_READ_ZOOM = 200;
export const READ_ZOOM_STEP = 10;
export const DEFAULT_READ_ZOOM = 100;

export interface ReadModeCommandArgs {
  'readmode.open': void;
  'readmode.close': void;
  'readmode.next': void;
  'readmode.previous': void;
  'readmode.zoom': { zoom: number };
  'readmode.layout': { layout: ReadLayout };
  'readmode.pageColor': { color: PageColor };
}

export interface ReadModeModule extends EditorModule {
  readonly isOpen: boolean;
  readonly layout: ReadLayout;
  readonly pageColor: PageColor;
  /** Zoom in percentuale. */
  readonly zoom: number;
  /** Prima pagina visibile (da 0). */
  readonly page: number;
  /** Pagine mostrate insieme con le pagine affiancate (lo decide la vista). */
  readonly perScreen: number;
  /** Impaginazione del documento attuale; null se manca il modulo di impaginazione. */
  documentLayout(): DocumentLayout | null;
  open(): void;
  close(): void;
  /** Va alla pagina data (da 0), allineata all'inizio della schermata. */
  goTo(page: number): void;
  next(): void;
  previous(): void;
  setZoom(zoom: number): void;
  setLayout(layout: ReadLayout): void;
  setPageColor(color: PageColor): void;
  setPerScreen(count: number): void;
  onChange(listener: () => void): () => void;
}

export const isReadLayout = (v: unknown): v is ReadLayout => READ_LAYOUTS.includes(v as ReadLayout);
export const isPageColor = (v: unknown): v is PageColor => PAGE_COLORS.includes(v as PageColor);

export function clampReadZoom(zoom: number): number {
  if (!Number.isFinite(zoom)) return DEFAULT_READ_ZOOM;
  return Math.min(MAX_READ_ZOOM, Math.max(MIN_READ_ZOOM, Math.round(zoom)));
}

/** Prima pagina della schermata che contiene `page`, con `perScreen` pagine per schermata. */
export function screenStart(page: number, perScreen: number, total: number): number {
  const n = Math.max(1, perScreen);
  const last = Math.max(0, total - 1);
  const p = Math.min(Math.max(0, page), last);
  return p - (p % n);
}

export function createReadModeModule(): ReadModeModule {
  let editor: Editor | undefined;
  let isOpen = false;
  let layout: ReadLayout = 'columns';
  let pageColor: PageColor = 'none';
  let zoom = DEFAULT_READ_ZOOM;
  let page = 0;
  let perScreen = 1;
  let cache: { doc: unknown; layout: DocumentLayout } | undefined;
  const listeners = new Set<() => void>();
  const notify = () => [...listeners].forEach((l) => l());

  const documentLayout = (): DocumentLayout | null => {
    const engine = editor?.services.layout;
    if (!editor || !engine) return null;
    const { doc } = editor.state;
    if (cache?.doc !== doc) cache = { doc, layout: engine.layout(doc) };
    return cache.layout;
  };
  const total = () => documentLayout()?.pages.length ?? 0;
  const step = () => (layout === 'columns' ? perScreen : 1);

  const module: ReadModeModule = {
    id: READ_MODE_MODULE_ID,
    get isOpen() {
      return isOpen;
    },
    get layout() {
      return layout;
    },
    get pageColor() {
      return pageColor;
    },
    get zoom() {
      return zoom;
    },
    get page() {
      return page;
    },
    get perScreen() {
      return perScreen;
    },
    documentLayout,
    open() {
      if (isOpen || !documentLayout()) return;
      // Si parte dalla pagina dove si trova il cursore, come in Word.
      const l = documentLayout()!;
      const current = editor?.services.layout?.caretRect(l, editor.state.selection.focus)?.page ?? 0;
      isOpen = true;
      page = screenStart(current, step(), l.pages.length);
      notify();
    },
    close() {
      if (!isOpen) return;
      isOpen = false;
      cache = undefined;
      notify();
    },
    goTo(target) {
      const next = screenStart(target, step(), total());
      if (next === page) return;
      page = next;
      notify();
    },
    next() {
      if (page + step() < total()) module.goTo(page + step());
    },
    previous() {
      module.goTo(page - step());
    },
    setZoom(value) {
      const next = clampReadZoom(value);
      if (next === zoom) return;
      zoom = next;
      notify();
    },
    setLayout(value) {
      if (value === layout) return;
      layout = value;
      page = screenStart(page, step(), total());
      notify();
    },
    setPageColor(value) {
      if (value === pageColor) return;
      pageColor = value;
      notify();
    },
    setPerScreen(count) {
      const n = Math.max(1, Math.floor(count));
      if (n === perScreen) return;
      perScreen = n;
      page = screenStart(page, step(), total());
      notify();
    },
    onChange(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    activate(ed) {
      editor = ed;
      const view = <A>(id: string, label: string, run: (args: A) => void, isEnabled?: (args: A) => boolean): AsyncCommand<A> => ({
        id,
        label,
        async: true,
        ...(isEnabled ? { isEnabled: (_ctx, args: A) => isEnabled(args) } : {}),
        async run(_ctx, args) {
          run(args);
          return null;
        },
      });
      const commands: AsyncCommand<any>[] = [
        view<void>('readmode.open', 'Modalità lettura', () => module.open(), () => !!ed.services.layout),
        view<void>('readmode.close', 'Modifica documento', () => module.close()),
        view<void>('readmode.next', 'Pagina successiva', () => module.next(), () => isOpen),
        view<void>('readmode.previous', 'Pagina precedente', () => module.previous(), () => isOpen),
        view<{ zoom: number }>('readmode.zoom', 'Zoom', (a) => module.setZoom(a.zoom), (a) => typeof a?.zoom === 'number'),
        view<{ layout: ReadLayout }>('readmode.layout', 'Layout', (a) => module.setLayout(a.layout), (a) => isReadLayout(a?.layout)),
        view<{ color: PageColor }>('readmode.pageColor', 'Colore pagina', (a) => module.setPageColor(a.color), (a) => isPageColor(a?.color)),
      ];
      const disposers = commands.map((c) => ed.commands.register(c));
      // Un documento aperto o cambiato mentre si legge: si ridisegna.
      disposers.push(
        ed.events.on('change', () => {
          if (!isOpen) return;
          page = screenStart(page, step(), total());
          notify();
        }),
      );
      return () => {
        disposers.forEach((d) => d());
        isOpen = false;
        cache = undefined;
        editor = undefined;
      };
    },
  };
  return module;
}
