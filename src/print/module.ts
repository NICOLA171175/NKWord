/**
 * Anteprima di stampa, come File > Stampa in Word: a sinistra le
 * impostazioni (copie, pagine, orientamento) e il pulsante Stampa, a destra
 * le pagine che verranno stampate. Il comando `view.print` (Ctrl+P) apre
 * questa anteprima quando il modulo è presente.
 */

import type { AsyncCommand, DocumentLayout, Editor, EditorModule } from '../core';
import { createDomPrinter, type Printer } from './output';
import { clampCopies, DEFAULT_SETTINGS, selectPages, withOrientation, type PrintSettings } from './pages';

export const PRINT_MODULE_ID = 'print';

export interface PrintCommandArgs {
  /** Apre l'anteprima di stampa. */
  'print.preview': void;
  /** Chiude l'anteprima e torna al documento. */
  'print.close': void;
  /** Stampa con le impostazioni scelte. */
  'print.print': void;
}

export type PrintPreview =
  | { readonly layout: DocumentLayout; readonly pages: readonly number[]; readonly sheets: number }
  | { readonly layout?: DocumentLayout; readonly error: string };

export interface PrintOptions {
  readonly printer?: Printer;
}

export interface PrintModule extends EditorModule {
  readonly isOpen: boolean;
  readonly settings: PrintSettings;
  update(changes: Partial<PrintSettings>): void;
  open(): void;
  close(): void;
  /** Impaginazione con l'orientamento scelto e pagine da stampare. */
  preview(): PrintPreview;
  /** Stampa; false se le impostazioni non sono valide. */
  print(): Promise<boolean>;
  onChange(listener: () => void): () => void;
}

export function createPrintModule(options: PrintOptions = {}): PrintModule {
  let editor: Editor | undefined;
  let settings = DEFAULT_SETTINGS;
  let isOpen = false;
  let printer = options.printer;
  let cache: { doc: unknown; orientation: PrintSettings['orientation']; layout: DocumentLayout } | undefined;
  const listeners = new Set<() => void>();
  const notify = () => listeners.forEach((l) => l());

  const preview = (): PrintPreview => {
    const engine = editor?.services.layout;
    if (!editor || !engine) return { error: "L'impaginazione non è disponibile: impossibile mostrare l'anteprima." };
    const { doc, selection } = editor.state;
    if (cache?.doc !== doc || cache.orientation !== settings.orientation) {
      cache = { doc, orientation: settings.orientation, layout: engine.layout(withOrientation(doc, settings.orientation)) };
    }
    const { layout } = cache;
    const current = engine.caretRect(layout, selection.focus)?.page ?? 0;
    const chosen = selectPages(settings, layout.pages.length, current);
    if ('error' in chosen) return { layout, error: chosen.error };
    return { layout, pages: chosen.pages, sheets: chosen.pages.length * settings.copies };
  };

  const module: PrintModule = {
    id: PRINT_MODULE_ID,
    get isOpen() {
      return isOpen;
    },
    get settings() {
      return settings;
    },
    update(changes) {
      settings = { ...settings, ...changes };
      if (changes.copies !== undefined) settings = { ...settings, copies: clampCopies(changes.copies) };
      notify();
    },
    open() {
      isOpen = true;
      notify();
    },
    close() {
      if (!isOpen) return;
      isOpen = false;
      cache = undefined;
      notify();
    },
    preview,
    async print() {
      const p = preview();
      if ('error' in p) return false;
      printer ??= createDomPrinter();
      await printer({ layout: p.layout, pages: p.pages, copies: settings.copies });
      return true;
    },
    onChange(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    activate(ed) {
      editor = ed;
      const commands: AsyncCommand<void>[] = [
        {
          id: 'print.preview',
          label: 'Stampa',
          async: true,
          isEnabled: () => !!ed.services.layout,
          async run() {
            module.open();
            return null;
          },
        },
        {
          id: 'print.close',
          label: 'Chiudi anteprima di stampa',
          async: true,
          async run() {
            module.close();
            return null;
          },
        },
        {
          id: 'print.print',
          label: 'Stampa',
          async: true,
          isEnabled: () => !('error' in preview()),
          async run() {
            if (await module.print()) module.close();
            return null;
          },
        },
      ];
      const disposers = commands.map((c) => ed.commands.register(c));
      // Un documento nuovo o modificato rende vecchia l'anteprima.
      disposers.push(ed.events.on('change', () => isOpen && notify()));
      return () => {
        disposers.forEach((d) => d());
        isOpen = false;
        editor = undefined;
        cache = undefined;
      };
    },
  };
  return module;
}
