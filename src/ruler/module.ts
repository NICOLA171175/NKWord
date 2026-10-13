/**
 * Modulo Righello: Visualizza > Righello mostra o nasconde il righello
 * orizzontale (la scelta resta fra un avvio e l'altro).
 *
 * Comandi (fuori dal catalogo del core, vedi `RulerCommandArgs`):
 * - `ruler.indent` imposta i rientri dei paragrafi selezionati;
 * - `ruler.tabs` aggiunge, sposta o toglie una tabulazione;
 * - `view.ruler` mostra o nasconde il righello (stato della vista, non entra
 *   in Annulla).
 * I margini trascinati sul righello usano `page.setup`.
 */

import type { AsyncCommand, Command, Editor, EditorModule, TabAlignment } from '../core';
import { addViewExtension } from '../ui/extensions';
import { setIndents, setTabs, type IndentChange, type TabChange } from './model';
import { RulerView } from './view';

export const RULER_MODULE_ID = 'ruler';
const KEY = 'nkoffice-righello';

export interface RulerCommandArgs {
  'ruler.indent': IndentChange;
  'ruler.tabs': TabChange;
  'view.ruler': { visible?: boolean } | undefined;
}

/** Dove si ricorda se il righello è visibile (predefinito: localStorage). */
export interface RulerStorage {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
}

export interface RulerOptions {
  readonly storage?: RulerStorage | null;
}

export interface RulerModule extends EditorModule {
  readonly visible: boolean;
  /** Tipo di tabulazione scelto nel quadratino a sinistra del righello. */
  tabType: TabAlignment;
  readonly view: RulerView | null;
  setVisible(visible: boolean): void;
  /** Mette il righello sopra il documento (chiamato da `src/ui/main.ts`). */
  mount(app: { readonly view: { readonly element: HTMLElement } }): () => void;
}

function browserStorage(): RulerStorage | null {
  try {
    return globalThis.localStorage ?? null;
  } catch {
    return null;
  }
}

export function createRulerModule(options: RulerOptions = {}): RulerModule {
  const storage = options.storage === undefined ? browserStorage() : options.storage;
  let visible = true;
  try {
    visible = storage?.getItem(KEY) !== 'no';
  } catch {
    // Memoria non disponibile: righello visibile.
  }
  const listeners = new Set<() => void>();
  let view: RulerView | null = null;
  let editor: Editor | null = null;

  const module: RulerModule = {
    id: RULER_MODULE_ID,
    tabType: 'left',
    get visible() {
      return visible;
    },
    get view() {
      return view;
    },
    setVisible(next) {
      if (next === visible) return;
      visible = next;
      try {
        storage?.setItem(KEY, next ? 'si' : 'no');
      } catch {
        // Vale solo per questa sessione.
      }
      view?.setVisible(visible);
      for (const l of [...listeners]) l();
    },
    activate(ed: Editor) {
      editor = ed;
      const indent: Command<IndentChange> = {
        id: 'ruler.indent',
        label: 'Rientro',
        isEnabled: (_ctx, a) => !!a,
        run: ({ state }, a) => setIndents(state, a),
      };
      const tabs: Command<TabChange> = {
        id: 'ruler.tabs',
        label: 'Tabulazioni',
        isEnabled: (_ctx, a) => !!a && (a.add !== undefined || a.remove !== undefined || a.move !== undefined),
        run: ({ state }, a) => setTabs(state, a),
      };
      const toggle: AsyncCommand<{ visible?: boolean } | undefined> = {
        id: 'view.ruler',
        label: 'Righello',
        async: true,
        run: async (_ctx, a) => {
          module.setVisible(a?.visible ?? !visible);
          return null;
        },
      };
      const offs = [
        ed.commands.register(indent),
        ed.commands.register(tabs),
        ed.commands.register(toggle),
        addViewExtension(ed, {
          id: RULER_MODULE_ID,
          isActive: (commandId) => commandId === 'view.ruler' && visible,
          decoratePages: (pages, layout, zoom) => view?.update(pages, layout, zoom),
          onChange: (listener) => {
            listeners.add(listener);
            return () => listeners.delete(listener);
          },
        }),
      ];
      return () => {
        offs.forEach((off) => off());
        view?.destroy();
        view = null;
        editor = null;
      };
    },
    mount(app) {
      if (!editor) return () => {};
      const v = new RulerView(editor, module);
      view = v;
      v.attach(app.view.element);
      v.setVisible(visible);
      // Primo disegno: le pagine sono già sullo schermo.
      for (const l of [...listeners]) l();
      return () => {
        v.destroy();
        if (view === v) view = null;
      };
    },
  };
  return module;
}
