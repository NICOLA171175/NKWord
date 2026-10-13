/**
 * Modulo UI: registra i comandi di proprietà `ui` in COMMAND_OWNERS
 * (view.*, page.headerFooter). Cambiano solo lo stato della vista, non il
 * documento, quindi sono `AsyncCommand` (comandi con effetti esterni) che
 * restituiscono `null`: nessuna transazione, nessun passo di annullamento.
 */

import type { AsyncCommand, CommandArgs, CommandId, Editor, EditorModule } from '../core';
import { ViewStore, clampZoom } from './view-state';

type UiCommand<K extends CommandId> = AsyncCommand<CommandArgs[K]> & { readonly id: K };

export interface UiModuleOptions {
  readonly view?: ViewStore;
  /** Stampa; di default `window.print`. Iniettabile per i test. */
  readonly print?: () => void;
}

export interface UiModule extends EditorModule {
  readonly view: ViewStore;
  /** Stato attivo dei comandi a interruttore della UI (l'editor non lo vede per i comandi asincroni). */
  isActive(commandId: string, args?: unknown): boolean;
}

export const UI_MODULE_ID = 'ui';

export function createUiModule(options: UiModuleOptions = {}): UiModule {
  const view = options.view ?? new ViewStore();
  const print = options.print ?? (() => globalThis.print?.());

  const zoom: UiCommand<'view.zoom'> = {
    id: 'view.zoom',
    label: 'Zoom',
    async: true,
    isEnabled: (_ctx, args) => typeof args?.zoom === 'number' && Number.isFinite(args.zoom),
    run: async (_ctx, { zoom }) => {
      view.update({ zoom: clampZoom(zoom) });
      return null;
    },
  };

  const formattingMarks: UiCommand<'view.toggleFormattingMarks'> = {
    id: 'view.toggleFormattingMarks',
    label: 'Mostra tutto',
    async: true,
    run: async () => {
      view.update({ formattingMarks: !view.state.formattingMarks });
      return null;
    },
  };

  let editor: Editor | undefined;

  const printCmd: UiCommand<'view.print'> = {
    id: 'view.print',
    label: 'Stampa',
    async: true,
    run: async () => {
      // Con il modulo stampa si apre la sua anteprima (File > Stampa), come in Word.
      if (editor?.commands.has('print.preview')) await editor.executeAsync('print.preview');
      else print();
      return null;
    },
  };

  const headerFooter: UiCommand<'page.headerFooter'> = {
    id: 'page.headerFooter',
    label: 'Intestazione e piè di pagina',
    async: true,
    isEnabled: (_ctx, args) => !!args && (args.part === 'header' || args.part === 'footer'),
    run: async (_ctx, { part, edit }) => {
      // Con il modulo intestazione e piè di pagina si apre la sua finestra.
      if (edit && editor?.commands.has('headerFooter.edit')) {
        await editor.executeAsync('headerFooter.edit', { part });
        return null;
      }
      if (edit) view.update({ headerFooter: part });
      else if (view.state.headerFooter === part) view.update({ headerFooter: null });
      return null;
    },
  };

  const commands = [zoom, formattingMarks, printCmd, headerFooter];

  return {
    id: UI_MODULE_ID,
    view,
    isActive(commandId, args) {
      switch (commandId) {
        case 'view.toggleFormattingMarks':
          return view.state.formattingMarks;
        case 'view.zoom':
          return (args as CommandArgs['view.zoom'] | undefined)?.zoom === view.state.zoom;
        case 'page.headerFooter':
          return view.state.headerFooter === (args as CommandArgs['page.headerFooter'] | undefined)?.part;
        default:
          return false;
      }
    },
    activate(ed: Editor) {
      editor = ed;
      const disposers = commands.map((c) => ed.commands.register(c));
      return () => {
        disposers.forEach((d) => d());
        editor = undefined;
      };
    },
  };
}
