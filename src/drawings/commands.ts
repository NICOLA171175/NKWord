/**
 * Comandi `drawing.*`: inserire, modificare, eliminare e riordinare forme,
 * caselle di testo e grafici. Senza `id` agiscono sul disegno selezionato.
 */

import type { Command, Drawing, Editor, EditorState, Id } from '../core';
import { createDrawing, getDrawing, reorderDrawing, removeDrawing, setDrawing, type NewDrawing } from './model';
import type { DrawingSelection } from './selection';

export type DrawingChanges = Partial<Omit<Drawing, 'id' | 'kind'>>;

export interface DrawingCommandArgs {
  'drawing.insert': NewDrawing;
  'drawing.update': { readonly id?: Id; readonly changes: DrawingChanges };
  'drawing.delete': { readonly id?: Id } | undefined;
  'drawing.order': { readonly id?: Id; readonly to: 'front' | 'back' };
}

export type DrawingCommandId = keyof DrawingCommandArgs;

type DrawingCommand<K extends DrawingCommandId> = Command<DrawingCommandArgs[K]> & { readonly id: K };

/** Dove mettere un disegno nuovo: sotto il cursore, dentro i margini della pagina. */
function placement(editor: Editor, state: EditorState, width: number): { x: number; y: number } {
  const section = state.doc.sections[0];
  const margins = section?.page.margins;
  let x = margins?.left ?? 72;
  let y = margins?.top ?? 72;
  const engine = editor.services.layout;
  if (engine) {
    const layout = engine.layout(state.doc);
    const c = engine.caretRect(layout, state.selection.focus);
    const page = c ? layout.pages[c.page] : undefined;
    if (c && page) {
      x = Math.max(page.contentRect.x, Math.min(c.rect.x, page.contentRect.x + page.contentRect.width - width));
      y = Math.min(c.rect.y + c.rect.height + 6, page.height - (margins?.bottom ?? 72) - 72);
    }
  }
  return { x, y };
}

export function createDrawingCommands(editor: Editor, selection: DrawingSelection): Command<any>[] {
  const target = (state: EditorState, id?: Id) => getDrawing(state.doc, id ?? selection.id ?? '');

  const insert: DrawingCommand<'drawing.insert'> = {
    id: 'drawing.insert',
    label: 'Inserisci forma',
    isEnabled: (_ctx, spec) => !!spec && ['shape', 'textbox', 'chart'].includes(spec.kind),
    run: ({ state }, spec) => {
      const width = spec.kind === 'chart' ? 360 : 108;
      const { x, y } = placement(editor, state, width);
      const d = createDrawing(spec, state.selection.focus.paragraphId, x, y);
      selection.select(d.id);
      const label = spec.kind === 'chart' ? 'Inserisci grafico' : spec.kind === 'textbox' ? 'Inserisci casella di testo' : 'Inserisci forma';
      return { state: { ...state, doc: setDrawing(state.doc, d) }, meta: { label } };
    },
  };

  const update: DrawingCommand<'drawing.update'> = {
    id: 'drawing.update',
    label: 'Modifica forma',
    isEnabled: ({ state }, args) => !!target(state, args?.id),
    run: ({ state }, { id, changes }) => {
      const d = target(state, id);
      if (!d) return null;
      const next: Drawing = { ...d, ...changes, id: d.id, kind: d.kind };
      if (next.width < 0 || next.height < 0) return null;
      return { state: { ...state, doc: setDrawing(state.doc, next) }, meta: { label: 'Modifica forma' } };
    },
  };

  const remove: DrawingCommand<'drawing.delete'> = {
    id: 'drawing.delete',
    label: 'Elimina forma',
    isEnabled: ({ state }, args) => !!target(state, args?.id),
    run: ({ state }, args) => {
      const d = target(state, args?.id);
      if (!d) return null;
      if (selection.id === d.id) selection.select(null);
      return { state: { ...state, doc: removeDrawing(state.doc, d.id) }, meta: { label: 'Elimina forma' } };
    },
  };

  const order: DrawingCommand<'drawing.order'> = {
    id: 'drawing.order',
    label: 'Disponi',
    isEnabled: ({ state }, args) => !!target(state, args?.id),
    run: ({ state }, { id, to }) => {
      const d = target(state, id);
      if (!d) return null;
      const doc = reorderDrawing(state.doc, d.id, to);
      return { state: { ...state, doc }, meta: { label: to === 'front' ? 'Porta in primo piano' : 'Porta in secondo piano' } };
    },
  };

  return [insert, update, remove, order];
}
