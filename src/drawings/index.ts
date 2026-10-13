/**
 * Modulo disegni (Inserisci > Forme, Casella di testo, Grafico): forme,
 * caselle di testo e grafici davanti al testo, spostabili e
 * ridimensionabili, salvati in DOCX e disegnati nel PDF.
 */

import type { EditorModule, RibbonItem, ShapeKind } from '../core';
import { addViewExtension } from '../ui/extensions';
import { createDrawingCommands } from './commands';
import { CHART_LABELS, reanchorOrphans, SHAPE_LABELS } from './model';
import { DrawingSelection } from './selection';
import { DrawingView } from './view';

export const DRAWINGS_MODULE_ID = 'drawings';

export interface DrawingsModule extends EditorModule {
  /** Disegno selezionato (stato della vista). */
  readonly selection: DrawingSelection;
  /** Vista attiva, dopo `activate`. */
  readonly view: DrawingView | null;
  /** Collega tastiera e clic all'app (chiamato da `src/ui/main.ts`). */
  mount(app: Parameters<DrawingView['mount']>[0]): () => void;
}

export function createDrawingsModule(): DrawingsModule {
  const selection = new DrawingSelection();
  let view: DrawingView | null = null;
  return {
    id: DRAWINGS_MODULE_ID,
    selection,
    get view() {
      return view;
    },
    activate(editor) {
      const v = new DrawingView(editor, selection);
      view = v;
      const disposers = [
        ...createDrawingCommands(editor, selection).map((c) => editor.commands.register(c)),
        addViewExtension(editor, v),
        // Un paragrafo con forme ancorate che sparisce (es. unito al precedente) le passa al cursore.
        editor.addTransactionFilter((tr, previous) => {
          const doc = tr.state.doc;
          if (doc === previous.doc || !doc.drawings) return tr;
          const fixed = reanchorOrphans(doc, tr.state.selection.focus.paragraphId);
          return fixed === doc ? tr : { ...tr, state: { ...tr.state, doc: fixed } };
        }),
      ];
      return () => {
        disposers.forEach((d) => d());
        v.destroy();
        if (view === v) view = null;
      };
    },
    mount(app) {
      return view?.mount(app) ?? (() => {});
    },
  };
}

type Item = RibbonItem & { readonly label: string; readonly options?: readonly { label: string; args: unknown }[] };

const SHAPES: readonly ShapeKind[] = [
  'line', 'arrow', 'doubleArrow', 'rect', 'roundRect', 'ellipse', 'triangle', 'rtTriangle', 'diamond', 'pentagon',
  'hexagon', 'star5', 'heart', 'rightArrow', 'leftArrow', 'upArrow', 'downArrow', 'wedgeRectCallout',
];

const FILLS: readonly [string, string | null][] = [
  ['Blu', '#4472c4'], ['Arancione', '#ed7d31'], ['Grigio', '#a5a5a5'], ['Oro', '#ffc000'], ['Azzurro', '#5b9bd5'],
  ['Verde', '#70ad47'], ['Rosso', '#c00000'], ['Bianco', '#ffffff'], ['Nessun riempimento', null],
];
const STROKES: readonly [string, string | null][] = [
  ['Nero', '#000000'], ['Blu scuro', '#2f528f'], ['Blu', '#4472c4'], ['Arancione', '#ed7d31'], ['Rosso', '#c00000'],
  ['Verde', '#70ad47'], ['Grigio', '#7f7f7f'], ['Nessun contorno', null],
];
const WIDTHS = [0.75, 1, 1.5, 2.25, 3, 4.5, 6];

/** Voci della barra multifunzione, raccolte dalla UI (vedi `collectRibbonItems`). */
export const ribbonItems: readonly Item[] = [
  {
    tab: 'insert', group: 'Illustrazioni', commandId: 'drawing.insert', label: 'Forme', icon: '⬠', kind: 'dropdown',
    args: { kind: 'shape' },
    options: SHAPES.map((shape) => ({ label: SHAPE_LABELS[shape], args: { kind: 'shape', shape } })),
  },
  {
    tab: 'insert', group: 'Illustrazioni', commandId: 'drawing.insert', label: 'Grafico', icon: '📊', kind: 'dropdown',
    args: { kind: 'chart' },
    options: (Object.keys(CHART_LABELS) as (keyof typeof CHART_LABELS)[]).map((type) => ({ label: CHART_LABELS[type], args: { kind: 'chart', type } })),
  },
  { tab: 'insert', group: 'Testo', commandId: 'drawing.insert', label: 'Casella di testo', icon: '▭', kind: 'button', args: { kind: 'textbox' } },
  {
    tab: 'insert', group: 'Formato forma', commandId: 'drawing.update', label: 'Riempimento', icon: '◧', kind: 'dropdown',
    options: FILLS.map(([label, fill]) => ({ label, args: { changes: { fill } } })),
  },
  {
    tab: 'insert', group: 'Formato forma', commandId: 'drawing.update', label: 'Contorno', icon: '▢', kind: 'dropdown',
    options: STROKES.map(([label, stroke]) => ({ label, args: { changes: { stroke } } })),
  },
  {
    tab: 'insert', group: 'Formato forma', commandId: 'drawing.update', label: 'Spessore', icon: '≡', kind: 'dropdown',
    options: WIDTHS.map((w) => ({ label: `${String(w).replace('.', ',')} pt`, args: { changes: { strokeWidth: w } } })),
  },
  { tab: 'insert', group: 'Formato forma', commandId: 'drawing.order', label: 'Porta avanti', icon: '⧉', kind: 'button', args: { to: 'front' } },
  { tab: 'insert', group: 'Formato forma', commandId: 'drawing.order', label: 'Porta indietro', icon: '⧈', kind: 'button', args: { to: 'back' } },
  { tab: 'insert', group: 'Formato forma', commandId: 'drawing.delete', label: 'Elimina forma', icon: '✕', kind: 'button', args: {} },
];

export { createDrawingCommands, type DrawingChanges, type DrawingCommandArgs, type DrawingCommandId } from './commands';
export { chartPrimitives, type ChartPrimitive } from './chart';
export { chartToGrid, gridToChart, openChartEditor, parseNumber, type ChartGrid } from './chart-editor';
export { drawingPaths, shapeOutline, type Segment, type ShapePath } from './geometry';
export * from './model';
export { drawingElement } from './render';
export { DrawingSelection } from './selection';
export { DrawingView, resizeGeometry } from './view';
