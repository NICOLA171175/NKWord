import type { EditorModule, TextMeasurer } from '../core';
import { createLayoutEngine, type LayoutOptions, type NikLayoutEngine } from './engine';
import { createApproxMeasurer, createCanvasMeasurer } from './measure';

export interface LayoutModuleOptions extends LayoutOptions {
  /** Default: canvas nel browser, approssimato altrove. */
  readonly measurer?: TextMeasurer;
}

function defaultMeasurer(): TextMeasurer {
  const hasCanvas = typeof OffscreenCanvas !== 'undefined' || typeof document !== 'undefined';
  try {
    if (hasCanvas) return createCanvasMeasurer();
  } catch {
    // Canvas 2D non disponibile (es. jsdom senza canvas): si ripiega.
  }
  return createApproxMeasurer();
}

/**
 * Modulo di impaginazione: pubblica `editor.services.layout`. Non registra
 * comandi (nessun id è assegnato a 'layout' in `COMMAND_OWNERS`) e non
 * disegna nulla: il DOM delle pagine è compito della UI, che legge il
 * `DocumentLayout`.
 *
 * Il nome `create…Module` senza argomenti obbligatori permette alla UI di
 * riconoscerlo e attivarlo automaticamente.
 */
export function createLayoutModule(options: LayoutModuleOptions = {}): EditorModule {
  return {
    id: 'layout',
    activate(editor) {
      const engine = createLayoutEngine(options.measurer ?? defaultMeasurer(), options);
      editor.services.layout = engine;
      return () => {
        if (editor.services.layout === engine) delete editor.services.layout;
      };
    },
  };
}
