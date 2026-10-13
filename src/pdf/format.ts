import type { DocumentFormat, DocumentModel, LayoutEngine, TextMeasurer } from '../core';
import { createApproxMeasurer, createCanvasMeasurer, createLayoutEngine } from '../layout';
import type { ImageRasterizer } from './images';
import { layoutToPdf } from './render';

export interface PdfFormatOptions {
  /**
   * Motore di impaginazione da usare (di solito `editor.services.layout`,
   * lo stesso delle pagine a schermo). Se manca se ne crea uno con `measurer`.
   */
  readonly layout?: () => LayoutEngine | undefined;
  /** Misuratore del testo; default: canvas nel browser, approssimato altrove (come il modulo layout). */
  readonly measurer?: TextMeasurer;
  readonly rasterize?: ImageRasterizer;
  readonly now?: () => Date;
}

/** Stessa scelta del modulo layout, così le larghezze coincidono con lo schermo. */
export function defaultMeasurer(): TextMeasurer {
  const hasCanvas = typeof OffscreenCanvas !== 'undefined' || typeof document !== 'undefined';
  try {
    if (hasCanvas) return createCanvasMeasurer();
  } catch {
    // Canvas 2D non disponibile: si ripiega.
  }
  return createApproxMeasurer();
}

/** Formato PDF, solo esportazione. */
export function createPdfFormat(options: PdfFormatOptions = {}): DocumentFormat {
  let measurer = options.measurer;
  let fallback: LayoutEngine | undefined;
  const getMeasurer = () => (measurer ??= defaultMeasurer());
  return {
    id: 'pdf',
    label: 'PDF',
    extensions: ['pdf'],
    mimeType: 'application/pdf',
    async export(doc: DocumentModel) {
      const engine = options.layout?.() ?? (fallback ??= createLayoutEngine(getMeasurer()));
      return layoutToPdf(engine.layout(doc), { measurer: getMeasurer(), rasterize: options.rasterize, now: options.now });
    },
  };
}
