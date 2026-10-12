/**
 * Stato della vista: zoom, segni di formattazione, modifica di intestazione
 * e piè di pagina. Non fa parte del documento né della cronologia, quindi
 * vive qui e non in `EditorState`: cambiarlo non crea un passo di
 * annullamento e non segna il documento come modificato.
 */

import { Emitter } from '../core';

export type HeaderFooterPart = 'header' | 'footer';

export interface ViewState {
  /** 1 = 100%. */
  readonly zoom: number;
  readonly formattingMarks: boolean;
  /** Parte in modifica (intestazione o piè di pagina), `null` = corpo del documento. */
  readonly headerFooter: HeaderFooterPart | null;
}

export const MIN_ZOOM = 0.1;
export const MAX_ZOOM = 5;
export const ZOOM_STEPS = [0.5, 0.75, 1, 1.25, 1.5, 2] as const;

export function clampZoom(zoom: number): number {
  if (!Number.isFinite(zoom)) return 1;
  return Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, Math.round(zoom * 100) / 100));
}

export const DEFAULT_VIEW: ViewState = { zoom: 1, formattingMarks: false, headerFooter: null };

export class ViewStore {
  readonly events = new Emitter<{ change: { readonly state: ViewState; readonly previous: ViewState } }>();
  private _state: ViewState;

  constructor(initial: Partial<ViewState> = {}) {
    this._state = { ...DEFAULT_VIEW, ...initial };
  }

  get state(): ViewState {
    return this._state;
  }

  /** Aggiorna lo stato; restituisce true se è cambiato qualcosa. */
  update(patch: Partial<ViewState>): boolean {
    const previous = this._state;
    const next: ViewState = { ...previous, ...patch };
    if (next.zoom === previous.zoom && next.formattingMarks === previous.formattingMarks && next.headerFooter === previous.headerFooter) {
      return false;
    }
    this._state = next;
    this.events.emit('change', { state: next, previous });
    return true;
  }
}
