/**
 * Collegamento del pennello con l'interfaccia: doppio clic sul pulsante,
 * applicazione al rilascio del mouse sul documento e puntatore a pennello.
 */

import type { Editor } from '../core';
import type { FormatPainterModule } from './module';
import './styles.css';

export interface FormatPainterHost {
  readonly editor: Editor;
  readonly ribbon: { readonly element: HTMLElement };
  readonly view: { readonly element: HTMLElement };
}

/** Classe sulla vista del documento mentre il pennello è acceso. */
export const PAINTER_ACTIVE_CLASS = 'formatpainter-active';

export function mountFormatPainter(host: FormatPainterHost, module: FormatPainterModule): () => void {
  const { editor } = host;
  const doc = host.view.element.ownerDocument;

  // Il doppio clic arriva dopo due clic (acceso, spento): lo si riaccende fisso.
  // La barra ridisegna i pulsanti cambiando scheda, quindi si ascolta sulla barra intera.
  const onDblClick = (e: MouseEvent) => {
    const button = (e.target as Element | null)?.closest?.('[data-command="formatPainter.toggle"]');
    if (button && !(button as HTMLButtonElement).disabled) void editor.executeAsync('formatPainter.toggle', { sticky: true });
  };

  // Si applica al rilascio del mouse, dopo che la vista ha aggiornato la selezione.
  let pressed = false;
  const onDown = (e: MouseEvent) => {
    pressed = e.button === 0 && module.mode !== 'off';
  };
  const onUp = () => {
    if (!pressed) return;
    pressed = false;
    if (editor.canExecute('formatPainter.apply')) editor.execute('formatPainter.apply');
  };

  const update = () => host.view.element.classList.toggle(PAINTER_ACTIVE_CLASS, module.mode !== 'off');

  host.ribbon.element.addEventListener('dblclick', onDblClick);
  host.view.element.addEventListener('mousedown', onDown);
  doc.addEventListener('mouseup', onUp);
  const off = module.onChange(update);
  update();
  return () => {
    host.ribbon.element.removeEventListener('dblclick', onDblClick);
    host.view.element.removeEventListener('mousedown', onDown);
    doc.removeEventListener('mouseup', onUp);
    off();
    host.view.element.classList.remove(PAINTER_ACTIVE_CLASS);
  };
}
