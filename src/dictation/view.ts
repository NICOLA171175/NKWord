/**
 * Indicatore della dettatura, in basso al centro della finestra come in
 * Word: un microfono che pulsa mentre si ascolta, la frase che si sta
 * dicendo e il pulsante per fermare. Mostra anche gli avvisi (microfono non
 * consentito, come usare Win+H). Esc ferma la dettatura.
 */

import type { DictationModule } from './module';
import './styles.css';

export interface DictationHost {
  /** Riporta il focus al documento. */
  readonly input?: { focus(): void };
}

/** Gli avvisi informativi spariscono da soli dopo un po'. */
export const INFO_TIMEOUT_MS = 12000;

export function mountDictation(host: DictationHost, module: DictationModule, d: Document = document): () => void {
  const root = d.createElement('div');
  root.className = 'dictation';
  root.setAttribute('role', 'status');
  root.setAttribute('aria-live', 'polite');
  root.hidden = true;

  const mic = d.createElement('span');
  mic.className = 'dictation-mic';
  mic.textContent = '🎤';
  mic.setAttribute('aria-hidden', 'true');
  const text = d.createElement('span');
  text.className = 'dictation-text';
  const close = d.createElement('button');
  close.type = 'button';
  close.className = 'dictation-close';
  root.append(mic, text, close);
  d.body.append(root);

  module.focusDocument = () => host.input?.focus();
  let timer: ReturnType<typeof setTimeout> | undefined;

  const render = () => {
    clearTimeout(timer);
    const { listening, message } = module;
    root.hidden = !listening && !message;
    root.classList.toggle('dictation-listening', listening);
    root.classList.toggle('dictation-error', !listening && message?.kind === 'error');
    if (listening) {
      text.textContent = module.interim || 'In ascolto… parla pure.';
      close.textContent = 'Ferma';
      close.title = 'Ferma la dettatura (Esc)';
    } else if (message) {
      text.textContent = message.text;
      close.textContent = '✕';
      close.title = 'Chiudi';
      if (message.kind === 'info') timer = setTimeout(() => module.dismissMessage(), INFO_TIMEOUT_MS);
    }
    close.setAttribute('aria-label', close.title);
  };

  close.addEventListener('mousedown', (e) => e.preventDefault());
  close.addEventListener('click', () => {
    if (module.listening) module.stop();
    else module.dismissMessage();
    host.input?.focus();
  });
  const onKey = (e: KeyboardEvent) => {
    if (e.key !== 'Escape' || (!module.listening && !module.message)) return;
    e.preventDefault();
    e.stopPropagation();
    if (module.listening) module.stop();
    else module.dismissMessage();
  };
  d.addEventListener('keydown', onKey, true);
  const off = module.onChange(render);
  render();

  return () => {
    clearTimeout(timer);
    off();
    d.removeEventListener('keydown', onKey, true);
    if (module.focusDocument) module.focusDocument = undefined;
    root.remove();
  };
}
