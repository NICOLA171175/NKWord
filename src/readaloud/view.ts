/**
 * Barra della Lettura ad alta voce, come in Word: compare in alto a destra
 * sul documento mentre si legge, con Indietro, Riproduci/Pausa, Avanti,
 * Impostazioni (velocità e voce) e Chiudi. Fa scorrere la pagina perché la
 * parola letta resti visibile.
 */

import { MAX_RATE, MIN_RATE, RATE_STEP, sortVoices, voiceLabel } from './text';
import type { ReadAloudModule } from './module';
import './styles.css';

export interface ReadAloudHost {
  /** Vista del documento: la barra si appoggia al suo angolo in alto a destra. */
  readonly view?: { readonly element: HTMLElement };
  /** Riporta il focus al documento. */
  readonly input?: { focus(): void };
}

function button(d: Document, className: string, text: string, title: string): HTMLButtonElement {
  const b = d.createElement('button');
  b.type = 'button';
  b.className = `readaloud-btn ${className}`;
  b.textContent = text;
  b.title = title;
  b.setAttribute('aria-label', title);
  return b;
}

/** Velocità scritta per l'utente, es. "1,5×". */
export function rateLabel(rate: number): string {
  return `${rate.toLocaleString('it-IT', { maximumFractionDigits: 1 })}×`;
}

export function mountReadAloud(host: ReadAloudHost, module: ReadAloudModule): () => void {
  const d = document;
  const bar = d.createElement('div');
  bar.className = 'readaloud-bar';
  bar.setAttribute('role', 'toolbar');
  bar.setAttribute('aria-label', 'Leggi ad alta voce');
  bar.hidden = true;

  const prev = button(d, 'readaloud-prev', '⏮', 'Indietro (paragrafo precedente)');
  const play = button(d, 'readaloud-play', '⏸', 'Pausa');
  const next = button(d, 'readaloud-next', '⏭', 'Avanti (paragrafo successivo)');
  const settings = button(d, 'readaloud-settings', '⚙', 'Impostazioni');
  const close = button(d, 'readaloud-close', '✕', 'Interrompi (Esc)');
  settings.setAttribute('aria-expanded', 'false');

  // Impostazioni: velocità e voce, come il menu dell'ingranaggio di Word.
  const panel = d.createElement('div');
  panel.className = 'readaloud-panel';
  panel.hidden = true;
  const rateRow = d.createElement('label');
  rateRow.className = 'readaloud-field';
  const rateTitle = d.createElement('span');
  rateTitle.textContent = 'Velocità di lettura';
  const rateValue = d.createElement('span');
  rateValue.className = 'readaloud-rate-value';
  const rateInput = d.createElement('input');
  rateInput.type = 'range';
  rateInput.className = 'readaloud-rate';
  rateInput.min = String(MIN_RATE);
  rateInput.max = String(MAX_RATE);
  rateInput.step = String(RATE_STEP);
  const rateScale = d.createElement('span');
  rateScale.className = 'readaloud-scale';
  rateScale.innerHTML = '<span>Più lenta</span><span>Più veloce</span>';
  const rateHead = d.createElement('span');
  rateHead.className = 'readaloud-field-head';
  rateHead.append(rateTitle, rateValue);
  rateRow.append(rateHead, rateInput, rateScale);

  const voiceRow = d.createElement('label');
  voiceRow.className = 'readaloud-field';
  const voiceTitle = d.createElement('span');
  voiceTitle.className = 'readaloud-field-head';
  voiceTitle.textContent = 'Selezione voce';
  const voiceSelect = d.createElement('select');
  voiceSelect.className = 'readaloud-voice';
  voiceRow.append(voiceTitle, voiceSelect);
  panel.append(rateRow, voiceRow);

  const notice = d.createElement('div');
  notice.className = 'readaloud-notice';
  notice.setAttribute('role', 'alert');
  notice.hidden = true;
  const noticeText = d.createElement('span');
  const noticeClose = button(d, 'readaloud-notice-close', '✕', 'Chiudi');
  notice.append(noticeText, noticeClose);

  bar.append(prev, play, next, settings, close, panel);
  d.body.append(bar, notice);

  let voiceKey = '';
  let lastWord: ReadAloudModule['word'] = null;
  let shownMessage: string | null = null;

  const place = () => {
    const rect = host.view?.element.getBoundingClientRect();
    const top = rect && rect.height ? rect.top + 8 : 120;
    const right = rect && rect.width ? Math.max(8, window.innerWidth - rect.right + 24) : 24;
    for (const e of [bar, notice]) Object.assign(e.style, { top: `${top}px`, right: `${right}px` });
  };

  const fillVoices = () => {
    const voices = sortVoices(module.voices());
    const key = voices.map((v) => v.name).join('\n');
    if (key !== voiceKey) {
      voiceKey = key;
      voiceSelect.replaceChildren();
      const auto = d.createElement('option');
      auto.value = '';
      auto.textContent = voices.length ? 'Automatica (italiana)' : 'Nessuna voce installata';
      voiceSelect.append(auto);
      for (const v of voices) {
        const o = d.createElement('option');
        o.value = v.name;
        o.textContent = voiceLabel(v);
        voiceSelect.append(o);
      }
    }
    voiceSelect.value = module.voiceName && module.voices().some((v) => v.name === module.voiceName) ? module.voiceName : '';
  };

  /** Fa scorrere la vista solo se la parola letta è uscita dallo schermo. */
  const reveal = () => {
    const el = host.view?.element.querySelector<HTMLElement>('.readaloud-word');
    const box = host.view?.element;
    if (!el || !box) return;
    const r = el.getBoundingClientRect();
    const b = box.getBoundingClientRect();
    if (!b.height) return;
    const margin = Math.min(60, b.height / 4);
    if (r.top < b.top + margin) box.scrollTop -= b.top + margin - r.top;
    else if (r.bottom > b.bottom - margin) box.scrollTop += r.bottom - (b.bottom - margin);
  };

  const render = () => {
    const status = module.status;
    const wasHidden = bar.hidden;
    bar.hidden = status === 'stopped';
    if (bar.hidden) {
      panel.hidden = true;
      settings.setAttribute('aria-expanded', 'false');
    } else if (wasHidden) place();
    const playing = status === 'playing';
    play.textContent = playing ? '⏸' : '▶';
    play.title = playing ? 'Pausa' : 'Riproduci';
    play.setAttribute('aria-label', play.title);
    play.classList.toggle('is-paused', !playing);
    rateInput.value = String(module.rate);
    rateValue.textContent = rateLabel(module.rate);
    fillVoices();
    if (module.message !== shownMessage) {
      shownMessage = module.message;
      noticeText.textContent = module.message ?? '';
      notice.hidden = !module.message;
      if (module.message) place();
    }
    const w = module.word;
    if (w && (w.paragraphId !== lastWord?.paragraphId || w.start !== lastWord.start)) {
      // Dopo che la vista ha ridisegnato l'evidenziazione.
      requestAnimationFrame(reveal);
    }
    lastWord = w;
  };

  const run = (f: () => void) => (e: Event) => {
    e.preventDefault();
    f();
  };
  // I pulsanti non rubano il focus al documento.
  bar.addEventListener('mousedown', (e) => {
    if (!(e.target instanceof HTMLSelectElement || e.target instanceof HTMLInputElement)) e.preventDefault();
  });
  prev.addEventListener('click', run(() => module.previous()));
  play.addEventListener('click', run(() => module.playPause()));
  next.addEventListener('click', run(() => module.next()));
  close.addEventListener('click', run(() => {
    module.stop();
    host.input?.focus();
  }));
  settings.addEventListener('click', run(() => {
    panel.hidden = !panel.hidden;
    settings.setAttribute('aria-expanded', String(!panel.hidden));
    settings.classList.toggle('is-open', !panel.hidden);
  }));
  rateInput.addEventListener('input', () => {
    rateValue.textContent = rateLabel(Number(rateInput.value));
  });
  rateInput.addEventListener('change', () => module.setRate(Number(rateInput.value)));
  voiceSelect.addEventListener('change', () => module.setVoice(voiceSelect.value || null));
  noticeClose.addEventListener('click', run(() => {
    notice.hidden = true;
  }));
  // Esc dentro la barra ferma la lettura anche se il focus è sul cursore della velocità.
  bar.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      e.preventDefault();
      module.stop();
      host.input?.focus();
    }
  });
  const onResize = () => {
    if (!bar.hidden || !notice.hidden) place();
  };
  window.addEventListener('resize', onResize);

  const off = module.onChange(render);
  render();
  return () => {
    off();
    window.removeEventListener('resize', onResize);
    bar.remove();
    notice.remove();
  };
}
