/**
 * Interfaccia del recupero: l'avviso "Documenti recuperati" sotto la barra
 * multifunzione e l'indicatore del salvataggio automatico nella barra di stato.
 */

import type { Editor } from '../core';
import type { AutosaveModule } from './module';
import type { RecoverySnapshot } from './store';
import './styles.css';

export interface AutosaveHost {
  readonly editor: Editor;
  readonly ribbon?: { readonly element: HTMLElement };
  readonly statusBar?: { readonly element: HTMLElement };
}

const time = (d: Date) => d.toLocaleTimeString('it-IT', { hour: '2-digit', minute: '2-digit' });
const dateTime = (iso: string) => {
  const d = new Date(iso);
  return `${d.toLocaleDateString('it-IT', { day: 'numeric', month: 'long' })} alle ${time(d)}`;
};

function el<K extends keyof HTMLElementTagNameMap>(tag: K, className: string, text = ''): HTMLElementTagNameMap[K] {
  const e = document.createElement(tag);
  e.className = className;
  e.textContent = text;
  return e;
}

export function mountAutosave(host: AutosaveHost, module: AutosaveModule): () => void {
  const { editor } = host;
  const doc = host.ribbon?.element.ownerDocument ?? document;
  const win = doc.defaultView;

  const banner = el('div', 'autosave-banner');
  banner.setAttribute('role', 'alert');
  banner.hidden = true;
  let closed = false;

  const restore = (s: RecoverySnapshot) => {
    if (module.dirty && win && !win.confirm('Il documento aperto ha modifiche non salvate. Sostituirlo con quello recuperato?')) return;
    editor.execute('autosave.restore', { id: s.id });
  };

  const renderBanner = () => {
    const items = module.recovered;
    banner.hidden = closed || items.length === 0;
    if (banner.hidden) return;
    banner.replaceChildren();
    const head = el('div', 'autosave-banner-head');
    head.append(
      el('strong', '', 'Documenti recuperati'),
      el('span', '', items.length === 1
        ? ' NKoffice si è chiuso prima che questo documento fosse salvato. Puoi riaprirlo da qui.'
        : ' NKoffice si è chiuso prima che questi documenti fossero salvati. Puoi riaprirli da qui.'),
    );
    const close = el('button', 'autosave-close', '✕');
    close.title = 'Chiudi (le copie restano disponibili in File)';
    close.addEventListener('click', () => {
      closed = true;
      renderBanner();
    });
    head.append(close);
    banner.append(head);
    for (const s of items) {
      const row = el('div', 'autosave-item');
      const open = el('button', 'autosave-restore', 'Ripristina');
      open.addEventListener('click', () => restore(s));
      const drop = el('button', 'autosave-discard', 'Elimina');
      drop.addEventListener('click', () => void editor.executeAsync('autosave.discard', { id: s.id }));
      row.append(el('span', 'autosave-title', s.title || 'Documento senza titolo'), el('span', 'autosave-date', ` · ultima modifica il ${dateTime(s.savedAt)}`), open, drop);
      banner.append(row);
    }
  };

  const indicator = el('span', 'autosave-status');
  const renderIndicator = () => {
    indicator.textContent = !module.dirty
      ? ''
      : module.lastSavedAt
        ? `Copia di recupero salvata alle ${time(module.lastSavedAt)}`
        : 'Modifiche non salvate';
  };

  const render = () => {
    renderBanner();
    renderIndicator();
  };
  // "Recupera" in File riapre l'avviso se era stato chiuso.
  const offExec = editor.events.on('executed', ({ commandId }) => {
    if (commandId === 'autosave.show') {
      closed = false;
      renderBanner();
    }
  });

  // File > Nuovo o Apri sostituiscono il documento: con modifiche non salvate
  // si chiede conferma, come fa Word prima di chiudere un documento.
  const offFilter = editor.addTransactionFilter((tr, previous) => {
    const label = tr.meta?.label;
    if (!module.dirty || !win || (label !== 'file.new' && label !== 'file.open')) return tr;
    const message =
      label === 'file.new'
        ? 'Il documento ha modifiche non salvate, che andranno perse. Creare comunque un nuovo documento?\n\nPer salvarlo prima scegli Annulla e poi File > Salva.'
        : 'Il documento ha modifiche non salvate, che andranno perse. Aprire comunque l\'altro documento?\n\nPer salvarlo prima scegli Annulla e poi File > Salva.';
    return win.confirm(message) ? tr : { ...tr, state: previous };
  });

  if (host.ribbon) host.ribbon.element.after(banner);
  else doc.body.prepend(banner);
  host.statusBar?.element.append(indicator);
  const off = module.onChange(render);
  render();

  // Chiudendo o nascondendo la finestra si scrive subito la copia.
  const flush = () => void module.flush();
  const onVisibility = () => {
    if (doc.visibilityState === 'hidden') flush();
  };
  win?.addEventListener('pagehide', flush);
  doc.addEventListener('visibilitychange', onVisibility);

  return () => {
    off();
    offExec();
    offFilter();
    win?.removeEventListener('pagehide', flush);
    doc.removeEventListener('visibilitychange', onVisibility);
    banner.remove();
    indicator.remove();
  };
}
