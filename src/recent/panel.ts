/**
 * Elenco "Documenti recenti": si apre sotto il pulsante della scheda File.
 * Un clic sul nome riapre il documento, la ✕ lo toglie dall'elenco.
 */

import type { Editor } from '../core';
import type { OpenedFile } from '../ui/shell';
import { dirName } from './desktop';
import type { RecentModule, RecentOpener } from './module';
import type { RecentEntry } from './store';
import './styles.css';

export interface RecentHost extends RecentOpener {
  readonly editor: Editor;
  readonly ribbon?: { readonly element: HTMLElement };
  onFileOpened(listener: (opened: OpenedFile) => void): () => void;
  /** Riporta il focus al documento. */
  readonly input?: { focus(): void };
}

function el<K extends keyof HTMLElementTagNameMap>(tag: K, className: string, text = ''): HTMLElementTagNameMap[K] {
  const e = document.createElement(tag);
  e.className = className;
  e.textContent = text;
  return e;
}

/** Quando è stato usato, in parole: "Oggi alle 10:30", "Ieri", "12 ottobre". */
export function whenUsed(iso: string, today: Date = new Date()): string {
  const d = new Date(iso);
  const day = (x: Date) => new Date(x.getFullYear(), x.getMonth(), x.getDate()).getTime();
  const days = Math.round((day(today) - day(d)) / 86400000);
  const time = d.toLocaleTimeString('it-IT', { hour: '2-digit', minute: '2-digit' });
  if (days === 0) return `Oggi alle ${time}`;
  if (days === 1) return `Ieri alle ${time}`;
  return d.toLocaleDateString('it-IT', { day: 'numeric', month: 'long', ...(d.getFullYear() === today.getFullYear() ? {} : { year: 'numeric' }) });
}

/** Dove si trova il file, sotto il nome. */
export function whereIs(entry: RecentEntry): string {
  if (entry.path) return dirName(entry.path);
  if (entry.handle) return 'Sul computer';
  return 'Scegli di nuovo il file per riaprirlo';
}

const MESSAGES: Partial<Record<string, string>> = {
  missing: 'Impossibile trovare il file. Potrebbe essere stato spostato, rinominato o eliminato.',
  denied: 'Il browser non ha dato il permesso di leggere il file.',
  failed: 'Impossibile aprire il file.',
};

export function mountRecent(host: RecentHost, module: RecentModule): () => void {
  const { editor } = host;
  const doc = host.ribbon?.element.ownerDocument ?? document;
  const win = doc.defaultView;

  const panel = el('div', 'recent-panel');
  panel.setAttribute('role', 'dialog');
  panel.setAttribute('aria-label', 'Documenti recenti');
  panel.hidden = true;

  const anchor = () => host.ribbon?.element.querySelector<HTMLElement>('[data-command="recent.show"]') ?? null;

  const close = () => {
    if (panel.hidden) return;
    panel.hidden = true;
    host.input?.focus();
  };

  const openEntry = async (entry: RecentEntry) => {
    close();
    const result = await module.open(entry.id);
    const message = MESSAGES[result];
    if (!message || !win) return;
    if (result === 'missing') {
      if (win.confirm(`${message}\n\n${entry.name}\n\nToglierlo dall'elenco dei documenti recenti?`)) await module.remove(entry.id);
    } else win.alert(message);
  };

  const render = () => {
    panel.replaceChildren();
    const head = el('div', 'recent-head');
    head.append(el('strong', '', 'Documenti recenti'));
    const x = el('button', 'recent-close', '✕');
    x.type = 'button';
    x.title = 'Chiudi';
    x.addEventListener('click', close);
    head.append(x);
    panel.append(head);
    if (!module.entries.length) {
      panel.append(el('p', 'recent-empty', 'Nessun documento recente. I file che apri o salvi compariranno qui.'));
      return;
    }
    const list = el('ul', 'recent-list');
    for (const entry of module.entries) {
      const row = el('li', 'recent-item');
      const open = el('button', 'recent-open');
      open.type = 'button';
      open.title = entry.path ?? entry.name;
      open.append(el('span', 'recent-name', entry.name), el('span', 'recent-where', `${whereIs(entry)} · ${whenUsed(entry.usedAt)}`));
      open.addEventListener('click', () => void openEntry(entry));
      const drop = el('button', 'recent-remove', '✕');
      drop.type = 'button';
      drop.title = "Rimuovi dall'elenco";
      drop.setAttribute('aria-label', `Rimuovi ${entry.name} dall'elenco`);
      drop.addEventListener('click', () => void module.remove(entry.id));
      row.append(open, drop);
      list.append(row);
    }
    const clearAll = el('button', 'recent-clear', 'Cancella elenco');
    clearAll.type = 'button';
    clearAll.addEventListener('click', () => {
      if (!win || win.confirm('Togliere tutti i documenti dall\'elenco? I file restano dove sono.')) void module.clear();
    });
    panel.append(list, clearAll);
  };

  const show = () => {
    render();
    const a = anchor();
    if (a) {
      const r = a.getBoundingClientRect();
      panel.style.left = `${Math.max(8, r.left + (win?.scrollX ?? 0))}px`;
      panel.style.top = `${r.bottom + (win?.scrollY ?? 0) + 4}px`;
    }
    panel.hidden = false;
  };

  const offExec = editor.events.on('executed', ({ commandId }) => {
    if (commandId === 'recent.show') {
      if (panel.hidden) show();
      else close();
    }
  });
  const offChange = module.onChange(() => {
    if (!panel.hidden) render();
  });
  const detach = module.attach(host);
  const offOpened = host.onFileOpened((opened) => void module.record(opened));

  // Clic fuori dall'elenco (ma non sul suo pulsante) o Esc: si chiude.
  const onPointer = (e: Event) => {
    const t = e.target as Node | null;
    if (panel.hidden || !t || panel.contains(t) || anchor()?.contains(t)) return;
    close();
  };
  const onKey = (e: KeyboardEvent) => {
    if (e.key === 'Escape' && !panel.hidden) close();
  };
  doc.addEventListener('mousedown', onPointer, true);
  doc.addEventListener('keydown', onKey, true);
  doc.body.append(panel);

  return () => {
    offExec();
    offChange();
    offOpened();
    detach();
    doc.removeEventListener('mousedown', onPointer, true);
    doc.removeEventListener('keydown', onKey, true);
    panel.remove();
  };
}
