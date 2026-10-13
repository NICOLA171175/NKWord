/**
 * Tabella dei simboli: si apre sotto il pulsante "Simbolo" della scheda
 * Inserisci. Un clic su un simbolo lo scrive nel documento; la tabella resta
 * aperta per inserirne altri, come la finestra Simbolo di Word.
 */

import type { Editor } from '../core';
import { charCode, charFromCode, SYMBOL_CATEGORIES, symbolName } from './catalog';
import type { SymbolsModule } from './module';
import './styles.css';

export interface SymbolsHost {
  readonly editor: Editor;
  readonly ribbon?: { readonly element: HTMLElement };
  /** Riporta il focus al documento. */
  readonly input?: { focus(): void };
}

function el<K extends keyof HTMLElementTagNameMap>(tag: K, className: string, text = ''): HTMLElementTagNameMap[K] {
  const e = document.createElement(tag);
  e.className = className;
  e.textContent = text;
  return e;
}

export function mountSymbols(host: SymbolsHost, module: SymbolsModule): () => void {
  const { editor } = host;
  const doc = host.ribbon?.element.ownerDocument ?? document;
  const win = doc.defaultView;
  let category = SYMBOL_CATEGORIES[0]!.id;
  let current: string | null = null;

  const panel = el('div', 'symbols-panel');
  panel.setAttribute('role', 'dialog');
  panel.setAttribute('aria-label', 'Simbolo');
  panel.hidden = true;

  const anchor = () => host.ribbon?.element.querySelector<HTMLElement>('[data-command="symbol.show"]') ?? null;

  // Parti fisse: intestazione, scelta del gruppo, anteprima con codice.
  const head = el('div', 'symbols-head');
  head.append(el('strong', '', 'Simbolo'));
  const x = el('button', 'symbols-close', '✕');
  x.type = 'button';
  x.title = 'Chiudi';
  head.append(x);

  const recentTitle = el('div', 'symbols-title', 'Simboli usati di recente');
  const recentGrid = el('div', 'symbols-grid symbols-recent');

  const pickRow = el('label', 'symbols-pick', 'Gruppo ');
  const select = el('select', 'symbols-select');
  for (const c of SYMBOL_CATEGORIES) {
    const o = el('option', '', c.label);
    o.value = c.id;
    select.append(o);
  }
  pickRow.append(select);
  const grid = el('div', 'symbols-grid symbols-main');

  const foot = el('div', 'symbols-foot');
  const preview = el('div', 'symbols-preview');
  const info = el('div', 'symbols-info');
  const name = el('div', 'symbols-name');
  const codeRow = el('label', 'symbols-code', 'Codice carattere ');
  const code = el('input', 'symbols-code-input');
  code.type = 'text';
  code.maxLength = 8;
  code.spellcheck = false;
  code.placeholder = 'es. 20AC';
  codeRow.append(code);
  info.append(name, codeRow);
  const insertBtn = el('button', 'symbols-insert', 'Inserisci');
  insertBtn.type = 'button';
  foot.append(preview, info, insertBtn);

  panel.append(head, recentTitle, recentGrid, pickRow, grid, foot);

  const describe = (char: string | null, fromCode = false) => {
    current = char;
    preview.textContent = char ?? '';
    name.textContent = char ? symbolName(char) : 'Codice non valido';
    if (char && !fromCode) code.value = charCode(char);
    insertBtn.disabled = !char;
  };

  const insert = (char: string) => {
    editor.execute('symbol.insert', { char });
    describe(char);
  };

  const cell = (char: string) => {
    const b = el('button', 'symbols-cell', char);
    b.type = 'button';
    b.title = `${symbolName(char)} (${charCode(char)})`;
    b.setAttribute('aria-label', symbolName(char));
    b.dataset.char = char;
    // Il clic non toglie il cursore dal documento.
    b.addEventListener('mousedown', (e) => e.preventDefault());
    b.addEventListener('click', () => insert(char));
    b.addEventListener('mouseenter', () => describe(char));
    b.addEventListener('focus', () => describe(char));
    return b;
  };

  const renderRecent = () => recentGrid.replaceChildren(...module.recent.map(cell));
  const renderGrid = () => {
    const chars = SYMBOL_CATEGORIES.find((c) => c.id === category)?.chars ?? [];
    grid.replaceChildren(...chars.map(cell));
  };

  const close = () => {
    if (panel.hidden) return;
    panel.hidden = true;
    host.input?.focus();
  };

  const show = () => {
    renderRecent();
    renderGrid();
    describe(current ?? module.recent[0] ?? null);
    const a = anchor();
    if (a) {
      const r = a.getBoundingClientRect();
      const width = 420;
      const maxLeft = (win?.innerWidth ?? width + 16) - width - 8 + (win?.scrollX ?? 0);
      panel.style.left = `${Math.max(8, Math.min(r.left + (win?.scrollX ?? 0), maxLeft))}px`;
      panel.style.top = `${r.bottom + (win?.scrollY ?? 0) + 4}px`;
    }
    panel.hidden = false;
  };

  x.addEventListener('click', close);
  select.addEventListener('change', () => {
    category = select.value;
    renderGrid();
  });
  code.addEventListener('input', () => describe(charFromCode(code.value), true));
  code.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && current) {
      e.preventDefault();
      insert(current);
    }
  });
  insertBtn.addEventListener('mousedown', (e) => e.preventDefault());
  insertBtn.addEventListener('click', () => current && insert(current));

  const offExec = editor.events.on('executed', ({ commandId }) => {
    if (commandId === 'symbol.show') {
      if (panel.hidden) show();
      else close();
    }
  });
  const offChange = module.onChange(() => {
    if (!panel.hidden) renderRecent();
  });

  // Clic fuori dalla tabella (ma non sul suo pulsante) o Esc: si chiude.
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
    doc.removeEventListener('mousedown', onPointer, true);
    doc.removeEventListener('keydown', onKey, true);
    panel.remove();
  };
}
