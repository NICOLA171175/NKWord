/**
 * Barra multifunzione: schede, gruppi e pulsanti generati da `RIBBON_ITEMS`.
 * Si aggiorna a ogni cambiamento dell'editor o della vista: un pulsante è
 * abilitato solo se il suo comando è registrato ed eseguibile.
 */

import type { Editor } from '../core';
import { sameArgs, type Keymap } from './keymap';
import type { UiModule } from './module';
import { extensionIsActive, onExtensionsChange } from './extensions';
import { RIBBON_ITEMS, RIBBON_TABS, type RibbonTab, type UiRibbonItem } from './ribbon-items';
import { charFormatAtCaret, paraFormatAtCaret, styleIdAtCaret } from './selection-format';

export interface RibbonOptions {
  readonly items?: readonly UiRibbonItem[];
  readonly keymap?: Keymap;
  readonly initialTab?: RibbonTab;
  /** Chiede un valore all'utente (default `window.prompt`). */
  readonly ask?: (message: string, initial?: string) => string | null;
  /** Comandi che la shell gestisce da sé (es. `file.open` apre il selettore di file). */
  readonly handlers?: Partial<Record<string, () => void>>;
  /** Dopo ogni azione: di solito riporta il focus al documento. */
  readonly afterAction?: () => void;
}

interface Control {
  readonly item: UiRibbonItem;
  readonly element: HTMLButtonElement | HTMLSelectElement;
}

export class Ribbon {
  readonly element: HTMLElement;
  private readonly items: readonly UiRibbonItem[];
  private readonly tabs: RibbonTab[];
  private tab: RibbonTab;
  private controls: Control[] = [];
  private readonly panel: HTMLElement;
  private readonly tabButtons = new Map<RibbonTab, HTMLButtonElement>();
  private readonly cleanup: (() => void)[] = [];

  constructor(
    private readonly editor: Editor,
    private readonly ui: UiModule,
    private readonly options: RibbonOptions = {},
  ) {
    this.items = options.items ?? RIBBON_ITEMS;
    this.tabs = RIBBON_TABS.map((t) => t.id).filter((id) => this.items.some((i) => i.tab === id));
    this.tab = options.initialTab && this.tabs.includes(options.initialTab) ? options.initialTab : this.tabs.includes('home') ? 'home' : this.tabs[0]!;

    this.element = document.createElement('div');
    this.element.className = 'ribbon';
    const tablist = document.createElement('div');
    tablist.className = 'ribbon-tabs';
    tablist.setAttribute('role', 'tablist');
    for (const id of this.tabs) {
      const b = document.createElement('button');
      b.type = 'button';
      b.className = 'ribbon-tab';
      b.setAttribute('role', 'tab');
      b.dataset.tab = id;
      b.textContent = RIBBON_TABS.find((t) => t.id === id)!.label;
      b.addEventListener('click', () => this.selectTab(id));
      this.tabButtons.set(id, b);
      tablist.append(b);
    }
    this.panel = document.createElement('div');
    this.panel.className = 'ribbon-panel';
    this.panel.setAttribute('role', 'tabpanel');
    this.element.append(tablist, this.panel);

    this.cleanup.push(
      editor.events.on('change', () => this.refresh()),
      ui.view.events.on('change', () => this.refresh()),
      onExtensionsChange(editor, () => this.refresh()),
    );
    this.renderPanel();
  }

  get activeTab(): RibbonTab {
    return this.tab;
  }

  selectTab(tab: RibbonTab): void {
    if (!this.tabs.includes(tab)) return;
    this.tab = tab;
    this.renderPanel();
  }

  destroy(): void {
    for (const c of this.cleanup.splice(0)) c();
    this.element.remove();
  }

  /** Aggiorna abilitazione, stato attivo e valori mostrati. */
  refresh(): void {
    for (const { item, element } of this.controls) {
      element.disabled = !this.isEnabled(item);
      if (element instanceof HTMLButtonElement) {
        if (item.kind === 'toggle') element.setAttribute('aria-pressed', String(this.isActive(item)));
      } else {
        element.value = this.currentValue(item);
      }
    }
    for (const [id, b] of this.tabButtons) {
      b.setAttribute('aria-selected', String(id === this.tab));
    }
  }

  private renderPanel(): void {
    this.panel.replaceChildren();
    this.controls = [];
    const groups = new Map<string, HTMLElement>();
    for (const item of this.items.filter((i) => i.tab === this.tab)) {
      let body = groups.get(item.group);
      if (!body) {
        const group = document.createElement('div');
        group.className = 'ribbon-group';
        body = document.createElement('div');
        body.className = 'ribbon-group-body';
        const title = document.createElement('div');
        title.className = 'ribbon-group-title';
        title.textContent = item.group;
        group.append(body, title);
        this.panel.append(group);
        groups.set(item.group, body);
      }
      const element = item.kind === 'dropdown' && item.options ? this.renderSelect(item) : this.renderButton(item);
      element.dataset.command = item.commandId;
      body.append(element);
      this.controls.push({ item, element });
    }
    this.refresh();
  }

  private renderButton(item: UiRibbonItem): HTMLButtonElement {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'ribbon-button';
    const shortcut = this.options.keymap?.shortcutFor(item.commandId, item.args);
    b.title = shortcut ? `${item.label} (${shortcut})` : item.label;
    b.setAttribute('aria-label', item.label);
    if (item.icon) {
      const icon = document.createElement('span');
      icon.className = 'ribbon-icon';
      icon.textContent = item.icon;
      icon.setAttribute('aria-hidden', 'true');
      b.append(icon);
    }
    const label = document.createElement('span');
    label.className = 'ribbon-label';
    label.textContent = item.label;
    b.append(label);
    // mousedown non deve togliere il focus (e la selezione) al documento.
    b.addEventListener('mousedown', (e) => e.preventDefault());
    b.addEventListener('click', () => this.invoke(item));
    return b;
  }

  private renderSelect(item: UiRibbonItem): HTMLSelectElement {
    const s = document.createElement('select');
    s.className = 'ribbon-select';
    s.title = item.label;
    s.setAttribute('aria-label', item.label);
    const placeholder = document.createElement('option');
    placeholder.value = '';
    placeholder.textContent = item.label;
    s.append(placeholder);
    item.options!.forEach((o, i) => {
      const opt = document.createElement('option');
      opt.value = String(i);
      opt.textContent = o.label;
      s.append(opt);
    });
    s.addEventListener('change', () => {
      const option = item.options![Number(s.value)];
      if (option) this.execute(item.commandId, option.args);
      this.refresh();
    });
    return s;
  }

  /** Esegue la voce come se fosse stata cliccata. */
  invoke(item: UiRibbonItem): void {
    const handler = this.options.handlers?.[item.commandId];
    if (handler) {
      handler();
      return;
    }
    let args = item.args;
    if (item.prompt) {
      const ask = this.options.ask ?? ((m: string, v?: string) => globalThis.prompt?.(m, v) ?? null);
      args = item.prompt(ask);
      if (args === undefined) return;
    }
    this.execute(item.commandId, args);
  }

  private execute(commandId: string, args: unknown): void {
    if (this.editor.commands.has(commandId) && this.editor.canExecute(commandId, args)) {
      void this.editor.executeAsync(commandId, args);
    }
    this.options.afterAction?.();
  }

  isEnabled(item: UiRibbonItem): boolean {
    if (!this.editor.commands.has(item.commandId)) return false;
    // Gli argomenti chiesti all'utente o gestiti dalla shell non sono ancora noti.
    if (item.prompt || this.options.handlers?.[item.commandId]) return true;
    const args = item.options ? item.options[0]?.args : item.args;
    return this.editor.canExecute(item.commandId, args);
  }

  isActive(item: UiRibbonItem): boolean {
    if (!this.editor.commands.has(item.commandId)) return false;
    const state = this.editor.state;
    switch (item.commandId) {
      case 'paragraph.align':
        return paraFormatAtCaret(state).align === (item.args as { align: string }).align;
      case 'style.apply':
        return styleIdAtCaret(state) === (item.args as { styleId: string }).styleId;
      default:
        break;
    }
    if (this.ui.isActive(item.commandId, item.args)) return true;
    if (extensionIsActive(this.editor, item.commandId, item.args)) return true;
    return item.args === undefined || sameArgs(item.args, {}) ? this.editor.isActive(item.commandId) : false;
  }

  /** Valore mostrato da un menu a tendina ('' = etichetta). */
  private currentValue(item: UiRibbonItem): string {
    const options = item.options ?? [];
    const state = this.editor.state;
    let index = -1;
    switch (item.commandId) {
      case 'format.fontFamily': {
        const f = charFormatAtCaret(state).fontFamily;
        index = options.findIndex((o) => (o.args as { fontFamily: string }).fontFamily === f);
        break;
      }
      case 'format.fontSize': {
        const s = charFormatAtCaret(state).fontSize;
        index = options.findIndex((o) => (o.args as { fontSize: number }).fontSize === s);
        break;
      }
      case 'view.zoom':
        index = options.findIndex((o) => this.ui.isActive('view.zoom', o.args));
        break;
      default:
        index = options.findIndex((o) => extensionIsActive(this.editor, item.commandId, o.args));
        break;
    }
    return index >= 0 ? String(index) : '';
  }
}
