/**
 * Barra multifunzione: schede, gruppi e pulsanti generati da `RIBBON_ITEMS`.
 * Si aggiorna a ogni cambiamento dell'editor o della vista: un pulsante è
 * abilitato solo se il suo comando è registrato ed eseguibile.
 */

import type { Editor } from '../core';
import { sameArgs, type Keymap } from './keymap';
import type { UiModule } from './module';
import { extensionIsActive, onExtensionsChange } from './extensions';
import { ColorButton, DEFAULT_COLORS } from './color-picker';
import { buttonSize, iconSvg } from './icons';
import { RIBBON_ITEMS, RIBBON_TABS, type RibbonTab, type UiRibbonItem } from './ribbon-items';
import { commonCharValue, paraFormatAtCaret, styleIdAtCaret } from './selection-format';

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
  readonly palette?: ColorButton;
}

/** Gruppi che, come in Word, vengono per primi anche se le loro voci arrivano dai moduli. */
const FIRST_GROUPS: Partial<Record<RibbonTab, readonly string[]>> = { home: ['Annulla', 'Appunti'] };

function orderGroups(items: readonly UiRibbonItem[]): UiRibbonItem[] {
  const first = FIRST_GROUPS[items[0]?.tab ?? 'home'] ?? [];
  const rank = (i: UiRibbonItem) => (first.includes(i.group) ? first.indexOf(i.group) : first.length);
  return [...items].sort((a, b) => rank(a) - rank(b));
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
  /** Ultimo colore scelto per ogni pulsante colore (resta cambiando scheda). */
  private readonly colors = new Map<string, string | null>();

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
      // Come in Word, cambiare scheda non toglie il cursore al documento.
      b.addEventListener('mousedown', (e) => e.preventDefault());
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
    for (const c of this.controls) c.palette?.destroy();
    for (const c of this.cleanup.splice(0)) c();
    this.element.remove();
  }

  /** Aggiorna abilitazione, stato attivo e valori mostrati. */
  refresh(): void {
    for (const { item, element, palette } of this.controls) {
      element.disabled = !this.isEnabled(item);
      if (palette) palette.setDisabled(element.disabled);
      else if (element instanceof HTMLButtonElement) {
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
    for (const c of this.controls) c.palette?.destroy();
    this.panel.replaceChildren();
    this.controls = [];
    const groups = new Map<string, HTMLElement>();
    for (const item of orderGroups(this.items.filter((i) => i.tab === this.tab))) {
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
      if (item.palette) {
        const palette = this.renderPalette(item);
        palette.main.dataset.command = item.commandId;
        body.append(palette.element);
        body.parentElement!.classList.add('ribbon-group--rows');
        this.controls.push({ item, element: palette.main, palette });
        continue;
      }
      const element = item.kind === 'dropdown' && item.options ? this.renderSelect(item) : this.renderButton(item);
      element.dataset.command = item.commandId;
      body.append(element instanceof HTMLSelectElement && item.icon ? this.wrapSelect(item, element) : element);
      // Pulsanti piccoli o galleria: disposti in righe invece che in colonne.
      if (item.commandId === 'style.apply' || buttonSize(item) === 'icon') body.parentElement!.classList.add('ribbon-group--rows');
      this.controls.push({ item, element });
    }
    this.refresh();
  }

  private renderButton(item: UiRibbonItem): HTMLButtonElement {
    const b = document.createElement('button');
    b.type = 'button';
    const shortcut = this.options.keymap?.shortcutFor(item.commandId, item.args);
    b.title = shortcut ? `${item.label} (${shortcut})` : item.label;
    b.setAttribute('aria-label', item.label);
    if (item.commandId === 'style.apply') {
      // Galleria degli stili: anteprima del testo con lo stile applicato.
      b.className = 'ribbon-button ribbon-style';
      const preview = document.createElement('span');
      preview.className = 'ribbon-style-preview';
      preview.dataset.style = (item.args as { styleId?: string } | undefined)?.styleId ?? '';
      preview.textContent = 'AaBbCc';
      preview.setAttribute('aria-hidden', 'true');
      b.append(preview);
    } else {
      const icon = this.renderIcon(item);
      b.className = `ribbon-button ribbon-button--${icon ? buttonSize(item) : 'small'}`;
      if (icon) b.append(icon);
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

  /** Pulsante colore diviso: applica l'ultimo colore o apre la tavolozza. */
  private renderPalette(item: UiRibbonItem): ColorButton {
    const kind = item.palette!;
    const palette = new ColorButton({
      kind,
      label: item.label,
      icon: this.renderIcon(item),
      initial: this.colors.has(item.commandId) ? this.colors.get(item.commandId)! : DEFAULT_COLORS[kind],
      onPick: (color) => {
        this.colors.set(item.commandId, color);
        this.execute(item.commandId, { color });
      },
    });
    palette.setTitle(item.label);
    return palette;
  }

  /** Icona SVG del set, o in mancanza il testo `icon` della voce. */
  private renderIcon(item: UiRibbonItem): Element | null {
    const svg = iconSvg(item);
    if (svg) return svg;
    if (!item.icon) return null;
    const icon = document.createElement('span');
    icon.className = 'ribbon-icon';
    icon.textContent = item.icon;
    icon.setAttribute('aria-hidden', 'true');
    return icon;
  }

  /**
   * Menu a tendina con icona: si presenta come un pulsante (icona, etichetta,
   * freccia) con il menu nativo trasparente sopra, che resta il controllo vero.
   */
  private wrapSelect(item: UiRibbonItem, select: HTMLSelectElement): HTMLElement {
    const size = buttonSize(item);
    const wrap = document.createElement('div');
    wrap.className = `ribbon-dropdown ribbon-button--${size}`;
    wrap.title = item.label;
    const icon = this.renderIcon(item);
    if (icon) wrap.append(icon);
    const label = document.createElement('span');
    label.className = 'ribbon-label';
    label.textContent = item.label;
    const chevron = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    chevron.setAttribute('viewBox', '0 0 10 10');
    chevron.setAttribute('class', 'ribbon-chevron');
    chevron.setAttribute('aria-hidden', 'true');
    chevron.innerHTML = '<path d="M2 3.5 5 6.5 8 3.5" fill="none" stroke="currentColor" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round"/>';
    select.classList.add('ribbon-select-overlay');
    wrap.append(label, chevron, select);
    return wrap;
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
    if (item.needsTarget) return this.editor.canExecute(item.commandId, undefined);
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
        const f = commonCharValue(state, 'fontFamily');
        index = options.findIndex((o) => (o.args as { fontFamily: string }).fontFamily === f);
        break;
      }
      case 'format.fontSize': {
        const s = commonCharValue(state, 'fontSize');
        index = options.findIndex((o) => (o.args as { fontSize: number }).fontSize === s);
        break;
      }
      case 'paragraph.lineSpacing': {
        const ls = paraFormatAtCaret(state).lineSpacing;
        index = options.findIndex((o) => {
          const want = (o.args as { lineSpacing: { rule: string; value: number } }).lineSpacing;
          return ls?.rule === want.rule && Math.abs(ls.value - want.value) < 1e-6;
        });
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
