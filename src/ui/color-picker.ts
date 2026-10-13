/**
 * Pulsanti colore della barra multifunzione, come in Word: la parte sinistra
 * applica l'ultimo colore scelto (mostrato nella barretta sotto l'icona), la
 * freccia apre la tavolozza con i colori del tema, quelli standard e
 * "Altri colori…".
 */

/** Tavolozza da mostrare: colore del testo, evidenziatore o sfondo del paragrafo. */
export type PaletteKind = 'text' | 'highlight' | 'shading';

export interface Swatch {
  readonly color: string;
  readonly name: string;
}

// ---------------------------------------------------------------------------
// Colori

/** Colori del tema Office con i nomi di Word in italiano. */
const THEME: readonly Swatch[] = [
  { color: '#ffffff', name: 'Bianco, Sfondo 1' },
  { color: '#000000', name: 'Nero, Testo 1' },
  { color: '#e7e6e6', name: 'Grigio chiaro, Sfondo 2' },
  { color: '#44546a', name: 'Blu-grigio, Testo 2' },
  { color: '#4472c4', name: 'Blu, Colore 1' },
  { color: '#ed7d31', name: 'Arancione, Colore 2' },
  { color: '#a5a5a5', name: 'Grigio, Colore 3' },
  { color: '#ffc000', name: 'Oro, Colore 4' },
  { color: '#5b9bd5', name: 'Blu, Colore 5' },
  { color: '#70ad47', name: 'Verde, Colore 6' },
];

export const STANDARD_COLORS: readonly Swatch[] = [
  { color: '#c00000', name: 'Rosso scuro' },
  { color: '#ff0000', name: 'Rosso' },
  { color: '#ffc000', name: 'Arancione' },
  { color: '#ffff00', name: 'Giallo' },
  { color: '#92d050', name: 'Verde chiaro' },
  { color: '#00b050', name: 'Verde' },
  { color: '#00b0f0', name: 'Azzurro' },
  { color: '#0070c0', name: 'Blu' },
  { color: '#002060', name: 'Blu scuro' },
  { color: '#7030a0', name: 'Viola' },
];

/** I 15 colori dell'evidenziatore di Word (gli stessi ammessi da DOCX). */
export const HIGHLIGHT_COLORS: readonly Swatch[] = [
  { color: '#ffff00', name: 'Giallo' },
  { color: '#00ff00', name: 'Verde brillante' },
  { color: '#00ffff', name: 'Turchese' },
  { color: '#ff00ff', name: 'Rosa' },
  { color: '#0000ff', name: 'Blu' },
  { color: '#ff0000', name: 'Rosso' },
  { color: '#000080', name: 'Blu scuro' },
  { color: '#008080', name: 'Verde acqua' },
  { color: '#008000', name: 'Verde' },
  { color: '#800080', name: 'Viola' },
  { color: '#800000', name: 'Rosso scuro' },
  { color: '#808000', name: 'Giallo scuro' },
  { color: '#808080', name: 'Grigio 50%' },
  { color: '#c0c0c0', name: 'Grigio 25%' },
  { color: '#000000', name: 'Nero' },
];

/** Colore applicato dalla parte sinistra del pulsante finché non se ne sceglie un altro. */
export const DEFAULT_COLORS: Record<PaletteKind, string> = {
  text: '#ff0000',
  highlight: '#ffff00',
  shading: '#d9d9d9',
};

const hex2 = (n: number) => Math.round(Math.max(0, Math.min(255, n))).toString(16).padStart(2, '0');

function toHsl(color: string): [number, number, number] {
  const n = parseInt(color.slice(1), 16);
  const [r, g, b] = [(n >> 16) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255];
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [h / 6, s, l];
}

function fromHsl(h: number, s: number, l: number): string {
  if (s === 0) return `#${hex2(l * 255).repeat(3)}`;
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const channel = (t: number) => {
    const x = t < 0 ? t + 1 : t > 1 ? t - 1 : t;
    if (x < 1 / 6) return p + (q - p) * 6 * x;
    if (x < 1 / 2) return q;
    if (x < 2 / 3) return p + (q - p) * (2 / 3 - x) * 6;
    return p;
  };
  return `#${hex2(channel(h + 1 / 3) * 255)}${hex2(channel(h) * 255)}${hex2(channel(h - 1 / 3) * 255)}`;
}

/** Variante più chiara (`amount` > 0) o più scura (< 0) di un colore, come le sfumature del tema di Word. */
export function shade(color: string, amount: number): string {
  const [h, s, l] = toHsl(color);
  return fromHsl(h, s, amount > 0 ? l + (1 - l) * amount : l * (1 + amount));
}

/** Le cinque sfumature sotto ogni colore del tema, con le percentuali usate da Word. */
function variants(color: string): number[] {
  const l = toHsl(color)[2];
  if (l > 0.95) return [-0.05, -0.15, -0.25, -0.35, -0.5];
  if (l > 0.8) return [-0.1, -0.25, -0.5, -0.75, -0.9];
  if (l < 0.05) return [0.5, 0.35, 0.25, 0.15, 0.05];
  return [0.8, 0.6, 0.4, -0.25, -0.5];
}

/** Griglia dei colori del tema: 6 righe da 10 (la prima con i colori base). */
export function themeGrid(): Swatch[][] {
  const rows: Swatch[][] = [THEME.slice()];
  for (let r = 0; r < 5; r++) {
    rows.push(
      THEME.map((t) => {
        const amount = variants(t.color)[r]!;
        const pct = Math.round(Math.abs(amount) * 100);
        return { color: shade(t.color, amount), name: `${t.name}, più ${amount > 0 ? 'chiaro' : 'scuro'} ${pct}%` };
      }),
    );
  }
  return rows;
}

// ---------------------------------------------------------------------------
// Pulsante

export interface ColorButtonOptions {
  readonly kind: PaletteKind;
  readonly label: string;
  /** Icona del pulsante; il suo elemento `.icon-bar` mostra il colore corrente. */
  readonly icon: Element | null;
  /** Colore iniziale della parte sinistra (default: `DEFAULT_COLORS`). */
  readonly initial?: string | null;
  /** Applica il colore (`null` = Automatico / Nessun colore). */
  readonly onPick: (color: string | null) => void;
}

const SVG_NS = 'http://www.w3.org/2000/svg';

function chevron(): SVGElement {
  const svg = document.createElementNS(SVG_NS, 'svg');
  svg.setAttribute('viewBox', '0 0 10 10');
  svg.setAttribute('class', 'ribbon-chevron');
  svg.setAttribute('aria-hidden', 'true');
  svg.innerHTML = '<path d="M2 3.5 5 6.5 8 3.5" fill="none" stroke="currentColor" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round"/>';
  return svg;
}

export class ColorButton {
  readonly element: HTMLElement;
  readonly main: HTMLButtonElement;
  readonly arrow: HTMLButtonElement;
  readonly menu: HTMLElement;
  private current: string | null;
  private readonly bar: Element | null;
  private readonly closeOnOutside = (e: Event) => {
    if (!this.element.contains(e.target as Node)) this.close();
  };

  constructor(private readonly options: ColorButtonOptions) {
    const { kind, label, icon } = options;
    this.current = options.initial !== undefined ? options.initial : DEFAULT_COLORS[kind];
    this.element = document.createElement('div');
    this.element.className = 'ribbon-split';
    this.element.dataset.palette = kind;

    this.main = document.createElement('button');
    this.main.type = 'button';
    this.main.className = 'ribbon-button ribbon-button--icon ribbon-split-main';
    this.main.setAttribute('aria-label', label);
    if (icon) this.main.append(icon);
    this.bar = icon?.querySelector('.icon-bar') ?? null;
    this.main.addEventListener('mousedown', (e) => e.preventDefault());
    this.main.addEventListener('click', () => this.pick(this.current));

    this.arrow = document.createElement('button');
    this.arrow.type = 'button';
    this.arrow.className = 'ribbon-button ribbon-split-arrow';
    this.arrow.title = `${label}: altri colori`;
    this.arrow.setAttribute('aria-label', `${label}: scegli il colore`);
    this.arrow.setAttribute('aria-haspopup', 'true');
    this.arrow.setAttribute('aria-expanded', 'false');
    this.arrow.append(chevron());
    this.arrow.addEventListener('mousedown', (e) => e.preventDefault());
    this.arrow.addEventListener('click', () => (this.isOpen ? this.close() : this.open()));

    this.menu = this.buildMenu();
    this.menu.hidden = true;
    this.element.append(this.main, this.arrow, this.menu);
    this.element.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && this.isOpen) {
        e.stopPropagation();
        this.close();
        this.arrow.focus();
      }
    });
    this.showCurrent();
  }

  /** Colore che applica la parte sinistra del pulsante. */
  get color(): string | null {
    return this.current;
  }

  get isOpen(): boolean {
    return !this.menu.hidden;
  }

  setTitle(title: string): void {
    this.main.title = title;
  }

  setDisabled(disabled: boolean): void {
    this.main.disabled = disabled;
    this.arrow.disabled = disabled;
    if (disabled) this.close();
  }

  open(): void {
    if (this.main.disabled) return;
    this.menu.hidden = false;
    // Posizione fissa: la barra multifunzione scorre in orizzontale e taglierebbe il menu.
    const r = this.element.getBoundingClientRect();
    const width = this.menu.offsetWidth;
    const max = (globalThis.innerWidth || 0) - width - 4;
    this.menu.style.left = `${Math.max(4, max > 0 ? Math.min(r.left, max) : r.left)}px`;
    this.menu.style.top = `${r.bottom + 2}px`;
    this.arrow.setAttribute('aria-expanded', 'true');
    document.addEventListener('mousedown', this.closeOnOutside, true);
  }

  close(): void {
    if (this.menu.hidden) return;
    this.menu.hidden = true;
    this.arrow.setAttribute('aria-expanded', 'false');
    document.removeEventListener('mousedown', this.closeOnOutside, true);
  }

  destroy(): void {
    this.close();
  }

  /** Applica un colore e lo ricorda per la parte sinistra del pulsante. */
  pick(color: string | null): void {
    this.current = color;
    this.showCurrent();
    this.close();
    this.options.onPick(color);
  }

  private showCurrent(): void {
    const c = this.current ?? (this.options.kind === 'text' ? '#000000' : null);
    if (this.bar instanceof SVGElement) this.bar.style.stroke = c ?? 'transparent';
    this.main.dataset.color = c ?? '';
  }

  private buildMenu(): HTMLElement {
    const { kind } = this.options;
    const menu = document.createElement('div');
    menu.className = 'color-menu';
    menu.setAttribute('role', 'menu');
    menu.setAttribute('aria-label', this.options.label);
    menu.addEventListener('mousedown', (e) => {
      // Il documento tiene il focus, tranne che per il selettore "Altri colori".
      if (!(e.target instanceof HTMLInputElement)) e.preventDefault();
    });

    if (kind === 'text') menu.append(this.wide('Automatico', null, '#000000'));
    if (kind === 'highlight') {
      menu.append(this.grid(chunk(HIGHLIGHT_COLORS, 5)), this.wide('Nessun colore', null));
      return menu;
    }
    if (kind === 'shading') menu.append(this.wide('Nessun colore', null));
    menu.append(this.heading('Colori tema'), this.grid(themeGrid(), true), this.heading('Colori standard'), this.grid([STANDARD_COLORS.slice()]));
    menu.append(this.more());
    return menu;
  }

  private heading(text: string): HTMLElement {
    const h = document.createElement('div');
    h.className = 'color-menu-heading';
    h.textContent = text;
    return h;
  }

  private swatch(s: Swatch): HTMLButtonElement {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'color-swatch';
    b.setAttribute('role', 'menuitem');
    b.title = s.name;
    b.setAttribute('aria-label', s.name);
    b.dataset.color = s.color;
    b.style.background = s.color;
    b.addEventListener('click', () => this.pick(s.color));
    return b;
  }

  private grid(rows: Swatch[][], theme = false): HTMLElement {
    const g = document.createElement('div');
    g.className = `color-grid${theme ? ' color-grid--theme' : ''}`;
    g.style.gridTemplateColumns = `repeat(${rows[0]!.length}, 1fr)`;
    for (const row of rows) for (const s of row) g.append(this.swatch(s));
    return g;
  }

  private wide(label: string, color: string | null, sample?: string): HTMLButtonElement {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'color-menu-item';
    b.setAttribute('role', 'menuitem');
    const box = document.createElement('span');
    box.className = `color-menu-sample${sample ? '' : ' color-menu-sample--none'}`;
    if (sample) box.style.background = sample;
    const text = document.createElement('span');
    text.textContent = label;
    b.append(box, text);
    b.addEventListener('click', () => this.pick(color));
    return b;
  }

  /** "Altri colori…": il selettore di colore del sistema. */
  private more(): HTMLElement {
    const label = document.createElement('label');
    label.className = 'color-menu-item color-menu-more';
    label.setAttribute('role', 'menuitem');
    const box = document.createElement('span');
    box.className = 'color-menu-sample color-menu-sample--rainbow';
    const text = document.createElement('span');
    text.textContent = 'Altri colori…';
    const input = document.createElement('input');
    input.type = 'color';
    input.className = 'color-menu-input';
    input.value = this.current ?? '#000000';
    input.setAttribute('aria-label', 'Altri colori');
    input.addEventListener('change', () => this.pick(input.value.toLowerCase()));
    label.append(box, text, input);
    return label;
  }
}

function chunk<T>(items: readonly T[], size: number): T[][] {
  const out: T[][] = [];
  for (let i = 0; i < items.length; i += size) out.push(items.slice(i, i + size));
  return out;
}
