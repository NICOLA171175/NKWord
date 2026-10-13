/** Normalizzazione dei colori in `#rrggbb` e tavolozza delle evidenziazioni di Word. */

const NAMED: Record<string, string> = {
  black: '#000000', white: '#ffffff', red: '#ff0000', green: '#008000', lime: '#00ff00', blue: '#0000ff',
  yellow: '#ffff00', cyan: '#00ffff', aqua: '#00ffff', magenta: '#ff00ff', fuchsia: '#ff00ff',
  gray: '#808080', grey: '#808080', silver: '#c0c0c0', maroon: '#800000', olive: '#808000',
  navy: '#000080', purple: '#800080', teal: '#008080', orange: '#ffa500', pink: '#ffc0cb',
  brown: '#a52a2a', darkblue: '#00008b', darkred: '#8b0000', darkgreen: '#006400', lightgray: '#d3d3d3',
  lightgrey: '#d3d3d3', darkgray: '#a9a9a9', darkgrey: '#a9a9a9',
};

const hex2 = (n: number) => Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, '0');

/** Converte un colore CSS o DOCX in `#rrggbb`; `undefined` se non riconosciuto o trasparente. */
export function normalizeColor(value: string | undefined): string | undefined {
  if (!value) return undefined;
  const v = value.trim().toLowerCase();
  if (!v || v === 'auto' || v === 'transparent' || v === 'none' || v === 'inherit') return undefined;
  let m = /^#?([0-9a-f]{6})$/.exec(v);
  if (m) return `#${m[1]}`;
  m = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/.exec(v);
  if (m) return `#${m[1]}${m[1]}${m[2]}${m[2]}${m[3]}${m[3]}`;
  m = /^rgba?\(\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)\s*(?:,\s*([\d.]+)\s*)?\)$/.exec(v);
  if (m) {
    if (m[4] !== undefined && Number(m[4]) === 0) return undefined;
    return `#${hex2(Number(m[1]))}${hex2(Number(m[2]))}${hex2(Number(m[3]))}`;
  }
  return NAMED[v];
}

/** Colori ammessi da `<w:highlight>` in DOCX. */
export const DOCX_HIGHLIGHTS: Record<string, string> = {
  black: '#000000',
  blue: '#0000ff',
  cyan: '#00ffff',
  green: '#00ff00',
  magenta: '#ff00ff',
  red: '#ff0000',
  yellow: '#ffff00',
  white: '#ffffff',
  darkBlue: '#000080',
  darkCyan: '#008080',
  darkGreen: '#008000',
  darkMagenta: '#800080',
  darkRed: '#800000',
  darkYellow: '#808000',
  darkGray: '#808080',
  lightGray: '#c0c0c0',
};

export function highlightName(color: string): string | undefined {
  const c = color.toLowerCase();
  for (const [name, hex] of Object.entries(DOCX_HIGHLIGHTS)) if (hex === c) return name;
  return undefined;
}
