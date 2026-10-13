/** Conversione fra formattazione del modello e dichiarazioni CSS in linea. */

import { pxToPt, type BorderLine, type CharFormat, type ParaFormat, type Points } from '../../core';
import { normalizeColor } from '../util/color';

export function parseStyleAttr(style: string | undefined): Map<string, string> {
  const out = new Map<string, string>();
  if (!style) return out;
  for (const decl of style.split(';')) {
    const i = decl.indexOf(':');
    if (i < 0) continue;
    const name = decl.slice(0, i).trim().toLowerCase();
    const value = decl.slice(i + 1).trim().replace(/\s*!important$/i, '');
    if (name) out.set(name, value);
  }
  return out;
}

/** Lunghezza CSS in punti; `fontSize` serve per `em`. */
export function parseLength(value: string | undefined, fontSize: Points = 12): Points | undefined {
  if (!value) return undefined;
  const m = /^(-?[\d.]+)\s*(pt|px|em|rem|cm|mm|in|pc)?$/i.exec(value.trim());
  if (!m) return undefined;
  const n = Number(m[1]);
  if (!Number.isFinite(n)) return undefined;
  switch ((m[2] ?? 'px').toLowerCase()) {
    case 'pt':
      return n;
    case 'px':
      return pxToPt(n);
    case 'em':
    case 'rem':
      return n * fontSize;
    case 'cm':
      return (n / 2.54) * 72;
    case 'mm':
      return (n / 25.4) * 72;
    case 'in':
      return n * 72;
    case 'pc':
      return n * 12;
  }
  return undefined;
}

const round = (n: number) => Math.round(n * 100) / 100;

const FONT_SIZE_KEYWORDS: Record<string, Points> = {
  'xx-small': 7, 'x-small': 7.5, small: 10, medium: 12, large: 13.5, 'x-large': 18, 'xx-large': 24,
};

/** Formattazione di carattere dichiarata da uno stile CSS (solo i campi presenti). */
export function charFormatFromCss(css: Map<string, string>, inherited: CharFormat): CharFormat {
  const f: { -readonly [K in keyof CharFormat]: CharFormat[K] } = {};
  const fw = css.get('font-weight');
  if (fw) f.bold = fw === 'bold' || fw === 'bolder' || Number(fw) >= 600;
  const fs = css.get('font-style');
  if (fs) f.italic = fs === 'italic' || fs === 'oblique';
  const deco = css.get('text-decoration') ?? css.get('text-decoration-line');
  if (deco) {
    if (/underline/.test(deco)) {
      const style = css.get('text-decoration-style') ?? deco;
      f.underline = /double/.test(style) ? 'double' : /dotted/.test(style) ? 'dotted' : /dashed/.test(style) ? 'dashed' : /wavy/.test(style) ? 'wave' : 'single';
    } else if (/none/.test(deco)) f.underline = null;
    if (/line-through/.test(deco)) f.strike = true;
    else if (/none/.test(deco)) f.strike = false;
  }
  const color = normalizeColor(css.get('color'));
  if (color) f.color = color;
  const bg = css.get('background-color') ?? css.get('background') ?? css.get('mso-highlight');
  if (bg) {
    const c = normalizeColor(bg.split(/\s+/)[0]);
    if (c) f.highlight = c;
  }
  const size = css.get('font-size');
  if (size) {
    const pt = FONT_SIZE_KEYWORDS[size] ?? parseLength(size, inherited.fontSize ?? 12);
    if (pt && pt > 0) f.fontSize = round(pt);
  }
  const family = css.get('font-family');
  if (family) {
    const first = family.split(',')[0]!.trim().replace(/^["']|["']$/g, '');
    if (first) f.fontFamily = first;
  }
  const va = css.get('vertical-align');
  if (va === 'super') f.verticalAlign = 'superscript';
  else if (va === 'sub') f.verticalAlign = 'subscript';
  else if (va === 'baseline') f.verticalAlign = 'baseline';
  if (css.get('font-variant') === 'small-caps') f.smallCaps = true;
  if (css.get('text-transform') === 'uppercase') f.allCaps = true;
  return f;
}

function parseBorder(value: string | undefined): BorderLine | undefined {
  if (!value) return undefined;
  const parts = value.trim().split(/\s+/);
  let style: BorderLine['style'] | undefined;
  let width: Points | undefined;
  let color: string | undefined;
  for (const part of parts) {
    const p = part.toLowerCase();
    if (['none', 'hidden'].includes(p)) style = 'none';
    else if (['solid', 'double', 'dotted', 'dashed'].includes(p)) style = p === 'solid' ? 'single' : (p as BorderLine['style']);
    else if (parseLength(p) !== undefined) width = parseLength(p);
    else if (normalizeColor(p)) color = normalizeColor(p);
  }
  if (!style || style === 'none') return undefined;
  return { style, width: width ?? 0.5, color: color ?? '#000000' };
}

/** Formattazione di paragrafo dichiarata da uno stile CSS (solo i campi presenti). */
export function paraFormatFromCss(css: Map<string, string>, fontSize: Points = 12): ParaFormat {
  const f: { -readonly [K in keyof ParaFormat]: ParaFormat[K] } = {};
  const align = css.get('text-align');
  if (align === 'center' || align === 'right' || align === 'justify') f.align = align;
  else if (align === 'left' || align === 'start') f.align = 'left';
  else if (align === 'end') f.align = 'right';

  const margin = css.get('margin')?.split(/\s+/);
  const side = (i: number) => {
    if (!margin) return undefined;
    const idx = margin.length === 1 ? 0 : margin.length === 2 ? i % 2 : margin.length === 3 ? (i === 3 ? 1 : i) : i;
    return margin[idx];
  };
  const len = (v: string | undefined) => (v === 'auto' ? undefined : parseLength(v, fontSize));
  const top = len(css.get('margin-top') ?? side(0));
  const right = len(css.get('margin-right') ?? side(1));
  const bottom = len(css.get('margin-bottom') ?? side(2));
  const left = len(css.get('margin-left') ?? side(3));
  if (top !== undefined) f.spaceBefore = round(Math.max(0, top));
  if (bottom !== undefined) f.spaceAfter = round(Math.max(0, bottom));
  const padLeft = len(css.get('padding-left'));
  if (left !== undefined || padLeft !== undefined) f.indentLeft = round((left ?? 0) + (padLeft ?? 0));
  if (right !== undefined) f.indentRight = round(right);
  const indent = len(css.get('text-indent'));
  if (indent !== undefined) f.indentFirstLine = round(indent);

  const lh = css.get('line-height');
  if (lh && lh !== 'normal') {
    if (/^[\d.]+$/.test(lh)) f.lineSpacing = { rule: 'multiple', value: Number(lh) };
    else if (lh.endsWith('%')) f.lineSpacing = { rule: 'multiple', value: Number(lh.slice(0, -1)) / 100 };
    else {
      const pt = parseLength(lh, fontSize);
      if (pt) f.lineSpacing = { rule: 'atLeast', value: round(pt) };
    }
  }
  if (css.get('page-break-before') === 'always' || css.get('break-before') === 'page') f.pageBreakBefore = true;
  if (css.get('page-break-after') === 'avoid' || css.get('break-after') === 'avoid') f.keepWithNext = true;
  const bg = normalizeColor(css.get('background-color') ?? css.get('background')?.split(/\s+/)[0]);
  if (bg) f.shading = bg;

  const all = parseBorder(css.get('border'));
  const borders = {
    top: parseBorder(css.get('border-top')) ?? all,
    right: parseBorder(css.get('border-right')) ?? all,
    bottom: parseBorder(css.get('border-bottom')) ?? all,
    left: parseBorder(css.get('border-left')) ?? all,
  };
  if (borders.top || borders.right || borders.bottom || borders.left) {
    f.borders = Object.fromEntries(Object.entries(borders).filter(([, v]) => v));
  }
  return f;
}

export { parseBorder };

// ---------------------------------------------------------------------------
// Verso CSS

const num = (n: number) => String(round(n));

export function borderCss(b: BorderLine | undefined): string | undefined {
  if (!b || b.style === 'none') return undefined;
  return `${num(b.width)}pt ${b.style === 'single' ? 'solid' : b.style} ${b.color}`;
}

export function charFormatToCss(f: CharFormat, skipTagged = false): string[] {
  const out: string[] = [];
  if (f.fontFamily) out.push(`font-family:${/[\s,]/.test(f.fontFamily) ? `'${f.fontFamily.replace(/'/g, '')}'` : f.fontFamily}`);
  if (f.fontSize) out.push(`font-size:${num(f.fontSize)}pt`);
  if (f.bold === false || (f.bold && !skipTagged)) out.push(`font-weight:${f.bold ? 'bold' : 'normal'}`);
  if (f.italic === false || (f.italic && !skipTagged)) out.push(`font-style:${f.italic ? 'italic' : 'normal'}`);
  const deco: string[] = [];
  if (f.underline && !(skipTagged && f.underline === 'single')) deco.push('underline');
  if (f.strike && !skipTagged) deco.push('line-through');
  if (deco.length) out.push(`text-decoration:${deco.join(' ')}`);
  if (f.underline && f.underline !== 'single') {
    out.push(`text-decoration-style:${f.underline === 'wave' ? 'wavy' : f.underline}`);
  }
  if (f.underline === null && f.strike !== true) out.push('text-decoration:none');
  if (f.color) out.push(`color:${f.color}`);
  if (f.highlight) out.push(`background-color:${f.highlight}`);
  if (f.verticalAlign && f.verticalAlign !== 'baseline' && !skipTagged) {
    out.push(`vertical-align:${f.verticalAlign === 'superscript' ? 'super' : 'sub'}`);
  }
  if (f.smallCaps) out.push('font-variant:small-caps');
  if (f.allCaps) out.push('text-transform:uppercase');
  return out;
}

export function paraFormatToCss(f: ParaFormat): string[] {
  const out: string[] = [];
  if (f.align) out.push(`text-align:${f.align}`);
  if (f.spaceBefore !== undefined) out.push(`margin-top:${num(f.spaceBefore)}pt`);
  if (f.spaceAfter !== undefined) out.push(`margin-bottom:${num(f.spaceAfter)}pt`);
  if (f.indentLeft !== undefined) out.push(`margin-left:${num(f.indentLeft)}pt`);
  if (f.indentRight !== undefined) out.push(`margin-right:${num(f.indentRight)}pt`);
  if (f.indentFirstLine !== undefined) out.push(`text-indent:${num(f.indentFirstLine)}pt`);
  if (f.lineSpacing) {
    out.push(`line-height:${f.lineSpacing.rule === 'multiple' ? num(f.lineSpacing.value) : `${num(f.lineSpacing.value)}pt`}`);
  }
  if (f.pageBreakBefore) out.push('page-break-before:always');
  if (f.keepWithNext) out.push('page-break-after:avoid');
  if (f.shading) out.push(`background-color:${f.shading}`);
  if (f.borders) {
    let any = false;
    for (const side of ['top', 'right', 'bottom', 'left'] as const) {
      const css = borderCss(f.borders[side]);
      if (css) out.push(`border-${side}:${css}`);
      any ||= !!css;
    }
    // Spazio fra testo e bordo, come in Word.
    if (any) out.push('padding:1pt 4pt');
  }
  return out;
}
