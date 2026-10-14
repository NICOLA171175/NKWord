/**
 * Da modello a HTML. Gli stili di paragrafo diventano classi CSS (con
 * `data-style` per il ritorno), i titoli 1-3 diventano `<h1>`-`<h3>`, la
 * formattazione diretta diventa tag semantici più `style` in linea.
 */

import {
  footnoteNumbers,
  ptToPx,
  resolveCharFormat,
  resolveParaFormat,
  STYLE_NORMAL,
  type Block,
  type CharFormat,
  type DocumentModel,
  type Inline,
  type ListLevel,
  type Paragraph,
  type Table,
  type TableCell,
} from '../../core';
import { escapeAttr, escapeText } from '../util/markup';
import { borderCss, charFormatToCss, paraFormatToCss } from './css';
import { internalAnchors } from './docx/toc';

const HEADING_TAGS: Record<string, string> = { Heading1: 'h1', Heading2: 'h2', Heading3: 'h3' };

const cssClass = (styleId: string) => `s-${styleId.replace(/[^A-Za-z0-9_-]/g, '_')}`;

export interface HtmlExportOptions {
  /** Solo il contenuto (per gli appunti), senza `<html>`/`<head>` ma con il foglio di stile. */
  readonly fragment?: boolean;
}

function styleSheetCss(doc: DocumentModel): string {
  const sheet = doc.styles;
  const rules: string[] = [];
  const body = [...charFormatToCss(sheet.defaults.char), ...paraFormatToCss(sheet.defaults.para)];
  rules.push(`.nikoffice-doc{${body.join(';')}}`);
  rules.push('.nikoffice-doc p,.nikoffice-doc h1,.nikoffice-doc h2,.nikoffice-doc h3,.nikoffice-doc li{margin:0}');
  for (const style of Object.values(sheet.styles)) {
    if (style.type !== 'paragraph') continue;
    const probe: Paragraph = { type: 'paragraph', id: '', format: {}, children: [], styleId: style.id };
    const para = resolveParaFormat(sheet, probe);
    const char = resolveCharFormat(sheet, probe);
    const decls = [...charFormatToCss(char), ...paraFormatToCss(para)];
    rules.push(`.nikoffice-doc .${cssClass(style.id)}{${decls.join(';')}}`);
  }
  rules.push('.nikoffice-doc table{border-collapse:collapse}');
  rules.push('.nikoffice-doc td,.nikoffice-doc th{vertical-align:top;padding:0 5.4pt}');
  return rules.join('\n');
}

function wrapTags(content: string, f: CharFormat): string {
  let out = content;
  if (f.verticalAlign === 'superscript') out = `<sup>${out}</sup>`;
  else if (f.verticalAlign === 'subscript') out = `<sub>${out}</sub>`;
  if (f.strike) out = `<s>${out}</s>`;
  if (f.underline === 'single') out = `<u>${out}</u>`;
  if (f.italic) out = `<i>${out}</i>`;
  if (f.bold) out = `<b>${out}</b>`;
  const css = charFormatToCss(f, true);
  if (css.length) out = `<span style="${escapeAttr(css.join(';'))}">${out}</span>`;
  return out;
}

class HtmlWriter {
  private readonly footnoteNumbers: ReadonlyMap<string, number>;
  /** Paragrafi a cui puntano i collegamenti interni (`#_Toc<id>`): ricevono l'attributo id. */
  private readonly anchors: ReadonlySet<string>;

  constructor(private readonly doc: DocumentModel) {
    this.footnoteNumbers = footnoteNumbers(doc);
    this.anchors = internalAnchors(doc);
  }

  inline(inline: Inline): string {
    switch (inline.type) {
      case 'text':
        return wrapTags(escapeText(inline.text).replace(/ {2}/g, ' &nbsp;'), inline.format);
      case 'tab':
        return '<span style="white-space:pre">\t</span>';
      case 'break':
        return inline.kind === 'line' ? '<br>' : '<br style="page-break-before:always">';
      case 'field':
        return wrapTags(`<span data-field="${inline.kind}"></span>`, inline.format);
      case 'footnoteRef':
        return wrapTags(escapeText(String(this.footnoteNumbers.get(inline.noteId) ?? '')), { ...inline.format, verticalAlign: 'superscript' });
      case 'image': {
        const src = this.doc.media[inline.src] ?? inline.src;
        const w = Math.round(ptToPx(inline.width));
        const h = Math.round(ptToPx(inline.height));
        const alt = inline.alt ? ` alt="${escapeAttr(inline.alt)}"` : ' alt=""';
        return `<img src="${escapeAttr(src)}" width="${w}" height="${h}"${alt}>`;
      }
    }
  }

  inlines(children: readonly Inline[]): string {
    let out = '';
    let i = 0;
    while (i < children.length) {
      const c = children[i]!;
      const link = c.type === 'text' ? c.format.link : undefined;
      if (!link) {
        out += this.inline(c);
        i++;
        continue;
      }
      // Run consecutivi con lo stesso collegamento in un solo <a>.
      let inner = '';
      while (i < children.length) {
        const n = children[i]!;
        if (n.type !== 'text' || n.format.link !== link) break;
        inner += this.inline(n);
        i++;
      }
      out += `<a href="${escapeAttr(link)}">${inner}</a>`;
    }
    return out;
  }

  paragraphAttrs(p: Paragraph, includeListIndent: boolean): string {
    const attrs: string[] = [];
    if (this.anchors.has(`_Toc${p.id}`)) attrs.push(`id="${escapeAttr(`_Toc${p.id}`)}"`);
    if (p.styleId && p.styleId !== STYLE_NORMAL) {
      attrs.push(`class="${cssClass(p.styleId)}"`, `data-style="${escapeAttr(p.styleId)}"`);
    }
    const format = includeListIndent ? p.format : { ...p.format, indentLeft: undefined, indentFirstLine: undefined };
    const css = paraFormatToCss(format);
    if (css.length) attrs.push(`style="${escapeAttr(css.join(';'))}"`);
    return attrs.length ? ` ${attrs.join(' ')}` : '';
  }

  paragraph(p: Paragraph): string {
    const tag = (p.styleId && HEADING_TAGS[p.styleId]) || 'p';
    const content = this.inlines(p.children);
    return `<${tag}${this.paragraphAttrs(p, true)}>${content || '<br>'}</${tag}>`;
  }

  listOpen(level: ListLevel | undefined): string {
    if (!level || level.format === 'bullet' || level.format === 'none') {
      return level?.format === 'none' ? '<ul style="list-style-type:none">' : '<ul>';
    }
    const type = { decimal: '1', lowerLetter: 'a', upperLetter: 'A', lowerRoman: 'i', upperRoman: 'I' }[level.format];
    const start = level.start !== 1 ? ` start="${level.start}"` : '';
    return `<ol type="${type}"${start}>`;
  }

  blocks(blocks: readonly Block[]): string {
    const out: string[] = [];
    /** Elenchi aperti: tag di chiusura e se c'è un <li> aperto. */
    const stack: { close: string; liOpen: boolean }[] = [];
    let currentList: string | undefined;
    const closeTo = (depth: number) => {
      while (stack.length > depth) {
        const top = stack.pop()!;
        out.push(`${top.liOpen ? '</li>' : ''}${top.close}`);
      }
    };
    for (const b of blocks) {
      const list = b.type === 'paragraph' ? b.format.list : null;
      if (!list || list.listId !== currentList) closeTo(0);
      if (!list || b.type !== 'paragraph') {
        currentList = undefined;
        out.push(b.type === 'paragraph' ? this.paragraph(b) : this.table(b));
        continue;
      }
      currentList = list.listId;
      const depth = list.level + 1;
      closeTo(depth);
      const top = stack[stack.length - 1];
      if (top && stack.length === depth && top.liOpen) {
        out.push('</li>');
        top.liOpen = false;
      }
      const levels = this.doc.lists[list.listId]?.levels;
      while (stack.length < depth) {
        const parent = stack[stack.length - 1];
        if (parent && !parent.liOpen) {
          out.push('<li style="list-style-type:none">');
          parent.liOpen = true;
        }
        const open = this.listOpen(levels?.[stack.length]);
        out.push(open);
        stack.push({ close: open.startsWith('<ol') ? '</ol>' : '</ul>', liOpen: false });
      }
      out.push(`<li${this.paragraphAttrs(b, false)}>${this.inlines(b.children)}`);
      stack[stack.length - 1]!.liOpen = true;
    }
    closeTo(0);
    return out.join('\n');
  }

  cell(cell: TableCell, rowSpan: number, innerBorder: string | undefined): string {
    // Sempre <td>: <th> renderebbe il testo in grassetto, che il modello non ha.
    const tag = 'td';
    const attrs: string[] = [];
    if (cell.colSpan && cell.colSpan > 1) attrs.push(`colspan="${cell.colSpan}"`);
    if (rowSpan > 1) attrs.push(`rowspan="${rowSpan}"`);
    const css: string[] = innerBorder ? [`border:${innerBorder}`] : [];
    if (cell.format.shading) css.push(`background-color:${cell.format.shading}`);
    if (cell.format.verticalAlign) css.push(`vertical-align:${cell.format.verticalAlign === 'center' ? 'middle' : cell.format.verticalAlign}`);
    for (const side of ['top', 'right', 'bottom', 'left'] as const) {
      const b = borderCss(cell.format.borders?.[side]);
      if (b) css.push(`border-${side}:${b}`);
    }
    if (css.length) attrs.push(`style="${escapeAttr(css.join(';'))}"`);
    return `<${tag}${attrs.length ? ` ${attrs.join(' ')}` : ''}>${this.blocks(cell.blocks)}</${tag}>`;
  }

  table(t: Table): string {
    const css: string[] = ['border-collapse:collapse'];
    const borders = t.format.borders;
    if (t.format.align === 'center') css.push('margin-left:auto', 'margin-right:auto');
    else if (t.format.align === 'right') css.push('margin-left:auto');
    const width = t.columnWidths.reduce((n, w) => n + w, 0);
    css.push(`width:${Math.round(width * 100) / 100}pt`);
    const inner = borderCss(borders?.insideH ?? borders?.insideV);
    const cols = t.columnWidths.map((w) => `<col style="width:${Math.round(w * 100) / 100}pt">`).join('');

    // rowspan dalle unioni verticali: per ogni cella 'restart' conta le 'continue' sotto.
    const gridStart = (cells: readonly TableCell[], idx: number) => cells.slice(0, idx).reduce((n, c) => n + (c.colSpan ?? 1), 0);
    const cellAt = (rowIdx: number, col: number): TableCell | undefined => {
      const row = t.rows[rowIdx];
      if (!row) return undefined;
      let g = 0;
      for (const c of row.cells) {
        if (g === col) return c;
        g += c.colSpan ?? 1;
      }
      return undefined;
    };
    const rows = t.rows.map((row, r) => {
      const cells = row.cells
        .map((cell, i) => {
          if (cell.vMerge === 'continue') return '';
          let span = 1;
          if (cell.vMerge === 'restart') {
            const col = gridStart(row.cells, i);
            while (cellAt(r + span, col)?.vMerge === 'continue') span++;
          }
          return this.cell(cell, span, inner);
        })
        .join('');
      const h = row.height && row.height.rule !== 'auto' ? ` style="height:${row.height.value}pt"` : '';
      return `<tr${row.isHeader ? ' data-header="1"' : ''}${h}>${cells}</tr>`;
    });
    const border = borders ? ` border="1"` : '';
    const edge = (['top', 'right', 'bottom', 'left'] as const)
      .map((s) => {
        const b = borderCss(borders?.[s]);
        return b ? `border-${s}:${b}` : '';
      })
      .filter(Boolean);
    return `<table${border} style="${escapeAttr([...css, ...edge].join(';'))}"><colgroup>${cols}</colgroup>\n${rows.join('\n')}\n</table>`;
  }

  body(): string {
    const sections = this.doc.sections
      .map((s, i) => {
        const brk = i > 0 && s.breakType !== 'continuous' ? ' style="page-break-before:always"' : '';
        return `<div class="section"${brk}>\n${this.blocks(s.blocks)}\n</div>`;
      })
      .join('\n');
    return sections + this.footnotes();
  }

  /** Le note a piè di pagina, in HTML, vanno in fondo al documento. */
  footnotes(): string {
    const notes = [...this.footnoteNumbers.keys()].map((id) => this.doc.footnotes?.[id]);
    const html = notes.map((n) => (n ? `<div class="footnote">\n${this.blocks(n.blocks)}\n</div>` : '')).join('\n');
    return html ? `\n<div class="footnotes">\n<hr>\n${html}\n</div>` : '';
  }
}

/** Converte un documento in HTML. */
export function documentToHtml(doc: DocumentModel, options: HtmlExportOptions = {}): string {
  const writer = new HtmlWriter(doc);
  const css = styleSheetCss(doc);
  const body = `<div class="nikoffice-doc">\n${writer.body()}\n</div>`;
  if (options.fragment) return `<style>\n${css}\n</style>\n${body}`;
  const page = doc.sections[0]?.page;
  const pageCss = page
    ? `@page{size:${Math.round(page.width * 100) / 100}pt ${Math.round(page.height * 100) / 100}pt;margin:${page.margins.top}pt ${page.margins.right}pt ${page.margins.bottom}pt ${page.margins.left}pt}`
    : '';
  const meta = doc.metadata;
  const head = [
    '<meta charset="utf-8">',
    '<meta name="generator" content="NKWord">',
    `<title>${escapeText(meta.title ?? 'Documento')}</title>`,
    meta.author ? `<meta name="author" content="${escapeAttr(meta.author)}">` : '',
    meta.subject ? `<meta name="description" content="${escapeAttr(meta.subject)}">` : '',
    meta.keywords?.length ? `<meta name="keywords" content="${escapeAttr(meta.keywords.join(', '))}">` : '',
    `<style>\n${pageCss}\n${css}\n</style>`,
  ].filter(Boolean);
  return `<!DOCTYPE html>\n<html>\n<head>\n${head.join('\n')}\n</head>\n<body>\n${body}\n</body>\n</html>\n`;
}
