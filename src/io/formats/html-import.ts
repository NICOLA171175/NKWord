/**
 * Da HTML a modello. Gestisce sia documenti completi sia frammenti degli
 * appunti (anche quelli prodotti da Word o dai browser): paragrafi e
 * titoli, formattazione in linea e CSS, collegamenti, immagini, elenchi
 * annidati e tabelle con celle unite.
 */

import {
  createDocument,
  defaultPageSetup,
  newId,
  normalizeInlines,
  paragraph,
  textRun,
  type Block,
  type CharFormat,
  type DocumentModel,
  type FieldKind,
  type Id,
  type Inline,
  type ListDefinition,
  type ListLevel,
  type NumberFormat,
  type ParaFormat,
  type Points,
  type Table,
  type TableCell,
  type TableRow,
} from '../../core';
import { normalizeColor } from '../util/color';
import { descendant, isElement, parseHtml, textContent, type MarkupElement, type MarkupNode } from '../util/markup';
import { charFormatFromCss, paraFormatFromCss, parseBorder, parseLength, parseStyleAttr } from './css';
import { defaultListDefinition, defaultListLevel } from './lists';

interface Ctx {
  readonly char: CharFormat;
  readonly para: ParaFormat;
  readonly styleId?: Id;
  readonly pre: boolean;
  readonly list?: { readonly listId: Id; readonly level: number };
}

const SKIP = new Set(['head', 'script', 'style', 'title', 'meta', 'link', 'template', 'noscript', 'object', 'iframe', 'svg', 'math', 'select', 'button', 'xml']);
const BLOCK = new Set([
  'html', 'body', 'div', 'section', 'article', 'header', 'footer', 'main', 'nav', 'aside', 'blockquote', 'center',
  'form', 'figure', 'figcaption', 'address', 'dl', 'dt', 'dd', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'li', 'pre',
  'fieldset', 'legend', 'details', 'summary', 'caption',
]);
/** Blocchi che producono sempre un paragrafo, anche se vuoti. */
const PARAGRAPH_LIKE = new Set(['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'li', 'pre', 'dt', 'dd']);

const HEADING_STYLES: Record<string, Id> = { h1: 'Heading1', h2: 'Heading2', h3: 'Heading3', h4: 'Heading3', h5: 'Heading3', h6: 'Heading3' };
const MSO_CLASS_STYLES: Record<string, Id> = { msotitle: 'Title', msoheading1: 'Heading1', msoheading2: 'Heading2', msoheading3: 'Heading3' };
const FIELD_KINDS = new Set<FieldKind>(['page', 'numPages', 'date', 'time']);

/** Larghezza utile di una pagina A4 con i margini predefiniti. */
function defaultContentWidth(): Points {
  const p = defaultPageSetup();
  return p.width - p.margins.left - p.margins.right;
}

class Importer {
  readonly lists: Record<Id, { id: Id; levels: ListLevel[] }> = {};
  title?: string;

  constructor(private readonly knownStyles: ReadonlySet<Id>) {}

  newList(kind: 'bullet' | 'numbered'): Id {
    const id = newId();
    this.lists[id] = { id, levels: [...defaultListDefinition(id, kind).levels] };
    return id;
  }

  styleFor(el: MarkupElement): Id | undefined {
    const explicit = el.attrs['data-style'];
    if (explicit && this.knownStyles.has(explicit)) return explicit;
    const heading = HEADING_STYLES[el.name];
    if (heading) return heading;
    for (const cls of (el.attrs.class ?? '').split(/\s+/)) {
      const s = MSO_CLASS_STYLES[cls.toLowerCase()];
      if (s) return s;
    }
    return undefined;
  }
}

class BlockBuilder {
  readonly blocks: Block[] = [];
  private inlines: Inline[] = [];
  private openCtx: Ctx | null = null;

  get hasOpenParagraph(): boolean {
    return this.openCtx !== null;
  }

  add(inline: Inline, ctx: Ctx): void {
    if (!this.openCtx) this.openCtx = ctx;
    this.inlines.push(inline);
  }

  addText(raw: string, ctx: Ctx): void {
    if (ctx.pre) {
      raw.replace(/\r\n?/g, '\n').split(/(\n|\t)/).forEach((part) => {
        if (part === '\n') this.add({ type: 'break', kind: 'line' }, ctx);
        else if (part === '\t') this.add({ type: 'tab' }, ctx);
        else if (part) this.add(textRun(part.replace(/ /g, ' '), ctx.char), ctx);
      });
      return;
    }
    let text = raw.replace(/[ \t\n\r\f]+/g, ' ');
    const last = this.inlines[this.inlines.length - 1];
    const atStart = !last || (last.type === 'break') || (last.type === 'text' && last.text.endsWith(' '));
    if (atStart) text = text.replace(/^ /, '');
    if (!text) return;
    this.add(textRun(text.replace(/ /g, ' '), ctx.char), ctx);
  }

  /** Chiude il paragrafo aperto; con `force` ne crea uno vuoto se non ce n'è. */
  flush(force?: Ctx): void {
    const ctx = this.openCtx ?? force;
    if (!ctx) return;
    let inlines = this.inlines;
    // Spazi finali e <br> finale non producono testo visibile in HTML.
    while (inlines.length) {
      const last = inlines[inlines.length - 1]!;
      if (last.type === 'break' && last.kind === 'line' && !ctx.pre) inlines = inlines.slice(0, -1);
      else if (last.type === 'text' && /\s$/.test(last.text) && !ctx.pre) {
        const text = last.text.replace(/\s+$/, '');
        inlines = text ? [...inlines.slice(0, -1), { ...last, text }] : inlines.slice(0, -1);
      } else break;
    }
    const children = normalizeInlines(inlines);
    const blank = children.every((c) => c.type === 'text' && !c.text.trim());
    const format: ParaFormat = ctx.list ? { ...ctx.para, list: { listId: ctx.list.listId, level: ctx.list.level } } : ctx.para;
    this.blocks.push(paragraph(blank ? [] : children, format, ctx.styleId));
    this.inlines = [];
    this.openCtx = null;
  }
}

/** Lo sfondo di un blocco o di una cella è ombreggiatura, non evidenziazione del testo. */
function withoutBackground(css: Map<string, string>): Map<string, string> {
  const out = new Map(css);
  out.delete('background');
  out.delete('background-color');
  return out;
}

function attrLength(v: string | undefined): Points | undefined {
  if (!v) return undefined;
  return /^\d+(\.\d+)?$/.test(v.trim()) ? parseLength(`${v}px`) : parseLength(v);
}

function listFormatFor(el: MarkupElement, css: Map<string, string>): NumberFormat {
  const type = css.get('list-style-type') ?? el.attrs.type;
  switch (type) {
    case 'a':
    case 'lower-alpha':
    case 'lower-latin':
      return 'lowerLetter';
    case 'A':
    case 'upper-alpha':
    case 'upper-latin':
      return 'upperLetter';
    case 'i':
    case 'lower-roman':
      return 'lowerRoman';
    case 'I':
    case 'upper-roman':
      return 'upperRoman';
    case '1':
    case 'decimal':
      return 'decimal';
    case 'none':
      return 'none';
    case 'disc':
    case 'circle':
    case 'square':
      return 'bullet';
  }
  return el.name === 'ol' ? 'decimal' : 'bullet';
}

function walk(nodes: readonly MarkupNode[], ctx: Ctx, b: BlockBuilder, imp: Importer): void {
  for (const node of nodes) {
    if (!isElement(node)) {
      b.addText(node, ctx);
      continue;
    }
    visit(node, ctx, b, imp);
  }
}

function visit(el: MarkupElement, ctx: Ctx, b: BlockBuilder, imp: Importer): void {
  const name = el.name;
  if (SKIP.has(name)) {
    const title = name === 'title' ? el : name === 'head' ? descendant(el, 'title') : undefined;
    if (title) imp.title ??= textContent(title).trim() || undefined;
    return;
  }
  const css = parseStyleAttr(el.attrs.style);
  if (css.get('display') === 'none' || /mso-list:\s*ignore/i.test(el.attrs.style ?? '')) return;

  // Formattazione di carattere ereditata.
  let char: CharFormat = { ...ctx.char };
  switch (name) {
    case 'b':
    case 'strong':
      char = { ...char, bold: true };
      break;
    case 'i':
    case 'em':
    case 'cite':
    case 'dfn':
    case 'var':
      char = { ...char, italic: true };
      break;
    case 'u':
    case 'ins':
      char = { ...char, underline: 'single' };
      break;
    case 's':
    case 'strike':
    case 'del':
      char = { ...char, strike: true };
      break;
    case 'sup':
      char = { ...char, verticalAlign: 'superscript' };
      break;
    case 'sub':
      char = { ...char, verticalAlign: 'subscript' };
      break;
    case 'code':
    case 'kbd':
    case 'samp':
    case 'tt':
      char = { ...char, fontFamily: 'Courier New' };
      break;
    case 'mark':
      char = { ...char, highlight: '#ffff00' };
      break;
    case 'th':
      char = { ...char, bold: true };
      break;
    case 'a':
      if (el.attrs.href && !el.attrs.href.startsWith('#')) char = { ...char, link: el.attrs.href };
      break;
    case 'font': {
      const color = normalizeColor(el.attrs.color);
      if (color) char = { ...char, color };
      if (el.attrs.face) char = { ...char, fontFamily: el.attrs.face.split(',')[0]!.trim() };
      const size = Number(el.attrs.size);
      if (size >= 1 && size <= 7) char = { ...char, fontSize: [7.5, 10, 12, 13.5, 18, 24, 36][size - 1]! };
      break;
    }
  }
  char = { ...char, ...charFormatFromCss(BLOCK.has(name) ? withoutBackground(css) : css, char) };
  const inner: Ctx = { ...ctx, char, pre: ctx.pre || /^pre/.test(css.get('white-space') ?? '') };

  switch (name) {
    case 'br':
      b.add({ type: 'break', kind: css.get('page-break-before') === 'always' || css.get('break-before') === 'page' ? 'page' : 'line' }, ctx);
      return;
    case 'img': {
      const src = el.attrs.src;
      if (!src) return;
      const width = parseLength(css.get('width')) ?? attrLength(el.attrs.width) ?? 100;
      const height = parseLength(css.get('height')) ?? attrLength(el.attrs.height) ?? width;
      b.add({ type: 'image', src, width, height, ...(el.attrs.alt ? { alt: el.attrs.alt } : {}) }, ctx);
      return;
    }
    case 'hr':
      b.flush();
      return;
    case 'table':
      b.flush();
      b.blocks.push(buildTable(el, inner, imp));
      return;
    case 'ul':
    case 'ol':
      visitList(el, css, inner, b, imp);
      return;
  }

  const field = el.attrs['data-field'] as FieldKind | undefined;
  if (field && FIELD_KINDS.has(field)) {
    b.add({ type: 'field', kind: field, format: char }, ctx);
    return;
  }

  if (!BLOCK.has(name)) {
    walk(el.children, inner, b, imp);
    return;
  }

  // Blocco: chiude il paragrafo aperto e ne apre uno nuovo con il formato del blocco.
  // Un contenitore (div, section...) con blocchi al suo interno passa ai figli
  // solo le proprietà che si ereditano o si sommano (allineamento, rientri);
  // margini, bordi e interruzioni valgono per il contenitore, non per ogni figlio.
  const own = paraFormatFromCss(css, char.fontSize ?? 12);
  const container = !PARAGRAPH_LIKE.has(name) && el.children.some((c) => isElement(c) && (BLOCK.has(c.name) || ['table', 'ul', 'ol'].includes(c.name)));
  const inherited: ParaFormat = container
    ? {
        ...(own.align ? { align: own.align } : {}),
        ...(own.lineSpacing ? { lineSpacing: own.lineSpacing } : {}),
        ...(own.indentLeft !== undefined ? { indentLeft: (ctx.para.indentLeft ?? 0) + own.indentLeft } : {}),
        ...(own.indentRight !== undefined ? { indentRight: (ctx.para.indentRight ?? 0) + own.indentRight } : {}),
      }
    : { ...own, ...(own.indentLeft !== undefined ? { indentLeft: (ctx.para.indentLeft ?? 0) + own.indentLeft } : {}) };
  let para: ParaFormat = { ...ctx.para, ...inherited };
  const align = el.attrs.align?.toLowerCase();
  if (align === 'center' || align === 'right' || align === 'justify' || align === 'left') para = { ...para, align };
  if (name === 'center') para = { ...para, align: 'center' };
  if (name === 'blockquote') para = { ...para, indentLeft: (para.indentLeft ?? 0) + 36 };
  const styleId = imp.styleFor(el) ?? ctx.styleId;
  const blockCtx: Ctx = {
    ...inner,
    para,
    ...(styleId ? { styleId } : {}),
    pre: ctx.pre || name === 'pre' || /^pre/.test(css.get('white-space') ?? ''),
  };
  const pre: Ctx = name === 'pre' ? { ...blockCtx, char: { fontFamily: 'Courier New', ...blockCtx.char } } : blockCtx;

  b.flush();
  const before = b.blocks.length;
  walk(el.children, pre, b, imp);
  if (PARAGRAPH_LIKE.has(name) && b.blocks.length === before && !b.hasOpenParagraph) b.flush(pre);
  else b.flush();
}

function visitList(el: MarkupElement, css: Map<string, string>, ctx: Ctx, b: BlockBuilder, imp: Importer): void {
  const format = listFormatFor(el, css);
  const start = Number(el.attrs.start);
  let listId: Id;
  let level: number;
  if (ctx.list) {
    listId = ctx.list.listId;
    level = Math.min(8, ctx.list.level + 1);
  } else {
    listId = imp.newList(format === 'bullet' ? 'bullet' : 'numbered');
    level = 0;
  }
  const def = imp.lists[listId]!;
  const current = def.levels[level]!;
  if (current.format !== format || (Number.isFinite(start) && start > 0)) {
    def.levels[level] = defaultListLevel(format, level, Number.isFinite(start) && start > 0 ? start : 1);
  }
  b.flush();
  const listCtx: Ctx = { ...ctx, list: { listId, level } };
  for (const c of el.children) {
    if (!isElement(c)) {
      if (c.trim()) b.addText(c, listCtx);
      continue;
    }
    visit(c, listCtx, b, imp);
  }
  b.flush();
}

function rowsOf(table: MarkupElement): MarkupElement[] {
  const rows: MarkupElement[] = [];
  for (const c of table.children) {
    if (!isElement(c)) continue;
    if (c.name === 'tr') rows.push(c);
    else if (c.name === 'thead' || c.name === 'tbody' || c.name === 'tfoot') rows.push(...rowsOf(c));
  }
  return rows;
}

function buildTable(el: MarkupElement, ctx: Ctx, imp: Importer): Table {
  const tableCss = parseStyleAttr(el.attrs.style);
  const cellCtx: Ctx = { char: ctx.char, para: {}, pre: false };
  const rows: TableRow[] = [];
  /** Celle che proseguono un'unione verticale: colonna → { righe rimaste, colSpan }. */
  const pending = new Map<number, { left: number; colSpan: number }>();
  const widths: (Points | undefined)[] = [];
  let hasBorders = Number(el.attrs.border) > 0 || !!parseBorder(tableCss.get('border'));

  for (const tr of rowsOf(el)) {
    const cells: TableCell[] = [];
    let col = 0;
    const takePending = () => {
      let p = pending.get(col);
      while (p) {
        cells.push({ id: newId(), blocks: [paragraph()], format: {}, vMerge: 'continue', ...(p.colSpan > 1 ? { colSpan: p.colSpan } : {}) });
        if (--p.left <= 0) pending.delete(col);
        col += p.colSpan;
        p = pending.get(col);
      }
    };
    const tds = tr.children.filter((c): c is MarkupElement => isElement(c) && (c.name === 'td' || c.name === 'th'));
    let allHeader = tds.length > 0;
    for (const td of tds) {
      takePending();
      if (td.name !== 'th') allHeader = false;
      const css = parseStyleAttr(td.attrs.style);
      const colSpan = Math.max(1, Math.min(63, Number(td.attrs.colspan) || 1));
      const rowSpan = Math.max(1, Number(td.attrs.rowspan) || 1);
      const width = parseLength(css.get('width')) ?? attrLength(td.attrs.width);
      if (width !== undefined && colSpan === 1 && widths[col] === undefined) widths[col] = width;
      if (parseBorder(css.get('border')) || parseBorder(css.get('border-top'))) hasBorders = true;

      const builder = new BlockBuilder();
      const inner: Ctx = { ...cellCtx, char: td.name === 'th' ? { ...cellCtx.char, bold: true } : cellCtx.char };
      walk(td.children, { ...inner, char: { ...inner.char, ...charFormatFromCss(withoutBackground(css), inner.char) } }, builder, imp);
      builder.flush();
      const shading = normalizeColor(css.get('background-color') ?? css.get('background')?.split(/\s+/)[0] ?? td.attrs.bgcolor);
      const va = css.get('vertical-align') ?? td.attrs.valign;
      const cell: TableCell = {
        id: newId(),
        blocks: builder.blocks.length ? builder.blocks : [paragraph()],
        format: {
          ...(shading ? { shading } : {}),
          ...(va === 'middle' ? { verticalAlign: 'center' as const } : va === 'bottom' ? { verticalAlign: 'bottom' as const } : {}),
        },
        ...(colSpan > 1 ? { colSpan } : {}),
        ...(rowSpan > 1 ? { vMerge: 'restart' as const } : {}),
      };
      cells.push(cell);
      if (rowSpan > 1) pending.set(col, { left: rowSpan - 1, colSpan });
      col += colSpan;
    }
    takePending();
    if (!cells.length) continue;
    rows.push({ id: newId(), cells, ...(allHeader || tr.attrs['data-header'] ? { isHeader: true } : {}) });
  }
  if (!rows.length) rows.push({ id: newId(), cells: [{ id: newId(), blocks: [paragraph()], format: {} }] });

  const gridCols = Math.max(...rows.map((r) => r.cells.reduce((n, c) => n + (c.colSpan ?? 1), 0)));
  const total = parseLength(tableCss.get('width')) ?? attrLength(el.attrs.width);
  const tableWidth = total && !String(el.attrs.width ?? '').endsWith('%') ? total : defaultContentWidth();
  const known = widths.slice(0, gridCols).reduce<number>((n, w) => n + (w ?? 0), 0);
  const unknownCount = Array.from({ length: gridCols }, (_, i) => widths[i]).filter((w) => w === undefined).length;
  const rest = unknownCount ? Math.max(18, (tableWidth - known) / unknownCount) : 0;
  const columnWidths = Array.from({ length: gridCols }, (_, i) => Math.round((widths[i] ?? rest) * 100) / 100);

  const line = { style: 'single' as const, width: 0.5, color: '#000000' };
  const align = tableCss.get('margin-left') === 'auto' && tableCss.get('margin-right') === 'auto' ? 'center' : el.attrs.align === 'center' || el.attrs.align === 'right' ? el.attrs.align : undefined;
  return {
    type: 'table',
    id: newId(),
    columnWidths,
    rows,
    format: {
      ...(align ? { align } : {}),
      ...(hasBorders ? { borders: { top: line, right: line, bottom: line, left: line, insideH: line, insideV: line } } : {}),
    },
  };
}

/** Converte HTML (documento o frammento) in un documento del modello. */
export function htmlToDocument(html: string): DocumentModel {
  const base = createDocument();
  const imp = new Importer(new Set(Object.keys(base.styles.styles)));
  const root = parseHtml(html);
  const b = new BlockBuilder();
  walk(root.children, { char: {}, para: {}, pre: false }, b, imp);
  b.flush();
  const blocks = b.blocks.length ? b.blocks : [paragraph()];
  const lists: Record<Id, ListDefinition> = {};
  for (const [id, def] of Object.entries(imp.lists)) lists[id] = { id, levels: def.levels };
  return {
    ...base,
    sections: [{ ...base.sections[0]!, blocks }],
    lists,
    metadata: { ...base.metadata, ...(imp.title ? { title: imp.title } : {}) },
  };
}
