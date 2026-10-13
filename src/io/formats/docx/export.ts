/**
 * Scrittura DOCX (WordprocessingML). L'ordine degli elementi dentro
 * `w:pPr`, `w:rPr`, `w:tblPr`, `w:tcPr` e `w:sectPr` segue lo schema: Word
 * rifiuta i file con elementi fuori ordine.
 */

import {
  ptToEmu,
  ptToHalfPoints,
  ptToTwips,
  footnoteNumbers,
  type Block,
  type BorderLine,
  type Borders,
  type CharFormat,
  type DocumentModel,
  type Drawing,
  type FieldKind,
  type HeaderFooter,
  type Id,
  type Inline,
  type ParaFormat,
  type Paragraph,
  type Revision,
  type Section,
  type Style,
  type Table,
} from '../../../core';
import { CHART_CT, CHART_REL, chartPartXml, drawingRunXml, WPS_NS } from '../../../drawings/docx';
import { highlightName } from '../../util/color';
import { extensionFromMime, parseDataUri } from '../../util/encoding';
import { escapeAttr, escapeText } from '../../util/markup';
import { writeZip } from '../../util/zip';
import {
  COMMENTS_CT,
  COMMENTS_EXT_CT,
  COMMENTS_EXT_REL,
  COMMENTS_REL,
  commentMarks,
  commentRangeEndXml,
  commentRangeStartXml,
  commentsExtendedXml,
  commentsXml,
} from './comments';
import {
  BUILTIN_STYLE_NAMES,
  CONTENT_TYPES,
  JC_TO_DOCX,
  NS,
  NUMFMT_TO_DOCX,
  REL,
  UNDERLINE_TO_DOCX,
} from './common';
import { internalAnchors, paragraphBookmark, tocFieldWraps, type FieldWrap } from './toc';

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const hex = (c: string) => c.replace('#', '').toUpperCase();
const val = (name: string, v: string | number) => `<w:${name} w:val="${escapeAttr(String(v))}"/>`;
const toggle = (name: string, v: boolean | undefined) => (v === undefined ? '' : v ? `<w:${name}/>` : `<w:${name} w:val="0"/>`);

interface Rel {
  readonly id: string;
  readonly type: string;
  readonly target: string;
  readonly external?: boolean;
}

/** Relazioni di una parte (document.xml, header1.xml...). */
class Rels {
  private readonly rels: Rel[] = [];
  add(type: string, target: string, external = false): string {
    const existing = this.rels.find((r) => r.type === type && r.target === target && !!r.external === external);
    if (existing) return existing.id;
    const id = `rId${this.rels.length + 1}`;
    this.rels.push({ id, type, target, external });
    return id;
  }
  xml(): string {
    const items = this.rels
      .map((r) => `<Relationship Id="${r.id}" Type="${r.type}" Target="${escapeAttr(r.target)}"${r.external ? ' TargetMode="External"' : ''}/>`)
      .join('');
    return `${XML_DECL}<Relationships xmlns="${NS.rels}">${items}</Relationships>`;
  }
  get size(): number {
    return this.rels.length;
  }
}

function borderXml(side: string, b: BorderLine | undefined, space = 0): string {
  if (!b) return '';
  const style = b.style === 'none' ? 'nil' : b.style;
  return `<w:${side} w:val="${style}" w:sz="${Math.max(2, Math.round(b.width * 8))}" w:space="${space}" w:color="${hex(b.color)}"/>`;
}

/** Distanza dal testo dei bordi di paragrafo, in punti (come l'impaginazione). */
const PARA_BORDER_SPACE: Record<string, number> = { top: 1, bottom: 1, left: 4, right: 4 };

function bordersXml(tag: string, b: Borders & { insideH?: BorderLine; insideV?: BorderLine } | undefined, order: readonly string[]): string {
  if (!b) return '';
  const space = (side: string) => (tag === 'pBdr' ? PARA_BORDER_SPACE[side] : 0);
  const inner = order.map((side) => borderXml(side, (b as Record<string, BorderLine | undefined>)[side], space(side))).join('');
  return inner ? `<w:${tag}>${inner}</w:${tag}>` : '';
}

export function charFormatXml(f: CharFormat, styleId?: string): string {
  const parts: string[] = [];
  if (styleId) parts.push(val('rStyle', styleId));
  if (f.fontFamily) {
    const font = escapeAttr(f.fontFamily);
    parts.push(`<w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:eastAsia="${font}" w:cs="${font}"/>`);
  }
  parts.push(toggle('b', f.bold), toggle('i', f.italic), toggle('caps', f.allCaps), toggle('smallCaps', f.smallCaps), toggle('strike', f.strike));
  if (f.color) parts.push(val('color', hex(f.color)));
  if (f.fontSize) parts.push(val('sz', ptToHalfPoints(f.fontSize)), val('szCs', ptToHalfPoints(f.fontSize)));
  const hl = f.highlight ? highlightName(f.highlight) : undefined;
  if (f.highlight === null) parts.push(val('highlight', 'none'));
  else if (hl) parts.push(val('highlight', hl));
  if (f.underline !== undefined) parts.push(val('u', f.underline ? UNDERLINE_TO_DOCX[f.underline] : 'none'));
  if (f.highlight && !hl) parts.push(`<w:shd w:val="clear" w:color="auto" w:fill="${hex(f.highlight)}"/>`);
  if (f.verticalAlign) parts.push(val('vertAlign', f.verticalAlign));
  const inner = parts.join('');
  return inner ? `<w:rPr>${inner}</w:rPr>` : '';
}

class DocxWriter {
  private readonly files: [string, Uint8Array][] = [];
  private readonly encoder = new TextEncoder();
  private readonly contentTypes = new Map<string, string>();
  private readonly defaults = new Map<string, string>([
    ['rels', CONTENT_TYPES.rels],
    ['xml', 'application/xml'],
  ]);
  private readonly mediaNames = new Map<string, string>();
  private readonly numIds = new Map<string, number>();
  private drawingId = 0;
  private headerCount = 0;
  private footerCount = 0;
  private usesEvenHeaders = false;
  /** Id DOCX delle note a piè di pagina (da 1; -1 e 0 sono i separatori). */
  private readonly footnoteIds: ReadonlyMap<Id, number>;
  /** Nota che si sta scrivendo in footnotes.xml. */
  private currentNote: Id | undefined;
  private revisionId = 0;
  private readonly anchors: ReadonlySet<string>;
  private bookmarkId = 0;

  private readonly comments: ReturnType<typeof commentMarks>;
  /** Forme, caselle di testo e grafici per paragrafo di ancoraggio, con la posizione nella pila. */
  private readonly drawings = new Map<Id, { drawing: Drawing; z: number }[]>();
  private chartCount = 0;

  constructor(private readonly doc: DocumentModel) {
    this.comments = commentMarks(doc);
    this.anchors = internalAnchors(doc);
    let n = 1;
    for (const id of Object.keys(doc.lists)) this.numIds.set(id, n++);
    this.footnoteIds = footnoteNumbers(doc);
    Object.values(doc.drawings ?? {}).forEach((drawing, z) => {
      const list = this.drawings.get(drawing.anchor);
      if (list) list.push({ drawing, z });
      else this.drawings.set(drawing.anchor, [{ drawing, z }]);
    });
  }

  /** Run dei disegni ancorati al paragrafo (all'inizio, come li mette Word). */
  private drawingsXml(p: Paragraph, rels: Rels): string {
    const list = p.id ? this.drawings.get(p.id) : undefined;
    if (!list) return '';
    return list
      .map(({ drawing, z }) => {
        let chartRel: string | undefined;
        if (drawing.kind === 'chart' && drawing.chart) {
          const name = `charts/chart${++this.chartCount}.xml`;
          this.addFile(`word/${name}`, chartPartXml(drawing.chart, drawing), CHART_CT);
          chartRel = rels.add(CHART_REL, name);
        }
        return drawingRunXml(drawing, ++this.drawingId, z, chartRel);
      })
      .join('');
  }

  private addFile(path: string, content: string | Uint8Array, contentType?: string): void {
    this.files.push([path, typeof content === 'string' ? this.encoder.encode(content) : content]);
    if (contentType) this.contentTypes.set(`/${path}`, contentType);
  }

  // -------------------------------------------------------------------------
  // Paragrafi e run

  /** Attributi di `w:ins`/`w:del`. */
  private revisionAttrs(r: Revision): string {
    const date = r.date ? ` w:date="${escapeAttr(r.date)}"` : '';
    return `w:id="${++this.revisionId}" w:author="${escapeAttr(r.author)}"${date}`;
  }

  private revisionTag(r: Revision): string {
    return r.kind === 'insert' ? 'w:ins' : 'w:del';
  }

  paraFormatXml(f: ParaFormat, styleId: string | undefined, mark?: CharFormat, sectPr = '', markRevision?: Revision): string {
    const p: string[] = [];
    if (styleId) p.push(val('pStyle', styleId));
    p.push(toggle('keepNext', f.keepWithNext), toggle('keepLines', f.keepLinesTogether), toggle('pageBreakBefore', f.pageBreakBefore), toggle('widowControl', f.widowControl));
    if (f.list === null) p.push('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>');
    else if (f.list) {
      const numId = this.numIds.get(f.list.listId);
      if (numId) p.push(`<w:numPr><w:ilvl w:val="${f.list.level}"/><w:numId w:val="${numId}"/></w:numPr>`);
    }
    p.push(bordersXml('pBdr', f.borders, ['top', 'left', 'bottom', 'right']));
    if (f.shading) p.push(`<w:shd w:val="clear" w:color="auto" w:fill="${hex(f.shading)}"/>`);
    if (f.tabs?.length) {
      const tabs = f.tabs
        .map((t) => `<w:tab w:val="${t.align}" w:pos="${ptToTwips(t.position)}"${t.leader && t.leader !== 'none' ? ` w:leader="${t.leader}"` : ''}/>`)
        .join('');
      p.push(`<w:tabs>${tabs}</w:tabs>`);
    }
    const spacing: string[] = [];
    if (f.spaceBefore !== undefined) spacing.push(`w:before="${ptToTwips(f.spaceBefore)}"`);
    if (f.spaceAfter !== undefined) spacing.push(`w:after="${ptToTwips(f.spaceAfter)}"`);
    if (f.lineSpacing) {
      const ls = f.lineSpacing;
      if (ls.rule === 'multiple') spacing.push(`w:line="${Math.round(ls.value * 240)}" w:lineRule="auto"`);
      else spacing.push(`w:line="${ptToTwips(ls.value)}" w:lineRule="${ls.rule}"`);
    }
    if (spacing.length) p.push(`<w:spacing ${spacing.join(' ')}/>`);
    const ind: string[] = [];
    if (f.indentLeft !== undefined) ind.push(`w:left="${ptToTwips(f.indentLeft)}"`);
    if (f.indentRight !== undefined) ind.push(`w:right="${ptToTwips(f.indentRight)}"`);
    if (f.indentFirstLine !== undefined) {
      ind.push(f.indentFirstLine >= 0 ? `w:firstLine="${ptToTwips(f.indentFirstLine)}"` : `w:hanging="${ptToTwips(-f.indentFirstLine)}"`);
    }
    if (ind.length) p.push(`<w:ind ${ind.join(' ')}/>`);
    if (f.align) p.push(val('jc', JC_TO_DOCX[f.align]));
    if (f.outlineLevel) p.push(val('outlineLvl', f.outlineLevel - 1));
    if (markRevision) {
      // Il segno di paragrafo revisionato: w:ins/w:del vanno all'inizio di w:rPr.
      const rPr = charFormatXml(mark ?? {});
      const rev = `<${this.revisionTag(markRevision)} ${this.revisionAttrs(markRevision)}/>`;
      p.push(rPr ? rPr.replace('<w:rPr>', `<w:rPr>${rev}`) : `<w:rPr>${rev}</w:rPr>`);
    } else if (mark) p.push(charFormatXml(mark));
    p.push(sectPr);
    const inner = p.join('');
    return inner ? `<w:pPr>${inner}</w:pPr>` : '';
  }

  private imageTarget(src: string, rels: Rels): { attr: string } | null {
    const data = this.doc.media[src] ?? src;
    const parsed = parseDataUri(data);
    if (parsed) {
      let name = this.mediaNames.get(data);
      if (!name) {
        const ext = extensionFromMime(parsed.mimeType);
        name = `image${this.mediaNames.size + 1}.${ext}`;
        this.mediaNames.set(data, name);
        this.addFile(`word/media/${name}`, parsed.bytes);
        if (!this.defaults.has(ext)) this.defaults.set(ext, parsed.mimeType);
      }
      return { attr: `r:embed="${rels.add(REL.image, `media/${name}`)}"` };
    }
    if (/^https?:/i.test(data)) return { attr: `r:link="${rels.add(REL.image, data, true)}"` };
    return null;
  }

  private fieldXml(kind: FieldKind, format: CharFormat): string {
    const instr = { page: 'PAGE', numPages: 'NUMPAGES', date: 'DATE \\@ "dd/MM/yyyy"', time: 'TIME \\@ "HH:mm"' }[kind];
    const placeholder = kind === 'page' || kind === 'numPages' ? '1' : '';
    return `<w:fldSimple w:instr=" ${escapeAttr(instr)} "><w:r>${charFormatXml(format)}<w:t>${placeholder}</w:t></w:r></w:fldSimple>`;
  }

  inlineXml(inline: Inline, rels: Rels): string {
    switch (inline.type) {
      case 'text': {
        const { link: _link, indexEntry: _entry, crossRef: _ref, ...format } = inline.format;
        const t = inline.revision?.kind === 'delete' ? 'w:delText' : 'w:t';
        return `<w:r>${charFormatXml(format, inline.styleId)}<${t} xml:space="preserve">${escapeText(inline.text)}</${t}></w:r>`;
      }
      case 'tab':
        return '<w:r><w:tab/></w:r>';
      case 'break':
        return inline.kind === 'line' ? '<w:r><w:br/></w:r>' : `<w:r><w:br w:type="${inline.kind}"/></w:r>`;
      case 'field':
        return this.fieldXml(inline.kind, inline.format);
      case 'footnoteRef': {
        const rPr = charFormatXml({ ...inline.format, verticalAlign: 'superscript' });
        if (inline.noteId === this.currentNote) return `<w:r>${rPr}<w:footnoteRef/></w:r>`;
        const id = this.footnoteIds.get(inline.noteId);
        return id === undefined ? '' : `<w:r>${rPr}<w:footnoteReference w:id="${id}"/></w:r>`;
      }
      case 'image': {
        const target = this.imageTarget(inline.src, rels);
        if (!target) return '';
        const id = ++this.drawingId;
        const cx = ptToEmu(inline.width);
        const cy = ptToEmu(inline.height);
        const descr = inline.alt ? ` descr="${escapeAttr(inline.alt)}"` : '';
        return (
          `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/>` +
          `<wp:docPr id="${id}" name="Immagine ${id}"${descr}/>` +
          `<a:graphic xmlns:a="${NS.a}"><a:graphicData uri="${NS.pic}"><pic:pic xmlns:pic="${NS.pic}">` +
          `<pic:nvPicPr><pic:cNvPr id="${id}" name="Immagine ${id}"/><pic:cNvPicPr/></pic:nvPicPr>` +
          `<pic:blipFill><a:blip ${target.attr}/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
          `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
          `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`
        );
      }
    }
  }

  /** Un contenuto in linea, dentro `w:ins`/`w:del` se ha una revisione (i campi no: lo schema non lo consente). */
  private revisedInlineXml(inline: Inline, rels: Rels): string {
    const xml = this.inlineXml(inline, rels);
    const r = inline.revision;
    if (!r || !xml || inline.type === 'field') return xml;
    const tag = this.revisionTag(r);
    return `<${tag} ${this.revisionAttrs(r)}>${xml}</${tag}>`;
  }

  paragraphXml(p: Paragraph, rels: Rels, sectPr = '', wrap: FieldWrap = {}): string {
    let runs = '';
    let i = 0;
    const children = p.children;
    const starts = this.comments?.starts.get(p.id);
    const ends = this.comments?.ends.get(p.id);
    // Run con gli eventuali inizi/fine di commento attorno.
    const runXml = (index: number) =>
      commentRangeStartXml(starts?.get(index)) + this.revisedInlineXml(children[index]!, rels) + indexEntryXml(children, index) + commentRangeEndXml(ends?.get(index));
    while (i < children.length) {
      const c = children[i]!;
      const link = c.type === 'text' ? c.format.link : undefined;
      if (!link) {
        runs += runXml(i);
        i++;
        continue;
      }
      let inner = '';
      while (i < children.length) {
        const n = children[i]!;
        if (n.type !== 'text' || n.format.link !== link) break;
        inner += runXml(i);
        i++;
      }
      runs += link.startsWith('#')
        ? `<w:hyperlink w:anchor="${escapeAttr(link.slice(1))}">${inner}</w:hyperlink>`
        : `<w:hyperlink r:id="${rels.add(REL.hyperlink, link, true)}">${inner}</w:hyperlink>`;
    }
    const bookmark = paragraphBookmark(p, this.anchors);
    if (bookmark) {
      const id = this.bookmarkId++;
      runs = `<w:bookmarkStart w:id="${id}" w:name="${escapeAttr(bookmark)}"/>${runs}<w:bookmarkEnd w:id="${id}"/>`;
    }
    runs = (wrap.before ?? '') + this.drawingsXml(p, rels) + runs + (wrap.after ?? '');
    return `<w:p>${this.paraFormatXml(p.format, p.styleId, p.markFormat, sectPr, p.revision)}${runs}</w:p>`;
  }

  tableXml(t: Table, rels: Rels): string {
    const pr: string[] = [];
    if (t.styleId) pr.push(val('tblStyle', t.styleId));
    pr.push('<w:tblW w:w="0" w:type="auto"/>');
    if (t.format.align) pr.push(val('jc', t.format.align));
    if (t.format.indentLeft) pr.push(`<w:tblInd w:w="${ptToTwips(t.format.indentLeft)}" w:type="dxa"/>`);
    pr.push(bordersXml('tblBorders', t.format.borders, ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']));
    pr.push('<w:tblLayout w:type="fixed"/>');
    pr.push('<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>');
    const grid = t.columnWidths.map((w) => `<w:gridCol w:w="${ptToTwips(w)}"/>`).join('');
    const rows = t.rows
      .map((row) => {
        const trPr: string[] = [];
        if (row.height && row.height.rule !== 'auto') trPr.push(`<w:trHeight w:val="${ptToTwips(row.height.value)}" w:hRule="${row.height.rule}"/>`);
        if (row.isHeader) trPr.push('<w:tblHeader/>');
        let col = 0;
        const cells = row.cells
          .map((cell) => {
            const span = cell.colSpan ?? 1;
            const width = t.columnWidths.slice(col, col + span).reduce((n, w) => n + w, 0);
            col += span;
            const tc: string[] = [`<w:tcW w:w="${ptToTwips(width)}" w:type="dxa"/>`];
            if (span > 1) tc.push(val('gridSpan', span));
            if (cell.vMerge === 'restart') tc.push(val('vMerge', 'restart'));
            else if (cell.vMerge === 'continue') tc.push('<w:vMerge/>');
            tc.push(bordersXml('tcBorders', cell.format.borders, ['top', 'left', 'bottom', 'right']));
            if (cell.format.shading) tc.push(`<w:shd w:val="clear" w:color="auto" w:fill="${hex(cell.format.shading)}"/>`);
            const pad = cell.format.padding;
            if (pad) {
              tc.push(
                `<w:tcMar><w:top w:w="${ptToTwips(pad.top)}" w:type="dxa"/><w:left w:w="${ptToTwips(pad.left)}" w:type="dxa"/>` +
                  `<w:bottom w:w="${ptToTwips(pad.bottom)}" w:type="dxa"/><w:right w:w="${ptToTwips(pad.right)}" w:type="dxa"/></w:tcMar>`,
              );
            }
            if (cell.format.verticalAlign) tc.push(val('vAlign', cell.format.verticalAlign));
            // Una cella deve terminare con un paragrafo.
            const blocks = cell.blocks[cell.blocks.length - 1]?.type === 'paragraph' ? cell.blocks : [...cell.blocks, emptyParagraph];
            return `<w:tc><w:tcPr>${tc.join('')}</w:tcPr>${this.blocksXml(blocks, rels)}</w:tc>`;
          })
          .join('');
        return `<w:tr>${trPr.length ? `<w:trPr>${trPr.join('')}</w:trPr>` : ''}${cells}</w:tr>`;
      })
      .join('');
    return `<w:tbl><w:tblPr>${pr.join('')}</w:tblPr><w:tblGrid>${grid}</w:tblGrid>${rows}</w:tbl>`;
  }

  blocksXml(blocks: readonly Block[], rels: Rels): string {
    const wraps = tocFieldWraps(blocks, this.doc.styles);
    return blocks.map((b) => (b.type === 'paragraph' ? this.paragraphXml(b, rels, '', wraps.get(b)) : this.tableXml(b, rels))).join('');
  }

  // -------------------------------------------------------------------------
  // Sezioni, intestazioni e piè di pagina

  private headerFooterRefs(kind: 'header' | 'footer', hf: HeaderFooter | undefined, rels: Rels): string {
    if (!hf) return '';
    const out: string[] = [];
    for (const type of ['default', 'first', 'even'] as const) {
      const blocks = hf[type];
      if (!blocks) continue;
      if (type === 'even') this.usesEvenHeaders = true;
      const n = kind === 'header' ? ++this.headerCount : ++this.footerCount;
      const name = `${kind}${n}.xml`;
      const partRels = new Rels();
      const root = kind === 'header' ? 'w:hdr' : 'w:ftr';
      const body = this.blocksXml(blocks.length && blocks[blocks.length - 1]!.type === 'paragraph' ? blocks : [...blocks, emptyParagraph], partRels);
      this.addFile(`word/${name}`, `${XML_DECL}<${root} ${namespaces()}>${body}</${root}>`, kind === 'header' ? CONTENT_TYPES.header : CONTENT_TYPES.footer);
      if (partRels.size) this.addFile(`word/_rels/${name}.rels`, partRels.xml());
      const id = rels.add(kind === 'header' ? REL.header : REL.footer, name);
      out.push(`<w:${kind}Reference w:type="${type}" r:id="${id}"/>`);
    }
    return out.join('');
  }

  sectPrXml(s: Section, rels: Rels): string {
    const p = s.page;
    const m = p.margins;
    const parts = [
      this.headerFooterRefs('header', s.header, rels),
      this.headerFooterRefs('footer', s.footer, rels),
      s.breakType ? val('type', s.breakType) : '',
      `<w:pgSz w:w="${ptToTwips(p.width)}" w:h="${ptToTwips(p.height)}"${p.orientation === 'landscape' ? ' w:orient="landscape"' : ''}/>`,
      `<w:pgMar w:top="${ptToTwips(m.top)}" w:right="${ptToTwips(m.right)}" w:bottom="${ptToTwips(m.bottom)}" w:left="${ptToTwips(m.left)}" w:header="${ptToTwips(m.header)}" w:footer="${ptToTwips(m.footer)}" w:gutter="0"/>`,
      s.pageNumberStart !== undefined ? `<w:pgNumType w:start="${s.pageNumberStart}"/>` : '',
      p.columns && p.columns.count > 1 ? `<w:cols w:num="${p.columns.count}" w:space="${ptToTwips(p.columns.gap)}"/>` : '<w:cols w:space="708"/>',
      s.titlePage ? '<w:titlePg/>' : '',
    ];
    return `<w:sectPr>${parts.join('')}</w:sectPr>`;
  }

  documentXml(rels: Rels): string {
    const body: string[] = [];
    const sections = this.doc.sections;
    sections.forEach((s, i) => {
      const last = i === sections.length - 1;
      const sectPr = this.sectPrXml(s, rels);
      if (last) {
        body.push(this.blocksXml(s.blocks, rels), sectPr);
        return;
      }
      // Le sezioni intermedie chiudono con il sectPr nel pPr dell'ultimo paragrafo.
      const blocks = s.blocks;
      const lastBlock = blocks[blocks.length - 1];
      if (lastBlock?.type === 'paragraph') {
        body.push(this.blocksXml(blocks.slice(0, -1), rels), this.paragraphXml(lastBlock, rels, sectPr));
      } else {
        body.push(this.blocksXml(blocks, rels), this.paragraphXml(emptyParagraph, rels, sectPr));
      }
    });
    return `${XML_DECL}<w:document ${namespaces()}><w:body>${body.join('')}</w:body></w:document>`;
  }

  // -------------------------------------------------------------------------
  // Note a piè di pagina

  footnotesXml(rels: Rels): string {
    const separator = (type: string, id: number, mark: string) =>
      `<w:footnote w:type="${type}" w:id="${id}"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:${mark}/></w:r></w:p></w:footnote>`;
    const notes = [...this.footnoteIds].map(([noteId, id]) => {
      const blocks = this.doc.footnotes?.[noteId]?.blocks ?? [];
      this.currentNote = noteId;
      const body = this.blocksXml(blocks.length && blocks[blocks.length - 1]!.type === 'paragraph' ? blocks : [...blocks, emptyParagraph], rels);
      this.currentNote = undefined;
      return `<w:footnote w:id="${id}">${body}</w:footnote>`;
    });
    return (
      `${XML_DECL}<w:footnotes ${namespaces()}>` +
      separator('separator', -1, 'separator') +
      separator('continuationSeparator', 0, 'continuationSeparator') +
      `${notes.join('')}</w:footnotes>`
    );
  }

  // -------------------------------------------------------------------------
  // Stili, numerazione, proprietà

  styleXml(style: Style, isDefault: boolean): string {
    const name = BUILTIN_STYLE_NAMES[style.id] ?? style.name;
    const parts = [
      val('name', name),
      style.basedOn ? val('basedOn', style.basedOn) : '',
      style.next ? val('next', style.next) : '',
      style.quickStyle ? '<w:qFormat/>' : '',
      style.paraFormat && style.type !== 'character' ? this.paraFormatXml(style.paraFormat, undefined) : '',
      style.charFormat ? charFormatXml(style.charFormat) : '',
    ];
    return `<w:style w:type="${style.type}"${isDefault ? ' w:default="1"' : ''} w:styleId="${escapeAttr(style.id)}">${parts.join('')}</w:style>`;
  }

  stylesXml(): string {
    const sheet = this.doc.styles;
    const rPr = charFormatXml(sheet.defaults.char);
    const pPr = this.paraFormatXml(sheet.defaults.para, undefined);
    const styles = Object.values(sheet.styles).map((s) => this.styleXml(s, s.id === sheet.defaultParagraphStyle));
    return (
      `${XML_DECL}<w:styles xmlns:w="${NS.w}" xmlns:r="${NS.r}">` +
      `<w:docDefaults><w:rPrDefault>${rPr || '<w:rPr/>'}</w:rPrDefault><w:pPrDefault>${pPr || '<w:pPr/>'}</w:pPrDefault></w:docDefaults>` +
      `${styles.join('')}</w:styles>`
    );
  }

  numberingXml(): string {
    const abstracts: string[] = [];
    const nums: string[] = [];
    for (const [listId, def] of Object.entries(this.doc.lists)) {
      const numId = this.numIds.get(listId)!;
      const lvls = def.levels
        .map((l, i) => {
          const ind = `<w:pPr><w:ind w:left="${ptToTwips(l.indentLeft)}" w:hanging="${ptToTwips(l.hanging)}"/></w:pPr>`;
          const rPr = charFormatXml(l.charFormat ?? {});
          return (
            `<w:lvl w:ilvl="${i}">${val('start', l.start)}${val('numFmt', NUMFMT_TO_DOCX[l.format])}` +
            `${val('lvlText', l.text)}${val('lvlJc', 'left')}${ind}${rPr}</w:lvl>`
          );
        })
        .join('');
      abstracts.push(`<w:abstractNum w:abstractNumId="${numId}"><w:multiLevelType w:val="hybridMultilevel"/>${lvls}</w:abstractNum>`);
      nums.push(`<w:num w:numId="${numId}"><w:abstractNumId w:val="${numId}"/></w:num>`);
    }
    return `${XML_DECL}<w:numbering xmlns:w="${NS.w}">${abstracts.join('')}${nums.join('')}</w:numbering>`;
  }

  coreXml(): string {
    const m = this.doc.metadata;
    const el = (tag: string, v: string | undefined, extra = '') => (v ? `<${tag}${extra}>${escapeText(v)}</${tag}>` : '');
    const now = new Date().toISOString().replace(/\.\d{3}Z$/, 'Z');
    const w3c = (d: string | undefined) => (d ? new Date(d).toISOString().replace(/\.\d{3}Z$/, 'Z') : now);
    return (
      `${XML_DECL}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
      `xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ` +
      `xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
      el('dc:title', m.title) +
      el('dc:subject', m.subject) +
      el('dc:creator', m.author) +
      el('cp:keywords', m.keywords?.join(', ')) +
      el('dcterms:created', w3c(m.created), ' xsi:type="dcterms:W3CDTF"') +
      el('dcterms:modified', w3c(m.modified), ' xsi:type="dcterms:W3CDTF"') +
      `</cp:coreProperties>`
    );
  }

  settingsXml(): string {
    return (
      `${XML_DECL}<w:settings xmlns:w="${NS.w}">` +
      (this.usesEvenHeaders ? '<w:evenAndOddHeaders/>' : '') +
      `<w:defaultTabStop w:val="708"/>` +
      (this.footnoteIds.size ? '<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>' : '') +
      `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>` +
      `</w:settings>`
    );
  }

  async write(): Promise<Uint8Array> {
    const rels = new Rels();
    rels.add(REL.styles, 'styles.xml');
    rels.add(REL.settings, 'settings.xml');
    const hasLists = Object.keys(this.doc.lists).length > 0;
    if (hasLists) rels.add(REL.numbering, 'numbering.xml');
    if (this.comments) {
      rels.add(COMMENTS_REL, 'comments.xml');
      rels.add(COMMENTS_EXT_REL, 'commentsExtended.xml');
    }

    // document.xml per primo: genera media, intestazioni e piè di pagina.
    const document = this.documentXml(rels);
    this.addFile('word/document.xml', document, CONTENT_TYPES.document);
    if (this.footnoteIds.size) {
      const noteRels = new Rels();
      this.addFile('word/footnotes.xml', this.footnotesXml(noteRels), CONTENT_TYPES.footnotes);
      if (noteRels.size) this.addFile('word/_rels/footnotes.xml.rels', noteRels.xml());
      rels.add(REL.footnotes, 'footnotes.xml');
    }
    this.addFile('word/_rels/document.xml.rels', rels.xml());
    this.addFile('word/styles.xml', this.stylesXml(), CONTENT_TYPES.styles);
    this.addFile('word/settings.xml', this.settingsXml(), CONTENT_TYPES.settings);
    if (hasLists) this.addFile('word/numbering.xml', this.numberingXml(), CONTENT_TYPES.numbering);
    if (this.comments) {
      this.addFile('word/comments.xml', commentsXml(this.comments.comments), COMMENTS_CT);
      this.addFile('word/commentsExtended.xml', commentsExtendedXml(this.comments.comments), COMMENTS_EXT_CT);
    }
    this.addFile('docProps/core.xml', this.coreXml(), CONTENT_TYPES.core);

    const pkgRels = new Rels();
    pkgRels.add(REL.officeDocument, 'word/document.xml');
    pkgRels.add(REL.coreProps, 'docProps/core.xml');
    this.addFile('_rels/.rels', pkgRels.xml());

    const defaults = [...this.defaults].map(([ext, ct]) => `<Default Extension="${ext}" ContentType="${ct}"/>`).join('');
    const overrides = [...this.contentTypes].map(([part, ct]) => `<Override PartName="${part}" ContentType="${ct}"/>`).join('');
    const types = `${XML_DECL}<Types xmlns="${NS.ct}">${defaults}${overrides}</Types>`;
    return writeZip([['[Content_Types].xml', this.encoder.encode(types)], ...this.files]);
  }
}

const emptyParagraph: Paragraph = { type: 'paragraph', id: '', format: {}, children: [] };

/** Campo XE di Word dopo l'ultimo run di un testo segnato come voce d'indice (vedi src/references). */
function indexEntryXml(children: readonly Inline[], index: number): string {
  const c = children[index]!;
  const entry = c.type === 'text' ? c.format.indexEntry : undefined;
  const next = children[index + 1];
  if (!entry || (next?.type === 'text' && next.format.indexEntry === entry)) return '';
  const instr = ` XE "${entry.replace(/"/g, '""')}" `;
  return `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">${escapeText(instr)}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
}

function namespaces(): string {
  return `xmlns:w="${NS.w}" xmlns:r="${NS.r}" xmlns:wp="${NS.wp}" xmlns:a="${NS.a}" xmlns:pic="${NS.pic}" xmlns:wps="${WPS_NS}"`;
}

export function exportDocx(doc: DocumentModel): Promise<Uint8Array> {
  return new DocxWriter(doc).write();
}
