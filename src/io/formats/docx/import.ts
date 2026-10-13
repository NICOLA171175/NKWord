/**
 * Lettura DOCX: documento, stili (con tema per i caratteri predefiniti),
 * numerazione, sezioni con intestazioni e piè di pagina, tabelle con celle
 * unite, collegamenti, immagini e campi (semplici e complessi). Le
 * revisioni vengono lette come accettate.
 */

import {
  defaultPageSetup,
  defaultStyleSheet,
  halfPointsToPt,
  emuToPt,
  newId,
  normalizeInlines,
  STYLE_NORMAL,
  twipsToPt,
  type Block,
  type BorderLine,
  type Borders,
  type CharFormat,
  type DocumentMetadata,
  type DocumentModel,
  type FieldKind,
  type HeaderFooter,
  type Id,
  type Inline,
  type ListDefinition,
  type ListLevel,
  type PageSetup,
  type ParaFormat,
  type Paragraph,
  type Section,
  type Style,
  type StyleSheet,
  type TabStop,
  type Table,
  type TableCell,
  type TableCellFormat,
  type TableRow,
} from '../../../core';
import { DOCX_HIGHLIGHTS, normalizeColor } from '../../util/color';
import { mimeFromExtension, toDataUri } from '../../util/encoding';
import { child, elements, isElement, parseXml, textContent, type MarkupElement, type MarkupNode } from '../../util/markup';
import { readZip, type ZipEntries } from '../../util/zip';
import { defaultListLevel } from '../lists';
import { JC_FROM_DOCX, NS, numFmtFromDocx, REL, TAB_FROM_DOCX, underlineFromDocx } from './common';

// ---------------------------------------------------------------------------
// XML con prefissi canonici

const CANONICAL_PREFIX: Record<string, string> = {
  [NS.w]: 'w',
  'http://purl.oclc.org/ooxml/wordprocessingml/main': 'w',
  [NS.r]: 'r',
  'http://purl.oclc.org/ooxml/officeDocument/relationships': 'r',
  [NS.wp]: 'wp',
  [NS.a]: 'a',
  [NS.pic]: 'pic',
  'urn:schemas-microsoft-com:vml': 'v',
  'http://schemas.openxmlformats.org/markup-compatibility/2006': 'mc',
  [NS.rels]: 'rel',
  'http://purl.org/dc/elements/1.1/': 'dc',
  'http://purl.org/dc/terms/': 'dcterms',
  'http://schemas.openxmlformats.org/package/2006/metadata/core-properties': 'cp',
};

/** Rinomina i prefissi secondo lo spazio dei nomi, così il resto del codice può usare 'w:p', 'r:id'... */
function canonicalize(el: MarkupElement, scope: ReadonlyMap<string, string> = new Map()): MarkupElement {
  let map = scope;
  for (const [k, v] of Object.entries(el.attrs)) {
    if (k === 'xmlns' || k.startsWith('xmlns:')) {
      if (map === scope) map = new Map(scope);
      (map as Map<string, string>).set(k === 'xmlns' ? '' : k.slice(6), CANONICAL_PREFIX[v] ?? (k === 'xmlns' ? '' : k.slice(6)));
    }
  }
  const rename = (name: string, isAttr: boolean) => {
    const i = name.indexOf(':');
    if (i < 0) {
      if (isAttr) return name;
      const def = map.get('');
      return def ? `${def}:${name}` : name;
    }
    const prefix = name.slice(0, i);
    if (prefix === 'xml' || prefix === 'xmlns') return name;
    const canon = map.get(prefix);
    return canon !== undefined && canon !== '' ? `${canon}:${name.slice(i + 1)}` : name;
  };
  const attrs: Record<string, string> = {};
  for (const [k, v] of Object.entries(el.attrs)) attrs[rename(k, true)] = v;
  const children: MarkupNode[] = el.children.map((c) => (isElement(c) ? canonicalize(c, map) : c));
  return { name: el.name === '#document' ? el.name : rename(el.name, false), attrs, children };
}

function rootElement(doc: MarkupElement): MarkupElement | undefined {
  return elements(doc)[0];
}

const attr = (el: MarkupElement | undefined, name: string): string | undefined => el?.attrs[name];
const wVal = (el: MarkupElement | undefined): string | undefined => attr(el, 'w:val');

function num(v: string | undefined): number | undefined {
  if (v === undefined) return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
}

function twips(v: string | undefined): number | undefined {
  const n = num(v);
  return n === undefined ? undefined : twipsToPt(n);
}

/** Proprietà on/off di WordprocessingML (`<w:b/>`, `<w:b w:val="0"/>`). */
function onOff(el: MarkupElement | undefined): boolean | undefined {
  if (!el) return undefined;
  const v = wVal(el);
  return !(v === '0' || v === 'false' || v === 'off');
}

type Mutable<T> = { -readonly [K in keyof T]: T[K] };

// ---------------------------------------------------------------------------
// Pacchetto e relazioni

interface Relationship {
  readonly type: string;
  readonly target: string;
  readonly external: boolean;
}

interface Part {
  readonly path: string;
  readonly rels: ReadonlyMap<string, Relationship>;
}

function resolvePath(base: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/').slice(0, -1);
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg && seg !== '.') parts.push(seg);
  }
  return parts.join('/');
}

class Package {
  private readonly decoder = new TextDecoder();
  private readonly xmlCache = new Map<string, MarkupElement | undefined>();

  constructor(readonly zip: ZipEntries) {}

  has(path: string): boolean {
    return this.zip.has(path);
  }

  xml(path: string): MarkupElement | undefined {
    if (this.xmlCache.has(path)) return this.xmlCache.get(path);
    const bytes = this.zip.get(path);
    const root = bytes ? rootElement(canonicalize(parseXml(this.decoder.decode(bytes)))) : undefined;
    this.xmlCache.set(path, root);
    return root;
  }

  part(path: string): Part {
    const dir = path.split('/').slice(0, -1).join('/');
    const name = path.split('/').pop()!;
    const relsPath = `${dir ? `${dir}/` : ''}_rels/${name}.rels`;
    const rels = new Map<string, Relationship>();
    const root = this.xml(relsPath);
    if (root) {
      for (const r of elements(root)) {
        if (!r.name.endsWith('Relationship')) continue;
        const id = r.attrs.Id;
        const target = r.attrs.Target;
        if (!id || target === undefined) continue;
        const external = r.attrs.TargetMode === 'External';
        rels.set(id, { type: r.attrs.Type ?? '', target: external ? target : resolvePath(path, target), external });
      }
    }
    return { path, rels };
  }

  findRel(part: Part, type: string): Relationship | undefined {
    for (const r of part.rels.values()) if (r.type === type || r.type.endsWith(`/${type.split('/').pop()}`)) return r;
    return undefined;
  }
}

// ---------------------------------------------------------------------------
// Formattazione

interface Theme {
  readonly minor?: string;
  readonly major?: string;
}

function readTheme(root: MarkupElement | undefined): Theme {
  if (!root) return {};
  const scheme = findDeep(root, 'a:fontScheme');
  const latin = (name: string) => attr(child(child(scheme, name), 'a:latin'), 'typeface') || undefined;
  return { minor: latin('a:minorFont'), major: latin('a:majorFont') };
}

function findDeep(el: MarkupElement, name: string): MarkupElement | undefined {
  for (const c of el.children) {
    if (!isElement(c)) continue;
    if (c.name === name) return c;
    const d = findDeep(c, name);
    if (d) return d;
  }
  return undefined;
}

function readBorder(el: MarkupElement | undefined): BorderLine | undefined {
  if (!el) return undefined;
  const v = wVal(el) ?? 'single';
  if (v === 'nil' || v === 'none') return { style: 'none', width: 0, color: '#000000' };
  const style: BorderLine['style'] = v === 'double' ? 'double' : v === 'dotted' ? 'dotted' : /dash/i.test(v) ? 'dashed' : 'single';
  const sz = num(attr(el, 'w:sz'));
  return { style, width: sz !== undefined ? sz / 8 : 0.5, color: normalizeColor(attr(el, 'w:color')) ?? '#000000' };
}

function readBorders(el: MarkupElement | undefined): (Borders & { insideH?: BorderLine; insideV?: BorderLine }) | undefined {
  if (!el) return undefined;
  const out: Record<string, BorderLine> = {};
  const map: Record<string, string> = { top: 'top', left: 'left', start: 'left', bottom: 'bottom', right: 'right', end: 'right', insideH: 'insideH', insideV: 'insideV' };
  for (const c of elements(el)) {
    const side = map[c.name.slice(2)];
    const b = readBorder(c);
    if (side && b) out[side] = b;
  }
  return Object.keys(out).length ? out : undefined;
}

class Formatter {
  constructor(private readonly theme: Theme) {}

  font(rFonts: MarkupElement | undefined): string | undefined {
    if (!rFonts) return undefined;
    const direct = attr(rFonts, 'w:ascii') ?? attr(rFonts, 'w:hAnsi');
    if (direct) return direct;
    const themed = attr(rFonts, 'w:asciiTheme') ?? attr(rFonts, 'w:hAnsiTheme');
    if (themed) return themed.startsWith('major') ? this.theme.major : this.theme.minor;
    return attr(rFonts, 'w:eastAsia') ?? attr(rFonts, 'w:cs');
  }

  char(rPr: MarkupElement | undefined): CharFormat {
    if (!rPr) return {};
    const f: Mutable<CharFormat> = {};
    const font = this.font(child(rPr, 'w:rFonts'));
    if (font) f.fontFamily = font;
    const set = <K extends keyof CharFormat>(key: K, v: CharFormat[K] | undefined) => {
      if (v !== undefined) f[key] = v;
    };
    set('bold', onOff(child(rPr, 'w:b')));
    set('italic', onOff(child(rPr, 'w:i')));
    set('allCaps', onOff(child(rPr, 'w:caps')));
    set('smallCaps', onOff(child(rPr, 'w:smallCaps')));
    set('strike', onOff(child(rPr, 'w:strike')) ?? onOff(child(rPr, 'w:dstrike')));
    const color = normalizeColor(wVal(child(rPr, 'w:color')));
    if (color) f.color = color;
    const sz = num(wVal(child(rPr, 'w:sz')));
    if (sz) f.fontSize = halfPointsToPt(sz);
    const hl = wVal(child(rPr, 'w:highlight'));
    if (hl === 'none') f.highlight = null;
    else if (hl && DOCX_HIGHLIGHTS[hl]) f.highlight = DOCX_HIGHLIGHTS[hl];
    else {
      const shd = child(rPr, 'w:shd');
      const fill = normalizeColor(attr(shd, 'w:fill'));
      if (fill && wVal(shd) !== 'nil') f.highlight = fill;
    }
    const u = child(rPr, 'w:u');
    if (u) f.underline = underlineFromDocx(wVal(u) ?? 'single');
    const va = wVal(child(rPr, 'w:vertAlign'));
    if (va === 'superscript' || va === 'subscript' || va === 'baseline') f.verticalAlign = va;
    return f;
  }

  para(pPr: MarkupElement | undefined, lists: ReadonlyMap<string, Id>): ParaFormat {
    if (!pPr) return {};
    const f: Mutable<ParaFormat> = {};
    const set = <K extends keyof ParaFormat>(key: K, v: ParaFormat[K] | undefined) => {
      if (v !== undefined) f[key] = v;
    };
    set('keepWithNext', onOff(child(pPr, 'w:keepNext')));
    set('keepLinesTogether', onOff(child(pPr, 'w:keepLines')));
    set('pageBreakBefore', onOff(child(pPr, 'w:pageBreakBefore')));
    set('widowControl', onOff(child(pPr, 'w:widowControl')));
    const numPr = child(pPr, 'w:numPr');
    if (numPr) {
      const numId = wVal(child(numPr, 'w:numId'));
      const level = num(wVal(child(numPr, 'w:ilvl'))) ?? 0;
      if (numId === '0') f.list = null;
      else if (numId && lists.has(numId)) f.list = { listId: lists.get(numId)!, level: Math.max(0, Math.min(8, level)) };
    }
    const borders = readBorders(child(pPr, 'w:pBdr'));
    if (borders) {
      const { top, left, bottom, right } = borders;
      f.borders = { ...(top ? { top } : {}), ...(left ? { left } : {}), ...(bottom ? { bottom } : {}), ...(right ? { right } : {}) };
    }
    const shd = normalizeColor(attr(child(pPr, 'w:shd'), 'w:fill'));
    if (shd) f.shading = shd;
    const tabs = child(pPr, 'w:tabs');
    if (tabs) {
      const stops: TabStop[] = [];
      for (const t of elements(tabs, 'w:tab')) {
        const align = TAB_FROM_DOCX[wVal(t) ?? ''];
        const pos = twips(attr(t, 'w:pos'));
        if (!align || pos === undefined) continue;
        const leader = attr(t, 'w:leader');
        stops.push({
          position: pos,
          align,
          ...(leader === 'dot' || leader === 'hyphen' || leader === 'underscore' ? { leader } : {}),
        });
      }
      if (stops.length) f.tabs = stops;
    }
    const spacing = child(pPr, 'w:spacing');
    if (spacing) {
      set('spaceBefore', twips(attr(spacing, 'w:before')));
      set('spaceAfter', twips(attr(spacing, 'w:after')));
      const line = num(attr(spacing, 'w:line'));
      if (line !== undefined) {
        const rule = attr(spacing, 'w:lineRule') ?? 'auto';
        if (rule === 'exact' || rule === 'atLeast') f.lineSpacing = { rule, value: twipsToPt(line) };
        else f.lineSpacing = { rule: 'multiple', value: Math.round((line / 240) * 100) / 100 };
      }
    }
    const ind = child(pPr, 'w:ind');
    if (ind) {
      set('indentLeft', twips(attr(ind, 'w:left') ?? attr(ind, 'w:start')));
      set('indentRight', twips(attr(ind, 'w:right') ?? attr(ind, 'w:end')));
      const first = twips(attr(ind, 'w:firstLine'));
      const hanging = twips(attr(ind, 'w:hanging'));
      if (hanging !== undefined) f.indentFirstLine = -hanging;
      else if (first !== undefined) f.indentFirstLine = first;
    }
    const jc = JC_FROM_DOCX[wVal(child(pPr, 'w:jc')) ?? ''];
    if (jc) f.align = jc;
    const outline = num(wVal(child(pPr, 'w:outlineLvl')));
    if (outline !== undefined && outline < 9) f.outlineLevel = outline + 1;
    return f;
  }
}

// ---------------------------------------------------------------------------
// Stili e numerazione

const BUILTIN_IDS: Record<string, Id> = {
  normal: 'Normal', title: 'Title', subtitle: 'Subtitle', hyperlink: 'Hyperlink',
  'heading 1': 'Heading1', 'heading 2': 'Heading2', 'heading 3': 'Heading3', 'heading 4': 'Heading4',
  'heading 5': 'Heading5', 'heading 6': 'Heading6', 'heading 7': 'Heading7', 'heading 8': 'Heading8', 'heading 9': 'Heading9',
};

function readStyles(root: MarkupElement | undefined, fmt: Formatter, theme: Theme, lists: ReadonlyMap<string, Id>): { sheet: StyleSheet; ids: Map<string, Id> } {
  const ids = new Map<string, Id>();
  if (!root) return { sheet: defaultStyleSheet(), ids };
  const builtins = defaultStyleSheet().styles;

  const raw = elements(root, 'w:style').filter((s) => ['paragraph', 'character', 'table'].includes(attr(s, 'w:type') ?? ''));
  for (const s of raw) {
    const id = attr(s, 'w:styleId');
    if (!id) continue;
    const name = (wVal(child(s, 'w:name')) ?? id).toLowerCase();
    const canon = BUILTIN_IDS[name];
    const isChar = attr(s, 'w:type') === 'character';
    ids.set(id, canon && isChar === (canon === 'Hyperlink') ? canon : id);
  }
  const mapId = (id: string | undefined) => (id ? ids.get(id) ?? id : undefined);

  const styles: Record<Id, Style> = {};
  let defaultParagraphStyle: Id = STYLE_NORMAL;
  for (const s of raw) {
    const rawId = attr(s, 'w:styleId');
    if (!rawId) continue;
    const id = ids.get(rawId)!;
    const type = attr(s, 'w:type') as Style['type'];
    const name = wVal(child(s, 'w:name')) ?? rawId;
    const basedOn = mapId(wVal(child(s, 'w:basedOn')));
    const next = mapId(wVal(child(s, 'w:next')));
    const paraFormat = fmt.para(child(s, 'w:pPr'), lists);
    const charFormat = fmt.char(child(s, 'w:rPr'));
    styles[id] = {
      id,
      name: builtins[id]?.name ?? name,
      type,
      ...(basedOn && basedOn !== id ? { basedOn } : {}),
      ...(next ? { next } : {}),
      ...(Object.keys(paraFormat).length ? { paraFormat } : {}),
      ...(Object.keys(charFormat).length ? { charFormat } : {}),
      ...(child(s, 'w:qFormat') ? { quickStyle: true } : {}),
    };
    const isDefault = attr(s, 'w:default');
    if (type === 'paragraph' && (isDefault === '1' || isDefault === 'true' || isDefault === 'on')) defaultParagraphStyle = id;
  }
  if (!styles[defaultParagraphStyle]) styles[defaultParagraphStyle] = { id: defaultParagraphStyle, name: 'Normale', type: 'paragraph', quickStyle: true };

  const defaults = child(root, 'w:docDefaults');
  const char: Mutable<CharFormat> = { ...fmt.char(child(child(defaults, 'w:rPrDefault'), 'w:rPr')) };
  char.fontFamily ??= theme.minor ?? 'Calibri';
  char.fontSize ??= 10;
  char.color ??= '#000000';
  const para = fmt.para(child(child(defaults, 'w:pPrDefault'), 'w:pPr'), lists);
  return { sheet: { styles, defaults: { char, para }, defaultParagraphStyle }, ids };
}

/** I punti elenco di Word usano spesso caratteri Symbol/Wingdings nell'area privata. */
function bulletText(text: string, font: string | undefined): string {
  if (!text) return '•';
  const code = text.codePointAt(0)!;
  if (code >= 0xf000 && code <= 0xf0ff) return code === 0xf0a7 || code === 0xf06e ? '▪' : '•';
  if (text === 'o' && font && /courier/i.test(font)) return '◦';
  if (font && /wingdings|symbol/i.test(font)) return text === '§' ? '▪' : '•';
  return text;
}

function readNumbering(root: MarkupElement | undefined, fmt: Formatter): { lists: Record<Id, ListDefinition>; ids: Map<string, Id> } {
  const lists: Record<Id, ListDefinition> = {};
  const ids = new Map<string, Id>();
  if (!root) return { lists, ids };
  const abstracts = new Map<string, MarkupElement>();
  for (const a of elements(root, 'w:abstractNum')) {
    const id = attr(a, 'w:abstractNumId');
    if (id !== undefined) abstracts.set(id, a);
  }
  for (const n of elements(root, 'w:num')) {
    const numId = attr(n, 'w:numId');
    const abs = abstracts.get(wVal(child(n, 'w:abstractNumId')) ?? '');
    if (numId === undefined || !abs) continue;
    const listId = `num${numId}`;
    const levels: ListLevel[] = Array.from({ length: 9 }, (_, l) => defaultListLevel('decimal', l));
    for (const lvl of elements(abs, 'w:lvl')) {
      const i = num(attr(lvl, 'w:ilvl')) ?? 0;
      if (i < 0 || i > 8) continue;
      const format = numFmtFromDocx(wVal(child(lvl, 'w:numFmt')));
      const rPr = child(lvl, 'w:rPr');
      const charFormat = fmt.char(rPr);
      const ind = child(child(lvl, 'w:pPr'), 'w:ind');
      const rawText = wVal(child(lvl, 'w:lvlText')) ?? '';
      const { fontFamily, ...restChar } = charFormat;
      const symbolFont = fontFamily && /symbol|wingdings/i.test(fontFamily);
      levels[i] = {
        format,
        text: format === 'bullet' ? bulletText(rawText, fontFamily) : rawText,
        start: num(wVal(child(lvl, 'w:start'))) ?? 1,
        indentLeft: twips(attr(ind, 'w:left') ?? attr(ind, 'w:start')) ?? 36 * (i + 1),
        hanging: twips(attr(ind, 'w:hanging')) ?? 18,
        ...(Object.keys(symbolFont ? restChar : charFormat).length ? { charFormat: symbolFont ? restChar : charFormat } : {}),
      };
    }
    for (const o of elements(n, 'w:lvlOverride')) {
      const i = num(attr(o, 'w:ilvl'));
      const start = num(wVal(child(o, 'w:startOverride')));
      if (i !== undefined && start !== undefined && levels[i]) levels[i] = { ...levels[i]!, start };
    }
    lists[listId] = { id: listId, levels };
    ids.set(numId, listId);
  }
  return { lists, ids };
}

function readCore(root: MarkupElement | undefined): DocumentMetadata {
  if (!root) return {};
  const text = (name: string) => {
    const el = child(root, name);
    const t = el ? textContent(el).trim() : '';
    return t || undefined;
  };
  const keywords = text('cp:keywords')?.split(/[,;]/).map((k) => k.trim()).filter(Boolean);
  const m: Mutable<DocumentMetadata> = {};
  const title = text('dc:title');
  if (title) m.title = title;
  const author = text('dc:creator');
  if (author) m.author = author;
  const subject = text('dc:subject');
  if (subject) m.subject = subject;
  if (keywords?.length) m.keywords = keywords;
  const created = text('dcterms:created');
  if (created) m.created = created;
  const modified = text('dcterms:modified');
  if (modified) m.modified = modified;
  return m;
}

// ---------------------------------------------------------------------------
// Contenuto

const FIELD_INSTR: Record<string, FieldKind> = { PAGE: 'page', NUMPAGES: 'numPages', SECTIONPAGES: 'numPages', DATE: 'date', TIME: 'time', CREATEDATE: 'date', SAVEDATE: 'date', PRINTDATE: 'date' };

function parseInstr(instr: string): { kind?: FieldKind; link?: string } {
  const trimmed = instr.trim();
  const word = trimmed.split(/\s+/)[0]?.toUpperCase() ?? '';
  if (FIELD_INSTR[word]) return { kind: FIELD_INSTR[word] };
  if (word === 'HYPERLINK') {
    const m = /HYPERLINK\s+(?:\\l\s+)?"([^"]*)"/i.exec(trimmed);
    const anchor = /\\l\s+"([^"]*)"/i.exec(trimmed);
    if (anchor && !m?.[1]) return { link: `#${anchor[1]}` };
    if (m?.[1]) return { link: /\\l/.test(trimmed.slice(0, m.index + 12)) ? `#${m[1]}` : m[1] };
  }
  return {};
}

interface ComplexField {
  instr: string;
  phase: 'instr' | 'result';
  kind?: FieldKind;
  link?: string;
  emitted: boolean;
  format: CharFormat;
}

interface SectionGroup {
  readonly blocks: Block[];
  readonly sectPr?: MarkupElement;
}

class Reader {
  private readonly fields: ComplexField[] = [];
  readonly media: Record<string, string> = {};

  constructor(
    private readonly pkg: Package,
    private readonly fmt: Formatter,
    private readonly styleIds: ReadonlyMap<string, Id>,
    private readonly listIds: ReadonlyMap<string, Id>,
  ) {}

  private styleId(raw: string | undefined): Id | undefined {
    return raw ? this.styleIds.get(raw) ?? raw : undefined;
  }

  // --- run --------------------------------------------------------------

  private mediaKey(part: Part, relId: string | undefined): string | undefined {
    if (!relId) return undefined;
    const rel = part.rels.get(relId);
    if (!rel) return undefined;
    if (rel.external) return rel.target;
    const key = rel.target.split('/').pop()!;
    if (!this.media[key]) {
      const bytes = this.pkg.zip.get(rel.target);
      if (!bytes) return undefined;
      this.media[key] = toDataUri(bytes, mimeFromExtension(rel.target));
    }
    return key;
  }

  private drawing(el: MarkupElement, part: Part): Inline | undefined {
    const container = child(el, 'wp:inline') ?? child(el, 'wp:anchor');
    if (!container) return undefined;
    const blip = findDeep(container, 'a:blip');
    const src = this.mediaKey(part, attr(blip, 'r:embed') ?? attr(blip, 'r:link'));
    if (!src) return undefined;
    const extent = child(container, 'wp:extent');
    const docPr = child(container, 'wp:docPr');
    const alt = attr(docPr, 'descr') || attr(docPr, 'title');
    return {
      type: 'image',
      src,
      width: emuToPt(num(attr(extent, 'cx')) ?? 0) || 100,
      height: emuToPt(num(attr(extent, 'cy')) ?? 0) || 100,
      ...(alt ? { alt } : {}),
    };
  }

  private vmlImage(el: MarkupElement, part: Part): Inline | undefined {
    const data = findDeep(el, 'v:imagedata');
    const src = this.mediaKey(part, attr(data, 'r:id'));
    if (!src) return undefined;
    const shape = findDeep(el, 'v:shape');
    const style = attr(shape, 'style') ?? '';
    const dim = (name: string) => {
      const m = new RegExp(`${name}:\\s*([\\d.]+)(pt|px|in)?`).exec(style);
      if (!m) return 100;
      const n = Number(m[1]);
      return m[2] === 'px' ? (n * 72) / 96 : m[2] === 'in' ? n * 72 : n;
    };
    return { type: 'image', src, width: dim('width'), height: dim('height') };
  }

  private get activeField(): ComplexField | undefined {
    return this.fields[this.fields.length - 1];
  }

  private get activeLink(): string | undefined {
    for (let i = this.fields.length - 1; i >= 0; i--) {
      const f = this.fields[i]!;
      if (f.phase === 'result' && f.link) return f.link;
    }
    return undefined;
  }

  /** Il contenuto corrente va scartato (istruzioni di campo o risultato di un campo calcolato)? */
  private get suppressed(): boolean {
    return this.fields.some((f) => f.phase === 'instr' || (f.kind && f.phase === 'result'));
  }

  private run(r: MarkupElement, part: Part, link: string | undefined, out: Inline[]): void {
    const rPr = child(r, 'w:rPr');
    const base = this.fmt.char(rPr);
    const rStyle = this.styleId(wVal(child(rPr, 'w:rStyle')));
    const effectiveLink = link ?? this.activeLink;
    const format: CharFormat = effectiveLink ? { ...base, link: effectiveLink } : base;
    const push = (inline: Inline) => {
      if (!this.suppressed) out.push(inline);
    };
    const pushText = (text: string) => {
      if (text) push({ type: 'text', text, format, ...(rStyle ? { styleId: rStyle } : {}) });
    };
    for (const c of r.children) {
      if (!isElement(c)) continue;
      switch (c.name) {
        case 'w:t':
          pushText(textContent(c).replace(/[\r\n]+/g, ' ').replace(/\t/g, ' '));
          break;
        case 'w:tab':
          push({ type: 'tab' });
          break;
        case 'w:br': {
          const type = attr(c, 'w:type');
          push({ type: 'break', kind: type === 'page' ? 'page' : type === 'column' ? 'column' : 'line' });
          break;
        }
        case 'w:cr':
          push({ type: 'break', kind: 'line' });
          break;
        case 'w:noBreakHyphen':
          pushText('‑');
          break;
        case 'w:sym': {
          const code = parseInt(attr(c, 'w:char') ?? '', 16);
          if (Number.isFinite(code) && code < 0xf000) pushText(String.fromCodePoint(code));
          break;
        }
        case 'w:drawing': {
          const img = this.drawing(c, part);
          if (img) push(img);
          break;
        }
        case 'w:pict':
        case 'w:object': {
          const img = this.vmlImage(c, part);
          if (img) push(img);
          break;
        }
        case 'mc:AlternateContent': {
          const branch = child(c, 'mc:Fallback') ?? child(c, 'mc:Choice');
          if (branch) this.run({ ...r, children: branch.children }, part, link, out);
          break;
        }
        case 'w:fldChar': {
          const type = attr(c, 'w:fldCharType');
          if (type === 'begin') this.fields.push({ instr: '', phase: 'instr', emitted: false, format: base });
          else if (type === 'separate') {
            const f = this.activeField;
            if (f) {
              Object.assign(f, parseInstr(f.instr));
              f.phase = 'result';
              if (f.kind && !this.suppressedAbove(f)) {
                out.push({ type: 'field', kind: f.kind, format: f.format });
                f.emitted = true;
              }
            }
          } else if (type === 'end') {
            const f = this.fields.pop();
            if (f && f.phase === 'instr') {
              Object.assign(f, parseInstr(f.instr));
              if (f.kind && !this.suppressed) out.push({ type: 'field', kind: f.kind, format: f.format });
            }
          }
          break;
        }
        case 'w:instrText': {
          const f = this.activeField;
          if (f && f.phase === 'instr') f.instr += textContent(c);
          break;
        }
      }
    }
  }

  private suppressedAbove(field: ComplexField): boolean {
    const i = this.fields.indexOf(field);
    return this.fields.slice(0, i).some((f) => f.phase === 'instr' || (f.kind && f.phase === 'result'));
  }

  private inlineContainer(nodes: readonly MarkupNode[], part: Part, link: string | undefined, out: Inline[]): void {
    for (const c of nodes) {
      if (!isElement(c)) continue;
      switch (c.name) {
        case 'w:r':
          this.run(c, part, link, out);
          break;
        case 'w:hyperlink': {
          const rel = part.rels.get(attr(c, 'r:id') ?? '');
          const anchor = attr(c, 'w:anchor');
          const href = rel ? rel.target : anchor ? `#${anchor}` : undefined;
          this.inlineContainer(c.children, part, href ?? link, out);
          break;
        }
        case 'w:fldSimple': {
          const { kind, link: fieldLink } = parseInstr(attr(c, 'w:instr') ?? '');
          if (kind) {
            const firstRun = child(c, 'w:r');
            if (!this.suppressed) out.push({ type: 'field', kind, format: this.fmt.char(child(firstRun, 'w:rPr')) });
          } else this.inlineContainer(c.children, part, fieldLink ?? link, out);
          break;
        }
        case 'w:sdt':
          this.inlineContainer(child(c, 'w:sdtContent')?.children ?? [], part, link, out);
          break;
        case 'w:ins':
        case 'w:moveTo':
        case 'w:smartTag':
        case 'w:customXml':
        case 'w:dir':
        case 'w:bdo':
          this.inlineContainer(c.children, part, link, out);
          break;
        case 'mc:AlternateContent': {
          const branch = child(c, 'mc:Fallback') ?? child(c, 'mc:Choice');
          if (branch) this.inlineContainer(branch.children, part, link, out);
          break;
        }
      }
    }
  }

  // --- blocchi ------------------------------------------------------------

  paragraph(p: MarkupElement, part: Part): Paragraph {
    const pPr = child(p, 'w:pPr');
    const inlines: Inline[] = [];
    this.inlineContainer(p.children, part, undefined, inlines);
    const styleId = this.styleId(wVal(child(pPr, 'w:pStyle')));
    const mark = this.fmt.char(child(pPr, 'w:rPr'));
    return {
      type: 'paragraph',
      id: newId(),
      format: this.fmt.para(pPr, this.listIds),
      children: normalizeInlines(inlines),
      ...(styleId ? { styleId } : {}),
      ...(Object.keys(mark).length ? { markFormat: mark } : {}),
    };
  }

  private cell(tc: MarkupElement, part: Part): TableCell {
    const tcPr = child(tc, 'w:tcPr');
    const blocks = this.blocks(tc.children, part);
    const span = num(wVal(child(tcPr, 'w:gridSpan')));
    const vMergeEl = child(tcPr, 'w:vMerge');
    const shading = normalizeColor(attr(child(tcPr, 'w:shd'), 'w:fill'));
    const va = wVal(child(tcPr, 'w:vAlign'));
    const borders = readBorders(child(tcPr, 'w:tcBorders'));
    const mar = child(tcPr, 'w:tcMar');
    const side = (n: string, alt: string) => twips(attr(child(mar, `w:${n}`) ?? child(mar, `w:${alt}`), 'w:w'));
    const format: Mutable<TableCellFormat> = {};
    if (shading) format.shading = shading;
    if (va === 'top' || va === 'center' || va === 'bottom') format.verticalAlign = va;
    if (borders) {
      const { top, left, bottom, right } = borders;
      format.borders = { ...(top ? { top } : {}), ...(left ? { left } : {}), ...(bottom ? { bottom } : {}), ...(right ? { right } : {}) };
    }
    if (mar) format.padding = { top: side('top', 'top') ?? 0, right: side('right', 'end') ?? 5.4, bottom: side('bottom', 'bottom') ?? 0, left: side('left', 'start') ?? 5.4 };
    return {
      id: newId(),
      blocks: blocks.length ? blocks : [emptyParagraph()],
      format,
      ...(span && span > 1 ? { colSpan: span } : {}),
      ...(vMergeEl ? { vMerge: wVal(vMergeEl) === 'restart' ? ('restart' as const) : ('continue' as const) } : {}),
    };
  }

  private rowsOf(nodes: readonly MarkupNode[]): MarkupElement[] {
    const out: MarkupElement[] = [];
    for (const c of nodes) {
      if (!isElement(c)) continue;
      if (c.name === 'w:tr') out.push(c);
      else if (c.name === 'w:sdt') out.push(...this.rowsOf(child(c, 'w:sdtContent')?.children ?? []));
      else if (c.name === 'w:customXml') out.push(...this.rowsOf(c.children));
    }
    return out;
  }

  private cellsOf(nodes: readonly MarkupNode[]): MarkupElement[] {
    const out: MarkupElement[] = [];
    for (const c of nodes) {
      if (!isElement(c)) continue;
      if (c.name === 'w:tc') out.push(c);
      else if (c.name === 'w:sdt') out.push(...this.cellsOf(child(c, 'w:sdtContent')?.children ?? []));
      else if (c.name === 'w:customXml') out.push(...this.cellsOf(c.children));
    }
    return out;
  }

  table(tbl: MarkupElement, part: Part): Table {
    const tblPr = child(tbl, 'w:tblPr');
    const rows: TableRow[] = this.rowsOf(tbl.children).map((tr) => {
      const trPr = child(tr, 'w:trPr');
      const h = child(trPr, 'w:trHeight');
      const hv = twips(wVal(h));
      const rule = attr(h, 'w:hRule');
      const cells = this.cellsOf(tr.children).map((tc) => this.cell(tc, part));
      return {
        id: newId(),
        cells: cells.length ? cells : [{ id: newId(), blocks: [emptyParagraph()], format: {} }],
        ...(hv !== undefined ? { height: { rule: rule === 'exact' ? ('exact' as const) : rule === 'auto' ? ('auto' as const) : ('atLeast' as const), value: hv } } : {}),
        ...(onOff(child(trPr, 'w:tblHeader')) ? { isHeader: true } : {}),
      };
    });
    let columnWidths = elements(child(tbl, 'w:tblGrid') ?? { name: '', attrs: {}, children: [] }, 'w:gridCol').map((g) => twips(attr(g, 'w:w')) ?? 0);
    const gridCount = Math.max(0, ...rows.map((r) => r.cells.reduce((n, c) => n + (c.colSpan ?? 1), 0)));
    if (columnWidths.length < gridCount || columnWidths.every((w) => w === 0)) {
      const page = defaultPageSetup();
      const w = (page.width - page.margins.left - page.margins.right) / Math.max(1, gridCount);
      columnWidths = Array.from({ length: gridCount }, (_, i) => columnWidths[i] || w);
    }
    const jc = wVal(child(tblPr, 'w:jc'));
    const align = jc === 'center' ? 'center' : jc === 'right' || jc === 'end' ? 'right' : jc ? 'left' : undefined;
    const ind = child(tblPr, 'w:tblInd');
    const indent = attr(ind, 'w:type') === 'dxa' || attr(ind, 'w:type') === undefined ? twips(attr(ind, 'w:w')) : undefined;
    const borders = readBorders(child(tblPr, 'w:tblBorders'));
    const styleId = this.styleId(wVal(child(tblPr, 'w:tblStyle')));
    return {
      type: 'table',
      id: newId(),
      columnWidths,
      rows: rows.length ? rows : [{ id: newId(), cells: [{ id: newId(), blocks: [emptyParagraph()], format: {} }] }],
      format: { ...(align ? { align } : {}), ...(indent ? { indentLeft: indent } : {}), ...(borders ? { borders } : {}) },
      ...(styleId ? { styleId } : {}),
    };
  }

  /** Blocchi divisi per sezione: un `sectPr` nel pPr chiude la sezione corrente. */
  sectionGroups(nodes: readonly MarkupNode[], part: Part, groups: SectionGroup[] = [{ blocks: [] }]): SectionGroup[] {
    for (const c of nodes) {
      if (!isElement(c)) continue;
      const current = groups[groups.length - 1]!;
      switch (c.name) {
        case 'w:p': {
          const p = this.paragraph(c, part);
          const sectPr = child(child(c, 'w:pPr'), 'w:sectPr');
          // Come in Word, il paragrafo che porta il sectPr resta l'ultimo della sezione.
          current.blocks.push(p);
          if (sectPr) {
            groups[groups.length - 1] = { blocks: current.blocks, sectPr };
            groups.push({ blocks: [] });
          }
          break;
        }
        case 'w:tbl':
          current.blocks.push(this.table(c, part));
          break;
        case 'w:sdt':
          this.sectionGroups(child(c, 'w:sdtContent')?.children ?? [], part, groups);
          break;
        case 'w:customXml':
        case 'w:ins':
        case 'w:moveTo':
          this.sectionGroups(c.children, part, groups);
          break;
        case 'mc:AlternateContent': {
          const branch = child(c, 'mc:Fallback') ?? child(c, 'mc:Choice');
          if (branch) this.sectionGroups(branch.children, part, groups);
          break;
        }
        case 'w:sectPr':
          groups[groups.length - 1] = { blocks: current.blocks, sectPr: c };
          break;
      }
    }
    return groups;
  }

  blocks(nodes: readonly MarkupNode[], part: Part): Block[] {
    return this.sectionGroups(nodes, part).flatMap((g) => g.blocks);
  }

  // --- sezioni -----------------------------------------------------------

  private headerFooter(sectPr: MarkupElement, kind: 'header' | 'footer', part: Part): HeaderFooter | undefined {
    const out: { -readonly [K in keyof HeaderFooter]: HeaderFooter[K] } = {};
    for (const ref of elements(sectPr, `w:${kind}Reference`)) {
      const type = attr(ref, 'w:type') ?? 'default';
      const rel = part.rels.get(attr(ref, 'r:id') ?? '');
      if (!rel || rel.external || (type !== 'default' && type !== 'first' && type !== 'even')) continue;
      const root = this.pkg.xml(rel.target);
      if (!root) continue;
      const blocks = this.blocks(root.children, this.pkg.part(rel.target));
      out[type] = blocks.length ? blocks : [emptyParagraph()];
    }
    return Object.keys(out).length ? out : undefined;
  }

  section(group: SectionGroup, part: Part): Section {
    const blocks = group.blocks.length ? group.blocks : [emptyParagraph()];
    const sectPr = group.sectPr;
    const page: Mutable<PageSetup> = { ...defaultPageSetup() };
    if (!sectPr) return { id: newId(), page, blocks };
    const pgSz = child(sectPr, 'w:pgSz');
    if (pgSz) {
      page.width = twips(attr(pgSz, 'w:w')) ?? page.width;
      page.height = twips(attr(pgSz, 'w:h')) ?? page.height;
      page.orientation = attr(pgSz, 'w:orient') === 'landscape' || page.width > page.height ? 'landscape' : 'portrait';
    }
    const pgMar = child(sectPr, 'w:pgMar');
    if (pgMar) {
      const m = page.margins;
      const get = (n: string, d: number) => twips(attr(pgMar, `w:${n}`)) ?? d;
      page.margins = {
        top: Math.abs(get('top', m.top)),
        right: get('right', m.right),
        bottom: Math.abs(get('bottom', m.bottom)),
        left: get('left', m.left),
        header: get('header', m.header),
        footer: get('footer', m.footer),
      };
    }
    const cols = child(sectPr, 'w:cols');
    const count = num(attr(cols, 'w:num'));
    if (count && count > 1) page.columns = { count, gap: twips(attr(cols, 'w:space')) ?? 36 };
    const type = wVal(child(sectPr, 'w:type'));
    const start = num(attr(child(sectPr, 'w:pgNumType'), 'w:start'));
    const header = this.headerFooter(sectPr, 'header', part);
    const footer = this.headerFooter(sectPr, 'footer', part);
    return {
      id: newId(),
      page,
      blocks,
      ...(type === 'nextPage' || type === 'continuous' || type === 'evenPage' || type === 'oddPage' ? { breakType: type } : {}),
      ...(start !== undefined ? { pageNumberStart: start } : {}),
      ...(onOff(child(sectPr, 'w:titlePg')) ? { titlePage: true } : {}),
      ...(header ? { header } : {}),
      ...(footer ? { footer } : {}),
    };
  }
}

function emptyParagraph(): Paragraph {
  return { type: 'paragraph', id: newId(), format: {}, children: [] };
}

/** Le intestazioni/piè di pagina in Word si ereditano dalla sezione precedente. */
function inheritHeaders(sections: Section[]): Section[] {
  let header: HeaderFooter | undefined;
  let footer: HeaderFooter | undefined;
  return sections.map((s) => {
    header = s.header ? { ...header, ...s.header } : header;
    footer = s.footer ? { ...footer, ...s.footer } : footer;
    return { ...s, ...(header ? { header } : {}), ...(footer ? { footer } : {}) };
  });
}

export async function importDocx(data: Uint8Array): Promise<DocumentModel> {
  const pkg = new Package(await readZip(data));
  // La parte '' è il pacchetto stesso: le sue relazioni sono in _rels/.rels.
  const pkgRels = pkg.part('');
  const officeRel = pkg.findRel(pkgRels, REL.officeDocument);
  const mainPath = officeRel?.target ?? 'word/document.xml';
  const documentRoot = pkg.xml(mainPath);
  if (!documentRoot || documentRoot.name !== 'w:document') throw new Error('Il file non è un documento Word valido');
  const main = pkg.part(mainPath);

  const themeRel = pkg.findRel(main, REL.theme);
  const theme = readTheme(themeRel ? pkg.xml(themeRel.target) : undefined);
  const fmt = new Formatter(theme);
  const numberingRel = pkg.findRel(main, REL.numbering);
  const { lists, ids: listIds } = readNumbering(numberingRel ? pkg.xml(numberingRel.target) : undefined, fmt);
  const stylesRel = pkg.findRel(main, REL.styles);
  const { sheet, ids: styleIds } = readStyles(stylesRel ? pkg.xml(stylesRel.target) : undefined, fmt, theme, listIds);
  const coreRel = pkg.findRel(pkgRels, REL.coreProps);
  const metadata = readCore(pkg.xml(coreRel?.target ?? 'docProps/core.xml'));

  const reader = new Reader(pkg, fmt, styleIds, listIds);
  const body = child(documentRoot, 'w:body');
  const groups = reader.sectionGroups(body?.children ?? [], main);
  // L'ultimo gruppo vuoto senza sectPr (dopo un sectPr finale nel pPr) si scarta.
  const last = groups[groups.length - 1];
  if (groups.length > 1 && last && !last.blocks.length && !last.sectPr) groups.pop();
  const sections = inheritHeaders(groups.map((g) => reader.section(g, main)));

  return { sections, styles: sheet, lists, metadata, media: reader.media };
}
