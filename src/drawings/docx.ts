/**
 * Forme, caselle di testo e grafici in DOCX. Si scrivono come in Word 2010
 * e successivi: forme `wps:wsp` ancorate alla pagina ("Davanti al testo") e
 * grafici DrawingML in una parte `word/charts/chartN.xml` con i dati in
 * cache. La lettura riconosce gli stessi elementi (anche salvati da Word).
 */

import {
  emuToPt,
  ptToEmu,
  type Alignment,
  type CharFormat,
  type ChartData,
  type ChartSeries,
  type Drawing,
  type DrawingText,
  type ShapeKind,
} from '../core';
import { child, elements, escapeAttr, escapeText, isElement, textContent, type MarkupElement } from '../io/util/markup';
import { LINE_SHAPES, newDrawingId } from './model';
import { seriesColor, pointColor } from './chart';
import { TEXT_INSET } from './render';

export const WPS_NS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
export const CHART_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
export const CHART_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
export const CHART_CT = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const hex = (c: string) => c.replace('#', '').toUpperCase();
const attr = (el: MarkupElement | undefined, name: string): string | undefined => el?.attrs[name];

// ---------------------------------------------------------------------------
// Scrittura

function solidFill(color: string | null | undefined): string {
  return color ? `<a:solidFill><a:srgbClr val="${hex(color)}"/></a:solidFill>` : '<a:noFill/>';
}

function runPr(f: CharFormat): string {
  const parts: string[] = [];
  if (f.fontFamily) parts.push(`<w:rFonts w:ascii="${escapeAttr(f.fontFamily)}" w:hAnsi="${escapeAttr(f.fontFamily)}"/>`);
  if (f.bold) parts.push('<w:b/>');
  if (f.italic) parts.push('<w:i/>');
  if (f.color) parts.push(`<w:color w:val="${hex(f.color)}"/>`);
  if (f.fontSize) parts.push(`<w:sz w:val="${Math.round(f.fontSize * 2)}"/>`);
  return parts.length ? `<w:rPr>${parts.join('')}</w:rPr>` : '';
}

const JC: Record<Alignment, string> = { left: 'left', center: 'center', right: 'right', justify: 'both' };

function textBoxContent(t: DrawingText): string {
  const rPr = runPr(t.format);
  const pPr = `<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>${t.align && t.align !== 'left' ? `<w:jc w:val="${JC[t.align]}"/>` : ''}${rPr}</w:pPr>`;
  const paras = t.text.split('\n').map((line) => `<w:p>${pPr}${line ? `<w:r>${rPr}<w:t xml:space="preserve">${escapeText(line)}</w:t></w:r>` : ''}</w:p>`);
  return `<wps:txbx><w:txbxContent>${paras.join('')}</w:txbxContent></wps:txbx>`;
}

function prstFor(d: Drawing): string {
  if (d.kind === 'textbox') return 'rect';
  const s = d.shape ?? 'rect';
  return LINE_SHAPES.has(s) ? 'line' : s;
}

function shapeXml(d: Drawing): string {
  const line = d.kind === 'shape' && !!d.shape && LINE_SHAPES.has(d.shape);
  const flip = `${d.flipH ? ' flipH="1"' : ''}${d.flipV ? ' flipV="1"' : ''}`;
  const ends = d.shape === 'arrow' ? '<a:tailEnd type="triangle"/>' : d.shape === 'doubleArrow' ? '<a:headEnd type="triangle"/><a:tailEnd type="triangle"/>' : '';
  const ln = d.stroke ? `<a:ln w="${ptToEmu(d.strokeWidth ?? 0.75)}">${solidFill(d.stroke)}${ends}</a:ln>` : '<a:ln><a:noFill/></a:ln>';
  const spPr =
    `<wps:spPr><a:xfrm${flip}><a:off x="0" y="0"/><a:ext cx="${ptToEmu(d.width)}" cy="${ptToEmu(d.height)}"/></a:xfrm>` +
    `<a:prstGeom prst="${prstFor(d)}"><a:avLst/></a:prstGeom>${line ? '' : solidFill(d.fill)}${ln}</wps:spPr>`;
  const anchor = { top: 't', middle: 'ctr', bottom: 'b' }[d.text?.verticalAlign ?? (d.kind === 'shape' ? 'middle' : 'top')];
  const ins = `lIns="${ptToEmu(TEXT_INSET.x)}" tIns="${ptToEmu(TEXT_INSET.y)}" rIns="${ptToEmu(TEXT_INSET.x)}" bIns="${ptToEmu(TEXT_INSET.y)}"`;
  const txbx = !line && d.text?.text ? textBoxContent(d.text) : '';
  return (
    `<a:graphicData uri="${WPS_NS}"><wps:wsp>` +
    (line ? '<wps:cNvCnPr/>' : `<wps:cNvSpPr${d.kind === 'textbox' ? ' txBox="1"' : ''}/>`) +
    spPr +
    txbx +
    `<wps:bodyPr rot="0" vert="horz" wrap="square" ${ins} anchor="${anchor}" anchorCtr="0"><a:noAutofit/></wps:bodyPr>` +
    `</wps:wsp></a:graphicData>`
  );
}

/**
 * Run con il disegno ancorato. `docPrId` è unico nel documento; per i
 * grafici `chartRelId` è la relazione verso la parte del grafico.
 */
export function drawingRunXml(d: Drawing, docPrId: number, z: number, chartRelId?: string): string {
  const name = d.kind === 'chart' ? 'Grafico' : d.kind === 'textbox' ? 'Casella di testo' : 'Forma';
  const graphic =
    d.kind === 'chart'
      ? `<a:graphicData uri="${CHART_NS}"><c:chart xmlns:c="${CHART_NS}" r:id="${chartRelId ?? ''}"/></a:graphicData>`
      : shapeXml(d);
  return (
    `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="${251659264 + z * 1024}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="page"><wp:posOffset>${ptToEmu(d.x)}</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="page"><wp:posOffset>${ptToEmu(d.y)}</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${ptToEmu(d.width)}" cy="${ptToEmu(d.height)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>` +
    `<wp:docPr id="${docPrId}" name="${name} ${docPrId}"/>` +
    `<wp:cNvGraphicFramePr/><a:graphic xmlns:a="${A_NS}">${graphic}</a:graphic></wp:anchor></w:drawing></w:r>`
  );
}

const strCache = (ref: string, values: readonly string[]) =>
  `<c:strRef><c:f>${escapeText(ref)}</c:f><c:strCache><c:ptCount val="${values.length}"/>${values
    .map((v, i) => `<c:pt idx="${i}"><c:v>${escapeText(v)}</c:v></c:pt>`)
    .join('')}</c:strCache></c:strRef>`;

const numCache = (ref: string, values: readonly number[]) =>
  `<c:numRef><c:f>${escapeText(ref)}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${values
    .map((v, i) => (Number.isFinite(v) ? `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>` : ''))
    .join('')}</c:numCache></c:numRef>`;

const col = (i: number) => String.fromCharCode(66 + (i % 25));

function seriesXml(chart: ChartData, s: ChartSeries, i: number): string {
  const n = chart.categories.length;
  const color = seriesColor(chart, i);
  const head = `<c:idx val="${i}"/><c:order val="${i}"/><c:tx>${strCache(`Foglio1!$${col(i)}$1`, [s.name])}</c:tx>`;
  const cat = `<c:cat>${strCache(`Foglio1!$A$2:$A$${n + 1}`, chart.categories)}</c:cat>`;
  const val = `<c:val>${numCache(`Foglio1!$${col(i)}$2:$${col(i)}$${n + 1}`, s.values.slice(0, n))}</c:val>`;
  switch (chart.type) {
    case 'line':
      return `<c:ser>${head}<c:spPr><a:ln w="28575" cap="rnd">${solidFill(color)}<a:round/></a:ln></c:spPr><c:marker><c:symbol val="none"/></c:marker>${cat}${val}<c:smooth val="0"/></c:ser>`;
    case 'pie': {
      const points = chart.categories
        .map((_, p) => `<c:dPt><c:idx val="${p}"/><c:bubble3D val="0"/><c:spPr>${solidFill(pointColor(p))}<a:ln w="12700">${solidFill('#ffffff')}</a:ln></c:spPr></c:dPt>`)
        .join('');
      return `<c:ser>${head}${points}${cat}${val}</c:ser>`;
    }
    default:
      return `<c:ser>${head}<c:spPr>${solidFill(color)}</c:spPr><c:invertIfNegative val="0"/>${cat}${val}</c:ser>`;
  }
}

/** Parte `word/charts/chartN.xml` di un grafico. */
export function chartPartXml(chart: ChartData, d?: Pick<Drawing, 'fill' | 'stroke' | 'strokeWidth'>): string {
  const axes = '<c:axId val="500000001"/><c:axId val="500000002"/>';
  const horizontal = chart.type === 'bar';
  let plot: string;
  if (chart.type === 'pie') {
    plot = `<c:pieChart><c:varyColors val="1"/>${chart.series.slice(0, 1).map((s, i) => seriesXml(chart, s, i)).join('')}<c:firstSliceAng val="0"/></c:pieChart>`;
  } else {
    const series = chart.series.map((s, i) => seriesXml(chart, s, i)).join('');
    const body =
      chart.type === 'line'
        ? `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${series}<c:marker val="1"/>${axes}</c:lineChart>`
        : `<c:barChart><c:barDir val="${horizontal ? 'bar' : 'col'}"/><c:grouping val="clustered"/><c:varyColors val="0"/>${series}<c:gapWidth val="150"/>${axes}</c:barChart>`;
    const common = '<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>';
    plot =
      body +
      `<c:catAx><c:axId val="500000001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${horizontal ? 'l' : 'b'}"/><c:numFmt formatCode="General" sourceLinked="1"/>${common}<c:crossAx val="500000002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>` +
      `<c:valAx><c:axId val="500000002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${horizontal ? 'b' : 'l'}"/><c:majorGridlines><c:spPr><a:ln w="9525">${solidFill('#D9D9D9')}</a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode="General" sourceLinked="1"/>${common}<c:crossAx val="500000001"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
  }
  const title = chart.title
    ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${escapeText(chart.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`
    : '<c:autoTitleDeleted val="1"/>';
  const frame = d ? `<c:spPr>${solidFill(d.fill)}${d.stroke ? `<a:ln w="${ptToEmu(d.strokeWidth ?? 0.75)}">${solidFill(d.stroke)}</a:ln>` : '<a:ln><a:noFill/></a:ln>'}</c:spPr>` : '';
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<c:chartSpace xmlns:c="${CHART_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}"><c:roundedCorners val="0"/>` +
    `<c:chart>${title}<c:plotArea><c:layout/>${plot}</c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
    `${frame}</c:chartSpace>`
  );
}

// ---------------------------------------------------------------------------
// Lettura

function findDeep(el: MarkupElement | undefined, name: string): MarkupElement | undefined {
  if (!el) return undefined;
  for (const c of el.children) {
    if (!isElement(c)) continue;
    if (c.name === name) return c;
    const found = findDeep(c, name);
    if (found) return found;
  }
  return undefined;
}

function readColor(el: MarkupElement | undefined): string | null | undefined {
  if (!el) return undefined;
  if (child(el, 'a:noFill')) return null;
  const clr = findDeep(child(el, 'a:solidFill'), 'a:srgbClr');
  const v = attr(clr, 'val');
  return v && /^[0-9a-f]{6}$/i.test(v) ? `#${v.toLowerCase()}` : undefined;
}

/** Riferimento orizzontale/verticale di una posizione (`relativeFrom`) e scostamento in punti. */
export interface AnchorPosition {
  readonly from: string;
  readonly offset: number;
}

/** Disegno letto, ancora da collocare: posizioni rispetto a pagina o margini. */
export interface ReadDrawing {
  readonly drawing: Omit<Drawing, 'anchor'>;
  readonly h: AnchorPosition;
  readonly v: AnchorPosition;
  /** Ordine di sovrapposizione (`relativeHeight`): più alto = più davanti. */
  readonly z: number;
}

const SHAPES = new Set<string>([
  'rect', 'roundRect', 'ellipse', 'triangle', 'rtTriangle', 'diamond', 'pentagon', 'hexagon', 'star5', 'heart',
  'rightArrow', 'leftArrow', 'upArrow', 'downArrow', 'wedgeRectCallout',
]);

function readText(txbx: MarkupElement | undefined, bodyPr: MarkupElement | undefined, isShape: boolean): DrawingText | undefined {
  const content = findDeep(txbx, 'w:txbxContent');
  const paras = content ? elements(content, 'w:p') : [];
  const anchor = attr(bodyPr, 'anchor');
  const verticalAlign = anchor === 'ctr' ? 'middle' : anchor === 'b' ? 'bottom' : anchor === 't' ? 'top' : isShape ? 'middle' : 'top';
  if (!paras.length) return undefined;
  const lines = paras.map((p) => elements(p, 'w:r').map((r) => elements(r).map((e) => (e.name === 'w:t' ? textContent(e) : e.name === 'w:tab' ? '\t' : '')).join('')).join(''));
  const firstRun = paras.flatMap((p) => elements(p, 'w:r'))[0];
  const rPr = child(firstRun, 'w:rPr') ?? child(child(paras[0], 'w:pPr'), 'w:rPr');
  const format: { -readonly [K in keyof CharFormat]: CharFormat[K] } = {};
  const sz = Number(attr(child(rPr, 'w:sz'), 'w:val'));
  if (sz) format.fontSize = sz / 2;
  const color = attr(child(rPr, 'w:color'), 'w:val');
  if (color && /^[0-9a-f]{6}$/i.test(color)) format.color = `#${color.toLowerCase()}`;
  const font = attr(child(rPr, 'w:rFonts'), 'w:ascii');
  if (font) format.fontFamily = font;
  const on = (name: string) => {
    const el = child(rPr, name);
    return !!el && !['0', 'false', 'off'].includes(attr(el, 'w:val') ?? '');
  };
  if (on('w:b')) format.bold = true;
  if (on('w:i')) format.italic = true;
  const jc = attr(child(child(paras[0], 'w:pPr'), 'w:jc'), 'w:val');
  const align: Alignment = jc === 'center' ? 'center' : jc === 'right' || jc === 'end' ? 'right' : jc === 'both' ? 'justify' : 'left';
  return { text: lines.join('\n'), format, align, verticalAlign };
}

function readPosition(el: MarkupElement | undefined): AnchorPosition {
  const from = attr(el, 'relativeFrom') ?? 'page';
  const off = child(el, 'wp:posOffset');
  return { from, offset: off ? emuToPt(Number(textContent(off)) || 0) : 0 };
}

/** Dati di un grafico da una parte `chartN.xml` (già con prefissi canonici). */
export function readChart(root: MarkupElement | undefined): ChartData | undefined {
  const chart = child(root, 'c:chart');
  const plot = child(chart, 'c:plotArea');
  if (!plot) return undefined;
  const kinds: [string, ChartData['type']][] = [['c:barChart', 'column'], ['c:bar3DChart', 'column'], ['c:lineChart', 'line'], ['c:line3DChart', 'line'], ['c:pieChart', 'pie'], ['c:pie3DChart', 'pie'], ['c:doughnutChart', 'pie']];
  let group: MarkupElement | undefined;
  let type: ChartData['type'] = 'column';
  for (const [name, t] of kinds) {
    group = child(plot, name);
    if (group) {
      type = t;
      break;
    }
  }
  if (!group) return undefined;
  if (type === 'column' && attr(child(group, 'c:barDir'), 'val') === 'bar') type = 'bar';
  const points = (el: MarkupElement | undefined): string[] => {
    const cache = findDeep(el, 'c:strCache') ?? findDeep(el, 'c:numCache') ?? findDeep(el, 'c:strLit') ?? findDeep(el, 'c:numLit');
    if (!cache) return [];
    const count = Number(attr(child(cache, 'c:ptCount'), 'val')) || 0;
    const out: string[] = Array.from({ length: count }, () => '');
    for (const pt of elements(cache, 'c:pt')) {
      const i = Number(attr(pt, 'idx'));
      if (Number.isInteger(i) && i >= 0) out[i] = textContent(child(pt, 'c:v') ?? pt);
    }
    return out;
  };
  let categories: string[] = [];
  const series: ChartSeries[] = elements(group, 'c:ser').map((ser, i) => {
    const cats = points(child(ser, 'c:cat'));
    if (cats.length > categories.length) categories = cats;
    const tx = child(ser, 'c:tx');
    const name = (tx && (points(tx)[0] ?? textContent(child(tx, 'c:v') ?? tx))) || `Serie ${i + 1}`;
    const values = points(child(ser, 'c:val')).map((v) => (v === '' ? 0 : Number(v)));
    const color = type === 'pie' ? undefined : readColor(child(ser, 'c:spPr')) ?? readColor(child(child(ser, 'c:spPr'), 'a:ln')) ?? undefined;
    return { name, values, ...(color ? { color } : {}) };
  });
  if (!categories.length) categories = Array.from({ length: Math.max(0, ...series.map((s) => s.values.length)) }, (_, i) => String(i + 1));
  const titleEl = child(chart, 'c:title');
  const title = titleEl ? collectTexts(titleEl).trim() : '';
  return { type, ...(title ? { title } : {}), categories, series };
}

function collectTexts(el: MarkupElement): string {
  let out = '';
  for (const c of el.children) {
    if (!isElement(c)) continue;
    out += c.name === 'a:t' ? textContent(c) : collectTexts(c);
  }
  return out;
}

/**
 * Forma, casella di testo o grafico da un `w:drawing` (null se è
 * un'immagine o un oggetto non gestito). `chart` carica la parte del
 * grafico dall'id di relazione.
 */
export function readDrawing(drawing: MarkupElement, chart: (relId: string) => MarkupElement | undefined): ReadDrawing | undefined {
  const container = child(drawing, 'wp:anchor') ?? child(drawing, 'wp:inline');
  if (!container) return undefined;
  const extent = child(container, 'wp:extent');
  const width = emuToPt(Number(attr(extent, 'cx')) || 0);
  const height = emuToPt(Number(attr(extent, 'cy')) || 0);
  const inline = container.name === 'wp:inline';
  const h: AnchorPosition = inline ? { from: 'margin', offset: 0 } : readPosition(child(container, 'wp:positionH'));
  const v: AnchorPosition = inline ? { from: 'paragraph', offset: 0 } : readPosition(child(container, 'wp:positionV'));
  const z = Number(attr(container, 'relativeHeight')) || 0;
  const data = findDeep(container, 'a:graphicData');
  const base = { id: newDrawingId(), x: 0, y: 0, width, height };

  const chartRef = child(data, 'c:chart');
  if (chartRef) {
    const root = chart(attr(chartRef, 'r:id') ?? '');
    const parsed = readChart(root);
    if (!parsed) return undefined;
    const spPr = child(root, 'c:spPr');
    const fill = readColor(spPr);
    const stroke = readColor(child(spPr, 'a:ln'));
    return { drawing: { ...base, kind: 'chart', chart: parsed, fill: fill === undefined ? '#ffffff' : fill, stroke: stroke === undefined ? '#d9d9d9' : stroke, strokeWidth: 0.75 }, h, v, z };
  }

  const wsp = child(data, 'wps:wsp');
  if (!wsp) return undefined;
  const spPr = child(wsp, 'wps:spPr');
  const xfrm = child(spPr, 'a:xfrm');
  const prst = attr(child(spPr, 'a:prstGeom'), 'prst') ?? 'rect';
  const ln = child(spPr, 'a:ln');
  const isTextBox = attr(child(wsp, 'wps:cNvSpPr'), 'txBox') === '1';
  const isLine = !!child(wsp, 'wps:cNvCnPr') || /^(line|straightConnector1|bentConnector\d|curvedConnector\d)$/.test(prst);
  const ends = (name: string) => {
    const t = attr(child(ln, name), 'type');
    return !!t && t !== 'none';
  };
  let shape: ShapeKind | undefined;
  if (isLine) {
    const head = ends('a:headEnd');
    const tail = ends('a:tailEnd');
    shape = head && tail ? 'doubleArrow' : head || tail ? 'arrow' : 'line';
  } else if (!isTextBox) shape = SHAPES.has(prst) ? (prst as ShapeKind) : 'rect';
  const lineWidth = Number(attr(ln, 'w'));
  const stroke = readColor(ln);
  const fill = readColor(spPr);
  const text = isLine ? undefined : readText(child(wsp, 'wps:txbx'), child(wsp, 'wps:bodyPr'), !isTextBox);
  let geometry = { flipH: attr(xfrm, 'flipH') === '1', flipV: attr(xfrm, 'flipV') === '1' };
  // Una linea con la testa all'inizio e senza coda è una freccia girata.
  if (shape === 'arrow' && ends('a:headEnd')) geometry = { flipH: !geometry.flipH, flipV: !geometry.flipV };
  const d: Omit<Drawing, 'anchor'> = {
    ...base,
    kind: isTextBox ? 'textbox' : 'shape',
    ...(shape ? { shape } : {}),
    ...(geometry.flipH ? { flipH: true } : {}),
    ...(geometry.flipV ? { flipV: true } : {}),
    fill: isLine ? null : fill === undefined ? (isTextBox ? '#ffffff' : '#4472c4') : fill,
    stroke: stroke === undefined ? (isTextBox ? '#000000' : isLine ? '#4472c4' : '#2f528f') : stroke,
    strokeWidth: lineWidth ? emuToPt(lineWidth) : 0.75,
    ...(text ? { text } : {}),
  };
  return { drawing: d, h, v, z };
}

/** Posizione dal bordo della pagina: i riferimenti diversi dalla pagina partono dal margine. */
export function resolvePosition(p: AnchorPosition, margin: number): number {
  return p.from === 'page' ? p.offset : margin + p.offset;
}
