/**
 * Icone della barra multifunzione: un set di icone SVG a tratto, coerenti
 * tra loro (stile Fluent), scelte per comando e argomenti. Le voci senza
 * icona qui usano il testo `icon` dichiarato dal modulo.
 */

import type { UiRibbonItem } from './ribbon-items';

/** Come appare un pulsante: grande (icona sopra l'etichetta), piccolo o solo icona. */
export type ButtonSize = 'large' | 'small' | 'icon';

const S = (d: string) => `<path d="${d}"/>`;
/** Testo come icona; `attrs` sostituisce le posizioni e la dimensione predefinite. */
const T = (text: string, attrs: Record<string, string | number> = {}) => {
  const all: Record<string, string | number> = { x: 10, y: 14.5, 'font-size': 12, 'text-anchor': 'middle', ...attrs };
  const list = Object.entries(all).map(([k, v]) => `${k}="${v}"`).join(' ');
  return `<text ${list} fill="currentColor" stroke="none">${text}</text>`;
};

const PAGE = S('M5 2.5h6.5L15 6v11.5H5z') + S('M11.5 2.5V6H15');
const LETTER_A = S('M3.5 15.5 7.5 4.5h1l4 11') + S('M5 11.5h6');
const BUBBLE = S('M3 4.5h14v9H9l-4 3v-3H3z');
const LINES = (xs: readonly [number, number][]) => xs.map(([x1, x2], i) => S(`M${x1} ${5 + i * 3.5}h${x2 - x1}`)).join('');

const ICONS: Record<string, string> = {
  'file.new': PAGE + S('M8 11h4M10 9v4'),
  'file.open': S('M2.5 5.5v10h12.5l2.5-6.5H6L3.5 15.5') + S('M2.5 5.5V4h5l1.5 1.5H15V9'),
  'file.export': PAGE + S('M10 8v6.5M7.5 12l2.5 2.5 2.5-2.5'),
  'recent.show': '<circle cx="10" cy="10" r="7"/>' + S('M10 6v4l2.5 2.5'),
  'view.print': S('M6 7V2.5h8V7') + '<rect x="2.5" y="7" width="15" height="7" rx="1.5"/>' + S('M6 11.5h8v6H6z'),

  'edit.undo': S('M7 4.5 3.5 8 7 11.5') + S('M3.5 8H12a4.5 4.5 0 0 1 0 9H8'),
  'edit.redo': S('M13 4.5 16.5 8 13 11.5') + S('M16.5 8H8a4.5 4.5 0 0 0 0 9h4'),
  'edit.find': '<circle cx="8.5" cy="8.5" r="5"/>' + S('m12.5 12.5 5 5'),
  'edit.replace': S('M3.5 7.5h11l-3-3') + S('M16.5 12.5h-11l3 3'),
  'edit.selectAll': '<rect x="3" y="3" width="14" height="14" rx="1" stroke-dasharray="2 2"/>' + S('M7 7.5h6M7 10h6M7 12.5h4'),

  'format.growFont': LETTER_A + S('m13.5 7 2.5-3 2.5 3'),
  'format.shrinkFont': LETTER_A + S('m13.5 4 2.5 3 2.5-3'),
  'format.clear': LETTER_A + S('m13 12 4.5 4.5M17.5 12 13 16.5'),
  'format.bold': T('G', { 'font-weight': 700, 'font-size': 16, y: 15.5 }),
  'format.italic': T('C', { 'font-style': 'italic', 'font-size': 16, y: 15.5, 'font-family': 'Georgia, serif' }),
  'format.underline': T('S', { 'font-size': 15, y: 13.5 }) + S('M5.5 17h9'),
  'format.strike': T('ab', { 'font-size': 14, y: 14 }) + S('M3 10h14'),
  'format.subscript': S('m3 5 6 8M9 5l-6 8') + T('2', { x: 15, y: 17.5, 'font-size': 8 }),
  'format.superscript': S('m3 7 6 8M9 7l-6 8') + T('2', { x: 15, y: 8, 'font-size': 8 }),
  'format.highlight': S('m11.5 3 5 5-6.5 6.5H5v-5z') + S('m5 14.5-2 2.5h4') + '<path class="icon-bar icon-bar-highlight" d="M2 18.5h16"/>',
  'format.color': S('M5 14 9.5 3.5h1L15 14') + S('M6.7 10h6.6') + '<path class="icon-bar icon-bar-color" d="M2 18h16"/>',
  'paragraph.shading':
    S('m4 9 5.5-5.5 6 6L10 15z') + S('M4 9h11.5') + S('M16.5 11.5s1.5 1.7 1.5 2.7a1.5 1.5 0 0 1-3 0c0-1 1.5-2.7 1.5-2.7z') + '<path class="icon-bar icon-bar-shading" d="M2 18.5h16"/>',
  'paragraph.borders': '<rect x="3" y="3" width="14" height="14" stroke-dasharray="1.5 2"/>' + S('M10 3v14M3 10h14').replace('<path', '<path stroke-dasharray="1.5 2"') + S('M2.5 17h15').replace('<path', '<path stroke-width="2"'),
  'formatPainter.toggle': S('M4 3h10v4H4z') + S('M14 5h2.5v4.5H9.5V12') + S('M8.5 12h2v5.5h-2z'),
  'format.link': S('M8.5 11.5a3 3 0 0 0 4.2 0l3-3a3 3 0 0 0-4.2-4.2l-1 1') + S('M11.5 8.5a3 3 0 0 0-4.2 0l-3 3a3 3 0 0 0 4.2 4.2l1-1'),

  'list.toggleBullet':
    '<circle cx="4" cy="5" r="1.2" fill="currentColor"/><circle cx="4" cy="10" r="1.2" fill="currentColor"/><circle cx="4" cy="15" r="1.2" fill="currentColor"/>' +
    S('M8 5h9M8 10h9M8 15h9'),
  'list.toggleNumbered':
    T('1', { x: 4, y: 7, 'font-size': 6 }) + T('2', { x: 4, y: 12, 'font-size': 6 }) + T('3', { x: 4, y: 17, 'font-size': 6 }) + S('M8 5h9M8 10h9M8 15h9'),
  'paragraph.outdent': S('M3 4h14M9 8h8M9 12h8M3 16h14') + S('M6 7.5 3.5 10 6 12.5'),
  'paragraph.indent': S('M3 4h14M9 8h8M9 12h8M3 16h14') + S('M3.5 7.5 6 10l-2.5 2.5'),
  'paragraph.align:left': LINES([[3, 17], [3, 12], [3, 17], [3, 12]]),
  'paragraph.align:center': LINES([[3, 17], [5.5, 14.5], [3, 17], [5.5, 14.5]]),
  'paragraph.align:right': LINES([[3, 17], [8, 17], [3, 17], [8, 17]]),
  'paragraph.align:justify': LINES([[3, 17], [3, 17], [3, 17], [3, 17]]),
  'paragraph.lineSpacing': S('M9 5h8M9 10h8M9 15h8') + S('M4.5 3v14M2.5 5l2-2 2 2M2.5 15l2 2 2-2'),
  'paragraph.spacing': S('M3 8h14M3 12h14') + S('M10 2.5v3M8.5 4 10 2.5 11.5 4M10 17.5v-3M8.5 16l1.5 1.5 1.5-1.5'),
  'view.toggleFormattingMarks': S('M11 16.5V3.5M14 16.5V3.5M16 3.5H9a3.5 3.5 0 0 0 0 7h2'),

  'insert.pageBreak': S('M5 8V2.5h10V8M5 12v5.5h10V12') + S('M3 10h2M7 10h2M11 10h2M15 10h2'),
  'insert.sectionBreak': S('M5 7V2.5h10V7M5 13v4.5h10V13') + S('M2.5 10h15') + S('M8 8.5 10 10l2-1.5'),
  'table.insert': '<rect x="2.5" y="3.5" width="15" height="13" rx="1.5"/>' + S('M2.5 8h15M2.5 12.5h15M7.5 3.5v13M12.5 3.5v13'),
  'drawing.insert:shape': '<rect x="2.5" y="8.5" width="8" height="8" rx="1"/><circle cx="12.5" cy="7.5" r="5"/>',
  'drawing.insert:chart': S('M3 17.5h14.5') + S('M5 17.5v-6h3v6M9.5 17.5v-11h3v11M14 17.5v-8h3v8'),
  'drawing.insert:textbox': '<rect x="2.5" y="3.5" width="15" height="13" rx="1"/>' + S('M6 7.5h8M10 7.5v6'),
  'insert.image': '<rect x="2.5" y="3.5" width="15" height="13" rx="1.5"/><circle cx="7" cy="8" r="1.5"/>' + S('m3 15 4.5-4.5 3 3 2.5-2.5 4.5 4.5'),
  'page.headerFooter:header': PAGE + '<path class="icon-fill" d="M6.5 4h4.5v2.5h-4.5z"/>',
  'page.headerFooter:footer': PAGE + '<path class="icon-fill" d="M6.5 13.5h7V16h-7z"/>',
  'insert.field:page': S('M7 3.5 5.5 16.5M13 3.5l-1.5 13M3.5 7.5h13M3 12.5h13'),
  'insert.field:numPages': S('M7 3.5 5.5 16.5M13 3.5l-1.5 13M3.5 7.5h13M3 12.5h13'),
  'insert.field:date': '<rect x="2.5" y="4" width="15" height="13" rx="1.5"/>' + S('M2.5 8h15M6.5 2.5v3M13.5 2.5v3M6 11h2M9 11h2M12 11h2M6 14h2M9 14h2'),
  'insert.field:time': '<circle cx="10" cy="10" r="7.5"/>' + S('M10 5.5V10l3 2'),
  'page.setup': S('M3 6.5h11v11H3z') + S('M8 2.5h9v9') + S('M12.5 2.5 14 1M12.5 2.5 14 4'),

  'footnote.insert': T('ab', { x: 8, 'font-size': 11 }) + T('1', { x: 16, y: 8, 'font-size': 7 }) + S('M3 17.5h14'),
  'footnote.edit': S('M3.5 16.5 4 13l9-9 3 3-9 9z') + S('m11 6 3 3'),
  'footnote.delete': T('1', { x: 8, y: 9, 'font-size': 8 }) + S('m11 11 5.5 5.5M16.5 11 11 16.5'),

  'comment.insert': BUBBLE + S('M10 6.5v5M7.5 9h5'),
  'comment.delete': BUBBLE + S('m8 7 4 4M12 7l-4 4'),
  'comment.previous': BUBBLE + S('M11.5 6.5 9 9l2.5 2.5'),
  'comment.next': BUBBLE + S('M8.5 6.5 11 9l-2.5 2.5'),

  'revision.toggleTracking': PAGE + S('m7.5 15 .3-2 4.5-4.5 1.7 1.7-4.5 4.5z'),
  'revision.setAuthor': '<circle cx="10" cy="6.5" r="3.5"/>' + S('M3.5 17.5a6.5 6.5 0 0 1 13 0'),
  'revision.previous': S('M16 10H4.5M9 5.5 4.5 10 9 14.5'),
  'revision.next': S('M4 10h11.5M11 5.5l4.5 4.5-4.5 4.5'),
  'revision.accept': S('m4 10.5 4 4 8-9'),
  'revision.reject': S('m5 5 10 10M15 5 5 15'),
  'revision.acceptAll': S('m2 10.5 3.5 3.5 7-8') + S('m9 13 1 1 7-8'),
  'revision.rejectAll': S('m3 6 8 8M11 6l-8 8') + S('m11.5 9.5 5.5 5.5M17 9.5 13.5 13'),

  'spelling.next': T('ABC', { 'font-size': 7.5, y: 9, 'font-weight': 600 }) + S('m6 14 2.5 2.5 5.5-5'),
  'spelling.toggle': T('ab', { 'font-size': 11, y: 11 }) + '<path class="icon-squiggle" d="m2.5 16 1.5-1.5 1.5 1.5 1.5-1.5 1.5 1.5 1.5-1.5 1.5 1.5 1.5-1.5 1.5 1.5 1.5-1.5 1.5 1.5"/>',

  'toc.insert': S('M3 4.5h14') + S('M3 8.5h9M3 12h9M3 15.5h9') + '<path d="M13.5 8.5h3.5M13.5 12h3.5M13.5 15.5h3.5" stroke-dasharray="1 1.2"/>',
  'toc.update': S('M15.5 6.5A6.5 6.5 0 1 0 16.5 11') + S('M16 2.5v4h-4'),
  'toc.remove': S('M3 4.5h9M3 8.5h6M3 12.5h6') + S('m11 11 5.5 5.5M16.5 11 11 16.5'),

  'view.zoom': '<circle cx="8.5" cy="8.5" r="5"/>' + S('m12.5 12.5 5 5M6.5 8.5h4M8.5 6.5v4'),
  'view.zoom:1': T('1:1', { 'font-size': 9, 'font-weight': 600 }),
};

const ARG_KEYS = ['align', 'part', 'kind', 'zoom'] as const;

/** Chiave dell'icona: id del comando, più l'argomento che distingue le varianti. */
export function iconKey(item: Pick<UiRibbonItem, 'commandId' | 'args'>): string[] {
  const keys: string[] = [];
  const args = item.args as Record<string, unknown> | undefined;
  if (args && typeof args === 'object') {
    for (const k of ARG_KEYS) if (k in args) keys.push(`${item.commandId}:${String(args[k])}`);
  }
  keys.push(item.commandId);
  return keys;
}

/** Icona SVG della voce, o `null` se il set non la prevede. */
export function iconSvg(item: Pick<UiRibbonItem, 'commandId' | 'args'>): SVGSVGElement | null {
  const key = iconKey(item).find((k) => k in ICONS);
  if (!key) return null;
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.setAttribute('viewBox', '0 0 20 20');
  svg.setAttribute('class', 'ribbon-svg');
  svg.setAttribute('aria-hidden', 'true');
  svg.setAttribute('fill', 'none');
  svg.setAttribute('stroke', 'currentColor');
  svg.setAttribute('stroke-width', '1.4');
  svg.setAttribute('stroke-linecap', 'round');
  svg.setAttribute('stroke-linejoin', 'round');
  svg.innerHTML = ICONS[key]!;
  return svg;
}

/** Comandi mostrati come pulsanti grandi, come in Word. */
const LARGE = new Set([
  'file.new',
  'file.open',
  'view.print',
  'insert.pageBreak',
  'table.insert',
  'insert.image',
  'drawing.insert',
  'format.link',
  'page.headerFooter',
  'page.setup',
  'insert.sectionBreak',
  'footnote.insert',
  'toc.insert',
  'comment.insert',
  'revision.toggleTracking',
  'revision.accept',
  'revision.reject',
  'spelling.next',
  'view.toggleFormattingMarks',
  'view.zoom',
]);

/** Gruppi della scheda Home dove, come in Word, i pulsanti mostrano solo l'icona. */
const ICON_ONLY_GROUPS = new Set(['Annulla', 'Carattere', 'Paragrafo']);

export function buttonSize(item: UiRibbonItem): ButtonSize {
  if (item.tab === 'home' && ICON_ONLY_GROUPS.has(item.group)) return 'icon';
  if (item.tab === 'home') return 'small';
  return LARGE.has(item.commandId) ? 'large' : 'small';
}
