/**
 * Collegamenti ipertestuali: indirizzi, ricerca nel testo e stile.
 *
 * Un collegamento è il campo `link` del formato dei run di testo:
 * - `https://…`, `mailto:…` (e altri schemi) per web ed email;
 * - `#_Toc<id del paragrafo>` per un punto del documento (stessa convenzione
 *   del sommario, che in DOCX diventa un segnalibro sul paragrafo);
 * - `#_top` per l'inizio del documento (segnalibro predefinito di Word).
 * I run collegati ricevono lo stile di carattere "Hyperlink" (blu,
 * sottolineato), come in Word.
 */

import {
  findParagraph,
  inlineLength,
  paragraphText,
  paragraphs,
  type DocumentModel,
  type Id,
  type Inline,
  type Paragraph,
  type Position,
  type Style,
  type StyleSheet,
} from '../core';
import { anchorTarget, collectHeadings, tocAnchor } from '../toc/toc';

export const HYPERLINK_STYLE = 'Hyperlink';
export const TOP_ANCHOR = '#_top';

export type LinkKind = 'web' | 'email' | 'place';

const EMAIL = /^[\w.!#$%&'*+/=?^`{|}~-]+@[\w-]+(\.[\w-]+)+$/;
const SCHEME = /^[a-z][a-z0-9+.-]*:/i;

export function linkKind(url: string): LinkKind {
  if (url.startsWith('#')) return 'place';
  if (/^mailto:/i.test(url)) return 'email';
  return 'web';
}

/**
 * Indirizzo scritto dall'utente → collegamento: "www.sito.it" diventa
 * "https://www.sito.it", "nome@sito.it" diventa "mailto:nome@sito.it".
 * Vuoto → `null`.
 */
export function normalizeAddress(input: string, kind: LinkKind = 'web'): string | null {
  const text = input.trim();
  if (!text) return null;
  if (kind === 'email' || EMAIL.test(text)) {
    const address = text.replace(/^mailto:/i, '');
    return address ? `mailto:${address}` : null;
  }
  if (text.startsWith('#') || SCHEME.test(text)) return text;
  return `https://${text}`;
}

/** Testo da mostrare quando l'utente non ne scrive uno: l'indirizzo senza "mailto:". */
export function displayText(url: string): string {
  return url.replace(/^mailto:/i, '');
}

/** Collegamenti che si possono aprire fuori dall'app (niente javascript: o file:). */
export function isExternalSafe(url: string): boolean {
  return /^(https?:|mailto:|ftp:)/i.test(url.trim());
}

// ---------------------------------------------------------------------------
// Stile

/** Id dello stile "Hyperlink" del foglio, se c'è (anche con un id localizzato). */
export function hyperlinkStyleId(sheet: StyleSheet): Id | undefined {
  if (sheet.styles[HYPERLINK_STYLE]) return HYPERLINK_STYLE;
  return Object.values(sheet.styles).find((s) => s.type === 'character' && /^hyperlink$/i.test(s.name))?.id;
}

const hyperlinkStyle: Style = {
  id: HYPERLINK_STYLE,
  name: 'Hyperlink',
  type: 'character',
  charFormat: { color: '#0563C1', underline: 'single' },
};

/** Aggiunge lo stile "Hyperlink" se manca. */
export function ensureHyperlinkStyle(sheet: StyleSheet): { sheet: StyleSheet; styleId: Id } {
  const existing = hyperlinkStyleId(sheet);
  if (existing) return { sheet, styleId: existing };
  return { sheet: { ...sheet, styles: { ...sheet.styles, [HYPERLINK_STYLE]: hyperlinkStyle } }, styleId: HYPERLINK_STYLE };
}

// ---------------------------------------------------------------------------
// Collegamenti nel testo

export interface LinkRange {
  readonly paragraphId: Id;
  readonly start: number;
  readonly end: number;
  readonly url: string;
  readonly text: string;
}

const linkOf = (i: Inline | undefined) => (i?.type === 'text' ? i.format.link : undefined);

/** Collegamento che contiene il carattere `index` del paragrafo. */
function linkAtChar(p: Paragraph, index: number): LinkRange | null {
  const spans: { from: number; to: number; inline: Inline }[] = [];
  let pos = 0;
  for (const inline of p.children) {
    spans.push({ from: pos, to: pos + inlineLength(inline), inline });
    pos += inlineLength(inline);
  }
  const hit = spans.findIndex((s) => s.from <= index && index < s.to && linkOf(s.inline));
  if (hit < 0) return null;
  const url = linkOf(spans[hit]!.inline)!;
  let a = hit;
  let b = hit;
  while (a > 0 && linkOf(spans[a - 1]!.inline) === url) a--;
  while (b + 1 < spans.length && linkOf(spans[b + 1]!.inline) === url) b++;
  const start = spans[a]!.from;
  const end = spans[b]!.to;
  return { paragraphId: p.id, start, end, url, text: paragraphText(p).slice(start, end) };
}

function findP(doc: DocumentModel, id: Id): Paragraph | undefined {
  return findParagraph(doc, id)?.paragraph;
}

/**
 * Collegamento nella posizione data: il carattere dopo il cursore o, se il
 * cursore è subito dopo un collegamento, quello prima (come il clic destro
 * di Word sulla fine di un collegamento).
 */
export function linkAt(doc: DocumentModel, pos: Position): LinkRange | null {
  const p = findP(doc, pos.paragraphId);
  if (!p) return null;
  return linkAtChar(p, pos.offset) ?? (pos.offset > 0 ? linkAtChar(p, pos.offset - 1) : null);
}

/** Collegamento strettamente intorno al cursore (non agli estremi): serve a sapere se si scrive dentro. */
export function linkInside(doc: DocumentModel, pos: Position): LinkRange | null {
  const link = linkAt(doc, pos);
  return link && link.start < pos.offset && pos.offset < link.end ? link : null;
}

// ---------------------------------------------------------------------------
// Punti del documento

export interface Place {
  readonly label: string;
  readonly url: string;
  readonly level: number;
}

/** Destinazioni interne: inizio del documento e titoli, come "Inserisci collegamento > Questo documento". */
export function documentPlaces(doc: DocumentModel): Place[] {
  return [
    { label: 'Inizio del documento', url: TOP_ANCHOR, level: 0 },
    ...collectHeadings(doc, 9).map((h) => ({ label: h.text, url: tocAnchor(h.paragraphId), level: h.level })),
  ];
}

/** Paragrafo a cui punta un collegamento interno, se esiste. */
export function placeTarget(doc: DocumentModel, url: string): Id | undefined {
  if (url.toLowerCase() === TOP_ANCHOR) return paragraphs(doc).next().value?.paragraph.id;
  const id = anchorTarget(url);
  return id && findP(doc, id) ? id : undefined;
}

/** Etichetta leggibile di un collegamento (per il suggerimento e la finestra). */
export function describeLink(doc: DocumentModel, url: string): string {
  if (linkKind(url) !== 'place') return displayText(url);
  const place = documentPlaces(doc).find((p) => p.url === url);
  return place ? place.label : url.slice(1);
}

// ---------------------------------------------------------------------------
// Riconoscimento automatico

const AUTO_URL = /^(?:(?:https?|ftp):\/\/|www\.)[^\s<>"]+$/i;
const TRAILING = /[.,;:!?)\]}'"»]+$/;

/**
 * Indirizzo web o email che termina in `end` nel paragrafo (subito prima di
 * uno spazio o di un a capo appena scritto), ancora senza collegamento.
 */
export function findAutoLink(p: Paragraph, end: number): { start: number; end: number; url: string } | null {
  const text = paragraphText(p);
  let start = end;
  while (start > 0 && !/\s/.test(text[start - 1]!) && text[start - 1] !== '￼') start--;
  let word = text.slice(start, end);
  const trailing = TRAILING.exec(word)?.[0].length ?? 0;
  word = word.slice(0, word.length - trailing);
  const stop = start + word.length;
  word = word.replace(/^[(\[{'"«]+/, (m) => {
    start += m.length;
    return '';
  });
  if (!word) return null;
  let url: string | null = null;
  if (AUTO_URL.test(word)) url = /^www\./i.test(word) ? `https://${word}` : word;
  else if (EMAIL.test(word)) url = `mailto:${word}`;
  else if (/^mailto:[^\s@]+@\S+$/i.test(word)) url = word;
  if (!url || /^(?:https?|ftp):\/\/$/i.test(word)) return null;
  // Solo testo semplice e non ancora collegato.
  let pos = 0;
  for (const inline of p.children) {
    const len = inlineLength(inline);
    if (pos < stop && start < pos + len && (inline.type !== 'text' || inline.format.link)) return null;
    pos += len;
  }
  return { start, end: stop, url };
}
