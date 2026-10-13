/**
 * Stili del sommario. Seguono i nomi interni di Word ("toc 1".."toc 9" e
 * "TOC Heading"), così un sommario importato da DOCX viene riconosciuto
 * anche quando l'id dello stile è localizzato (es. "Sommario1").
 */

import { STYLE_NORMAL, type Id, type Style, type StyleSheet } from '../core';

export const TOC_HEADING_STYLE = 'TOCHeading';
export const MAX_TOC_LEVEL = 9;

export function tocEntryStyleId(level: number): Id {
  return `TOC${level}`;
}

/** Livello (1..9) se lo stile è una voce di sommario, altrimenti 0. */
export function tocEntryLevel(sheet: StyleSheet, styleId: Id | undefined): number {
  if (!styleId) return 0;
  const byId = /^TOC([1-9])$/.exec(styleId);
  if (byId) return Number(byId[1]);
  const byName = /^toc ([1-9])$/i.exec(sheet.styles[styleId]?.name ?? '');
  return byName ? Number(byName[1]) : 0;
}

export function isTocHeadingStyle(sheet: StyleSheet, styleId: Id | undefined): boolean {
  if (!styleId) return false;
  return styleId === TOC_HEADING_STYLE || /^toc heading$/i.test(sheet.styles[styleId]?.name ?? '');
}

/** Il paragrafo con questo stile fa parte di un sommario (titolo o voce). */
export function isTocStyle(sheet: StyleSheet, styleId: Id | undefined): boolean {
  return tocEntryLevel(sheet, styleId) > 0 || isTocHeadingStyle(sheet, styleId);
}

function entryStyle(level: number): Style {
  return {
    id: tocEntryStyleId(level),
    name: `toc ${level}`,
    type: 'paragraph',
    basedOn: STYLE_NORMAL,
    next: STYLE_NORMAL,
    paraFormat: { spaceAfter: 5, ...(level > 1 ? { indentLeft: (level - 1) * 11 } : {}) },
  };
}

const headingStyle: Style = {
  id: TOC_HEADING_STYLE,
  name: 'TOC Heading',
  type: 'paragraph',
  basedOn: STYLE_NORMAL,
  next: STYLE_NORMAL,
  paraFormat: { spaceBefore: 12, spaceAfter: 6, keepWithNext: true },
  charFormat: { fontSize: 16, color: '#2f5496' },
};

/** Aggiunge gli stili del sommario mancanti (fino a `maxLevel`); non tocca quelli esistenti, anche se con un altro id. */
export function ensureTocStyles(sheet: StyleSheet, maxLevel: number): StyleSheet {
  const ids = Object.keys(sheet.styles);
  const missing: Style[] = [];
  if (!ids.some((id) => isTocHeadingStyle(sheet, id))) missing.push(headingStyle);
  for (let level = 1; level <= maxLevel; level++) {
    if (!ids.some((id) => tocEntryLevel(sheet, id) === level)) missing.push(entryStyle(level));
  }
  if (!missing.length) return sheet;
  const styles = { ...sheet.styles };
  for (const s of missing) styles[s.id] = s;
  return { ...sheet, styles };
}
