import {
  createDocument,
  paragraph,
  type Block,
  type DocumentLayout,
  type DocumentModel,
  type LineBox,
  type Section,
} from '../../src/core';
import { createApproxMeasurer, createLayoutEngine } from '../../src/layout';

/** Misuratore deterministico: carattere = 0,5 × corpo, spazio = 0,25 × corpo. */
export const measurer = createApproxMeasurer();
export const engine = createLayoutEngine(measurer, { now: () => new Date(2026, 9, 12, 9, 30) });

export function doc(blocks: Block[], section: Partial<Section> = {}): DocumentModel {
  const d = createDocument(blocks);
  return { ...d, sections: [{ ...d.sections[0]!, ...section }] };
}

/** Paragrafi con testo e senza spaziatura, per conti semplici. */
export function plain(text: string, format = {}) {
  return paragraph(text, { spaceAfter: 0, lineSpacing: { rule: 'multiple', value: 1 }, ...format });
}

export function lineText(l: LineBox): string {
  return l.runs.filter((r) => r.kind !== 'listMarker').map((r) => r.text).join('');
}

export function pageTexts(layout: DocumentLayout): string[][] {
  return layout.pages.map((p) => p.lines.map(lineText));
}
