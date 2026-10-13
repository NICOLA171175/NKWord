/**
 * Visualizzazione delle revisioni ("tutti i commenti e le modifiche" di
 * Word): il testo inserito è sottolineato e quello eliminato barrato, nel
 * colore dell'autore. Non tocca il modulo di impaginazione né la UI: avvolge
 * il `LayoutEngine` pubblicato e ritocca la formattazione dei pezzi di
 * testo che coprono una revisione.
 */

import type { Block, CharFormat, DocumentLayout, DocumentModel, GlyphRun, Id, LayoutEngine, LineBox, Paragraph, Revision } from '../core';
import { paragraphs } from '../core';

/** Colori degli autori, come la tavolozza di Word. */
export const AUTHOR_COLORS = ['#b5082e', '#2e75b6', '#7030a0', '#00796b', '#c55a11', '#bf3f7f', '#4a7c1f', '#806000'] as const;

export function authorColor(author: string): string {
  let h = 0;
  for (let i = 0; i < author.length; i++) h = (h * 31 + author.charCodeAt(i)) | 0;
  return AUTHOR_COLORS[Math.abs(h) % AUTHOR_COLORS.length]!;
}

export function revisionFormat(format: CharFormat, revision: Revision): CharFormat {
  const color = authorColor(revision.author);
  return revision.kind === 'insert' ? { ...format, color, underline: 'single' } : { ...format, color, strike: true };
}

/** Revisione del contenuto all'offset dato (offset dei segni, come `Position`). */
function revisionAt(p: Paragraph, offset: number): Revision | undefined {
  let at = 0;
  for (const c of p.children) {
    const len = c.type === 'text' ? c.text.length : 1;
    if (offset < at + len) return c.revision;
    at += len;
  }
  return undefined;
}

function revisedParagraphs(doc: DocumentModel): Map<Id, Paragraph> {
  const out = new Map<Id, Paragraph>();
  for (const { paragraph: p } of paragraphs(doc)) {
    if (p.children.some((c) => c.revision)) out.set(p.id, p);
  }
  for (const s of doc.sections) {
    for (const part of [s.header, s.footer]) {
      for (const blocks of Object.values(part ?? {})) {
        for (const b of (blocks ?? []) as readonly Block[]) {
          if (b.type === 'paragraph' && b.children.some((c) => c.revision)) out.set(b.id, b);
        }
      }
    }
  }
  return out;
}

function markLines(lines: readonly LineBox[], revised: Map<Id, Paragraph>): readonly LineBox[] {
  let changed = false;
  const out = lines.map((line) => {
    const p = revised.get(line.paragraphId);
    if (!p) return line;
    let lineChanged = false;
    const runs = line.runs.map((run): GlyphRun => {
      if (run.kind === 'listMarker' || run.end <= run.start) return run;
      const rev = revisionAt(p, run.start);
      if (!rev) return run;
      lineChanged = true;
      return { ...run, format: revisionFormat(run.format, rev) };
    });
    if (!lineChanged) return line;
    changed = true;
    return { ...line, runs };
  });
  return changed ? out : lines;
}

export function markRevisions(layout: DocumentLayout): DocumentLayout {
  const revised = revisedParagraphs(layout.doc);
  if (!revised.size) return layout;
  return {
    ...layout,
    pages: layout.pages.map((page) => ({
      ...page,
      lines: markLines(page.lines, revised),
      header: markLines(page.header, revised),
      footer: markLines(page.footer, revised),
    })),
  };
}

/** Avvolge un motore di impaginazione; `enabled()` false mostra il documento senza evidenziazioni. */
export function withRevisionMarkup(engine: LayoutEngine, enabled: () => boolean = () => true): LayoutEngine {
  return new Proxy(engine, {
    get(target, prop, receiver) {
      if (prop === 'layout') {
        return (doc: DocumentModel) => {
          const layout = target.layout(doc);
          return enabled() ? markRevisions(layout) : layout;
        };
      }
      const value = Reflect.get(target, prop, receiver);
      return typeof value === 'function' ? value.bind(target) : value;
    },
  });
}
