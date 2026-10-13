/**
 * Riempitivi delle tabulazioni (puntini, trattini, linea), come nelle voci
 * del sommario. L'impaginazione dà solo la posizione della tabulazione: il
 * riempitivo si ricava dalle tabulazioni del paragrafo.
 */

import { paragraphs, resolveParaFormat, type DocumentModel, type GlyphRun, type Id, type TabStop } from '../core';

const LEADER_CHARS: Record<string, string> = { dot: '.', hyphen: '-', underscore: '_' };

export class TabLeaders {
  private tabs: Map<Id, readonly TabStop[]> | null = null;

  constructor(private readonly doc: DocumentModel) {}

  /**
   * Carattere di riempimento per una tabulazione che inizia a `start`
   * (punti dal bordo sinistro della colonna), o `undefined` se non ne ha.
   */
  leader(run: GlyphRun, start: number): string | undefined {
    if (run.kind !== 'tab') return undefined;
    if (!this.tabs) {
      this.tabs = new Map();
      for (const { paragraph: p } of paragraphs(this.doc)) {
        const tabs = resolveParaFormat(this.doc.styles, p).tabs;
        if (tabs?.some((t) => t.leader && t.leader !== 'none')) this.tabs.set(p.id, tabs);
      }
    }
    const tabs = this.tabs.get(run.paragraphId);
    if (!tabs) return undefined;
    // Stessa scelta dell'impaginazione: la prima tabulazione oltre l'inizio.
    const stop = tabs.filter((t) => t.position > start + 0.01).sort((a, b) => a.position - b.position)[0];
    return stop?.leader ? LEADER_CHARS[stop.leader] : undefined;
  }
}
