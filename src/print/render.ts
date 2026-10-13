/**
 * Disegna una pagina impaginata come elementi HTML, per l'anteprima e per
 * la stampa. Usa la stessa geometria e gli stessi stili di carattere delle
 * pagine a schermo, ma senza cursore, selezione, segni di correzione o
 * sfondo dei campi: solo quello che finisce sulla carta.
 */

import { ptToPx, type DocumentModel, type GlyphRun, type LineBox, type PageLayout } from '../core';
import { charStyle } from '../ui/document-view';
import { TabLeaders } from '../ui/tab-leaders';

export interface PageRenderOptions {
  /** 1 = dimensioni reali (stampa); meno di 1 per le miniature. */
  readonly zoom: number;
  readonly doc: DocumentModel;
  /** Riempitivi delle tabulazioni, condivisi fra le pagine dello stesso documento. */
  readonly leaders?: TabLeaders;
  readonly document?: Document;
}

export function renderPage(page: PageLayout, options: PageRenderOptions): HTMLElement {
  const d = options.document ?? document;
  const { zoom, doc } = options;
  const leaders = options.leaders ?? new TabLeaders(doc);
  const px = (pt: number) => `${ptToPx(pt, zoom)}px`;
  const box = (className: string, x: number, y: number, w: number, h: number) => {
    const e = d.createElement('div');
    e.className = className;
    Object.assign(e.style, { left: px(x), top: px(y), width: px(w), height: px(h) });
    return e;
  };

  const el = d.createElement('div');
  el.className = 'print-page';
  el.dataset.page = String(page.index);
  el.style.width = px(page.width);
  el.style.height = px(page.height);

  for (const dec of page.decorations) {
    const b = box(`print-decoration print-decoration-${dec.kind}`, dec.rect.x, dec.rect.y, dec.rect.width, dec.rect.height);
    b.style.background = dec.color;
    el.append(b);
  }

  const run = (r: GlyphRun, line: LineBox): HTMLElement => {
    let e: HTMLElement;
    if (r.kind === 'image' && r.src) {
      const img = d.createElement('img');
      img.src = doc.media[r.src] ?? r.src;
      img.alt = '';
      img.style.height = px(r.height ?? line.height);
      e = img;
    } else {
      e = d.createElement('span');
      const leader = leaders.leader(r, line.x - page.contentRect.x + r.x);
      if (leader) {
        e.textContent = leader.repeat(Math.ceil(r.width / 2) + 1);
        e.classList.add('print-leader');
      } else e.textContent = r.kind === 'tab' ? '' : r.text;
      Object.assign(e.style, charStyle(r.format, zoom));
      if (r.format.link && !r.format.color) e.style.color = '#0563c1';
      if (r.format.link) e.style.textDecorationLine = 'underline';
    }
    e.classList.add('print-run');
    e.style.left = px(r.x);
    e.style.width = px(r.width);
    return e;
  };

  for (const lines of [page.header, page.lines, page.footnotes ?? [], page.footer]) {
    for (const line of lines) {
      const l = box('print-line', line.x, line.y, line.width, line.height);
      for (const r of line.runs) l.append(run(r, line));
      el.append(l);
    }
  }
  return el;
}
