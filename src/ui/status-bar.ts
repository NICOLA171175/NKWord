/** Barra di stato: pagine, parole, zoom. */

import { paragraphText, paragraphs, type DocumentModel, type Editor } from '../core';
import { runFirstExecutable } from './input';
import { MAX_ZOOM, MIN_ZOOM, type ViewStore } from './view-state';

export function countWords(doc: DocumentModel): number {
  let n = 0;
  for (const { paragraph } of paragraphs(doc)) {
    n += paragraphText(paragraph).split(/\s+/).filter((w) => /[\p{L}\p{N}]/u.test(w)).length;
  }
  return n;
}

export class StatusBar {
  readonly element: HTMLElement;
  private readonly pages: HTMLElement;
  private readonly words: HTMLElement;
  private readonly zoom: HTMLInputElement;
  private readonly zoomLabel: HTMLElement;
  private readonly cleanup: (() => void)[] = [];
  private lastDoc: DocumentModel | null = null;

  constructor(
    private readonly editor: Editor,
    private readonly view: ViewStore,
  ) {
    this.element = document.createElement('footer');
    this.element.className = 'status-bar';
    this.pages = document.createElement('span');
    this.pages.className = 'status-pages';
    this.words = document.createElement('span');
    this.words.className = 'status-words';
    const spacer = document.createElement('span');
    spacer.className = 'status-spacer';
    this.zoom = document.createElement('input');
    this.zoom.type = 'range';
    this.zoom.className = 'status-zoom';
    this.zoom.min = String(MIN_ZOOM * 100);
    this.zoom.max = String(MAX_ZOOM * 100);
    this.zoom.step = '10';
    this.zoom.setAttribute('aria-label', 'Zoom');
    this.zoom.addEventListener('input', () => {
      runFirstExecutable(editor, [{ commandId: 'view.zoom', args: { zoom: Number(this.zoom.value) / 100 } }]);
    });
    this.zoomLabel = document.createElement('span');
    this.zoomLabel.className = 'status-zoom-label';
    this.element.append(this.pages, this.words, spacer, this.zoom, this.zoomLabel);
    this.cleanup.push(
      editor.events.on('change', () => this.refresh()),
      view.events.on('change', () => this.refresh()),
    );
    this.refresh();
  }

  refresh(): void {
    const { doc } = this.editor.state;
    if (doc !== this.lastDoc) {
      this.lastDoc = doc;
      const layout = this.editor.services.layout;
      const pages = layout ? layout.layout(doc).pages.length : doc.sections.length;
      this.pages.textContent = `Pagine: ${pages}`;
      const words = countWords(doc);
      this.words.textContent = `${words} ${words === 1 ? 'parola' : 'parole'}`;
    }
    const z = Math.round(this.view.state.zoom * 100);
    this.zoom.value = String(z);
    this.zoomLabel.textContent = `${z}%`;
  }

  destroy(): void {
    for (const c of this.cleanup.splice(0)) c();
    this.element.remove();
  }
}
