// @vitest-environment happy-dom
import { afterEach, describe, expect, it } from 'vitest';
import {
  Editor,
  createDocument,
  createState,
  paragraph,
  textRun,
  type CharFormat,
  type DocumentLayout,
  type LayoutEngine,
  type LineBox,
} from '../../src/core';
import { DocumentView, runVerticalStyle, selectionByParagraph } from '../../src/ui/document-view';
import { ViewStore } from '../../src/ui/view-state';

afterEach(() => document.body.replaceChildren());

function sample() {
  const a = paragraph([textRun('Ciao '), textRun('mondo', { bold: true })]);
  const b = paragraph('seconda riga');
  const doc = createDocument([a, b]);
  return { a, b, doc };
}

describe('selectionByParagraph', () => {
  it('calcola gli intervalli selezionati per paragrafo, in ordine di documento', () => {
    const { a, b, doc } = sample();
    const sel = { anchor: { paragraphId: b.id, offset: 3 }, focus: { paragraphId: a.id, offset: 2 } };
    const map = selectionByParagraph(doc, sel);
    expect(map.get(a.id)).toEqual([2, 11]); // fino al segno di paragrafo incluso
    expect(map.get(b.id)).toEqual([0, 3]);
  });
});

describe('DocumentView senza impaginazione', () => {
  it('disegna i paragrafi con formattazione, cursore e selezione', () => {
    const { a, doc } = sample();
    const editor = new Editor({ ...createState(doc), selection: { anchor: { paragraphId: a.id, offset: 1 }, focus: { paragraphId: a.id, offset: 3 } } });
    const view = new DocumentView(editor, new ViewStore());
    document.body.append(view.element);
    const paras = view.element.querySelectorAll('.para');
    expect(paras).toHaveLength(2);
    expect(paras[0]!.textContent).toBe('Ciao mondo​');
    const bold = Array.from(paras[0]!.querySelectorAll<HTMLElement>('.run-text')).find((s) => s.textContent === 'mondo')!;
    expect(bold.style.fontWeight).toBe('bold');
    expect(bold.dataset.start).toBe('5');
    const selected = Array.from(paras[0]!.querySelectorAll('.selected')).map((s) => s.textContent);
    expect(selected).toEqual(['ia']);
    expect(paras[0]!.querySelectorAll('.caret')).toHaveLength(1);
  });

  it('mostra i segni di formattazione e attenua il corpo in modifica intestazione', () => {
    const { doc } = sample();
    const editor = new Editor(createState(doc));
    const store = new ViewStore();
    const view = new DocumentView(editor, store);
    store.update({ formattingMarks: true, headerFooter: 'header' });
    const pages = view.element.querySelector<HTMLElement>('.doc-pages')!;
    expect(pages.classList.contains('show-marks')).toBe(true);
    expect(pages.dataset.headerFooter).toBe('header');
    expect(view.element.querySelector('.para')!.textContent).toBe('Ciao·mondo¶');
    // In modifica intestazione il cursore del corpo non si vede.
    expect(view.element.querySelectorAll('.caret')).toHaveLength(0);
  });

  it('applica lo zoom alle dimensioni della pagina', () => {
    const editor = new Editor();
    const store = new ViewStore({ zoom: 2 });
    const view = new DocumentView(editor, store);
    const page = view.element.querySelector<HTMLElement>('.page')!;
    // A4 = 210 mm = 793.7 px a 96 dpi, al 200%.
    expect(parseFloat(page.style.width)).toBeCloseTo(1587.4, 0);
  });

  it('un clic sul segno di paragrafo sposta il cursore con edit.setSelection', () => {
    const { a, doc } = sample();
    const editor = new Editor(createState(doc));
    const got: unknown[] = [];
    editor.commands.register({ id: 'edit.setSelection', label: 'Selezione', run: (_c, args) => (got.push(args), null) });
    const view = new DocumentView(editor, new ViewStore());
    document.body.append(view.element);
    const mark = view.element.querySelector<HTMLElement>('.para-mark')!;
    mark.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }));
    expect(got).toEqual([{ selection: { anchor: { paragraphId: a.id, offset: 10 }, focus: { paragraphId: a.id, offset: 10 } } }]);
  });
});

describe('DocumentView con impaginazione', () => {
  it('disegna le pagine del LayoutEngine e usa hitTest per i clic', () => {
    const { a, doc } = sample();
    const hits: [number, number, number][] = [];
    const engine: LayoutEngine = {
      layout: (d): DocumentLayout => ({
        doc: d,
        pages: [0, 1].map((i) => ({
          index: i,
          number: i + 1,
          sectionId: d.sections[0]!.id,
          width: 600,
          height: 800,
          contentRect: { x: 50, y: 50, width: 500, height: 700 },
          lines: [{
            paragraphId: a.id, start: 0, end: 5, x: 50, y: 50, width: 500, height: 14, baseline: 11,
            runs: [{ kind: 'text', paragraphId: a.id, start: 0, end: 5, text: 'Ciao ', format: { bold: true }, x: 0, width: 30 }],
          }],
          decorations: [{ kind: 'border', rect: { x: 50, y: 49.75, width: 500, height: 0.5 }, color: '#ff0000', width: 0.5 }],
          header: [],
          footer: [],
        })),
      }),
      caretRect: () => ({ page: 1, rect: { x: 50, y: 50, width: 0, height: 14 } }),
      hitTest: (_l, page, x, y) => (hits.push([page, x, y]), { paragraphId: a.id, offset: 2 }),
      selectionRects: () => [],
    };
    const editor = new Editor(createState(doc));
    editor.services.layout = engine;
    const got: unknown[] = [];
    editor.commands.register({ id: 'edit.setSelection', label: 'Selezione', run: (_c, args) => (got.push(args), null) });
    const view = new DocumentView(editor, new ViewStore());
    document.body.append(view.element);

    const pages = view.element.querySelectorAll<HTMLElement>('.page-laid-out');
    expect(pages).toHaveLength(2);
    expect(pages[0]!.style.width).toBe('800px'); // 600 pt
    expect(pages[0]!.querySelector('.run')!.textContent).toBe('Ciao ');
    expect(pages[1]!.querySelector('.caret')).not.toBeNull();
    const border = pages[0]!.querySelector<HTMLElement>('.decoration-border')!;
    expect(border.style.background).toContain('#ff0000');
    expect(border.style.border).toBe('');
    expect(parseFloat(border.style.height)).toBeCloseTo(0.667, 2);

    pages[1]!.querySelector('.run')!.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0, clientX: 0, clientY: 0 }));
    expect(hits[0]![0]).toBe(1);
    expect(got).toEqual([{ selection: { anchor: { paragraphId: a.id, offset: 2 }, focus: { paragraphId: a.id, offset: 2 } } }]);
  });
});

describe('posizione verticale dei run nella pagina impaginata', () => {
  const line = { baseline: 12 } as LineBox;
  const top = (f: CharFormat) => parseFloat(runVerticalStyle(f, line, 1).top);
  const bottom = (f: CharFormat) => {
    const v = runVerticalStyle(f, line, 1);
    return parseFloat(v.top) + parseFloat(v.lineHeight);
  };

  it('il pedice scende sotto la linea di base, non resta in cima alla riga come un apice', () => {
    const normale = { fontSize: 11 };
    const pedice = { fontSize: 11, verticalAlign: 'subscript' } as const;
    const apice = { fontSize: 11, verticalAlign: 'superscript' } as const;
    expect(bottom(pedice)).toBeGreaterThan(bottom(normale));
    expect(top(apice)).toBeLessThan(top(normale));
    expect(top(pedice)).toBeGreaterThan(top(apice));
  });

  it('caratteri di dimensioni diverse condividono la stessa linea di base', () => {
    const base = (f: CharFormat) => parseFloat(runVerticalStyle(f, line, 1).top) + 0.8 * parseFloat(runVerticalStyle(f, line, 1).lineHeight);
    expect(base({ fontSize: 24 })).toBeCloseTo(base({ fontSize: 11 }), 5);
  });

  it('anche senza dimensione esplicita apice e pedice sono più piccoli', () => {
    const view = runVerticalStyle({ verticalAlign: 'subscript' }, line, 1);
    expect(parseFloat(view.lineHeight)).toBeLessThan(parseFloat(runVerticalStyle({}, line, 1).lineHeight));
  });
});
