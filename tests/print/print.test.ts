// @vitest-environment happy-dom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { caret, createDocument, createState, Editor, paragraph, type Block } from '../../src/core';
import { createApproxMeasurer, createLayoutModule } from '../../src/layout';
import {
  buildPrintOutput,
  createDomPrinter,
  createPrintModule,
  mountPrint,
  parsePageRanges,
  PRINT_ACTIVE_CLASS,
  selectPages,
  withOrientation,
  DEFAULT_SETTINGS,
  type PrintJob,
} from '../../src/print';
import { collectModules } from '../../src/ui/modules';
import { createUiModule } from '../../src/ui/module';
import * as printExports from '../../src/print';

afterEach(() => document.body.replaceChildren());

const tick = () => new Promise((r) => setTimeout(r, 0));

function longDoc(n = 200) {
  const blocks: Block[] = Array.from({ length: n }, (_, i) => paragraph(`Paragrafo numero ${i + 1}`, { spaceAfter: 0 }));
  return createDocument(blocks);
}

function setup(doc = longDoc()) {
  const editor = new Editor(createState(doc));
  editor.use(createLayoutModule({ measurer: createApproxMeasurer() }));
  const jobs: PrintJob[] = [];
  const print = createPrintModule({ printer: async (job) => void jobs.push(job) });
  editor.use(createUiModule({ print: () => {} }));
  editor.use(print);
  return { editor, print, jobs };
}

describe('intervalli di pagine', () => {
  it('legge numeri e intervalli come in Word', () => {
    expect(parsePageRanges('1-3, 5', 6)).toEqual({ pages: [0, 1, 2, 4] });
    expect(parsePageRanges(' 4 ; 2 - 3 ', 6)).toEqual({ pages: [3, 1, 2] });
    expect(parsePageRanges('5-', 6)).toEqual({ pages: [4, 5] });
    expect(parsePageRanges('3-1', 6)).toEqual({ pages: [2, 1, 0] });
    expect(parsePageRanges('2,2,1-2', 6)).toEqual({ pages: [1, 0] });
  });

  it('spiega gli errori', () => {
    expect(parsePageRanges('', 3)).toHaveProperty('error');
    expect(parsePageRanges('a', 3)).toEqual({ error: '"a" non è un numero di pagina valido.' });
    expect(parsePageRanges('2-7', 3)).toEqual({ error: 'La pagina 7 non esiste: il documento ha 3 pagine.' });
    expect(parsePageRanges('0', 1)).toEqual({ error: 'La pagina 0 non esiste: il documento ha 1 pagina.' });
  });

  it('sceglie tutte le pagine, la corrente o un intervallo', () => {
    expect(selectPages(DEFAULT_SETTINGS, 3, 1)).toEqual({ pages: [0, 1, 2] });
    expect(selectPages({ ...DEFAULT_SETTINGS, pages: 'current' }, 3, 1)).toEqual({ pages: [1] });
    expect(selectPages({ ...DEFAULT_SETTINGS, pages: 'custom', range: '3' }, 3, 1)).toEqual({ pages: [2] });
  });
});

describe('orientamento', () => {
  it('scambia larghezza e altezza delle sezioni da girare', () => {
    const doc = createDocument();
    const { width, height } = doc.sections[0]!.page;
    expect(withOrientation(doc, null)).toBe(doc);
    expect(withOrientation(doc, 'portrait')).toBe(doc);
    const turned = withOrientation(doc, 'landscape').sections[0]!.page;
    expect(turned).toMatchObject({ orientation: 'landscape', width: height, height: width });
    expect(withOrientation(withOrientation(doc, 'landscape'), 'portrait').sections[0]!.page).toMatchObject({ width, height });
  });
});

describe('modulo stampa', () => {
  it('Stampa (Ctrl+P) apre l\'anteprima invece di stampare subito', async () => {
    const ui = vi.fn();
    const editor = new Editor();
    editor.use(createLayoutModule({ measurer: createApproxMeasurer() }));
    editor.use(createUiModule({ print: ui }));
    const print = createPrintModule({ printer: async () => {} });
    editor.use(print);
    await editor.executeAsync('view.print');
    expect(print.isOpen).toBe(true);
    expect(ui).not.toHaveBeenCalled();
    await editor.executeAsync('print.close');
    expect(print.isOpen).toBe(false);
  });

  it('stampa le pagine scelte con le copie richieste', async () => {
    const { editor, print, jobs } = setup();
    const p = print.preview();
    if ('error' in p) throw new Error(p.error);
    const total = p.layout.pages.length;
    expect(total).toBeGreaterThan(2);
    expect(p.pages).toHaveLength(total);

    print.update({ pages: 'custom', range: '1, 3', copies: 2 });
    const q = print.preview();
    expect(q).toMatchObject({ pages: [0, 2], sheets: 4 });

    print.open();
    await editor.executeAsync('print.print');
    expect(jobs).toHaveLength(1);
    expect(jobs[0]).toMatchObject({ pages: [0, 2], copies: 2 });
    expect(print.isOpen).toBe(false);
  });

  it('non stampa con un intervallo sbagliato', async () => {
    const { editor, print, jobs } = setup();
    print.update({ pages: 'custom', range: '999' });
    expect(print.preview()).toHaveProperty('error');
    expect(editor.canExecute('print.print')).toBe(false);
    expect(await print.print()).toBe(false);
    expect(jobs).toHaveLength(0);
  });

  it('la pagina corrente è quella del cursore', () => {
    const doc = longDoc();
    const last = doc.sections[0]!.blocks.at(-1)!;
    const editor = new Editor({ ...createState(doc), selection: caret(last.id, 0) });
    editor.use(createLayoutModule({ measurer: createApproxMeasurer() }));
    const print = createPrintModule({ printer: async () => {} });
    editor.use(print);
    print.update({ pages: 'current' });
    const p = print.preview();
    if ('error' in p) throw new Error(p.error);
    expect(p.pages).toEqual([p.layout.pages.length - 1]);
  });

  it('impagina in orizzontale senza toccare il documento', () => {
    const { editor, print } = setup();
    const doc = editor.state.doc;
    print.update({ orientation: 'landscape' });
    const p = print.preview();
    if ('error' in p) throw new Error(p.error);
    expect(p.layout.pages[0]!.width).toBeGreaterThan(p.layout.pages[0]!.height);
    expect(editor.state.doc).toBe(doc);
  });

  it('limita le copie fra 1 e 999', () => {
    const { print } = setup();
    print.update({ copies: 0 });
    expect(print.settings.copies).toBe(1);
    print.update({ copies: 5000 });
    expect(print.settings.copies).toBe(999);
  });

  it('è caricato automaticamente dalla UI', () => {
    const modules = collectModules([printExports as Record<string, unknown>]);
    expect(modules.map((m) => m.id)).toEqual(['print']);
  });
});

describe('schermata Stampa', () => {
  it('mostra le pagine affiancate e reagisce alle impostazioni', async () => {
    const { editor, print, jobs } = setup();
    const focus = vi.fn();
    const off = mountPrint({ editor, input: { focus } }, print);
    const root = document.querySelector<HTMLElement>('.print-backstage')!;
    expect(root.hidden).toBe(true);

    await editor.executeAsync('view.print');
    expect(root.hidden).toBe(false);
    const total = print.preview().layout!.pages.length;
    expect(root.querySelectorAll('.print-sheet')).toHaveLength(total);
    expect(root.querySelector('.print-caption')!.textContent).toBe('Pagina 1');
    expect(root.querySelector('.print-page')!.textContent).toContain('Paragrafo numero 1');

    // Scrivere un intervallo passa a "Intervallo personalizzato".
    const range = root.querySelector<HTMLInputElement>('.print-range')!;
    range.value = '2';
    range.dispatchEvent(new Event('input'));
    expect(print.settings.pages).toBe('custom');
    expect(root.querySelectorAll('.print-sheet')).toHaveLength(1);
    expect(root.querySelector('.print-caption')!.textContent).toBe('Pagina 2');

    range.value = 'x';
    range.dispatchEvent(new Event('input'));
    expect(root.querySelector('.print-error')!.textContent).toContain('non è un numero di pagina');
    expect(root.querySelector<HTMLButtonElement>('.print-button')!.disabled).toBe(true);

    range.value = '1-2';
    range.dispatchEvent(new Event('input'));
    const copies = root.querySelector<HTMLInputElement>('.print-copies')!;
    copies.value = '3';
    copies.dispatchEvent(new Event('change'));
    expect(root.querySelector('.print-summary')!.textContent).toBe(`2 pagine di ${total}, 3 copie: 6 pagine stampate`);

    root.querySelector<HTMLButtonElement>('.print-button')!.click();
    await tick();
    expect(jobs[0]).toMatchObject({ pages: [0, 1], copies: 3 });
    expect(root.hidden).toBe(true);
    expect(focus).toHaveBeenCalled();

    // Esc chiude.
    print.open();
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
    await tick();
    expect(print.isOpen).toBe(false);
    off();
    expect(document.querySelector('.print-backstage')).toBeNull();
  });
});

describe('stampa nel browser', () => {
  it('prepara le pagine a grandezza reale e la misura del foglio', () => {
    const { print } = setup();
    const p = print.preview();
    if ('error' in p) throw new Error(p.error);
    const { root, style } = buildPrintOutput({ layout: p.layout, pages: [1, 0], copies: 2 });
    const pages = Array.from(root.querySelectorAll<HTMLElement>('.print-page'));
    expect(pages.map((e) => e.dataset.page)).toEqual(['1', '0', '1', '0']);
    expect(parseFloat(pages[0]!.style.width)).toBeCloseTo((p.layout.pages[0]!.width * 96) / 72, 3);
    expect(style.textContent).toMatch(/@page \{ size: [\d.]+px [\d.]+px; margin: 0; \}/);
  });

  it('mostra solo le pagine durante la stampa e poi pulisce', async () => {
    const { print } = setup();
    const p = print.preview();
    if ('error' in p) throw new Error(p.error);
    let seen = 0;
    const printer = createDomPrinter(window, () => {
      seen = document.querySelectorAll('.print-output .print-page').length;
      expect(document.body.classList.contains(PRINT_ACTIVE_CLASS)).toBe(true);
    });
    await printer({ layout: p.layout, pages: [0], copies: 1 });
    expect(seen).toBe(1);
    window.dispatchEvent(new Event('afterprint'));
    expect(document.querySelector('.print-output')).toBeNull();
    expect(document.body.classList.contains(PRINT_ACTIVE_CLASS)).toBe(false);
  });
});
