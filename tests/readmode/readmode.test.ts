// @vitest-environment happy-dom
import { afterEach, describe, expect, it } from 'vitest';
import { caret, createDocument, createState, Editor, paragraph, type Block } from '../../src/core';
import { createApproxMeasurer, createLayoutModule } from '../../src/layout';
import { createReadModeModule, mountReadMode, pageLabel, pagesPerScreen, ribbonItems, screenStart, clampReadZoom } from '../../src/readmode';
import * as readExports from '../../src/readmode';
import { collectModules, collectRibbonItems } from '../../src/ui/modules';
import { createUiModule } from '../../src/ui/module';

afterEach(() => document.body.replaceChildren());

const tick = () => new Promise((r) => setTimeout(r, 0));

function longDoc(n = 200) {
  const blocks: Block[] = Array.from({ length: n }, (_, i) => paragraph(`Paragrafo numero ${i + 1}`, { spaceAfter: 0 }));
  return createDocument(blocks);
}

function setup(doc = longDoc(), atEnd = false) {
  const state = createState(doc);
  const last = doc.sections[0]!.blocks.at(-1)!;
  const editor = new Editor(atEnd ? { ...state, selection: caret(last.id, 0) } : state);
  editor.use(createLayoutModule({ measurer: createApproxMeasurer() }));
  editor.use(createUiModule({ print: () => {} }));
  const read = createReadModeModule();
  editor.use(read);
  return { editor, read };
}

describe('calcoli', () => {
  it('allinea la pagina all’inizio della schermata', () => {
    expect(screenStart(3, 2, 10)).toBe(2);
    expect(screenStart(4, 2, 10)).toBe(4);
    expect(screenStart(20, 2, 10)).toBe(8);
    expect(screenStart(-1, 3, 10)).toBe(0);
  });

  it('limita lo zoom', () => {
    expect(clampReadZoom(10)).toBe(50);
    expect(clampReadZoom(500)).toBe(200);
    expect(clampReadZoom(Number.NaN)).toBe(100);
  });

  it('scrive la pagina come in Word', () => {
    expect(pageLabel(0, 1, 5)).toBe('Pagina 1 di 5');
    expect(pageLabel(2, 2, 5)).toBe('Pagine 3-4 di 5');
    expect(pageLabel(4, 2, 5)).toBe('Pagina 5 di 5');
  });

  it('affianca al massimo tre pagine', () => {
    expect(pagesPerScreen(500, 600)).toBe(1);
    expect(pagesPerScreen(1300, 600)).toBe(2);
    expect(pagesPerScreen(10000, 600)).toBe(3);
  });
});

describe('modulo', () => {
  it('si carica da solo con la sua voce in Visualizza', () => {
    const mods = collectModules([readExports]);
    expect(mods.map((m) => m.id)).toContain('readmode');
    expect(collectRibbonItems([readExports])).toEqual(ribbonItems);
    expect(ribbonItems[0]).toMatchObject({ tab: 'view', group: 'Visualizzazioni', commandId: 'readmode.open' });
  });

  it('apre, sfoglia e chiude senza toccare il documento né Annulla', async () => {
    const { editor, read } = setup();
    const doc = editor.state.doc;
    const total = read.documentLayout()!.pages.length;
    expect(total).toBeGreaterThan(3);
    await editor.executeAsync('readmode.open');
    expect(read.isOpen).toBe(true);
    expect(read.page).toBe(0);
    read.setPerScreen(2);
    await editor.executeAsync('readmode.next');
    expect(read.page).toBe(2);
    await editor.executeAsync('readmode.previous');
    expect(read.page).toBe(0);
    await editor.executeAsync('readmode.previous');
    expect(read.page).toBe(0);
    read.goTo(total + 5);
    expect(read.page).toBe(screenStart(total - 1, 2, total));
    await editor.executeAsync('readmode.zoom', { zoom: 130 });
    await editor.executeAsync('readmode.layout', { layout: 'scroll' });
    await editor.executeAsync('readmode.pageColor', { color: 'sepia' });
    expect([read.zoom, read.layout, read.pageColor]).toEqual([130, 'scroll', 'sepia']);
    await editor.executeAsync('readmode.close');
    expect(read.isOpen).toBe(false);
    expect(editor.state.doc).toBe(doc);
    expect(editor.canExecute('edit.undo')).toBe(false);
  });

  it('si apre alla pagina del cursore', async () => {
    const { editor, read } = setup(longDoc(), true);
    const layout = read.documentLayout()!;
    await editor.executeAsync('readmode.open');
    expect(read.page).toBe(layout.pages.length - 1);
  });
});

describe('schermata', () => {
  it('mostra le pagine, sfoglia con i tasti e chiude con Esc', async () => {
    const { editor, read } = setup();
    let focused = false;
    const off = mountReadMode({ editor, input: { focus: () => (focused = true) } }, read);
    const root = document.querySelector('.readmode') as HTMLElement;
    expect(root.hidden).toBe(true);

    await editor.executeAsync('readmode.open');
    expect(root.hidden).toBe(false);
    const shown = root.querySelectorAll('.print-page').length;
    expect(shown).toBe(read.perScreen);
    expect(root.querySelector('.readmode-status')!.textContent).toMatch(/^Pagin[ae] 1/);

    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true }));
    await tick();
    expect(read.page).toBe(read.perScreen);

    // Le lettere non arrivano al documento.
    const letter = new KeyboardEvent('keydown', { key: 'a', bubbles: true, cancelable: true });
    document.dispatchEvent(letter);
    expect(letter.defaultPrevented).toBe(true);

    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    await tick();
    expect(read.isOpen).toBe(false);
    expect(root.hidden).toBe(true);
    expect(focused).toBe(true);
    off();
    expect(document.querySelector('.readmode')).toBeNull();
  });

  it('con lo scorrimento mostra tutte le pagine e cambia colore', async () => {
    const { editor, read } = setup();
    mountReadMode({ editor }, read);
    await editor.executeAsync('readmode.layout', { layout: 'scroll' });
    await editor.executeAsync('readmode.pageColor', { color: 'inverse' });
    await editor.executeAsync('readmode.open');
    const root = document.querySelector('.readmode') as HTMLElement;
    expect(root.querySelectorAll('.print-page').length).toBe(read.documentLayout()!.pages.length);
    expect(root.dataset.layout).toBe('scroll');
    expect(root.dataset.pageColor).toBe('inverse');
    const checked = Array.from(root.querySelectorAll('.readmode-menu-item[aria-checked="true"]')).map((b) => b.textContent);
    expect(checked).toEqual(['✓Scorrimento verticale', '✓Inverso']);
  });
});
