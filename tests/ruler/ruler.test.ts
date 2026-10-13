// @vitest-environment happy-dom
import { afterEach, describe, expect, it } from 'vitest';
import {
  caret,
  cmToPt,
  createState,
  Editor,
  firstParagraph,
  paragraph,
  ptToPx,
  textRun,
  type DocumentModel,
  type EditorState,
  type Paragraph,
} from '../../src/core';
import { editingModule } from '../../src/editing';
import { exportDocx, importDocx } from '../../src/io';
import { createApproxMeasurer, createLayoutModule } from '../../src/layout';
import { changeTabs, createRulerModule, effectiveIndents, ribbonItems, setIndents, setTabs, snap } from '../../src/ruler';
import * as rulerExports from '../../src/ruler';
import { structuresModule } from '../../src/structures';
import { collectModules, collectRibbonItems } from '../../src/ui/modules';
import { createApp } from '../../src/ui/shell';
import { doc, engine, plain } from '../layout/helpers';

afterEach(() => document.body.replaceChildren());

const first = (d: DocumentModel) => firstParagraph(d);

class MapStorage {
  readonly data = new Map<string, string>();
  getItem(k: string) {
    return this.data.get(k) ?? null;
  }
  setItem(k: string, v: string) {
    this.data.set(k, v);
  }
}

function stateWith(...ps: Paragraph[]): EditorState {
  const d = doc(ps);
  const s = createState(d);
  return { ...s, selection: { anchor: { paragraphId: ps[0]!.id, offset: 0 }, focus: { paragraphId: ps[ps.length - 1]!.id, offset: 0 } } };
}

describe('rientri e tabulazioni (modello)', () => {
  it('aggancia a 0,25 cm', () => {
    expect(snap(cmToPt(1.1))).toBeCloseTo(cmToPt(1), 1);
    expect(snap(cmToPt(1.2))).toBeCloseTo(cmToPt(1.25), 1);
  });

  it('imposta i rientri su tutti i paragrafi selezionati', () => {
    const s = stateWith(plain('uno'), plain('due'));
    const tr = setIndents(s, { left: 36, firstLine: -18, right: 10 })!;
    for (const p of tr.state.doc.sections[0]!.blocks as Paragraph[]) {
      expect(p.format).toMatchObject({ indentLeft: 36, indentFirstLine: -18, indentRight: 10 });
    }
    expect(setIndents(tr.state, { left: 36 })).toBeNull();
  });

  it('su una voce di elenco parte dai rientri del livello', () => {
    const editor = new Editor();
    editor.use(structuresModule);
    editor.reset(stateWith(plain('voce')));
    editor.execute('list.toggleNumbered', {});
    const s = editor.state;
    expect(first(s.doc).format.list).toBeTruthy();
    const before = effectiveIndents(s.doc, first(s.doc));
    const tr = setIndents(s, { left: before.left + 20 })!;
    const after = effectiveIndents(tr.state.doc, first(tr.state.doc));
    expect(after.left).toBeCloseTo(before.left + 20);
    expect(after.firstLine).toBeCloseTo(before.firstLine);
  });

  it('aggiunge, sposta e toglie tabulazioni', () => {
    let tabs = changeTabs([], { add: { position: 72, align: 'center' } });
    tabs = changeTabs(tabs, { add: { position: 36, align: 'left' } });
    expect(tabs.map((t) => t.position)).toEqual([36, 72]);
    tabs = changeTabs(tabs, { move: { from: 72, to: 144 } });
    expect(tabs).toEqual([{ position: 36, align: 'left' }, { position: 144, align: 'center' }]);
    tabs = changeTabs(tabs, { remove: 36 });
    expect(tabs).toEqual([{ position: 144, align: 'center' }]);
    const s = stateWith(plain('a'));
    const tr = setTabs(s, { add: { position: 100, align: 'decimal' } })!;
    expect(first(tr.state.doc).format.tabs).toEqual([{ position: 100, align: 'decimal' }]);
    expect(setTabs(tr.state, { add: { position: 100, align: 'decimal' } })).toBeNull();
  });
});

describe('Tab nel testo', () => {
  const tabbed = (text: string, align: 'left' | 'center' | 'right' | 'decimal'): Paragraph => ({
    ...plain('', { tabs: [{ position: 200, align }] }),
    children: [{ type: 'tab' }, textRun(text)],
  });
  const textX = (p: Paragraph) => {
    const layout = engine.layout(doc([p]));
    const line = layout.pages[0]!.lines[0]!;
    const run = line.runs.find((r) => r.kind === 'text')!;
    return { x: run.x, width: run.width, left: layout.pages[0]!.contentRect.x, line };
  };

  it('salta alla tabulazione sinistra, destra e centrata', () => {
    expect(textX(tabbed('12', 'left')).x).toBeCloseTo(200, 0);
    const r = textX(tabbed('12', 'right'));
    expect(r.x + r.width).toBeCloseTo(200, 0);
    const c = textX(tabbed('12', 'center'));
    expect(c.x + c.width / 2).toBeCloseTo(200, 0);
  });

  it('la tabulazione decimale allinea la virgola', () => {
    const m = createApproxMeasurer();
    for (const value of ['3,5', '1234,75', '7']) {
      const { x } = textX(tabbed(value, 'decimal'));
      const intPart = value.split(',')[0]!;
      const w = m.measure(intPart, {}).width;
      // Senza virgola il numero finisce sulla tabulazione (come a destra).
      const expected = value.includes(',') ? 200 - w : 200 - m.measure(value, {}).width;
      expect(x).toBeCloseTo(expected, 0);
    }
  });
});

describe('DOCX', () => {
  it('salva e riapre rientri e tabulazioni', async () => {
    const p = paragraph('testo', {
      indentLeft: 36,
      indentRight: 18,
      indentFirstLine: -14.2,
      tabs: [
        { position: 56.7, align: 'left' },
        { position: 113.4, align: 'center' },
        { position: 170.1, align: 'right' },
        { position: 226.8, align: 'decimal' },
      ],
    });
    const back = first(await importDocx(await exportDocx(doc([p]))));
    expect(back.format.indentLeft).toBeCloseTo(36);
    expect(back.format.indentRight).toBeCloseTo(18);
    expect(back.format.indentFirstLine).toBeCloseTo(-14.2, 1);
    expect(back.format.tabs!.map((t) => t.align)).toEqual(['left', 'center', 'right', 'decimal']);
    back.format.tabs!.forEach((t, i) => expect(t.position).toBeCloseTo(p.format.tabs![i]!.position, 0));
  });
});

describe('righello nell’app', () => {
  function setup(storage = new MapStorage()) {
    const editor = new Editor();
    editor.use(createLayoutModule({ measurer: createApproxMeasurer() }));
    editor.use(editingModule);
    editor.use(structuresModule);
    const ruler = createRulerModule({ storage });
    editor.use(ruler);
    const p = paragraph('Ciao mondo');
    editor.reset({ ...createState(doc([p])), selection: caret(p.id, 0) });
    const root = document.createElement('div');
    document.body.append(root);
    const app = createApp(root, editor, { ribbonItems: collectRibbonItems([rulerExports as Record<string, unknown>]) });
    ruler.mount(app);
    const c = editor.services.layout!.layout(editor.state.doc).pages[0]!.contentRect;
    const mouse = (target: EventTarget, type: string, pt: number, init: MouseEventInit = {}) =>
      target.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, button: 0, clientX: ptToPx(pt), ...init }));
    return { editor, ruler, app, root, c, mouse, storage };
  }

  it('compare sopra il documento con margini, numeri e marcatori', () => {
    const { root, app } = setup();
    const bar = root.querySelector('.ruler-bar')!;
    expect(bar).toBeTruthy();
    expect(bar.nextElementSibling).toBe(app.view.element);
    expect(bar.querySelectorAll('.ruler-number').length).toBeGreaterThan(10);
    for (const k of ['first', 'hanging', 'left', 'right']) expect(bar.querySelector(`.ruler-${k}`)).toBeTruthy();
    expect(bar.querySelector('.ruler-text-area')).toBeTruthy();
  });

  it('Visualizza > Righello lo nasconde e se lo ricorda', async () => {
    const { editor, root, storage } = setup();
    const bar = root.querySelector<HTMLElement>('.ruler-bar')!;
    expect(collectModules([rulerExports as Record<string, unknown>]).map((m) => m.id)).toEqual(['ruler']);
    expect(ribbonItems[0]).toMatchObject({ tab: 'view', commandId: 'view.ruler', kind: 'toggle' });
    await editor.executeAsync('view.ruler');
    expect(bar.hidden).toBe(true);
    expect(storage.getItem('nkoffice-righello')).toBe('no');
    expect(editor.canExecute('edit.undo')).toBe(false);
    document.body.replaceChildren();
    expect(setup(storage).root.querySelector<HTMLElement>('.ruler-bar')!.hidden).toBe(true);
  });

  it('trascinando i marcatori cambiano i rientri, con un solo Annulla', () => {
    const { editor, root, c, mouse } = setup();
    mouse(root.querySelector('.ruler-first')!, 'mousedown', c.x);
    mouse(document, 'mousemove', c.x + cmToPt(1));
    expect(document.querySelector('.ruler-guide')).toBeTruthy();
    mouse(document, 'mouseup', c.x + cmToPt(1));
    expect(document.querySelector('.ruler-guide')).toBeNull();
    expect(first(editor.state.doc).format.indentFirstLine).toBeCloseTo(cmToPt(1), 1);

    // Sporgente: il rientro sinistro va a 2 cm, la prima riga resta a 1 cm.
    mouse(root.querySelector('.ruler-hanging')!, 'mousedown', c.x);
    mouse(document, 'mouseup', c.x + cmToPt(2));
    let ind = effectiveIndents(editor.state.doc, first(editor.state.doc));
    expect(ind.left).toBeCloseTo(cmToPt(2), 1);
    expect(ind.left + ind.firstLine).toBeCloseTo(cmToPt(1), 1);

    // Rettangolino: si spostano insieme.
    mouse(root.querySelector('.ruler-left')!, 'mousedown', c.x + cmToPt(2));
    mouse(document, 'mouseup', c.x + cmToPt(3));
    ind = effectiveIndents(editor.state.doc, first(editor.state.doc));
    expect(ind.left).toBeCloseTo(cmToPt(3), 1);
    expect(ind.left + ind.firstLine).toBeCloseTo(cmToPt(2), 1);

    // Destro.
    mouse(root.querySelector('.ruler-right')!, 'mousedown', c.x + c.width);
    mouse(document, 'mouseup', c.x + c.width - cmToPt(1.5));
    expect(first(editor.state.doc).format.indentRight).toBeCloseTo(cmToPt(1.5), 1);

    editor.undo();
    expect(first(editor.state.doc).format.indentRight).toBeUndefined();
  });

  it('clic sul righello aggiunge la tabulazione del tipo scelto; trascinata fuori si toglie', () => {
    const { editor, root, c, mouse } = setup();
    const typeButton = root.querySelector<HTMLButtonElement>('.ruler-tab-type')!;
    typeButton.click();
    typeButton.click();
    expect(typeButton.dataset.type).toBe('right');
    const strip = root.querySelector('.ruler-strip')!;
    mouse(strip, 'mousedown', c.x + cmToPt(5.1));
    mouse(document, 'mouseup', c.x + cmToPt(5.1));
    expect(first(editor.state.doc).format.tabs).toEqual([{ position: expect.closeTo(cmToPt(5), 1), align: 'right' }]);
    expect(root.querySelectorAll('.ruler-tab-right')).toHaveLength(1);

    // Spostata a 8 cm.
    mouse(root.querySelector('.ruler-tab')!, 'mousedown', c.x + cmToPt(5));
    mouse(document, 'mousemove', c.x + cmToPt(8));
    mouse(document, 'mouseup', c.x + cmToPt(8));
    expect(first(editor.state.doc).format.tabs![0]!.position).toBeCloseTo(cmToPt(8), 1);

    // Trascinata giù, fuori dal righello: sparisce.
    mouse(root.querySelector('.ruler-tab')!, 'mousedown', c.x + cmToPt(8));
    mouse(document, 'mousemove', c.x + cmToPt(8), { clientY: 200 });
    mouse(document, 'mouseup', c.x + cmToPt(8), { clientY: 200 });
    expect(first(editor.state.doc).format.tabs).toEqual([]);

    // Nel margine grigio non si aggiungono tabulazioni.
    mouse(strip, 'mousedown', c.x / 2);
    mouse(document, 'mouseup', c.x / 2);
    expect(first(editor.state.doc).format.tabs).toEqual([]);
  });

  it('trascinando il bordo grigio cambia il margine', () => {
    const { editor, root, c, mouse } = setup();
    mouse(root.querySelector('.ruler-margin-left')!, 'mousedown', c.x);
    mouse(document, 'mousemove', cmToPt(3));
    mouse(document, 'mouseup', cmToPt(3));
    expect(editor.state.doc.sections[0]!.page.margins.left).toBeCloseTo(cmToPt(3), 1);
  });
});
