// @vitest-environment happy-dom
import { afterEach, describe, expect, it } from 'vitest';
import { createDocument, createState, Editor, paragraph, ptToPx } from '../../src/core';
import { createDrawingsModule, drawingList, ribbonItems } from '../../src/drawings';
import { createApproxMeasurer, createLayoutEngine } from '../../src/layout';
import { collectRibbonItems } from '../../src/ui/modules';
import { createApp } from '../../src/ui/shell';

afterEach(() => document.body.replaceChildren());

function setup() {
  const editor = new Editor(createState(createDocument([paragraph('Testo della pagina')])));
  editor.services.layout = createLayoutEngine(createApproxMeasurer());
  const module = createDrawingsModule();
  editor.use(module);
  const root = document.createElement('div');
  document.body.append(root);
  const app = createApp(root, editor, { ribbonItems: collectRibbonItems([{ ribbonItems }]) });
  module.mount(app);
  const drawingEl = () => root.querySelector<HTMLElement>('.page-laid-out .drawing');
  const d = () => drawingList(editor.state.doc)[0]!;
  const mouse = (target: EventTarget, type: string, x: number, y: number, extra: MouseEventInit = {}) =>
    target.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, clientX: x, clientY: y, button: 0, ...extra }));
  const key = (k: string, extra: KeyboardEventInit = {}) => app.input.element.dispatchEvent(new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true, ...extra }));
  return { editor, module, app, root, drawingEl, d, mouse, key };
}

describe('forme sulla pagina', () => {
  it('compare sulla pagina, selezionata con le maniglie', () => {
    const { editor, drawingEl, d } = setup();
    editor.execute('drawing.insert', { kind: 'shape', shape: 'ellipse' });
    const el = drawingEl()!;
    expect(el.dataset.drawing).toBe(d().id);
    expect(el.querySelector('svg path')!.getAttribute('fill')).toBe('#4472c4');
    expect(el.classList.contains('selected')).toBe(true);
    expect(el.querySelectorAll('.drawing-handle')).toHaveLength(8);
    expect(parseFloat(el.style.width)).toBeCloseTo(ptToPx(72), 3);
  });

  it('si sposta e si ridimensiona trascinando, in un solo passo di Annulla', () => {
    const { editor, drawingEl, d, mouse } = setup();
    editor.execute('drawing.insert', { kind: 'shape', shape: 'rect' });
    const start = d();
    mouse(drawingEl()!, 'mousedown', 100, 100);
    mouse(document, 'mousemove', 120, 110);
    // Durante il trascinamento si vede l'anteprima, il documento non cambia.
    expect(d().x).toBe(start.x);
    expect(parseFloat(drawingEl()!.style.left)).toBeCloseTo(ptToPx(start.x + 15), 3);
    mouse(document, 'mousemove', 140, 130);
    mouse(document, 'mouseup', 140, 130);
    expect(d().x).toBeCloseTo(start.x + 30, 5);
    expect(d().y).toBeCloseTo(start.y + 22.5, 5);

    const handle = drawingEl()!.querySelector<HTMLElement>('[data-handle="se"]')!;
    mouse(handle, 'mousedown', 0, 0);
    mouse(document, 'mousemove', 40, 40);
    mouse(document, 'mouseup', 40, 40);
    expect(d().width).toBeCloseTo(start.width + 30, 5);
    expect(d().height).toBeCloseTo(start.height + 30, 5);
    editor.undo();
    expect(d().width).toBe(start.width);
    expect(d().x).toBeCloseTo(start.x + 30, 5);
  });

  it('un clic senza movimento seleziona soltanto; un clic sul testo deseleziona', () => {
    const { editor, module, app, drawingEl, mouse } = setup();
    editor.execute('drawing.insert', { kind: 'shape', shape: 'rect' });
    module.selection.select(null);
    const doc = editor.state.doc;
    mouse(drawingEl()!, 'mousedown', 50, 50);
    mouse(document, 'mouseup', 51, 50);
    expect(editor.state.doc).toBe(doc);
    expect(module.selection.id).not.toBeNull();
    expect(drawingEl()!.classList.contains('selected')).toBe(true);
    mouse(app.view.element.querySelector('.page')!, 'mousedown', 5, 5);
    expect(module.selection.id).toBeNull();
    expect(drawingEl()!.classList.contains('selected')).toBe(false);
  });

  it('con la tastiera si sposta, si elimina e si deseleziona', () => {
    const { editor, module, d, key } = setup();
    editor.execute('drawing.insert', { kind: 'shape', shape: 'rect' });
    const x = d().x;
    key('ArrowRight');
    expect(d().x).toBe(x + 6);
    key('ArrowLeft', { ctrlKey: true });
    expect(d().x).toBe(x + 5);
    key('Escape');
    expect(module.selection.id).toBeNull();
    // Senza selezione Canc torna al testo.
    key('Delete');
    expect(drawingList(editor.state.doc)).toHaveLength(1);
    module.selection.select(d().id);
    key('Delete');
    expect(editor.state.doc.drawings).toBeUndefined();
  });

  it('con un doppio clic si scrive nella casella di testo', () => {
    const { editor, drawingEl, d, key } = setup();
    editor.execute('drawing.insert', { kind: 'textbox' });
    drawingEl()!.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
    const area = drawingEl()!.querySelector<HTMLTextAreaElement>('textarea.drawing-editor')!;
    expect(area).not.toBeNull();
    area.value = 'Prima riga\nSeconda';
    area.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(d().text?.text).toBe('Prima riga\nSeconda');
    const lines = drawingEl()!.querySelectorAll('.drawing-text > div');
    expect(Array.from(lines, (l) => l.textContent)).toEqual(['Prima riga', 'Seconda']);
    // Scrivere con la forma selezionata riapre la modifica con il carattere digitato.
    key('!');
    expect(drawingEl()!.querySelector<HTMLTextAreaElement>('textarea')!.value).toBe('Prima riga\nSeconda!');
  });

  it('i dati del grafico si modificano nella finestra', () => {
    const { editor, drawingEl, d } = setup();
    editor.execute('drawing.insert', { kind: 'chart', type: 'column' });
    expect(drawingEl()!.querySelectorAll('svg rect').length).toBeGreaterThan(12);
    drawingEl()!.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
    const dialog = document.querySelector<HTMLElement>('.chart-editor')!;
    expect(dialog.querySelectorAll('.chart-editor-value')).toHaveLength(12);
    dialog.querySelector<HTMLSelectElement>('.chart-editor-type')!.value = 'pie';
    dialog.querySelector<HTMLInputElement>('.chart-editor-title')!.value = 'Spese';
    dialog.querySelector<HTMLButtonElement>('.chart-remove-series')!.click();
    expect(dialog.querySelectorAll('.chart-editor-value')).toHaveLength(8);
    dialog.querySelector<HTMLInputElement>('.chart-editor-value')!.value = '10,5';
    dialog.querySelector<HTMLButtonElement>('.drawing-dialog-ok')!.click();
    expect(document.querySelector('.chart-editor')).toBeNull();
    expect(d().chart).toMatchObject({ type: 'pie', title: 'Spese', series: [{ values: [10.5, 2.5, 3.5, 4.5] }, { name: 'Serie 2' }] });
  });

  it('ha un menu contestuale e i pulsanti di formato si attivano con una forma selezionata', () => {
    const { editor, module, root, drawingEl, d, mouse } = setup();
    editor.execute('drawing.insert', { kind: 'shape', shape: 'rect' });
    mouse(drawingEl()!, 'contextmenu', 30, 30, { button: 2 });
    const items = Array.from(document.querySelectorAll('.context-menu-item'), (b) => b.textContent);
    expect(items).toEqual(['Modifica testo', 'Porta in primo piano', 'Porta in secondo piano', 'Elimina']);
    (document.querySelector('.context-menu') as HTMLElement).remove();

    const ribbonTab = root.querySelector<HTMLButtonElement>('.ribbon-tab[data-tab="insert"]')!;
    ribbonTab.click();
    const fill = () => root.querySelector<HTMLSelectElement>('select[data-command="drawing.update"]')!;
    expect(fill().disabled).toBe(false);
    fill().value = '1';
    fill().dispatchEvent(new Event('change'));
    expect(d().fill).toBe('#ed7d31');
    module.selection.select(null);
    expect(fill().disabled).toBe(true);
    expect(root.querySelector<HTMLSelectElement>('select[data-command="drawing.insert"]')!.disabled).toBe(false);
  });
});
