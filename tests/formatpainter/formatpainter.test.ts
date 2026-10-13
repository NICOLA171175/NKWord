// @vitest-environment happy-dom
import { afterEach, describe, expect, it } from 'vitest';
import { createDocument, createState, Editor, paragraph, textRun, type EditorState, type Paragraph, type TextRun } from '../../src/core';
import { editingModule } from '../../src/editing';
import {
  copyFormat,
  createFormatPainterModule,
  pasteFormat,
  PAINTER_ACTIVE_CLASS,
  ribbonItems,
} from '../../src/formatpainter';
import * as painterExports from '../../src/formatpainter';
import { collectModules, collectRibbonItems } from '../../src/ui/modules';
import { createApp } from '../../src/ui/shell';
import { createUiModule } from '../../src/ui/module';

afterEach(() => document.body.replaceChildren());

/** "Grassetto rosso" (grassetto, rosso) + " e normale" nel primo paragrafo; un secondo paragrafo centrato e Titolo 1. */
function sample() {
  const p1 = paragraph([textRun('Grassetto', { bold: true, color: '#ff0000', fontSize: 14 }), textRun(' e normale')]);
  const p2 = paragraph('Altro testo qui', { align: 'center' }, 'Heading1');
  const doc = createDocument([p1, p2]);
  return { doc, p1, p2 };
}

function select(state: EditorState, a: [string, number], b: [string, number]): EditorState {
  return { ...state, selection: { anchor: { paragraphId: a[0], offset: a[1] }, focus: { paragraphId: b[0], offset: b[1] } } };
}

const para = (state: EditorState, i: number) => state.doc.sections[0]!.blocks[i] as Paragraph;
const runs = (p: Paragraph) => p.children as TextRun[];

describe('copia e incolla del formato', () => {
  it('copia il formato del carattere e lo applica alla selezione', () => {
    const { doc, p1 } = sample();
    let state = select(createState(doc), [p1.id, 0], [p1.id, 4]);
    const copied = copyFormat(state);
    expect(copied.char).toMatchObject({ bold: true, color: '#ff0000', fontSize: 14 });
    expect(copied.paragraph).toBeUndefined();
    state = select(state, [p1.id, 12], [p1.id, 19]);
    state = pasteFormat(state, copied)!.state;
    const painted = runs(para(state, 0)).find((r) => r.text === 'normale')!;
    expect(painted.format).toMatchObject({ bold: true, color: '#ff0000', fontSize: 14 });
  });

  it('toglie il grassetto se il testo di partenza non lo ha', () => {
    const { doc, p1 } = sample();
    let state = select(createState(doc), [p1.id, 12], [p1.id, 19]);
    const copied = copyFormat(state);
    state = select(state, [p1.id, 0], [p1.id, 9]);
    state = pasteFormat(state, copied)!.state;
    const first = runs(para(state, 0))[0]!;
    expect(first.format.bold).toBeUndefined();
    expect(first.format.color).toBeUndefined();
    expect(runs(para(state, 0)).length).toBe(1);
  });

  it('con selezione vuota applica alla parola sotto il cursore', () => {
    const { doc, p1, p2 } = sample();
    let state = select(createState(doc), [p1.id, 0], [p1.id, 9]);
    const copied = copyFormat(state);
    state = select(state, [p2.id, 7], [p2.id, 7]);
    state = pasteFormat(state, copied)!.state;
    const word = runs(para(state, 1)).find((r) => r.text === 'testo')!;
    expect(word.format).toMatchObject({ color: '#ff0000', fontSize: 14 });
  });

  it('copia anche il paragrafo se la selezione è vuota', () => {
    const { doc, p1, p2 } = sample();
    let state = select(createState(doc), [p2.id, 2], [p2.id, 2]);
    const copied = copyFormat(state);
    expect(copied.paragraph).toEqual({ styleId: 'Heading1', format: { align: 'center' } });
    state = select(state, [p1.id, 0], [p1.id, 4]);
    state = pasteFormat(state, copied)!.state;
    expect(para(state, 0).styleId).toBe('Heading1');
    expect(para(state, 0).format.align).toBe('center');
  });

  it('non tocca collegamenti e commenti', () => {
    const p = paragraph([textRun('sito', { link: 'https://example.com' }), textRun(' x'), textRun('bold', { bold: true })]);
    let state = select(createState(createDocument([p])), [p.id, 7], [p.id, 11]);
    const copied = copyFormat(state);
    state = select(state, [p.id, 0], [p.id, 4]);
    state = pasteFormat(state, copied)!.state;
    expect(runs(para(state, 0))[0]!.format).toMatchObject({ link: 'https://example.com', bold: true });
  });
});

function setup() {
  const { doc, p1, p2 } = sample();
  const editor = new Editor(createState(doc));
  editor.use(editingModule);
  const painter = createFormatPainterModule();
  editor.use(painter);
  return { editor, painter, p1, p2 };
}

describe('pennello', () => {
  it('un clic: si applica una volta e si spegne', async () => {
    const { editor, painter, p1 } = setup();
    editor.apply({ state: select(editor.state, [p1.id, 0], [p1.id, 4]) });
    await editor.executeAsync('formatPainter.toggle');
    expect(painter.mode).toBe('once');
    expect(editor.isActive('formatPainter.toggle')).toBe(false);
    editor.apply({ state: select(editor.state, [p1.id, 12], [p1.id, 19]) });
    expect(editor.execute('formatPainter.apply')).toBe(true);
    expect(painter.mode).toBe('off');
    expect(editor.canExecute('formatPainter.apply')).toBe(false);
    // Annulla riporta il testo com'era.
    editor.undo();
    expect(runs(para(editor.state, 0)).some((r) => r.text.includes('normale') && !r.format.bold)).toBe(true);
  });

  it('doppio clic: resta acceso finché non si preme Esc', async () => {
    const { editor, painter, p1, p2 } = setup();
    editor.apply({ state: select(editor.state, [p1.id, 0], [p1.id, 4]) });
    await editor.executeAsync('formatPainter.toggle', { sticky: true });
    editor.apply({ state: select(editor.state, [p1.id, 12], [p1.id, 19]) });
    editor.execute('formatPainter.apply');
    editor.apply({ state: select(editor.state, [p2.id, 0], [p2.id, 5]) });
    editor.execute('formatPainter.apply');
    expect(painter.mode).toBe('sticky');
    await editor.executeAsync('formatPainter.cancel');
    expect(painter.mode).toBe('off');
  });

  it('un secondo clic sul pennello lo spegne', async () => {
    const { editor, painter } = setup();
    await editor.executeAsync('formatPainter.toggle');
    await editor.executeAsync('formatPainter.toggle');
    expect(painter.mode).toBe('off');
  });

  it('Ctrl+Maiusc+C e Ctrl+Maiusc+V copiano e incollano senza pennello', async () => {
    const { editor, painter, p1 } = setup();
    expect(editor.canExecute('formatPainter.paste')).toBe(false);
    editor.apply({ state: select(editor.state, [p1.id, 0], [p1.id, 4]) });
    await editor.executeAsync('formatPainter.copy');
    editor.apply({ state: select(editor.state, [p1.id, 12], [p1.id, 19]) });
    expect(editor.execute('formatPainter.paste')).toBe(true);
    expect(painter.mode).toBe('off');
    expect(runs(para(editor.state, 0)).find((r) => r.text === 'normale')!.format.bold).toBe(true);
  });
});

describe("nell'app", () => {
  function app() {
    const { editor, painter, p1 } = setup();
    const ui = createUiModule();
    editor.use(ui);
    const root = document.createElement('div');
    document.body.append(root);
    const a = createApp(root, editor, { ui, ribbonItems: collectRibbonItems([painterExports]) });
    const dispose = (painter as typeof painter & { mount(h: typeof a): () => void }).mount(a);
    const button = () => a.ribbon.element.querySelector<HTMLButtonElement>('[data-command="formatPainter.toggle"]')!;
    return { editor, painter, p1, a, button, dispose };
  }

  it('il modulo e il pulsante vengono trovati dalla UI', () => {
    expect(collectModules([painterExports]).map((m) => m.id)).toContain('formatpainter');
    expect(collectRibbonItems([painterExports])).toEqual(ribbonItems);
  });

  it('il pulsante nella scheda Home accende il pennello e cambia il puntatore', async () => {
    const { painter, a, button } = app();
    expect(button()).toBeTruthy();
    expect(button().disabled).toBe(false);
    button().click();
    await Promise.resolve();
    expect(painter.mode).toBe('once');
    expect(button().getAttribute('aria-pressed')).toBe('true');
    expect(a.view.element.classList.contains(PAINTER_ACTIVE_CLASS)).toBe(true);
  });

  it('doppio clic tiene acceso il pennello', async () => {
    const { painter, button } = app();
    button().click();
    await Promise.resolve();
    button().click();
    await Promise.resolve();
    button().dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
    await Promise.resolve();
    expect(painter.mode).toBe('sticky');
  });

  it('rilasciando il mouse sul documento il formato si applica', async () => {
    const { editor, painter, p1, a, button } = app();
    editor.apply({ state: select(editor.state, [p1.id, 0], [p1.id, 4]) });
    button().click();
    await Promise.resolve();
    a.view.element.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }));
    editor.apply({ state: select(editor.state, [p1.id, 12], [p1.id, 19]) });
    document.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, button: 0 }));
    expect(runs(para(editor.state, 0)).find((r) => r.text === 'normale')!.format.bold).toBe(true);
    expect(painter.mode).toBe('off');
    expect(a.view.element.classList.contains(PAINTER_ACTIVE_CLASS)).toBe(false);
  });

  it('Esc spegne il pennello', async () => {
    const { painter, a, button } = app();
    button().click();
    await Promise.resolve();
    a.input.focus();
    document.activeElement!.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    await Promise.resolve();
    expect(painter.mode).toBe('off');
  });
});
