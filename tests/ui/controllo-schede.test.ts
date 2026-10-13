// @vitest-environment happy-dom
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { Editor, cmToPt, firstParagraph, paragraphText } from '../../src/core';
import { editingModule } from '../../src/editing';
import { footnotesModule } from '../../src/footnotes';
import { createHeaderFooterModule, ribbonItems as headerFooterItems } from '../../src/headerfooter';
import { structuresModule } from '../../src/structures';
import { collectRibbonItems } from '../../src/ui/modules';
import { createApp } from '../../src/ui/shell';
import { recordingEditor } from './helpers';

// Controllo dei comandi delle schede Inserisci, Layout, Riferimenti, Revisione, Visualizza e File.
afterEach(() => document.body.replaceChildren());

function mount(editor: Editor) {
  const root = document.createElement('div');
  document.body.append(root);
  return createApp(root, editor, { ribbonItems: collectRibbonItems([{ ribbonItems: headerFooterItems }]) });
}

const button = (root: ParentNode, label: string) => root.querySelector<HTMLButtonElement>(`button[aria-label="${label}"]`)!;
const select = (root: ParentNode, label: string) => root.querySelector<HTMLSelectElement>(`select[aria-label="${label}"]`)!;
const choose = (s: HTMLSelectElement, label: string) => {
  s.value = Array.from(s.options).find((o) => o.textContent === label)!.value;
  s.dispatchEvent(new Event('change'));
};
const key = (target: HTMLElement, k: string, init: KeyboardEventInit = {}) =>
  target.dispatchEvent(new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true, ...init }));

describe('controllo delle schede', () => {
  it('cambiare scheda non toglie il cursore al documento', () => {
    const app = mount(new Editor());
    const tab = app.ribbon.element.querySelector<HTMLButtonElement>('.ribbon-tab[data-tab="insert"]')!;
    const down = new MouseEvent('mousedown', { bubbles: true, cancelable: true });
    tab.dispatchEvent(down);
    expect(down.defaultPrevented).toBe(true);
  });

  it('Inserisci > Intestazione apre la finestra, che tiene il cursore; Esc la chiude senza toccare il documento', async () => {
    const editor = new Editor();
    editor.use(editingModule);
    editor.use(structuresModule);
    editor.use(createHeaderFooterModule());
    const app = mount(editor);
    app.ribbon.selectTab('insert');
    button(app.ribbon.element, 'Intestazione').click();
    await new Promise((r) => setTimeout(r, 0));
    const text = document.querySelector<HTMLTextAreaElement>('.hf-dialog textarea')!;
    expect(text).not.toBeNull();
    expect(document.activeElement).toBe(text);
    key(text, 'Escape');
    await new Promise((r) => setTimeout(r, 0));
    expect(document.querySelector('.hf-dialog')).toBeNull();
    expect(editor.state.doc.sections[0]!.header).toBeUndefined();

    // Con OK il testo va nell'intestazione e non nel corpo.
    button(app.ribbon.element, 'Intestazione').click();
    await new Promise((r) => setTimeout(r, 0));
    document.querySelector<HTMLTextAreaElement>('.hf-dialog textarea')!.value = 'Relazione';
    document.querySelector<HTMLButtonElement>('.hf-primary')!.click();
    await new Promise((r) => setTimeout(r, 0));
    expect(editor.state.doc.sections[0]!.header!.default![0]).toMatchObject({ children: [{ text: 'Relazione' }] });
    expect(paragraphText(firstParagraph(editor.state.doc))).toBe('');
  });

  it('Modifica nota è spento finché il cursore non è su una nota', () => {
    const editor = new Editor();
    editor.use(editingModule);
    editor.use(footnotesModule);
    const app = mount(editor);
    app.ribbon.selectTab('references');
    expect(button(app.ribbon.element, 'Modifica nota…').disabled).toBe(true);
    editor.execute('footnote.insert', { text: 'Nota' });
    expect(button(app.ribbon.element, 'Modifica nota…').disabled).toBe(false);
  });

  it('Layout ha Margini e Dimensioni come Word', () => {
    const editor = new Editor();
    editor.use(structuresModule);
    const app = mount(editor);
    app.ribbon.selectTab('layout');
    const labels = Array.from(app.ribbon.element.querySelectorAll('[data-command="page.setup"]')).map((e) => e.getAttribute('aria-label'));
    expect(labels).toEqual(['Margini', 'Orientamento', 'Dimensioni']);
    choose(select(app.ribbon.element, 'Margini'), 'Stretti');
    expect(editor.state.doc.sections[0]!.page.margins.left).toBeCloseTo(cmToPt(1.27));
    choose(select(app.ribbon.element, 'Orientamento'), 'Orizzontale');
    choose(select(app.ribbon.element, 'Dimensioni'), 'A5 (14,8 × 21 cm)');
    const page = editor.state.doc.sections[0]!.page;
    expect(page.orientation).toBe('landscape');
    expect(page.width).toBeGreaterThan(page.height);
    editor.undo();
    expect(editor.state.doc.sections[0]!.page.width).toBeGreaterThan(cmToPt(29));
  });

  it('File ha Salva; Ctrl+S salva in Word (.docx) invece di "Salva pagina" del browser', () => {
    const { editor, calls } = recordingEditor(['file.export']);
    const app = mount(editor);
    app.ribbon.selectTab('file');
    expect(button(app.ribbon.element, 'Salva')).not.toBeNull();
    const e = new KeyboardEvent('keydown', { key: 's', ctrlKey: true, bubbles: true, cancelable: true });
    app.input.element.dispatchEvent(e);
    expect(e.defaultPrevented).toBe(true);
    expect(calls).toEqual([{ id: 'file.export', args: { formatId: 'docx' } }]);
  });

  it('i menu spenti restano invisibili sotto la loro etichetta', () => {
    const css = readFileSync(resolve(__dirname, '../../src/ui/styles.css'), 'utf8');
    expect(css).toMatch(/\.ribbon-select-overlay:disabled\s*\{\s*opacity:\s*0;/);
  });
});
