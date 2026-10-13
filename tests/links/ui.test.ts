// @vitest-environment happy-dom
import { afterEach, describe, expect, it } from 'vitest';
import { caret, createDocument, createState, Editor, paragraph, paragraphs, textRun, type Paragraph, type TextRun } from '../../src/core';
import { editingModule } from '../../src/editing';
import { createApproxMeasurer, createLayoutModule } from '../../src/layout';
import * as links from '../../src/links';
import { tocAnchor } from '../../src/toc';
import { collectRibbonItems } from '../../src/ui/modules';
import { createApp } from '../../src/ui/shell';

afterEach(() => document.body.replaceChildren());

function setup(blocks: Paragraph[]) {
  const doc = createDocument(blocks);
  const editor = new Editor({ ...createState(doc), selection: caret(blocks[0]!.id, 0) });
  editor.use(editingModule);
  editor.use(createLayoutModule({ measurer: createApproxMeasurer() }));
  const opened: string[] = [];
  const module = links.createLinksModule({ openExternal: (u) => opened.push(u) });
  editor.use(module);
  const root = document.createElement('div');
  document.body.append(root);
  const app = createApp(root, editor, { ribbonItems: collectRibbonItems([links]) });
  module.mount(app);
  const first = () => [...paragraphs(editor.state.doc)][0]!.paragraph;
  const linkedRuns = () => first().children.filter((c): c is TextRun => c.type === 'text' && !!c.format.link);
  const dialog = () => document.querySelector<HTMLFormElement>('.link-dialog');
  const flush = () => new Promise((r) => setTimeout(r, 0));
  return { editor, app, root, opened, first, linkedRuns, dialog, flush };
}

const field = (form: HTMLFormElement, name: string) => form.querySelector<HTMLInputElement>(`[name="${name}"]`)!;

describe('collegamenti nell\'interfaccia', () => {
  it('Ctrl+K apre la finestra, OK inserisce il collegamento sulla selezione', async () => {
    const p = paragraph('il mio sito');
    const { editor, app, dialog, linkedRuns, flush } = setup([p]);
    editor.execute('edit.setSelection', { selection: { anchor: { paragraphId: p.id, offset: 7 }, focus: { paragraphId: p.id, offset: 11 } } });
    app.input.element.dispatchEvent(new KeyboardEvent('keydown', { key: 'k', ctrlKey: true, bubbles: true, cancelable: true }));
    const form = dialog()!;
    expect(form).not.toBeNull();
    expect(field(form, 'text').value).toBe('sito');

    // Senza indirizzo non si chiude.
    form.dispatchEvent(new Event('submit', { cancelable: true }));
    expect(form.querySelector('.link-dialog-error')!.textContent).toContain('indirizzo');
    field(form, 'address').value = 'www.nik.it';
    form.dispatchEvent(new Event('submit', { cancelable: true }));
    await flush();
    expect(dialog()).toBeNull();
    expect(linkedRuns().map((r) => [r.text, r.format.link])).toEqual([['sito', 'https://www.nik.it']]);
  });

  it('email con oggetto e punto del documento dalla finestra', async () => {
    const title = paragraph('Capitolo', {}, 'Heading1');
    const p = paragraph('');
    const { editor, root, dialog, linkedRuns, flush, app } = setup([p, title]);
    app.ribbon.selectTab('insert');
    const button = Array.from(root.querySelectorAll<HTMLButtonElement>('.ribbon-button')).find((b) => b.textContent?.includes('Collegamento'))!;
    button.click();
    let form = dialog()!;
    form.querySelector<HTMLButtonElement>('[data-kind="email"]')!.click();
    field(form, 'text').value = 'Scrivimi';
    field(form, 'email').value = 'anna@posta.it';
    field(form, 'subject').value = 'Ciao a te';
    form.dispatchEvent(new Event('submit', { cancelable: true }));
    await flush();
    expect(linkedRuns().map((r) => [r.text, r.format.link])).toEqual([['Scrivimi', 'mailto:anna@posta.it?subject=Ciao%20a%20te']]);

    editor.execute('edit.insertText', { text: ' ' });
    // La finestra precedente si è chiusa; la nuova elenca i titoli.
    button.click();
    form = dialog()!;
    form.querySelector<HTMLButtonElement>('[data-kind="place"]')!.click();
    const list = form.querySelector<HTMLSelectElement>('select[name="place"]')!;
    expect(Array.from(list.options).map((o) => o.textContent)).toEqual(['Inizio del documento', 'Capitolo']);
    list.value = tocAnchor(title.id);
    field(form, 'text').value = 'vai';
    form.dispatchEvent(new Event('submit', { cancelable: true }));
    await flush();
    expect(linkedRuns().map((r) => r.format.link)).toContain(tocAnchor(title.id));
  });

  it('Ctrl+clic apre il collegamento, il clic normale sposta solo il cursore', async () => {
    const title = paragraph('Capitolo', {}, 'Heading1');
    const p = paragraph([textRun('vai al '), textRun('sito', { link: 'https://nik.it' }), textRun(' o al '), textRun('capitolo', { link: tocAnchor(title.id) })]);
    const { editor, root, opened, flush } = setup([p, title]);
    const run = (url: string) => root.querySelector<HTMLElement>(`.page-laid-out [data-link="${url}"]`)!;
    expect(run('https://nik.it').title).toBe('https://nik.it\nCtrl+clic per aprire il collegamento');
    expect(run(tocAnchor(title.id)).title).toContain('Capitolo');

    run('https://nik.it').dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0 }));
    await flush();
    expect(opened).toEqual([]);

    run('https://nik.it').dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0, ctrlKey: true }));
    await flush();
    expect(opened).toEqual(['https://nik.it']);

    run(tocAnchor(title.id)).dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0, ctrlKey: true }));
    await flush();
    expect(editor.state.selection.focus).toEqual({ paragraphId: title.id, offset: 0 });
  });

  it('il clic destro su un collegamento lo modifica, apre o rimuove', async () => {
    const p = paragraph([textRun('vai al '), textRun('sito', { link: 'https://nik.it' })]);
    const { editor, opened, linkedRuns, dialog, flush } = setup([p]);
    const ext = editor.services.viewExtensions!.find((e) => e.id === links.LINKS_MODULE_ID)!;
    const items = ext.contextMenu!({ paragraphId: p.id, offset: 9 }, editor.state);
    expect(items.map((i) => i.label)).toEqual(['Modifica collegamento…', 'Apri collegamento', 'Copia collegamento', 'Rimuovi collegamento']);
    expect(ext.contextMenu!({ paragraphId: p.id, offset: 2 }, editor.state).map((i) => i.label)).toEqual(['Collegamento…']);

    items[1]!.run!();
    await flush();
    expect(opened).toEqual(['https://nik.it']);

    items[0]!.run!();
    const form = dialog()!;
    expect(field(form, 'address').value).toBe('https://nik.it');
    expect(field(form, 'text').value).toBe('sito');
    form.querySelector<HTMLButtonElement>('.link-dialog-remove')!.click();
    await flush();
    expect(linkedRuns()).toEqual([]);

    editor.undo();
    items[3]!.run!();
    await flush();
    expect(linkedRuns()).toEqual([]);
  });
});
