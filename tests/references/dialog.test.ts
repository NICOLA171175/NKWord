// @vitest-environment happy-dom
import { afterEach, describe, expect, it } from 'vitest';
import { caret, createDocument, createState, Editor, paragraph, paragraphText, paragraphs } from '../../src/core';
import { editingModule } from '../../src/editing';
import { collectCaptions, collectMarks, createReferencesModule, openCrossRefDialog, openMarkEntryDialog } from '../../src/references';

afterEach(() => document.body.replaceChildren());

const dialog = () => document.querySelector<HTMLFormElement>('.refs-dialog')!;
const control = <T extends HTMLElement>(name: string) => dialog().querySelector<T>(`[name="${name}"]`)!;
const buttonNamed = (text: string) => Array.from(dialog().querySelectorAll('button')).find((b) => b.textContent === text)!;
const flush = () => new Promise((r) => setTimeout(r, 0));

function setup() {
  const fig = paragraph('[immagine]');
  const editor = new Editor({ ...createState(createDocument([fig, paragraph('Vedi ')])), selection: caret(fig.id, 0) });
  editor.use(editingModule);
  editor.use(createReferencesModule());
  return editor;
}

describe('finestre di Riferimenti', () => {
  it('"Inserisci didascalia" mostra l\'anteprima e inserisce la didascalia', async () => {
    const editor = setup();
    const done = editor.executeAsync('references.captionDialog');
    expect(dialog().querySelector('.refs-preview')!.textContent).toBe('Figura 1');
    const text = control<HTMLInputElement>('text');
    text.value = 'Il porto';
    text.dispatchEvent(new Event('input'));
    expect(dialog().querySelector('.refs-preview')!.textContent).toBe('Figura 1: Il porto');
    // Tabella propone la posizione "sopra".
    const label = control<HTMLSelectElement>('label');
    label.value = 'Tabella';
    label.dispatchEvent(new Event('change'));
    expect(control<HTMLSelectElement>('position').value).toBe('above');
    label.value = 'Figura';
    label.dispatchEvent(new Event('change'));
    buttonNamed('OK').click();
    await done;
    expect(document.querySelector('.refs-dialog')).toBeNull();
    expect(collectCaptions(editor.state.doc).map((c) => c.text)).toEqual(['Figura 1: Il porto']);
  });

  it('una nuova etichetta deve essere una sola parola', async () => {
    const editor = setup();
    const done = editor.executeAsync('references.captionDialog');
    const custom = control<HTMLInputElement>('newLabel');
    custom.value = 'Due parole';
    custom.dispatchEvent(new Event('input'));
    expect(buttonNamed('OK').disabled).toBe(true);
    custom.value = 'Grafico';
    custom.dispatchEvent(new Event('input'));
    expect(buttonNamed('OK').disabled).toBe(false);
    buttonNamed('Annulla').click();
    await done;
    expect(collectCaptions(editor.state.doc)).toHaveLength(0);
  });

  it('"Riferimento incrociato" elenca le didascalie e inserisce il riferimento', async () => {
    const editor = setup();
    editor.execute('references.insertCaption', { label: 'Figura', text: 'Mappa' });
    const vedi = [...paragraphs(editor.state.doc)].at(-1)!.paragraph;
    editor.execute('edit.setSelection', { selection: caret(vedi.id, 5) });

    const done = editor.executeAsync('references.crossRefDialog');
    expect(dialog().querySelector('.refs-empty')!.textContent).toContain('non ci sono titoli');
    const type = control<HTMLSelectElement>('type');
    type.value = 'caption:Figura';
    type.dispatchEvent(new Event('change'));
    const list = control<HTMLSelectElement>('target');
    expect(Array.from(list.options).map((o) => o.textContent)).toEqual(['Figura 1: Mappa']);
    expect(Array.from(control<HTMLSelectElement>('kind').options).map((o) => o.textContent)).toContain('Solo etichetta e numero');
    buttonNamed('Inserisci').click();
    await done;
    expect(paragraphText([...paragraphs(editor.state.doc)].at(-1)!.paragraph)).toBe('Vedi Figura 1');
  });

  it('"Segna voce" propone la parola e segna con "Segna tutto"', async () => {
    const p = paragraph('Roma e Roma');
    const editor = new Editor({ ...createState(createDocument([p])), selection: caret(p.id, 2) });
    editor.use(editingModule);
    editor.use(createReferencesModule());
    const done = editor.executeAsync('references.markEntryDialog');
    expect(control<HTMLInputElement>('main').value).toBe('Roma');
    control<HTMLInputElement>('sub').value = 'storia';
    buttonNamed('Segna tutto').click();
    await done;
    expect(collectMarks(editor.state.doc).map((m) => m.entry)).toEqual(['Roma:storia', 'Roma:storia']);
  });

  it('Esc chiude le finestre senza fare nulla', async () => {
    const result = openMarkEntryDialog({ entry: 'x' });
    dialog().dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(await result).toBeNull();
    const ref = openCrossRefDialog({ targets: [] });
    await flush();
    buttonNamed('Annulla').click();
    expect(await ref).toBeNull();
  });
});
