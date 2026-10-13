// @vitest-environment happy-dom
import { afterEach, describe, expect, it } from 'vitest';
import { Editor, createDocument, createState, paragraph, paragraphs, type TextRun } from '../../src/core';
import { editingModule } from '../../src/editing';
import { HIGHLIGHT_COLORS, shade, themeGrid } from '../../src/ui/color-picker';
import { createUiModule } from '../../src/ui/module';
import { Ribbon } from '../../src/ui/ribbon';

function setup() {
  const p = paragraph('uno due');
  const editor = new Editor(createState(createDocument([p])));
  editor.use(editingModule);
  const ui = createUiModule({ print: () => {} });
  editor.use(ui);
  editor.execute('edit.setSelection', { selection: { anchor: { paragraphId: p.id, offset: 0 }, focus: { paragraphId: p.id, offset: 3 } } });
  const ribbon = new Ribbon(editor, ui);
  document.body.append(ribbon.element);
  const q = <T extends Element>(sel: string) => ribbon.element.querySelector<T>(sel)!;
  const para = () => [...paragraphs(editor.state.doc)][0]!.paragraph;
  const firstRun = () => para().children[0] as TextRun;
  return { editor, ribbon, q, para, firstRun };
}

const tick = () => new Promise((r) => setTimeout(r, 0));

afterEach(() => document.body.replaceChildren());

describe('tavolozze dei colori', () => {
  it('la griglia del tema ha 6 righe da 10 con le sfumature di Word', () => {
    const grid = themeGrid();
    expect(grid.map((r) => r.length)).toEqual([10, 10, 10, 10, 10, 10]);
    expect(grid[1]![0]!).toEqual({ color: '#f2f2f2', name: 'Bianco, Sfondo 1, più scuro 5%' });
    expect(grid[1]![4]!.name).toBe('Blu, Colore 1, più chiaro 80%');
    expect(shade('#4472c4', 0.8)).toBe('#dae3f3');
    expect(shade('#4472c4', -0.5)).toBe('#203864');
    expect(HIGHLIGHT_COLORS).toHaveLength(15);
  });

  it('la freccia apre la tavolozza e un colore si applica al testo selezionato', async () => {
    const { q, firstRun } = setup();
    const arrow = q<HTMLButtonElement>('[data-palette="text"] .ribbon-split-arrow');
    const menu = q<HTMLElement>('[data-palette="text"] .color-menu');
    expect(arrow.disabled).toBe(false);
    expect(menu.hidden).toBe(true);
    arrow.click();
    expect(menu.hidden).toBe(false);
    expect(menu.querySelectorAll('.color-swatch')).toHaveLength(70);
    menu.querySelector<HTMLButtonElement>('[aria-label="Blu"]')!.click();
    await tick();
    expect(menu.hidden).toBe(true);
    expect(firstRun()).toMatchObject({ text: 'uno', format: { color: '#0070c0' } });
  });

  it('il pulsante principale riapplica l\'ultimo colore, Automatico lo toglie', async () => {
    const { q, firstRun } = setup();
    const main = q<HTMLButtonElement>('[data-palette="text"] .ribbon-split-main');
    expect(main.getAttribute('aria-label')).toBe('Colore carattere');
    main.click();
    await tick();
    expect(firstRun().format.color).toBe('#ff0000');
    q<HTMLButtonElement>('[data-palette="text"] .ribbon-split-arrow').click();
    Array.from(q<HTMLElement>('[data-palette="text"] .color-menu').querySelectorAll<HTMLButtonElement>('.color-menu-item'))
      .find((b) => b.textContent === 'Automatico')!
      .click();
    await tick();
    expect(firstRun().format).toEqual({});
    expect(main.dataset.color).toBe('#000000');
  });

  it('evidenziatore con i 15 colori e Nessun colore', async () => {
    const { q, firstRun } = setup();
    q<HTMLButtonElement>('[data-palette="highlight"] .ribbon-split-arrow').click();
    const menu = q<HTMLElement>('[data-palette="highlight"] .color-menu');
    expect(menu.querySelectorAll('.color-swatch')).toHaveLength(15);
    menu.querySelector<HTMLButtonElement>('[aria-label="Verde brillante"]')!.click();
    await tick();
    expect(firstRun().format.highlight).toBe('#00ff00');
  });

  it('sfondo e bordi del paragrafo dalla scheda Home', async () => {
    const { q, para } = setup();
    q<HTMLButtonElement>('[data-palette="shading"] .ribbon-split-main').click();
    await tick();
    expect(para().format.shading).toBe('#d9d9d9');
    const select = q<HTMLSelectElement>('select[data-command="paragraph.borders"]');
    select.value = String(Array.from(select.options).findIndex((o) => o.textContent === 'Bordi esterni') - 1);
    select.dispatchEvent(new Event('change'));
    await tick();
    expect(Object.keys(para().format.borders ?? {}).sort()).toEqual(['bottom', 'left', 'right', 'top']);
  });

  it('"Altri colori…" usa il selettore di sistema; Esc e un clic fuori chiudono il menu', async () => {
    const { q, firstRun } = setup();
    const arrow = q<HTMLButtonElement>('[data-palette="shading"] .ribbon-split-arrow');
    const menu = q<HTMLElement>('[data-palette="shading"] .color-menu');
    arrow.click();
    menu.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(menu.hidden).toBe(true);
    arrow.click();
    document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
    expect(menu.hidden).toBe(true);

    q<HTMLButtonElement>('[data-palette="text"] .ribbon-split-arrow').click();
    const input = q<HTMLInputElement>('[data-palette="text"] input[type="color"]');
    input.value = '#123456';
    input.dispatchEvent(new Event('change'));
    await tick();
    expect(firstRun().format.color).toBe('#123456');
  });
});
