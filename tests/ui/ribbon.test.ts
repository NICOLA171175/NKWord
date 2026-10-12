// @vitest-environment happy-dom
import { afterEach, describe, expect, it } from 'vitest';
import { COMMAND_IDS, Editor, findParagraph, updateParagraph, type Command } from '../../src/core';
import { Keymap, defaultKeybindings } from '../../src/ui/keymap';
import { createUiModule } from '../../src/ui/module';
import { Ribbon } from '../../src/ui/ribbon';
import { RIBBON_ITEMS } from '../../src/ui/ribbon-items';

function setup(extra: (editor: Editor) => void = () => {}, opts: ConstructorParameters<typeof Ribbon>[2] = {}) {
  const editor = new Editor();
  const ui = createUiModule({ print: () => {} });
  editor.use(ui);
  extra(editor);
  const ribbon = new Ribbon(editor, ui, { keymap: new Keymap(defaultKeybindings(false), false), ...opts });
  document.body.append(ribbon.element);
  const button = (label: string) =>
    ribbon.element.querySelector<HTMLButtonElement>(`button[aria-label="${label}"]`)!;
  return { editor, ui, ribbon, button };
}

afterEach(() => document.body.replaceChildren());

describe('Ribbon', () => {
  it('collega ogni voce a un comando del catalogo', () => {
    for (const item of RIBBON_ITEMS) expect(COMMAND_IDS).toContain(item.commandId);
  });

  it('disabilita i pulsanti dei comandi non ancora registrati', () => {
    const { button } = setup();
    expect(button('Grassetto').disabled).toBe(true);
    expect(button('Grassetto').title).toBe('Grassetto (Ctrl+B)');
    expect(button('Mostra tutto').disabled).toBe(false);
    expect(button('Annulla').disabled).toBe(true); // cronologia vuota
  });

  it('abilita un pulsante quando il modulo registra il comando e mostra lo stato attivo', () => {
    let bold = false;
    const cmd: Command = {
      id: 'format.bold',
      label: 'Grassetto',
      isActive: () => bold,
      run: ({ state }) => {
        bold = !bold;
        return { state: { ...state }, meta: { addToHistory: false } };
      },
    };
    const { editor, button } = setup((e) => e.commands.register(cmd));
    const b = button('Grassetto');
    expect(b.disabled).toBe(false);
    expect(b.getAttribute('aria-pressed')).toBe('false');
    b.click();
    expect(button('Grassetto').getAttribute('aria-pressed')).toBe('true');
    expect(editor.isActive('format.bold')).toBe(true);
  });

  it('esegue i comandi della UI e aggiorna gli interruttori', async () => {
    const { ui, ribbon, button } = setup();
    button('Mostra tutto').click();
    await Promise.resolve();
    expect(ui.view.state.formattingMarks).toBe(true);
    expect(button('Mostra tutto').getAttribute('aria-pressed')).toBe('true');

    ribbon.selectTab('view');
    const zoom = ribbon.element.querySelector<HTMLSelectElement>('select[aria-label="Zoom"]')!;
    zoom.value = '0'; // 50%
    zoom.dispatchEvent(new Event('change'));
    await Promise.resolve();
    expect(ui.view.state.zoom).toBe(0.5);
    expect(zoom.value).toBe('0');
  });

  it('mostra allineamento e stile del paragrafo al cursore', () => {
    const { editor, button } = setup((e) => {
      e.commands.register({ id: 'paragraph.align', label: 'Allinea', run: () => null });
      e.commands.register({ id: 'style.apply', label: 'Stile', run: () => null });
    });
    expect(button('Allinea a sinistra').getAttribute('aria-pressed')).toBe('true');
    expect(button('Normale').getAttribute('aria-pressed')).toBe('true');
    const p = findParagraph(editor.state.doc, editor.state.selection.focus.paragraphId)!.paragraph;
    editor.apply({ state: { ...editor.state, doc: updateParagraph(editor.state.doc, { ...p, styleId: 'Heading1', format: { align: 'center' } }) } });
    expect(button('Centra').getAttribute('aria-pressed')).toBe('true');
    expect(button('Allinea a sinistra').getAttribute('aria-pressed')).toBe('false');
    expect(button('Titolo 1').getAttribute('aria-pressed')).toBe('true');
  });

  it('chiede gli argomenti quando servono e annulla se l\'utente rinuncia', () => {
    const got: unknown[] = [];
    const answers = ['https://example.com', null];
    const { ribbon, button } = setup(
      (e) => e.commands.register({ id: 'format.link', label: 'Link', run: (_c, args) => (got.push(args), null) }),
      { ask: () => answers.shift() ?? null },
    );
    ribbon.selectTab('insert');
    button('Collegamento…').click();
    button('Collegamento…').click();
    expect(got).toEqual([{ url: 'https://example.com' }]);
  });

  it('lascia alla shell i comandi con un gestore proprio', () => {
    let opened = 0;
    const { ribbon, button } = setup(
      (e) => e.commands.register({ id: 'file.open', label: 'Apri', isEnabled: (_c, a) => !!a, run: () => null }),
      { handlers: { 'file.open': () => opened++ } },
    );
    ribbon.selectTab('file');
    expect(button('Apri…').disabled).toBe(false);
    button('Apri…').click();
    expect(opened).toBe(1);
  });
});
