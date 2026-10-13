// @vitest-environment happy-dom
import { afterEach, describe, expect, it } from 'vitest';
import { Editor, createDocument, createState, paragraph, type DocumentLayout, type LayoutEngine } from '../../src/core';
import { DocumentView } from '../../src/ui/document-view';
import { addViewExtension, type ContextMenuItem, type TextMark, type ViewExtension } from '../../src/ui/extensions';
import type { UiRibbonItem } from '../../src/ui/ribbon-items';
import { createApp } from '../../src/ui/shell';
import { ViewStore } from '../../src/ui/view-state';

afterEach(() => document.body.replaceChildren());

function setup(extra: Partial<ViewExtension> = {}) {
  const p = paragraph('ciao gatot nero');
  const editor = new Editor(createState(createDocument([p])));
  const listeners = new Set<() => void>();
  let marks: TextMark[] = [{ paragraphId: p.id, start: 5, end: 10, kind: 'squiggle', color: '#e81123', className: 'spelling-error' }];
  const ext: ViewExtension = {
    id: 'prova',
    marks: () => marks,
    onChange: (l) => (listeners.add(l), () => listeners.delete(l)),
    ...extra,
  };
  addViewExtension(editor, ext);
  const setMarks = (m: TextMark[]) => {
    marks = m;
    listeners.forEach((l) => l());
  };
  return { p, editor, setMarks };
}

describe('segni delle estensioni', () => {
  it('nella vista a scorrimento sottolineano in rosso ondulato solo il testo indicato', () => {
    const { editor, setMarks } = setup();
    const view = new DocumentView(editor, new ViewStore());
    const marked = view.element.querySelectorAll<HTMLElement>('.spelling-error');
    expect(Array.from(marked).map((s) => s.textContent)).toEqual(['gatot']);
    expect(marked[0]!.classList.contains('mark-squiggle')).toBe(true);
    expect(marked[0]!.style.textDecorationLine).toBe('underline');
    expect(marked[0]!.style.textDecorationStyle).toBe('wavy');
    expect(view.element.querySelector('.para')!.textContent).toBe('ciao gatot nero​');
    // Un cambiamento annunciato dall'estensione ridisegna la vista.
    setMarks([]);
    expect(view.element.querySelectorAll('.spelling-error')).toHaveLength(0);
  });

  it('nella vista impaginata diventano rettangoli sotto il testo, dai rettangoli del motore', () => {
    const { editor } = setup();
    const asked: [number, number][] = [];
    const engine: LayoutEngine = {
      layout: (d): DocumentLayout => ({
        doc: d,
        pages: [{ index: 0, number: 1, sectionId: d.sections[0]!.id, width: 600, height: 800, contentRect: { x: 0, y: 0, width: 600, height: 800 }, lines: [], decorations: [], header: [], footer: [] }],
      }),
      caretRect: () => null,
      hitTest: () => null,
      selectionRects: (_l, from, to) => (asked.push([from.offset, to.offset]), [{ page: 0, rect: { x: 30, y: 60, width: 30, height: 15 } }]),
    };
    editor.services.layout = engine;
    const view = new DocumentView(editor, new ViewStore());
    const mark = view.element.querySelector<HTMLElement>('.page-laid-out .text-mark.spelling-error')!;
    expect(asked).toEqual([[5, 10]]);
    expect(mark.classList.contains('mark-squiggle')).toBe(true);
    expect(mark.style.getPropertyValue('--mark-color')).toBe('#e81123');
    expect(mark.style.left).toBe('40px'); // 30 pt
    expect(parseFloat(mark.style.height)).toBeCloseTo(3, 5);
    expect(parseFloat(mark.style.top) + parseFloat(mark.style.height)).toBeCloseTo(100, 5); // fondo della riga: 75 pt
  });
});

describe('shell con estensioni', () => {
  it('mostra lo stato dei pulsanti dei moduli e aggiunge le loro scorciatoie', async () => {
    const ran: string[] = [];
    let on = true;
    const ribbonItems = [
      { tab: 'review', group: 'Prova', commandId: 'prova.toggle', label: 'Interruttore', kind: 'toggle' },
      { tab: 'review', group: 'Prova', commandId: 'prova.lingua', label: 'Lingua', kind: 'dropdown', options: [{ label: 'A', args: { l: 'a' } }, { label: 'B', args: { l: 'b' } }] },
    ] as unknown as UiRibbonItem[];
    const { editor } = setup({
      keybindings: [{ key: 'F7', commandId: 'prova.toggle' }],
      isActive: (id, args) => (id === 'prova.toggle' ? on : (args as { l: string } | undefined)?.l === 'b'),
    });
    editor.commands.register({ id: 'prova.toggle', label: 'Interruttore', async: true, run: async () => (ran.push('toggle'), (on = !on), null) });
    editor.commands.register({ id: 'prova.lingua', label: 'Lingua', async: true, run: async () => null });
    const root = document.createElement('div');
    document.body.append(root);
    const app = createApp(root, editor, { ribbonItems });
    app.ribbon.selectTab('review');
    const button = root.querySelector<HTMLButtonElement>('[data-command="prova.toggle"]')!;
    expect(button.disabled).toBe(false);
    expect(button.title).toBe('Interruttore (F7)');
    expect(button.getAttribute('aria-pressed')).toBe('true');
    expect(root.querySelector<HTMLSelectElement>('[data-command="prova.lingua"]')!.value).toBe('1');
    app.input.element.dispatchEvent(new KeyboardEvent('keydown', { key: 'F7', cancelable: true }));
    await Promise.resolve();
    expect(ran).toEqual(['toggle']);
    app.destroy();
  });

  it('il clic destro apre il menu delle estensioni ed esegue la voce scelta', () => {
    const chosen: string[] = [];
    const items: ContextMenuItem[] = [
      { label: 'gatto', strong: true, run: () => chosen.push('gatto') },
      { label: '(Nessuno)', disabled: true },
      { label: 'Ignora tutto', separator: true, run: () => chosen.push('ignora') },
    ];
    const { editor, p } = setup({ contextMenu: (pos) => (pos.offset >= 5 && pos.offset <= 10 ? items : []) });
    const root = document.createElement('div');
    document.body.append(root);
    const app = createApp(root, editor);
    const word = Array.from(root.querySelectorAll<HTMLElement>('.run-text')).find((s) => s.textContent === 'gatot')!;
    const event = new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 10, clientY: 10 });
    word.dispatchEvent(event);
    expect(event.defaultPrevented).toBe(true);
    const menu = document.querySelector<HTMLElement>('.context-menu')!;
    expect(app.contextMenu).not.toBeNull();
    const buttons = Array.from(menu.querySelectorAll<HTMLButtonElement>('.context-menu-item'));
    expect(buttons.map((b) => b.textContent)).toEqual(['gatto', '(Nessuno)', 'Ignora tutto']);
    expect(buttons[0]!.classList.contains('context-menu-strong')).toBe(true);
    expect(buttons[1]!.disabled).toBe(true);
    expect(menu.querySelectorAll('.context-menu-separator')).toHaveLength(1);
    buttons[0]!.click();
    expect(chosen).toEqual(['gatto']);
    expect(document.querySelector('.context-menu')).toBeNull();
    expect(app.contextMenu).toBeNull();

    // Fuori da una parola segnalata resta il menu del browser.
    const other = Array.from(root.querySelectorAll<HTMLElement>('.run-text')).find((s) => s.textContent === 'ciao ')!;
    const plain = new MouseEvent('contextmenu', { bubbles: true, cancelable: true });
    other.dispatchEvent(plain);
    expect(plain.defaultPrevented).toBe(false);
    expect(p.id).toBeTruthy();
  });

  it('Esc o un clic fuori chiudono il menu', () => {
    const { editor } = setup({ contextMenu: () => [{ label: 'voce', run: () => {} }] });
    const root = document.createElement('div');
    document.body.append(root);
    const app = createApp(root, editor);
    const word = root.querySelector<HTMLElement>('.run-text')!;
    word.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true }));
    expect(document.querySelector('.context-menu')).not.toBeNull();
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
    expect(document.querySelector('.context-menu')).toBeNull();
    word.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true }));
    document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
    expect(document.querySelector('.context-menu')).toBeNull();
    app.destroy();
  });
});
