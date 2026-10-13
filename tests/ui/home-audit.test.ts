// @vitest-environment happy-dom
import { afterEach, describe, expect, it } from 'vitest';
import { createDocument, createState, Editor, paragraph, textRun } from '../../src/core';
import { editingModule } from '../../src/editing';
import { Keymap, defaultKeybindings } from '../../src/ui/keymap';
import { createUiModule } from '../../src/ui/module';
import { Ribbon } from '../../src/ui/ribbon';
import { RIBBON_ITEMS, type UiRibbonItem } from '../../src/ui/ribbon-items';
import { createApp } from '../../src/ui/shell';

afterEach(() => document.body.replaceChildren());

function setup(items?: readonly UiRibbonItem[]) {
  const p = paragraph([textRun('aa', { fontFamily: 'Arial', fontSize: 20 }), textRun('bb')]);
  const editor = new Editor(createState(createDocument([p])));
  editor.use(editingModule);
  const ui = createUiModule({ print: () => {} });
  editor.use(ui);
  const ribbon = new Ribbon(editor, ui, { keymap: new Keymap(defaultKeybindings(false), false), ...(items ? { items } : {}) });
  document.body.append(ribbon.element);
  const select = (a: number, b: number) =>
    editor.execute('edit.setSelection', {
      selection: { anchor: { paragraphId: p.id, offset: a }, focus: { paragraphId: p.id, offset: b } },
    });
  const shown = (label: string) => {
    const s = ribbon.element.querySelector<HTMLSelectElement>(`select[aria-label="${label}"]`)!;
    return s.options[s.selectedIndex]!.textContent;
  };
  return { editor, ribbon, select, shown };
}

describe('scheda Home: menu e gruppi come in Word', () => {
  it('carattere e dimensione restano vuoti se la selezione ne contiene più d\'uno', () => {
    const { select, shown } = setup();
    select(0, 2);
    expect(shown('Tipo di carattere')).toBe('Arial');
    expect(shown('Dimensione carattere')).toBe('20');
    select(0, 4);
    expect(shown('Tipo di carattere')).toBe('Tipo di carattere');
    expect(shown('Dimensione carattere')).toBe('Dimensione carattere');
  });

  it('l\'interlinea si legge con la virgola e mostra il valore del paragrafo', () => {
    const { editor, shown } = setup();
    const item = RIBBON_ITEMS.find((i) => i.commandId === 'paragraph.lineSpacing')!;
    expect(item.options!.map((o) => o.label)).toEqual(['1,0', '1,15', '1,5', '2,0', '2,5', '3,0']);
    editor.execute('paragraph.lineSpacing', { lineSpacing: { rule: 'multiple', value: 1.5 } });
    expect(shown('Interlinea')).toBe('1,5');
  });

  it('il gruppo Appunti viene subito dopo Annulla anche se le sue voci arrivano dai moduli', () => {
    const extra: UiRibbonItem = { tab: 'home', group: 'Appunti', commandId: 'edit.selectAll', label: 'Prova', kind: 'button' };
    const { ribbon } = setup([...RIBBON_ITEMS, extra]);
    const groups = Array.from(ribbon.element.querySelectorAll('.ribbon-group-title')).map((g) => g.textContent);
    expect(groups.slice(0, 3)).toEqual(['Annulla', 'Appunti', 'Carattere']);
  });

  it('Ctrl+F apre Trova invece della ricerca del browser', () => {
    const editor = new Editor(createState(createDocument([paragraph('uno due')])));
    editor.use(editingModule);
    const asked: string[] = [];
    const root = document.createElement('div');
    document.body.append(root);
    const app = createApp(root, editor, { ask: (m) => (asked.push(m), 'due') });
    const e = new KeyboardEvent('keydown', { key: 'f', code: 'KeyF', ctrlKey: true, cancelable: true });
    app.input.element.dispatchEvent(e);
    expect(e.defaultPrevented).toBe(true);
    expect(asked).toEqual(['Trova:']);
    expect(editor.state.selection.anchor.offset).toBe(4);
    app.destroy();
  });
});
