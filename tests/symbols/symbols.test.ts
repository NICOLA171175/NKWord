// @vitest-environment happy-dom
import { afterEach, describe, expect, it } from 'vitest';
import { createDocument, createState, Editor, firstParagraph, paragraph, paragraphText, textRun, type TextRun } from '../../src/core';
import { editingModule } from '../../src/editing';
import {
  charCode,
  charFromCode,
  createSymbolsModule,
  DEFAULT_RECENT,
  MAX_RECENT_SYMBOLS,
  pushRecent,
  ribbonItems,
  SYMBOL_CATEGORIES,
  symbolName,
  type SymbolStorage,
} from '../../src/symbols';
import * as symbolExports from '../../src/symbols';
import { collectModules, collectRibbonItems } from '../../src/ui/modules';
import { createApp } from '../../src/ui/shell';
import { createUiModule } from '../../src/ui/module';

afterEach(() => document.body.replaceChildren());

class MapStorage implements SymbolStorage {
  readonly data = new Map<string, string>();
  getItem(k: string) {
    return this.data.get(k) ?? null;
  }
  setItem(k: string, v: string) {
    this.data.set(k, v);
  }
}

function setup(storage: SymbolStorage | null = new MapStorage()) {
  const p = paragraph([textRun('Prezzo 10', { bold: true })]);
  const editor = new Editor(createState(createDocument([p])));
  editor.use(editingModule);
  const symbols = createSymbolsModule({ storage });
  editor.use(symbols);
  editor.apply({ state: { ...editor.state, selection: { anchor: { paragraphId: p.id, offset: 9 }, focus: { paragraphId: p.id, offset: 9 } } } });
  return { editor, symbols, p };
}

const text = (editor: Editor) => paragraphText(firstParagraph(editor.state.doc));

describe('tabella dei simboli', () => {
  it('ha i gruppi richiesti, senza doppioni nello stesso gruppo', () => {
    const all = SYMBOL_CATEGORIES.flatMap((c) => c.chars);
    for (const ch of ['€', '©', '®', '™', '°', '±', '½', '→', 'È', 'ñ', 'ß']) expect(all).toContain(ch);
    for (const c of SYMBOL_CATEGORIES) expect(new Set(c.chars).size).toBe(c.chars.length);
  });

  it('nomi e codici come in Word', () => {
    expect(symbolName('€')).toBe('Euro');
    expect(charCode('€')).toBe('20AC');
    expect(charCode('È')).toBe('00C8');
    expect(symbolName('ⅶ')).toBe('Carattere U+2176');
    expect(charFromCode('20ac')).toBe('€');
    expect(charFromCode('U+00A9')).toBe('©');
    expect(charFromCode('0x2192')).toBe('→');
    expect(charFromCode('zz')).toBeNull();
    expect(charFromCode('0007')).toBeNull();
    expect(charFromCode('D800')).toBeNull();
  });

  it('i recenti vanno in testa, senza doppioni, al massimo 20', () => {
    let list: string[] = [...DEFAULT_RECENT];
    list = pushRecent(list, 'È');
    list = pushRecent(list, '€');
    expect(list.slice(0, 2)).toEqual(['€', 'È']);
    expect(list.filter((c) => c === '€')).toHaveLength(1);
    expect(list).toHaveLength(MAX_RECENT_SYMBOLS);
  });
});

describe('comando symbol.insert', () => {
  it('scrive il simbolo al cursore con la formattazione del testo intorno', () => {
    const { editor } = setup();
    expect(editor.execute('symbol.insert', { char: '€' })).toBe(true);
    expect(text(editor)).toBe('Prezzo 10€');
    const runs = firstParagraph(editor.state.doc).children as readonly TextRun[];
    expect(runs.every((r) => r.format.bold)).toBe(true);
  });

  it('ogni simbolo si annulla da solo', () => {
    const { editor } = setup();
    editor.execute('edit.insertText', { text: ' ' });
    editor.execute('symbol.insert', { char: '©' });
    editor.execute('symbol.insert', { char: '®' });
    editor.undo();
    expect(text(editor)).toBe('Prezzo 10 ©');
    editor.undo();
    expect(text(editor)).toBe('Prezzo 10 ');
  });

  it('ricorda i simboli usati, anche alla riapertura', () => {
    const storage = new MapStorage();
    const { editor, symbols } = setup(storage);
    editor.execute('symbol.insert', { char: '½' });
    editor.execute('symbol.insert', { char: 'È' });
    expect(symbols.recent.slice(0, 2)).toEqual(['È', '½']);
    const again = createSymbolsModule({ storage });
    expect(again.recent.slice(0, 2)).toEqual(['È', '½']);
  });

  it('funziona anche senza memoria del browser', () => {
    const { editor, symbols } = setup(null);
    editor.execute('symbol.insert', { char: '°' });
    expect(symbols.recent[0]).toBe('°');
  });

  it('senza carattere non fa nulla', () => {
    const { editor } = setup();
    expect(editor.canExecute('symbol.insert', { char: '' })).toBe(false);
  });
});

describe("nell'app", () => {
  function app() {
    const { editor, symbols } = setup();
    const ui = createUiModule();
    editor.use(ui);
    const root = document.createElement('div');
    document.body.append(root);
    const a = createApp(root, editor, { ui, ribbonItems: collectRibbonItems([symbolExports]) });
    const dispose = symbols.mount(a);
    a.ribbon.selectTab('insert');
    const button = () => a.ribbon.element.querySelector<HTMLButtonElement>('[data-command="symbol.show"]')!;
    const panel = () => document.querySelector<HTMLElement>('.symbols-panel')!;
    return { editor, symbols, a, button, panel, dispose };
  }
  const tick = () => new Promise((r) => setTimeout(r, 0));

  it('il modulo e il pulsante vengono trovati dalla UI', () => {
    expect(collectModules([symbolExports]).map((m) => m.id)).toContain('symbols');
    expect(collectRibbonItems([symbolExports])).toEqual(ribbonItems);
  });

  it('il pulsante Simbolo apre la tabella; un clic inserisce e aggiorna i recenti', async () => {
    const { editor, button, panel } = app();
    expect(panel().hidden).toBe(true);
    button().click();
    await tick();
    expect(panel().hidden).toBe(false);
    const euro = panel().querySelector<HTMLButtonElement>('.symbols-main [data-char="€"]')!;
    euro.click();
    expect(text(editor)).toBe('Prezzo 10€');
    expect(panel().hidden).toBe(false);
    expect(panel().querySelector('.symbols-recent .symbols-cell')!.textContent).toBe('€');
    expect(panel().querySelector('.symbols-name')!.textContent).toBe('Euro');
  });

  it('cambia gruppo e inserisce da codice carattere', async () => {
    const { editor, button, panel } = app();
    button().click();
    await tick();
    const select = panel().querySelector<HTMLSelectElement>('.symbols-select')!;
    select.value = 'maiuscole';
    select.dispatchEvent(new Event('change'));
    expect(panel().querySelector('.symbols-main [data-char="È"]')).toBeTruthy();
    const code = panel().querySelector<HTMLInputElement>('.symbols-code-input')!;
    code.value = '2122';
    code.dispatchEvent(new Event('input'));
    expect(panel().querySelector('.symbols-preview')!.textContent).toBe('™');
    panel().querySelector<HTMLButtonElement>('.symbols-insert')!.click();
    expect(text(editor)).toBe('Prezzo 10™');
    code.value = 'xyz';
    code.dispatchEvent(new Event('input'));
    expect(panel().querySelector<HTMLButtonElement>('.symbols-insert')!.disabled).toBe(true);
  });

  it('Esc chiude la tabella', async () => {
    const { button, panel } = app();
    button().click();
    await tick();
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
    expect(panel().hidden).toBe(true);
  });
});
