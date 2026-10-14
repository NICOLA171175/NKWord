// @vitest-environment happy-dom
import { afterEach, describe, expect, it } from 'vitest';
import { createDocument, createState, Editor, firstParagraph, paragraph, paragraphs, paragraphText, textRun, type TextRun } from '../../src/core';
import { editingModule } from '../../src/editing';
import {
  correctWord,
  createAutoCorrectModule,
  DEFAULT_SETTINGS,
  diffReplacements,
  isSentenceStart,
  loadSettings,
  openAutoCorrectDialog,
  ribbonItems,
  smartQuote,
  type AutoCorrectSettings,
  type AutoCorrectStorage,
} from '../../src/autocorrect';
import * as autocorrectExports from '../../src/autocorrect';
import { collectModules, collectRibbonItems } from '../../src/ui/modules';

afterEach(() => document.body.replaceChildren());

class MapStorage implements AutoCorrectStorage {
  readonly data = new Map<string, string>();
  getItem(k: string) {
    return this.data.get(k) ?? null;
  }
  setItem(k: string, v: string) {
    this.data.set(k, v);
  }
}

function setup(initial = '', settings?: Partial<AutoCorrectSettings>) {
  const p = paragraph(initial ? [textRun(initial)] : []);
  const editor = new Editor(createState(createDocument([p])));
  editor.use(editingModule);
  const module = createAutoCorrectModule({ storage: new MapStorage() });
  if (settings) module.setSettings({ ...DEFAULT_SETTINGS, ...settings });
  editor.use(module);
  const at = { paragraphId: p.id, offset: initial.length };
  editor.apply({ state: { ...editor.state, selection: { anchor: at, focus: at } }, meta: { addToHistory: false } });
  return { editor, module };
}

/** Scrive come da tastiera, un carattere alla volta. */
function type(editor: Editor, s: string) {
  for (const ch of s) {
    if (ch === '\n') editor.execute('edit.splitParagraph');
    else if (ch === '\t') editor.execute('edit.insertTab');
    else editor.execute('edit.insertText', { text: ch });
    editor.history.seal();
  }
}

const text = (editor: Editor) => [...paragraphs(editor.state.doc)].map((l) => paragraphText(l.paragraph)).join('\n');

const word = (s: string, settings = DEFAULT_SETTINGS) => {
  const fix = correctWord(s, s.length, settings);
  return fix ? s.slice(0, fix.start) + fix.text + s.slice(fix.end) : s;
};

describe('regole', () => {
  it('maiuscola a inizio frase, ma non dopo le abbreviazioni', () => {
    expect(word('ciao')).toBe('Ciao');
    expect(word('Bene. poi')).toBe('Bene. Poi');
    expect(word('Davvero? sì')).toBe('Davvero? Sì');
    expect(word('«Vieni.» allora')).toBe('«Vieni.» Allora');
    expect(word('Mele, pere ecc. e')).toBe('Mele, pere ecc. e');
    expect(word('Il dott. rossi')).toBe('Il dott. rossi');
    expect(word('Di G. verdi')).toBe('Di G. verdi');
    expect(word('Aspetta... poi')).toBe('Aspetta... poi');
    expect(word('Il sole')).toBe('Il sole');
    expect(word('(ciao')).toBe('(Ciao');
    expect(word('iPhone')).toBe('iPhone');
    expect(isSentenceStart('Fine! ', 6)).toBe(true);
  });

  it('DUe INiziali e bLOC mAIUSC', () => {
    expect(word('Il DUe')).toBe('Il Due');
    expect(word('Il cIAO')).toBe('Il Ciao');
    expect(word('La NASA')).toBe('La NASA');
    expect(word('Un OK')).toBe('Un OK');
  });

  it('simboli ed errori frequenti, rispettando le maiuscole', () => {
    expect(word('Marchio (c)')).toBe('Marchio ©');
    expect(word('Marchio (TM)')).toBe('Marchio ™');
    expect(word('Costa 5 (e)')).toBe('Costa 5 €');
    expect(word('Bene--male')).toBe('Bene—male');
    expect(word('Bene --')).toBe('Bene —');
    expect(word('Bene -')).toBe('Bene –');
    expect(word('---')).toBe('---');
    expect(word('Aspetta...')).toBe('Aspetta…');
    expect(word('Vai -->')).toBe('Vai →');
    expect(word('Un 1/2')).toBe('Un ½');
    expect(word('Il 11/2')).toBe('Il 11/2');
    expect(word('Sì :)')).toBe('Sì ☺');
    expect(word('Ma perchè')).toBe('Ma perché');
    expect(word('Ma perchè?')).toBe('Ma perché?');
    expect(word('Ma PERCHÈ')).toBe('Ma PERCHÉ');
    expect(word('Perchè')).toBe('Perché');
    expect(word('Un pò')).toBe('Un po’');
    expect(word("Qual'è")).toBe('Qual è');
    expect(word('Qual’è')).toBe('Qual è');
    expect(word('È sopratutto')).toBe('È soprattutto');
    expect(word('Un pò', { ...DEFAULT_SETTINGS, smartQuotes: false })).toBe("Un po'");
  });

  it('non tocca indirizzi web ed email (ci pensa il modulo collegamenti)', () => {
    for (const s of ['www.esempio.it', 'https://esempio.it/a--b', 'mario@esempio.it', 'esempio.com']) expect(word(s)).toBe(s);
  });

  it('ogni opzione si può spegnere', () => {
    const off = { ...DEFAULT_SETTINGS, capitalizeSentences: false, replaceText: false, twoInitialCaps: false, capsLock: false };
    for (const s of ['ciao', 'Ma perchè', 'Il DUe', 'Il cIAO', 'Marchio (c)']) expect(word(s, off)).toBe(s);
  });

  it('voci aggiunte e tolte dall’utente', () => {
    const s = { ...DEFAULT_SETTINGS, custom: [{ from: 'nk', to: 'NKWord' }], removed: ['(c)'] };
    expect(word('Uso nk', s)).toBe('Uso NKWord');
    expect(word('Marchio (c)', s)).toBe('Marchio (c)');
  });

  it('virgolette tipografiche aperte e chiuse', () => {
    const q = (s: string, settings = DEFAULT_SETTINGS) => {
      const fix = smartQuote(s, s.length - 1, settings)!;
      return s.slice(0, fix.start) + fix.text;
    };
    expect(q('"')).toBe('“');
    expect(q('Disse "ciao"')).toBe('Disse "ciao”');
    expect(q("l'")).toBe('l’');
    expect(q(" '")).toBe(' ‘');
    expect(q('Disse "', { ...DEFAULT_SETTINGS, quoteStyle: 'angle' })).toBe('Disse «');
  });
});

describe('mentre si scrive', () => {
  it('corregge dopo lo spazio, la punteggiatura e l’Invio', () => {
    const { editor } = setup();
    type(editor, 'ciao, perchè non vieni? domani. sì\nfine');
    expect(text(editor)).toBe('Ciao, perché non vieni? Domani. Sì\nfine');
    type(editor, ' ');
    expect(text(editor)).toBe('Ciao, perché non vieni? Domani. Sì\nFine ');
  });

  it('il cursore resta dopo la parola corretta e si continua a scrivere', () => {
    const { editor } = setup();
    type(editor, 'Un pò di (c) qual\'è ok');
    expect(text(editor)).toBe('Un po’ di © qual è ok');
    expect(editor.state.selection.focus.offset).toBe(text(editor).length);
  });

  it('virgolette e apostrofi diventano tipografici', () => {
    const { editor } = setup();
    type(editor, 'Disse "l\'albero" ');
    expect(text(editor)).toBe('Disse “l’albero” ');
  });

  it('Ctrl+Z annulla solo la correzione, poi la digitazione', () => {
    const { editor } = setup('Ma perchè');
    type(editor, ' ');
    expect(text(editor)).toBe('Ma perché ');
    editor.undo();
    expect(text(editor)).toBe('Ma perchè ');
    editor.undo();
    expect(text(editor)).toBe('Ma perchè');

    type(editor, '"');
    expect(text(editor)).toBe('Ma perchè”');
    editor.undo();
    expect(text(editor)).toBe('Ma perchè"');
  });

  it('mantiene la formattazione della parola', () => {
    const p = paragraph([textRun('Ma '), textRun('perchè', { bold: true })]);
    const editor = new Editor(createState(createDocument([p])));
    editor.use(editingModule);
    editor.use(createAutoCorrectModule({ storage: null }));
    const at = { paragraphId: p.id, offset: 9 };
    editor.apply({ state: { ...editor.state, selection: { anchor: at, focus: at } } });
    type(editor, ' ');
    const runs = firstParagraph(editor.state.doc).children as TextRun[];
    expect(runs.map((r) => [r.text, !!r.format.bold])).toEqual([
      ['Ma ', false],
      ['perché ', true],
    ]);
  });

  it('non fa nulla con le opzioni spente', () => {
    const { editor } = setup('', { capitalizeSentences: false, smartQuotes: false, replaceText: false });
    type(editor, 'ciao "perchè" ');
    expect(text(editor)).toBe('ciao "perchè" ');
  });

  it('ricorda le opzioni', () => {
    const storage = new MapStorage();
    const module = createAutoCorrectModule({ storage });
    module.setSettings({ ...DEFAULT_SETTINGS, smartQuotes: false, quoteStyle: 'angle', custom: [{ from: 'nk', to: 'NKWord' }] });
    const again = loadSettings(storage);
    expect(again.smartQuotes).toBe(false);
    expect(again.quoteStyle).toBe('angle');
    expect(again.custom).toEqual([{ from: 'nk', to: 'NKWord' }]);
    storage.setItem('nkoffice-correzione-automatica', '{rotto');
    expect(loadSettings(storage)).toEqual(DEFAULT_SETTINGS);
  });
});

describe('finestra delle opzioni', () => {
  it('compare nella barra e si carica da sola', () => {
    expect(collectRibbonItems([autocorrectExports]).map((i) => i.commandId)).toEqual(['autocorrect.options']);
    expect(ribbonItems[0]!.tab).toBe('review');
    expect(collectModules([autocorrectExports]).map((m) => m.id)).toEqual(['autocorrect']);
  });

  it('spegne un’opzione e aggiunge una voce', async () => {
    const promise = openAutoCorrectDialog(DEFAULT_SETTINGS);
    const dialog = document.querySelector<HTMLFormElement>('.autocorrect-dialog')!;
    expect(dialog).toBeTruthy();
    const box = dialog.querySelector<HTMLInputElement>('input[name="capitalizeSentences"]')!;
    box.checked = false;
    const from = dialog.querySelector<HTMLInputElement>('input[name="from"]')!;
    const to = dialog.querySelector<HTMLInputElement>('input[name="to"]')!;
    from.value = 'nk';
    from.dispatchEvent(new Event('input'));
    to.value = 'NKWord';
    to.dispatchEvent(new Event('input'));
    dialog.querySelector<HTMLButtonElement>('button[name="add"]')!.click();
    expect(dialog.querySelector('.autocorrect-row[data-from="nk"]')).toBeTruthy();
    // Toglie una voce predefinita.
    dialog.querySelector<HTMLElement>('.autocorrect-row[data-from="(c)"]')!.click();
    dialog.querySelector<HTMLButtonElement>('button[name="delete"]')!.click();
    dialog.requestSubmit();
    const result = (await promise)!;
    expect(result.capitalizeSentences).toBe(false);
    expect(result.custom).toEqual([{ from: 'nk', to: 'NKWord' }]);
    expect(result.removed).toEqual(['(c)']);
    expect(document.querySelector('.autocorrect-dialog')).toBeNull();
  });

  it('Annulla non cambia nulla', async () => {
    const promise = openAutoCorrectDialog(DEFAULT_SETTINGS);
    Array.from(document.querySelectorAll<HTMLButtonElement>('.autocorrect-btn')).find((b) => b.textContent === 'Annulla')!.click();
    expect(await promise).toBeNull();
  });

  it('il comando apre la finestra e applica le opzioni', async () => {
    const module = createAutoCorrectModule({ storage: null, openDialog: async (s) => ({ ...s, replaceText: false }) });
    const editor = new Editor();
    editor.use(module);
    await editor.executeAsync('autocorrect.options');
    expect(module.settings.replaceText).toBe(false);
    expect(diffReplacements([])).toMatchObject({ custom: [] });
  });
});
