// @vitest-environment happy-dom
import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  createState,
  Editor,
  firstParagraph,
  paragraph,
  paragraphText,
  textRun,
  type DocumentModel,
  type Paragraph,
} from '../../src/core';
import { editingModule } from '../../src/editing';
import * as translateExports from '../../src/translate';
import {
  batches,
  collectTexts,
  createGoogleTranslator,
  mountTranslatePanel,
  parseGoogleResponse,
  ribbonItems,
  rewriteDocument,
  rewriteParagraph,
  rewriteSelection,
  selectionText,
  SETTINGS_KEY,
  TranslationError,
  type FetchLike,
  type Translator,
} from '../../src/translate';
import { createTranslateModule } from '../../src/translate/module';
import { collectModules, collectRibbonItems } from '../../src/ui/modules';
import { doc } from '../layout/helpers';

afterEach(() => document.body.replaceChildren());

/** Traduttore finto: mette il testo in maiuscolo e ricorda le richieste. */
function fakeTranslator(detected = 'it') {
  const calls: { texts: readonly string[]; to: string; from?: string }[] = [];
  const translator: Translator = {
    async translate(texts, to, from) {
      calls.push({ texts, to, from });
      return { texts: texts.map((t) => t.toUpperCase()), detected };
    },
  };
  return { translator, calls };
}

const upper = (t: string) => t.toUpperCase();
const para = (...children: Paragraph['children']): Paragraph => ({ ...paragraph(), children });

function setup(model: DocumentModel, translator: Translator) {
  const editor = new Editor(createState(model));
  editor.use(editingModule);
  const module = createTranslateModule({ translator, storage: null, delayMs: 0 });
  editor.use(module);
  return { editor, module };
}

function select(editor: Editor, a: [string, number], b: [string, number]) {
  editor.apply({
    state: { ...editor.state, selection: { anchor: { paragraphId: a[0], offset: a[1] }, focus: { paragraphId: b[0], offset: b[1] } } },
    meta: { addToHistory: false },
  });
}

describe('pezzi di testo da tradurre', () => {
  it('traduce il paragrafo intero e tiene la formattazione del run più lungo', () => {
    const p = para(textRun('Nota: ', { bold: true }), textRun('questo è il testo lungo.', { italic: true }));
    const out = rewriteParagraph(p, upper);
    expect(out.children).toEqual([{ type: 'text', text: 'NOTA: QUESTO È IL TESTO LUNGO.', format: { italic: true } }]);
    expect(out.id).toBe(p.id);
  });

  it('tabulazioni, immagini e note restano al loro posto', () => {
    const p = para(textRun('uno'), { type: 'tab' }, textRun(' due '), { type: 'footnoteRef', noteId: 'n1', format: {} });
    const out = rewriteParagraph(p, upper);
    expect(out.children.map((c) => (c.type === 'text' ? c.text : c.type))).toEqual(['UNO', 'tab', ' DUE ', 'footnoteRef']);
  });

  it('numeri e punteggiatura non si mandano al servizio', () => {
    const texts = collectTexts((m) => rewriteDocument(doc([paragraph('12,5 %'), paragraph('Ciao'), paragraph('Ciao')]), m));
    expect(texts).toEqual(['Ciao']);
  });

  it('il testo eliminato in revisione resta com\'è', () => {
    const p = para(textRun('vecchio', {}), { type: 'text', text: 'tolto', format: {}, revision: { kind: 'delete', author: 'A' } });
    const out = rewriteParagraph(p, upper);
    expect(out.children.map((c) => (c.type === 'text' ? c.text : ''))).toEqual(['VECCHIO', 'tolto']);
  });

  it('traduce corpo, tabelle, intestazioni, note e caselle di testo', () => {
    const base = doc([paragraph('corpo')]);
    const cell = { ...base.sections[0]! };
    const model: DocumentModel = {
      ...base,
      sections: [{ ...cell, header: { default: [paragraph('intestazione')] } }],
      footnotes: { n1: { id: 'n1', blocks: [paragraph('nota')] } },
      drawings: {
        d1: { id: 'd1', kind: 'textbox', anchor: 'x', x: 0, y: 0, width: 10, height: 10, text: { text: 'riga uno\nriga due', format: {} } },
      },
    };
    const out = rewriteDocument(model, upper);
    expect(paragraphText(firstParagraph(out))).toBe('CORPO');
    expect(paragraphText(out.sections[0]!.header!.default![0] as Paragraph)).toBe('INTESTAZIONE');
    expect(paragraphText(out.footnotes!.n1!.blocks[0] as Paragraph)).toBe('NOTA');
    expect(out.drawings!.d1!.text!.text).toBe('RIGA UNO\nRIGA DUE');
  });
});

describe('selezione', () => {
  it('sostituisce solo la parte selezionata, anche su più paragrafi, e la lascia selezionata', () => {
    const a = paragraph('prima parte');
    const b = paragraph('seconda parte');
    const { editor } = setup(doc([a, b]), fakeTranslator().translator);
    select(editor, [a.id, 6], [b.id, 7]);
    expect(selectionText(editor.state)).toBe('parte\nseconda');
    const tr = rewriteSelection(editor.state, (t) => `[${t}]`)!;
    editor.apply(tr);
    const [p1, p2] = editor.state.doc.sections[0]!.blocks as Paragraph[];
    expect(paragraphText(p1!)).toBe('prima [parte]');
    expect(paragraphText(p2!)).toBe('[seconda] parte');
    expect(editor.state.selection.focus).toEqual({ paragraphId: b.id, offset: 9 });
    expect(selectionText(editor.state)).toBe('[parte]\n[seconda]');
  });
});

describe('modulo', () => {
  it('Traduci selezione mostra la traduzione e Inserisci la mette nel documento (si annulla)', async () => {
    const p = paragraph('ciao mondo');
    const { translator, calls } = fakeTranslator();
    const { editor, module } = setup(doc([p]), translator);
    select(editor, [p.id, 0], [p.id, 4]);
    await editor.executeAsync('translate.menu', { choice: 'selection' });
    await vi.waitFor(() => expect(module.status).toBe('done'));
    expect(module.open).toBe(true);
    expect(module.source).toBe('ciao');
    expect(module.result).toBe('CIAO');
    expect(module.detected).toBe('it');
    expect(calls[0]).toEqual({ texts: ['ciao'], to: 'en', from: 'auto' });
    expect(module.insert()).toBe(true);
    expect(paragraphText(firstParagraph(editor.state.doc))).toBe('CIAO mondo');
    editor.undo();
    expect(paragraphText(firstParagraph(editor.state.doc))).toBe('ciao mondo');
  });

  it('senza selezione chiede di selezionare del testo', async () => {
    const { editor, module, } = setup(doc([paragraph('ciao')]), fakeTranslator().translator);
    await editor.executeAsync('translate.selection');
    expect(module.status).toBe('idle');
    expect(module.message).toMatch(/Seleziona del testo/);
  });

  it('cambiando selezione con il riquadro aperto si ritraduce', async () => {
    const p = paragraph('uno due');
    const { translator, calls } = fakeTranslator();
    const { editor, module } = setup(doc([p]), translator);
    select(editor, [p.id, 0], [p.id, 3]);
    module.show('selection');
    await vi.waitFor(() => expect(module.result).toBe('UNO'));
    select(editor, [p.id, 4], [p.id, 7]);
    await vi.waitFor(() => expect(module.result).toBe('DUE'));
    expect(calls).toHaveLength(2);
  });

  it('Traduci documento traduce tutto in un passo che si annulla', async () => {
    const { translator, calls } = fakeTranslator();
    const { editor, module } = setup(doc([paragraph('primo'), paragraph('secondo')]), translator);
    module.show('document');
    module.setLanguages({ to: 'de' });
    expect(await module.translateDocument()).toBe(true);
    expect(calls).toEqual([{ texts: ['primo', 'secondo'], to: 'de', from: 'auto' }]);
    expect((editor.state.doc.sections[0]!.blocks as Paragraph[]).map(paragraphText)).toEqual(['PRIMO', 'SECONDO']);
    expect(module.message).toMatch(/Annulla/);
    editor.undo();
    expect((editor.state.doc.sections[0]!.blocks as Paragraph[]).map(paragraphText)).toEqual(['primo', 'secondo']);
  });

  it('mostra l\'errore del servizio senza toccare il documento', async () => {
    const translator: Translator = { translate: async () => Promise.reject(new TranslationError('Niente rete.')) };
    const { editor, module } = setup(doc([paragraph('ciao')]), translator);
    const before = editor.state.doc;
    module.show('document');
    expect(await module.translateDocument()).toBe(false);
    expect(module.status).toBe('error');
    expect(module.message).toBe('Niente rete.');
    expect(editor.state.doc).toBe(before);
  });

  it('ricorda le lingue scelte', () => {
    const store = new Map<string, string>();
    const storage = { getItem: (k: string) => store.get(k) ?? null, setItem: (k: string, v: string) => void store.set(k, v) };
    createTranslateModule({ translator: fakeTranslator().translator, storage }).setLanguages({ from: 'it', to: 'fr' });
    expect(JSON.parse(store.get(SETTINGS_KEY)!)).toEqual({ from: 'it', to: 'fr' });
    const again = createTranslateModule({ translator: fakeTranslator().translator, storage });
    expect([again.from, again.to]).toEqual(['it', 'fr']);
  });

  it('il clic destro su una selezione offre Traduci', () => {
    const p = paragraph('ciao');
    const { editor, module } = setup(doc([p]), fakeTranslator().translator);
    select(editor, [p.id, 0], [p.id, 4]);
    const items = editor.services.viewExtensions!.flatMap((e) => e.contextMenu?.({ paragraphId: p.id, offset: 1 }, editor.state) ?? []);
    const item = items.find((i) => i.label === 'Traduci')!;
    item.run!();
    expect(module.open).toBe(true);
    expect(module.mode).toBe('selection');
  });

  it('compare in Revisione con il menu Traduci', () => {
    expect(collectModules([translateExports]).map((m) => m.id)).toEqual(['translate']);
    const item = collectRibbonItems([{ ribbonItems }]).find((i) => (i.commandId as string) === 'translate.menu')!;
    expect(item.tab).toBe('review');
    expect(item.options!.map((o) => o.label)).toEqual(['Traduci selezione', 'Traduci documento']);
  });
});

describe('riquadro', () => {
  it('si affianca al documento, mostra la traduzione e si chiude', async () => {
    const p = paragraph('ciao');
    const { editor, module } = setup(doc([p]), fakeTranslator().translator);
    const view = document.createElement('div');
    view.className = 'doc-scroll';
    document.body.append(view);
    const dispose = mountTranslatePanel({ view: { element: view } }, module);
    const pane = document.querySelector<HTMLElement>('.translate-pane')!;
    expect(view.parentElement!.className).toBe('translate-workspace');
    expect(pane.hidden).toBe(true);
    select(editor, [p.id, 0], [p.id, 4]);
    module.show('selection');
    await vi.waitFor(() => expect(pane.querySelector('.translate-result')!.textContent).toBe('CIAO'));
    expect(pane.hidden).toBe(false);
    expect(pane.querySelector('.translate-status')!.textContent).toBe('Lingua rilevata: Italiano.');
    pane.querySelector<HTMLButtonElement>('.translate-primary')!.click();
    expect(paragraphText(firstParagraph(editor.state.doc))).toBe('CIAO');
    pane.querySelector<HTMLButtonElement>('.translate-close')!.click();
    expect(pane.hidden).toBe(true);
    dispose();
    expect(view.parentElement).toBe(document.body);
  });
});

describe('servizio Google', () => {
  const response = (body: unknown, status = 200) => ({ ok: status === 200, status, json: async () => body });

  it('legge la risposta e la lingua rilevata', () => {
    expect(parseGoogleResponse([[['Hello. ', 'Ciao. '], ['World', 'Mondo']], null, 'it'])).toEqual({ text: 'Hello. World', detected: 'it' });
    expect(() => parseGoogleResponse({})).toThrow(TranslationError);
  });

  it('manda più testi in una richiesta, uno per riga', async () => {
    const fetchFn = vi.fn<FetchLike>(async () => response([[['One\n', 'uno\n'], ['Two', 'due']], null, 'it']));
    const out = await createGoogleTranslator(fetchFn).translate(['uno', 'due'], 'en');
    expect(out).toEqual({ texts: ['One', 'Two'], detected: 'it' });
    expect(fetchFn).toHaveBeenCalledTimes(1);
    const [url, init] = fetchFn.mock.calls[0]!;
    expect(url).toContain('sl=auto');
    expect(url).toContain('tl=en');
    expect(new URLSearchParams(init.body).get('q')).toBe('uno\ndue');
  });

  it('se le righe non tornano uguali traduce un testo alla volta', async () => {
    const replies = [[[['One two', 'uno\ndue']]], [[['One', 'uno']]], [[['Two', 'due']]]];
    const fetchFn = vi.fn<FetchLike>(async () => response(replies.shift()));
    const out = await createGoogleTranslator(fetchFn).translate(['uno', 'due'], 'en', 'it');
    expect(out.texts).toEqual(['One', 'Two']);
    expect(fetchFn).toHaveBeenCalledTimes(3);
  });

  it('spiega gli errori di rete in italiano', async () => {
    const offline = createGoogleTranslator(async () => Promise.reject(new TypeError('Failed to fetch')));
    await expect(offline.translate(['ciao'], 'en')).rejects.toThrow(/connessione a Internet/);
    const busy = createGoogleTranslator(async () => response(null, 429));
    await expect(busy.translate(['ciao'], 'en')).rejects.toThrow(/troppe richieste/);
  });

  it('divide i testi lunghi in più richieste', () => {
    expect(batches(['aaaa', 'bbbb', 'cc'], 10)).toEqual([[0, 1], [2]]);
    expect(batches(['x'.repeat(20)], 10)).toEqual([[0]]);
  });
});
