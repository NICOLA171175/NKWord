import { describe, expect, it } from 'vitest';
import {
  Editor,
  caret,
  createDocument,
  createState,
  findParagraph,
  paragraph,
  paragraphText,
  textRun,
  updateParagraph,
  type Command,
} from '../../src/core';
import { createSpellcheckModule, replaceInParagraph, SpellChecker } from '../../src/spellcheck';
import { collectModules, collectRibbonItems } from '../../src/ui/modules';
import { collectMarks, viewExtensions } from '../../src/ui/extensions';
import * as spellcheckExports from '../../src/spellcheck';
import { miniLanguages } from './helpers';

async function setup(text: string | ReturnType<typeof paragraph>[] = 'ciao gatot e casa') {
  const blocks = typeof text === 'string' ? [paragraph(text)] : text;
  const doc = createDocument(blocks);
  const editor = new Editor(createState(doc));
  const notes: string[] = [];
  const checker = new SpellChecker({ languages: miniLanguages, storage: null });
  const module = createSpellcheckModule({ checker, notify: (m) => notes.push(m) });
  editor.use(module);
  await checker.load();
  return { editor, checker, module, notes, p: blocks[0]! };
}

describe('replaceInParagraph', () => {
  it('sostituisce un intervallo che attraversa più run tenendo il formato del primo', () => {
    const p = paragraph([textRun('il ga', { bold: true }), textRun('tot nero')]);
    const out = replaceInParagraph(p, 3, 8, 'gatto');
    expect(paragraphText(out)).toBe('il gatto nero');
    expect(out.children).toEqual([textRun('il gatto', { bold: true }), textRun(' nero')]);
  });

  it('inserisce in un paragrafo vuoto o alla fine', () => {
    expect(paragraphText(replaceInParagraph(paragraph(), 0, 0, 'ciao'))).toBe('ciao');
    expect(paragraphText(replaceInParagraph(paragraph('a'), 1, 1, 'b'))).toBe('ab');
  });
});

describe('modulo controllo ortografico', () => {
  it('viene trovato dalla UI tra le esportazioni del suo index.ts, con i suoi pulsanti', () => {
    const modules = collectModules([spellcheckExports]);
    expect(modules.map((m) => m.id)).toEqual(['spellcheck']);
    const items = collectRibbonItems([spellcheckExports]);
    expect(items.map((i) => i.commandId)).toEqual(['spelling.next', 'spelling.toggle', 'spelling.setLanguage']);
    expect(items[2]!.options!.map((o) => o.label)).toEqual(['Italiano', 'Inglese (Stati Uniti)']);
  });

  it('registra i comandi e pubblica i segni rossi sulle parole sbagliate', async () => {
    const { editor, p } = await setup();
    for (const id of ['spelling.next', 'spelling.toggle', 'spelling.setLanguage', 'spelling.replace', 'spelling.ignoreAll', 'spelling.addToDictionary']) {
      expect(editor.commands.has(id), id).toBe(true);
    }
    expect(viewExtensions(editor).map((e) => e.id)).toEqual(['spellcheck']);
    expect(collectMarks(editor)).toEqual([
      { paragraphId: p.id, start: 5, end: 10, kind: 'squiggle', color: '#e81123', className: 'spelling-error' },
      { paragraphId: p.id, start: 11, end: 12, kind: 'squiggle', color: '#e81123', className: 'spelling-error' },
    ]);
  });

  it('non segnala la parola che si sta scrivendo finché non la si lascia', async () => {
    const { editor, p } = await setup('ciao');
    const type: Command<{ text: string }> = {
      id: 'edit.insertText',
      label: 'Scrivi',
      run: ({ state }, { text }) => {
        const para = findParagraph(state.doc, state.selection.focus.paragraphId)!.paragraph;
        const at = state.selection.focus.offset + text.length;
        const doc = updateParagraph(state.doc, replaceInParagraph(para, state.selection.focus.offset, state.selection.focus.offset, text));
        return { state: { doc, selection: caret(para.id, at) }, meta: { mergeKey: 'typing' } };
      },
    };
    editor.commands.register(type);
    editor.apply({ state: { ...editor.state, selection: caret(p.id, 4) } });
    editor.execute('edit.insertText', { text: ' gatot' });
    expect(collectMarks(editor)).toEqual([]);
    editor.execute('edit.insertText', { text: ' ' });
    expect(collectMarks(editor).map((m) => [m.start, m.end])).toEqual([[5, 10]]);
  });

  it('il menu contestuale propone i suggerimenti e li applica con un passo annullabile', async () => {
    const { editor, p } = await setup();
    const ext = viewExtensions(editor)[0]!;
    expect(ext.contextMenu!({ paragraphId: p.id, offset: 1 }, editor.state)).toEqual([]);
    const items = ext.contextMenu!({ paragraphId: p.id, offset: 7 }, editor.state);
    expect(items.map((i) => i.label)).toEqual(['gatto', 'Ignora tutto', 'Aggiungi al dizionario']);
    expect(items[0]!.strong).toBe(true);
    expect(items[1]!.separator).toBe(true);
    items[0]!.run!();
    await Promise.resolve();
    expect(paragraphText(editor.state.doc.sections[0]!.blocks[0] as never)).toBe('ciao gatto e casa');
    expect(editor.state.selection.focus).toEqual({ paragraphId: p.id, offset: 10 });
    editor.undo();
    expect(paragraphText(editor.state.doc.sections[0]!.blocks[0] as never)).toBe('ciao gatot e casa');
  });

  it('senza suggerimenti mostra una voce disabilitata; Ignora tutto e Aggiungi tolgono il segno', async () => {
    const { editor, checker, p } = await setup('zzzz qqqq');
    const ext = viewExtensions(editor)[0]!;
    const items = ext.contextMenu!({ paragraphId: p.id, offset: 2 }, editor.state);
    expect(items[0]).toEqual({ label: '(Nessun suggerimento)', disabled: true });
    await editor.executeAsync('spelling.ignoreAll', { word: 'zzzz' });
    await editor.executeAsync('spelling.addToDictionary', { word: 'qqqq' });
    expect(collectMarks(editor)).toEqual([]);
    expect(checker.personalWords).toEqual(['qqqq']);
  });

  it('F7 / Ortografia seleziona l\'errore successivo, poi ricomincia; avvisa se non ce ne sono', async () => {
    const { editor, notes, p } = await setup('gatot ciao gatto ciaoo');
    expect(await editor.executeAsync('spelling.next')).toBe(true);
    expect(editor.state.selection).toEqual({ anchor: { paragraphId: p.id, offset: 0 }, focus: { paragraphId: p.id, offset: 5 } });
    await editor.executeAsync('spelling.next');
    expect(editor.state.selection.anchor.offset).toBe(17);
    await editor.executeAsync('spelling.next');
    expect(editor.state.selection.anchor.offset).toBe(0);
    expect(editor.history.canUndo).toBe(false);

    const clean = await setup('ciao gatto');
    expect(await clean.editor.executeAsync('spelling.next')).toBe(false);
    expect(clean.notes).toEqual(['Controllo ortografia completato. Nessun errore trovato.']);
    expect(notes).toEqual([]);
  });

  it('Mostra errori e Lingua cambiano lo stato mostrato dalla barra multifunzione', async () => {
    const { editor, checker } = await setup('hello gatto');
    const ext = viewExtensions(editor)[0]!;
    expect(ext.isActive!('spelling.toggle')).toBe(true);
    expect(ext.isActive!('spelling.setLanguage', { language: 'it' })).toBe(true);
    await editor.executeAsync('spelling.setLanguage', { language: 'xx' });
    expect(ext.isActive!('spelling.setLanguage', { language: 'xx' })).toBe(true);
    expect(collectMarks(editor).map((m) => m.start)).toEqual([6]);
    expect(editor.canExecute('spelling.setLanguage', { language: 'nessuna' })).toBe(false);
    await editor.executeAsync('spelling.toggle');
    expect(checker.settings.enabled).toBe(false);
    expect(ext.isActive!('spelling.toggle')).toBe(false);
    expect(collectMarks(editor)).toEqual([]);
    expect(ext.keybindings).toEqual([{ key: 'F7', commandId: 'spelling.next' }]);
  });

  it('si disattiva pulendo comandi ed estensione', async () => {
    const editor = new Editor();
    const dispose = editor.use(createSpellcheckModule({ checker: new SpellChecker({ languages: miniLanguages, storage: null }) }));
    dispose();
    expect(editor.commands.has('spelling.next')).toBe(false);
    expect(viewExtensions(editor)).toEqual([]);
  });
});
