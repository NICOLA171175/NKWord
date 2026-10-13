import { describe, expect, it } from 'vitest';
import { COMMAND_IDS, COMMAND_OWNERS, Editor, paragraphText, firstParagraph, type Command } from '../../src/core';
import { createIoModule, ioModule, type SaveRequest } from '../../src/io';

const typing: Command<string> = {
  id: 'test.type',
  label: 'Scrivi',
  run: ({ state }) => {
    const p = firstParagraph(state.doc);
    const doc = { ...state.doc, sections: [{ ...state.doc.sections[0]!, blocks: [{ ...p, children: [{ type: 'text' as const, text: 'x', format: {} }] }] }] };
    return { state: { ...state, doc } };
  },
};

describe('modulo io', () => {
  it('pubblica formati e appunti e registra tutti e soli i suoi comandi', () => {
    const editor = new Editor();
    const dispose = editor.use(createIoModule({ save: () => {} }));
    expect(editor.services.formats!.list().map((f) => f.id)).toEqual(['docx', 'html', 'txt']);
    expect(editor.services.formats!.byExtension('Relazione.DOCX')?.id).toBe('docx');
    expect(editor.services.formats!.byExtension('.htm')?.id).toBe('html');
    expect(editor.services.clipboard).toBeDefined();
    const own = COMMAND_IDS.filter((id) => COMMAND_OWNERS[id] === 'io');
    for (const id of own) expect(editor.commands.has(id)).toBe(true);
    expect(editor.commands.list().map((c) => c.id).filter((id) => id.startsWith('file.')).sort()).toEqual([...own].sort());

    dispose();
    for (const id of own) expect(editor.commands.has(id)).toBe(false);
    expect(editor.services.formats).toBeUndefined();
    expect(editor.services.clipboard).toBeUndefined();
  });

  it('file.new crea un documento vuoto e azzera la cronologia', () => {
    const editor = new Editor();
    editor.use(ioModule);
    editor.commands.register(typing);
    editor.execute('test.type');
    expect(editor.history.canUndo).toBe(true);
    expect(editor.execute('file.new')).toBe(true);
    expect(paragraphText(firstParagraph(editor.state.doc))).toBe('');
    expect(editor.history.canUndo).toBe(false);
    expect(editor.canExecute('edit.undo')).toBe(false);
  });

  it('file.open importa con il formato indicato e azzera la cronologia', async () => {
    const editor = new Editor();
    editor.use(ioModule);
    editor.commands.register(typing);
    editor.execute('test.type');
    expect(await editor.executeAsync('file.open', { data: '<p>Aperto</p>', formatId: 'html' })).toBe(true);
    expect(paragraphText(firstParagraph(editor.state.doc))).toBe('Aperto');
    expect(editor.state.selection.focus.paragraphId).toBe(firstParagraph(editor.state.doc).id);
    expect(editor.history.canUndo).toBe(false);
    expect(editor.canExecute('file.open', { data: '', formatId: 'pdf' })).toBe(false);
  });

  it('file.open segnala i file non validi senza cambiare documento', async () => {
    const editor = new Editor();
    editor.use(ioModule);
    const errors: unknown[] = [];
    editor.events.on('error', ({ error }) => errors.push(error));
    const before = editor.state;
    expect(await editor.executeAsync('file.open', { data: new Uint8Array([1, 2, 3]), formatId: 'docx' })).toBe(false);
    expect(errors).toHaveLength(1);
    expect(editor.state).toBe(before);
  });

  it('file.export consegna i dati al gestore di salvataggio', async () => {
    const saved: SaveRequest[] = [];
    const editor = new Editor();
    editor.use(createIoModule({ save: (f) => void saved.push(f) }));
    await editor.executeAsync('file.open', { data: 'riga uno\nriga due', formatId: 'txt' });
    const before = editor.state;
    await editor.executeAsync('file.export', { formatId: 'txt' });
    expect(saved).toEqual([{ data: 'riga uno\r\nriga due', fileName: 'Documento.txt', mimeType: 'text/plain', formatId: 'txt' }]);
    expect(editor.state).toBe(before);

    await editor.executeAsync('file.export', { formatId: 'docx' });
    expect(saved[1]!.fileName).toBe('Documento.docx');
    expect(saved[1]!.data).toBeInstanceOf(Uint8Array);
  });

  it('file.export è disabilitato senza destinazione (nessun DOM) o per formati sconosciuti', () => {
    const editor = new Editor();
    editor.use(ioModule);
    expect(editor.canExecute('file.export', { formatId: 'docx' })).toBe(false);
    const other = new Editor();
    other.use(createIoModule({ save: () => {} }));
    expect(other.canExecute('file.export', { formatId: 'docx' })).toBe(true);
    expect(other.canExecute('file.export', { formatId: 'odt' })).toBe(false);
  });
});
