import { describe, expect, it } from 'vitest';
import { Editor, caret, createDocument, createState, paragraph, textRun, type DocumentModel } from '../../src/core';
import { createHeaderFooterModule, readHeaderFooter, ribbonItems, type HeaderFooterDialog } from '../../src/headerfooter';
import { exportDocx, importDocx } from '../../src/io';
import { structuresModule } from '../../src/structures';

function editorWith(doc: DocumentModel = createDocument([paragraph('Corpo del documento')]), dialog?: HeaderFooterDialog): Editor {
  const first = doc.sections[0]!.blocks[0]!;
  const editor = new Editor({ ...createState(doc), selection: caret(first.id, 4) });
  editor.use(structuresModule);
  editor.use(createHeaderFooterModule({ dialog }));
  editor.events.on('error', (e) => {
    throw e.error;
  });
  return editor;
}

const header = (e: Editor) => readHeaderFooter(e.state.doc.sections[0]!, 'header');
const footer = (e: Editor) => readHeaderFooter(e.state.doc.sections[0]!, 'footer');

describe('intestazione e piè di pagina', () => {
  it("Inserisci > Intestazione scrive nell'intestazione, non nel corpo, e si annulla", async () => {
    const seen: unknown[] = [];
    const editor = editorWith(undefined, async (data) => {
      seen.push(data);
      return { text: 'Relazione\nanno 2026', align: 'right', pageNumber: false };
    });
    const body = JSON.stringify(editor.state.doc.sections[0]!.blocks);
    await editor.executeAsync('headerFooter.edit', { part: 'header' });
    expect(seen).toEqual([{ part: 'header', current: { text: '', align: 'left', pageNumber: false, complex: false } }]);
    expect(header(editor)).toEqual({ text: 'Relazione\nanno 2026', align: 'right', pageNumber: false, complex: false });
    expect(JSON.stringify(editor.state.doc.sections[0]!.blocks)).toBe(body);
    editor.undo();
    expect(editor.state.doc.sections[0]!.header).toBeUndefined();
  });

  it('Annulla nella finestra non cambia nulla; Rimuovi svuota la parte', async () => {
    let answer: Awaited<ReturnType<HeaderFooterDialog>> = null;
    const editor = editorWith(undefined, async () => answer);
    await editor.executeAsync('headerFooter.edit', { part: 'footer' });
    expect(editor.state.doc.sections[0]!.footer).toBeUndefined();
    editor.execute('headerFooter.set', { part: 'footer', text: 'Mario Rossi', align: 'left', pageNumber: true });
    expect(footer(editor)).toMatchObject({ text: 'Mario Rossi', pageNumber: true });
    answer = { text: '', align: 'left', pageNumber: false };
    await editor.executeAsync('headerFooter.edit', { part: 'footer' });
    expect(editor.state.doc.sections[0]!.footer).toBeUndefined();
  });

  it('Numero di pagina: in fondo al centro, in alto accanto al testo, oppure al cursore', () => {
    const editor = editorWith();
    editor.execute('headerFooter.pageNumber', { position: 'footer' });
    expect(footer(editor)).toMatchObject({ text: '', align: 'center', pageNumber: true });
    expect(editor.state.doc.sections[0]!.footer!.default![0]).toMatchObject({ children: [{ type: 'field', kind: 'page' }] });
    // Un secondo clic non ne aggiunge un altro.
    const before = editor.state.doc;
    editor.execute('headerFooter.pageNumber', { position: 'footer' });
    expect(editor.state.doc).toBe(before);

    editor.execute('headerFooter.set', { part: 'header', text: 'Titolo', align: 'left', pageNumber: false });
    editor.execute('headerFooter.pageNumber', { position: 'header' });
    expect(editor.state.doc.sections[0]!.header!.default![0]).toMatchObject({
      children: [{ type: 'text', text: 'Titolo ' }, { type: 'field', kind: 'page' }],
    });

    editor.execute('headerFooter.pageNumber', { position: 'here' });
    expect(editor.state.doc.sections[0]!.blocks[0]).toMatchObject({ children: [{ text: 'Corp' }, { type: 'field', kind: 'page' }, { text: 'o del documento' }] });
  });

  it('vale per tutte le sezioni e conserva prima pagina e pagine pari', () => {
    const doc = createDocument([paragraph('uno')]);
    const two: DocumentModel = {
      ...doc,
      sections: [
        { ...doc.sections[0]!, header: { first: [paragraph('Copertina')] } },
        { ...doc.sections[0]!, id: 's2', blocks: [paragraph('due')] },
      ],
    };
    const editor = editorWith(two);
    editor.execute('headerFooter.set', { part: 'header', text: 'Capitolo', align: 'center', pageNumber: false });
    for (const s of editor.state.doc.sections) expect(readHeaderFooter(s, 'header').text).toBe('Capitolo');
    expect(editor.state.doc.sections[0]!.header!.first).toHaveLength(1);
  });

  it('un contenuto ricco (es. da DOCX) è segnalato e il numero di pagina si aggiunge senza toccarlo', () => {
    const doc = createDocument([paragraph('x')]);
    const rich = paragraph([textRun('Ditta', { bold: true })]);
    const editor = editorWith({ ...doc, sections: [{ ...doc.sections[0]!, header: { default: [rich] } }] });
    expect(header(editor)).toMatchObject({ text: 'Ditta', complex: true });
    editor.execute('headerFooter.pageNumber', { position: 'header' });
    const blocks = editor.state.doc.sections[0]!.header!.default!;
    expect(blocks[0]).toBe(rich);
    expect(blocks[1]).toMatchObject({ format: { align: 'center' }, children: [{ type: 'field', kind: 'page' }] });
  });

  it('si salva in DOCX e si riapre uguale', async () => {
    const editor = editorWith();
    editor.execute('headerFooter.set', { part: 'header', text: 'Relazione', align: 'right', pageNumber: false });
    editor.execute('headerFooter.pageNumber', { position: 'footer' });
    const reopened = await importDocx(await exportDocx(editor.state.doc));
    expect(readHeaderFooter(reopened.sections[0]!, 'header')).toEqual({ text: 'Relazione', align: 'right', pageNumber: false, complex: false });
    expect(readHeaderFooter(reopened.sections[0]!, 'footer')).toMatchObject({ align: 'center', pageNumber: true });
  });

  it('il pulsante Numero di pagina offre in alto, in basso e al cursore', () => {
    expect(ribbonItems[0]!.options.map((o) => o.label)).toEqual(['Inizio pagina', 'Fine pagina', 'Posizione corrente']);
  });
});
