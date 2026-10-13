import { describe, expect, it } from 'vitest';
import { paragraph, paragraphText, textRun, type Paragraph, type TextRun } from '../../src/core';
import { documentToHtml, exportDocx, importDocx } from '../../src/io';
import { readZip, writeZip } from '../../src/io/util/zip';
import { createApproxMeasurer, createLayoutEngine } from '../../src/layout';
import { createLinksModule, findAutoLink, linkAt, normalizeAddress, TOP_ANCHOR, type LinkDialogInput, type LinkDialogResult } from '../../src/links';
import { layoutToPdf } from '../../src/pdf';
import { tocAnchor } from '../../src/toc';
import { paras, select, setup } from '../editing/helpers';

const runs = (p: Paragraph) => p.children.filter((c): c is TextRun => c.type === 'text');
const linked = (p: Paragraph) => runs(p).filter((r) => r.format.link).map((r) => [r.text, r.format.link, r.styleId]);

function linkEditor(blocks: Paragraph[], options: Parameters<typeof createLinksModule>[0] = {}) {
  const editor = setup(blocks);
  editor.use(createLinksModule(options));
  return editor;
}

describe('indirizzi', () => {
  it('completa gli indirizzi scritti a mano', () => {
    expect(normalizeAddress('www.nik.it')).toBe('https://www.nik.it');
    expect(normalizeAddress('  http://a.it ')).toBe('http://a.it');
    expect(normalizeAddress('nome@posta.it')).toBe('mailto:nome@posta.it');
    expect(normalizeAddress('mario@x.it', 'email')).toBe('mailto:mario@x.it');
    expect(normalizeAddress('#_top')).toBe('#_top');
    expect(normalizeAddress('   ')).toBeNull();
  });

  it('riconosce indirizzi web ed email appena scritti, senza la punteggiatura finale', () => {
    const p = paragraph('vedi (www.nik.it), poi');
    expect(findAutoLink(p, 17)).toEqual({ start: 6, end: 16, url: 'https://www.nik.it' });
    expect(findAutoLink(paragraph('scrivi a anna@posta.it'), 22)).toMatchObject({ url: 'mailto:anna@posta.it' });
    expect(findAutoLink(paragraph('https://a.it/x?y=1'), 18)).toMatchObject({ start: 0, url: 'https://a.it/x?y=1' });
    expect(findAutoLink(paragraph('parola normale'), 14)).toBeNull();
    expect(findAutoLink(paragraph('http://'), 7)).toBeNull();
  });
});

describe('comandi dei collegamenti', () => {
  it('collega il testo selezionato con lo stile Hyperlink', () => {
    const p = paragraph('il mio sito oggi');
    const editor = linkEditor([p]);
    select(editor, [p, 7], [p, 11]);
    editor.execute('link.insert', { url: 'https://nik.it' });
    expect(linked(paras(editor)[0]!)).toEqual([['sito', 'https://nik.it', 'Hyperlink']]);
    expect(editor.state.doc.styles.styles.Hyperlink).toMatchObject({ type: 'character', charFormat: { color: '#0563C1', underline: 'single' } });
  });

  it('senza selezione inserisce il testo, o l\'indirizzo stesso', () => {
    const p = paragraph('Vai: ');
    const editor = linkEditor([p]);
    select(editor, [p, 5]);
    editor.execute('link.insert', { url: 'mailto:a@b.it' });
    expect(paragraphText(paras(editor)[0]!)).toBe('Vai: a@b.it');
    editor.execute('link.insert', { url: 'https://x.it', text: ' qui' });
    expect(paragraphText(paras(editor)[0]!)).toBe('Vai: a@b.it qui');
    expect(editor.state.selection.focus.offset).toBe(15);
  });

  it('modifica testo e indirizzo di un collegamento e lo rimuove', () => {
    const p = paragraph([textRun('vai al '), textRun('sito', { link: 'https://a.it' }), textRun(' ora')]);
    const editor = linkEditor([p]);
    select(editor, [p, 9]);
    expect(linkAt(editor.state.doc, editor.state.selection.focus)).toMatchObject({ start: 7, end: 11, text: 'sito' });
    editor.execute('link.insert', { url: 'https://b.it', text: 'nuovo sito' });
    expect(paragraphText(paras(editor)[0]!)).toBe('vai al nuovo sito ora');
    expect(linked(paras(editor)[0]!)).toEqual([['nuovo sito', 'https://b.it', 'Hyperlink']]);

    select(editor, [paras(editor)[0]!, 9]);
    editor.execute('link.remove');
    expect(runs(paras(editor)[0]!)).toEqual([textRun('vai al nuovo sito ora')]);
  });

  it('il testo scritto alla fine di un collegamento non ne fa parte', () => {
    const p = paragraph('a ');
    const editor = linkEditor([p]);
    select(editor, [p, 2]);
    editor.execute('link.insert', { url: 'https://a.it', text: 'sito' });
    editor.execute('edit.insertText', { text: 'x' });
    const ps = runs(paras(editor)[0]!);
    expect(ps.map((r) => [r.text, r.format.link, r.styleId])).toEqual([
      ['a ', undefined, undefined],
      ['sito', 'https://a.it', 'Hyperlink'],
      ['x', undefined, undefined],
    ]);
    // Dentro il collegamento invece si allarga.
    select(editor, [paras(editor)[0]!, 4]);
    editor.execute('edit.insertText', { text: 'I' });
    expect(linked(paras(editor)[0]!)).toEqual([['siIto', 'https://a.it', 'Hyperlink']]);
  });

  it('riconosce l\'indirizzo dopo lo spazio o l\'Invio; Ctrl+Z toglie solo il collegamento', () => {
    const p = paragraph('');
    const editor = linkEditor([p]);
    select(editor, [p, 0]);
    editor.execute('edit.insertText', { text: 'www.nik.it' });
    expect(linked(paras(editor)[0]!)).toEqual([]);
    editor.execute('edit.insertText', { text: ' ' });
    expect(linked(paras(editor)[0]!)).toEqual([['www.nik.it', 'https://www.nik.it', 'Hyperlink']]);
    editor.undo();
    expect(paragraphText(paras(editor)[0]!)).toBe('www.nik.it ');
    expect(linked(paras(editor)[0]!)).toEqual([]);

    editor.execute('edit.insertText', { text: 'mario@posta.it' });
    editor.execute('edit.splitParagraph');
    expect(linked(paras(editor)[0]!)).toEqual([['mario@posta.it', 'mailto:mario@posta.it', 'Hyperlink']]);
  });

  it('si può spegnere il riconoscimento automatico', () => {
    const p = paragraph('');
    const editor = linkEditor([p], { autoLink: false });
    select(editor, [p, 0]);
    editor.execute('edit.insertText', { text: 'www.nik.it' });
    editor.execute('edit.insertText', { text: ' ' });
    expect(linked(paras(editor)[0]!)).toEqual([]);
  });

  it('apre gli indirizzi fuori dall\'app e salta ai punti del documento', async () => {
    const opened: string[] = [];
    const title = paragraph('Capitolo', {}, 'Heading1');
    const first = paragraph('inizio');
    const editor = linkEditor([first, paragraph('testo'), title], { openExternal: (u) => opened.push(u) });
    await editor.executeAsync('link.open', { url: 'https://nik.it' });
    expect(opened).toEqual(['https://nik.it']);
    expect(editor.canExecute('link.open', { url: 'javascript:alert(1)' })).toBe(false);

    await editor.executeAsync('link.open', { url: tocAnchor(title.id) });
    expect(editor.state.selection.focus).toEqual({ paragraphId: title.id, offset: 0 });
    await editor.executeAsync('link.open', { url: TOP_ANCHOR });
    expect(editor.state.selection.focus).toEqual({ paragraphId: first.id, offset: 0 });
    expect(editor.canExecute('link.open', { url: '#_Tocnonc\'è' })).toBe(false);
  });

  it('Ctrl+K: la finestra riceve il collegamento esistente e ne applica le modifiche', async () => {
    const seen: LinkDialogInput[] = [];
    let answer: LinkDialogResult = { url: 'https://b.it', text: 'B' };
    const p = paragraph([textRun('x '), textRun('A', { link: 'https://a.it' })]);
    const editor = linkEditor([p], { openDialog: async (input) => (seen.push(input), answer) });
    // Il collegamento selezionato (come fa "Modifica collegamento" del clic destro).
    select(editor, [p, 2], [p, 3]);
    await editor.executeAsync('link.dialog');
    expect(seen[0]).toMatchObject({ url: 'https://a.it', text: 'A' });
    expect(linked(paras(editor)[0]!)).toEqual([['B', 'https://b.it', 'Hyperlink']]);

    // Con il cursore subito dopo il collegamento se ne inserisce uno nuovo.
    answer = null;
    await editor.executeAsync('link.dialog');
    expect(seen[1]).toEqual({ doc: editor.state.doc, text: '' });

    answer = { remove: true };
    select(editor, [paras(editor)[0]!, 2], [paras(editor)[0]!, 3]);
    await editor.executeAsync('link.dialog');
    expect(linked(paras(editor)[0]!)).toEqual([]);

    // Selezione su più paragrafi: niente testo da modificare.
    const q = paragraph('due');
    const other = linkEditor([paragraph('uno'), q], { openDialog: async (input) => (seen.push(input), null) });
    other.execute('edit.selectAll');
    await other.executeAsync('link.dialog');
    expect(seen[3]!.text).toBeUndefined();
  });
});

describe('collegamenti nei file', () => {
  it('DOCX: stile Hyperlink e collegamenti ai titoli che sopravvivono alla riapertura', async () => {
    const title = paragraph('Capitolo', {}, 'Heading1');
    const p = paragraph('vai al capitolo o al sito');
    const editor = linkEditor([p, title]);
    select(editor, [p, 7], [p, 15]);
    editor.execute('link.insert', { url: tocAnchor(title.id) });
    select(editor, [paras(editor)[0]!, 21], [paras(editor)[0]!, 25]);
    editor.execute('link.insert', { url: 'https://nik.it' });

    const bytes = await exportDocx(editor.state.doc);
    const zip = await readZip(bytes);
    const styles = new TextDecoder().decode(zip.get('word/styles.xml'));
    expect(styles).toContain('<w:name w:val="Hyperlink"/>');

    const back = await importDocx(bytes);
    const [bp, bt] = back.sections[0]!.blocks as Paragraph[];
    expect(linked(bp!).map(([text, url]) => [text, url])).toEqual([
      ['capitolo', tocAnchor(bt!.id)],
      ['sito', 'https://nik.it'],
    ]);
    expect(runs(bp!).find((r) => r.text === 'sito')!.styleId).toBe('Hyperlink');
  });

  it('DOCX di Word: i collegamenti a un segnalibro puntano al suo paragrafo', async () => {
    const title = paragraph('Capitolo', {}, 'Heading1');
    const p = paragraph([textRun('vai', { link: tocAnchor(title.id) })]);
    const zip = await readZip(await exportDocx({ ...setup([p, title]).state.doc }));
    const xml = new TextDecoder().decode(zip.get('word/document.xml'));
    const renamed = xml.replaceAll(`_Toc${title.id}`, 'Introduzione');
    zip.set('word/document.xml', new TextEncoder().encode(renamed));
    const back = await importDocx(await writeZip([...zip]));
    const [bp, bt] = back.sections[0]!.blocks as Paragraph[];
    expect(runs(bp!)[0]!.format.link).toBe(tocAnchor(bt!.id));
  });

  it('HTML: i punti del documento ricevono l\'id del collegamento', () => {
    const title = paragraph('Capitolo', {}, 'Heading1');
    const p = paragraph([textRun('vai', { link: tocAnchor(title.id) })]);
    const html = documentToHtml(setup([p, title]).state.doc);
    expect(html).toContain(`<a href="#_Toc${title.id}">vai</a>`);
    expect(html).toContain(`id="_Toc${title.id}"`);
  });

  it('PDF: collegamenti web e interni cliccabili', async () => {
    const measurer = createApproxMeasurer();
    const title = paragraph('Capitolo', {}, 'Heading1');
    const doc = setup([
      paragraph([textRun('sito', { link: 'https://nik.it' }), textRun(' e '), textRun('capitolo', { link: tocAnchor(title.id) })]),
      title,
    ]).state.doc;
    const pdf = new TextDecoder('latin1').decode(await layoutToPdf(createLayoutEngine(measurer).layout(doc), { measurer }));
    expect(pdf).toContain('/URI (https://nik.it)');
    expect(pdf).toMatch(/\/Subtype \/Link \/Rect \[[^\]]+\] \/Border \[0 0 0\] \/Dest \[\d+ 0 R \/XYZ null [\d.]+ null\]/);
  });
});
