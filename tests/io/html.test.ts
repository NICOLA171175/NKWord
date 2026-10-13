import { describe, expect, it } from 'vitest';
import { paragraphText, type Paragraph, type TextRun } from '../../src/core';
import { documentToHtml, htmlFormat, htmlToDocument } from '../../src/io';
import { firstTable, plainParagraphs, PNG_DATA_URI, richDocument } from './fixtures';

const runs = (p: Paragraph) => p.children.filter((c): c is TextRun => c.type === 'text');

describe('HTML → modello', () => {
  it('converte paragrafi, titoli e formattazione in linea', () => {
    const doc = htmlToDocument(`
      <h1>Titolo</h1>
      <p style="text-align:center; margin-left: 36pt">Testo <b>grassetto</b> e <span style="color:#00f;font-size:16px;font-family:'Arial', sans-serif">blu</span></p>
      <p><a href="https://x.it">link</a> <s>barrato</s><sup>2</sup></p>`);
    const [h, p1, p2] = plainParagraphs(doc);
    expect(h!.styleId).toBe('Heading1');
    expect(paragraphText(h!)).toBe('Titolo');
    expect(p1!.format).toMatchObject({ align: 'center', indentLeft: 36 });
    expect(paragraphText(p1!)).toBe('Testo grassetto e blu');
    expect(runs(p1!)[1]).toMatchObject({ text: 'grassetto', format: { bold: true } });
    expect(runs(p1!)[3]!.format).toEqual({ color: '#0000ff', fontSize: 12, fontFamily: 'Arial' });
    expect(runs(p2!)[0]).toMatchObject({ text: 'link', format: { link: 'https://x.it' } });
    expect(runs(p2!).find((r) => r.text === 'barrato')!.format.strike).toBe(true);
    expect(runs(p2!).find((r) => r.text === '2')!.format.verticalAlign).toBe('superscript');
  });

  it('comprime gli spazi e gestisce br, nbsp e pre', () => {
    const doc = htmlToDocument('<p>  a \n   b<br>c<br></p><p>&nbsp;</p><pre>x\n\ty</pre>');
    const [p, blank, pre] = plainParagraphs(doc);
    expect(paragraphText(p!)).toBe('a b\nc');
    expect(blank!.children).toEqual([]);
    expect(paragraphText(pre!)).toBe('x\n\ty');
  });

  it('converte elenchi annidati in un solo elenco a più livelli', () => {
    const doc = htmlToDocument('<ol start="3"><li>uno<ul><li>a</li><li>b</li></ul></li><li>due</li></ol><ul><li>altro</li></ul>');
    const ps = plainParagraphs(doc);
    expect(ps.map(paragraphText)).toEqual(['uno', 'a', 'b', 'due', 'altro']);
    const [uno, a, , due, altro] = ps;
    expect(uno!.format.list).toEqual({ listId: due!.format.list!.listId, level: 0 });
    expect(a!.format.list).toEqual({ listId: uno!.format.list!.listId, level: 1 });
    expect(altro!.format.list!.listId).not.toBe(uno!.format.list!.listId);
    const def = doc.lists[uno!.format.list!.listId]!;
    expect(def.levels[0]).toMatchObject({ format: 'decimal', start: 3 });
    expect(def.levels[1]!.format).toBe('bullet');
  });

  it('converte tabelle con celle unite', () => {
    const doc = htmlToDocument(`
      <table border="1">
        <tr><th colspan="2">A</th><th rowspan="2" style="background:#ff0">C</th></tr>
        <tr><td>a2</td><td>b2</td></tr>
      </table>`);
    const t = firstTable(doc);
    expect(t.columnWidths).toHaveLength(3);
    expect(t.format.borders?.insideH).toBeDefined();
    expect(t.rows[0]!.isHeader).toBe(true);
    expect(t.rows[0]!.cells[0]!.colSpan).toBe(2);
    expect(t.rows[0]!.cells[1]!.vMerge).toBe('restart');
    expect(t.rows[0]!.cells[1]!.format.shading).toBe('#ffff00');
    expect(t.rows[1]!.cells.map((c) => c.vMerge)).toEqual([undefined, undefined, 'continue']);
    expect(runs(t.rows[0]!.cells[0]!.blocks[0] as Paragraph)[0]!.format.bold).toBe(true);
  });

  it('legge immagini e ignora script, stili e marcatori di elenco di Word', () => {
    const doc = htmlToDocument(`<html><head><title>T</title><style>p{color:red}</style></head><body>
      <script>alert(1)</script>
      <p class=MsoListParagraph><span style='mso-list:Ignore'>1.<span>&nbsp;&nbsp;</span></span>voce<o:p></o:p></p>
      <p><img src="${PNG_DATA_URI}" width="96" height="48" alt="logo"></p></body></html>`);
    const [voce, img] = plainParagraphs(doc);
    expect(doc.metadata.title).toBe('T');
    expect(paragraphText(voce!)).toBe('voce');
    expect(img!.children[0]).toEqual({ type: 'image', src: PNG_DATA_URI, width: 72, height: 36, alt: 'logo' });
  });

  it('non propaga margini e interruzioni di un contenitore ai paragrafi interni', () => {
    const doc = htmlToDocument('<div style="page-break-before:always;margin-top:20pt;text-align:right"><p>a</p><p>b</p></div>');
    const [a, b] = plainParagraphs(doc);
    expect(a!.format).toEqual({ align: 'right' });
    expect(b!.format).toEqual({ align: 'right' });
  });
});

describe('modello → HTML', () => {
  it('produce un documento completo con titoli, elenchi, tabelle e collegamenti', () => {
    const html = documentToHtml(richDocument());
    expect(html).toMatch(/^<!DOCTYPE html>/);
    expect(html).toContain('<title>Prova</title>');
    expect(html).toMatch(/<h1 class="s-Heading1" data-style="Heading1">Capitolo uno<\/h1>/);
    expect(html).toContain('<b>grassetto</b>');
    expect(html).toContain('<a href="https://example.com/a?b=1&amp;c=2">collegamento</a>');
    expect(html.replace(/\n/g, '')).toContain('<ul><li>Primo punto<ol type="a"><li>Sottopunto</li></ol></li><li>Secondo punto</li></ul>');
    expect(html).toContain('colspan="2"');
    expect(html).toContain('rowspan="2"');
    expect(html).toContain(`src="${PNG_DATA_URI}"`);
    expect(html).toContain('data-field="page"');
  });

  it('esporta e reimporta senza perdere contenuto e struttura', async () => {
    const original = richDocument();
    const html = (await htmlFormat.export!(original)) as string;
    const doc = await htmlFormat.import!(html);
    const before = plainParagraphs(original).map(paragraphText);
    const after = plainParagraphs(doc).map(paragraphText);
    expect(after).toEqual(before);

    const ps = plainParagraphs(doc);
    expect(ps[0]!.styleId).toBe('Title');
    expect(ps[1]!.styleId).toBe('Heading1');
    expect(ps[2]!.format).toMatchObject({ align: 'justify', indentLeft: 18, indentFirstLine: -9, spaceBefore: 6, spaceAfter: 12, lineSpacing: { rule: 'multiple', value: 1.5 } });
    const formats = Object.fromEntries(runs(ps[2]!).map((r) => [r.text, r.format]));
    expect(formats['grassetto']).toEqual({ bold: true });
    expect(formats[' rosso 14']).toEqual({ color: '#ff0000', fontSize: 14, fontFamily: 'Times New Roman' });
    expect(formats['evidenziato']).toEqual({ highlight: '#ffff00' });
    expect(formats['collegamento']).toEqual({ link: 'https://example.com/a?b=1&c=2' });
    expect(ps[3]!.format.list?.level).toBe(0);
    expect(ps[4]!.format.list?.level).toBe(1);
    expect(ps.at(-2)!.children.filter((c) => c.type === 'field').map((c) => c.type === 'field' && c.kind)).toEqual(['page', 'numPages']);

    const t = firstTable(doc);
    expect(t.rows[0]!.isHeader).toBe(true);
    expect(t.rows[0]!.cells[0]!.colSpan).toBe(2);
    expect(t.rows[0]!.cells[0]!.format.shading).toBe('#d9e2f3');
    expect(runs(t.rows[0]!.cells[0]!.blocks[0] as Paragraph)[0]!.format).toEqual({});
    expect(t.rows[1]!.cells[2]!.vMerge).toBe('continue');
    expect(t.columnWidths).toEqual([150, 150, 150]);
  });
});
