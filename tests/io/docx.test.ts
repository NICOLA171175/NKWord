import { describe, expect, it } from 'vitest';
import { paragraphText, resolveCharFormat, type Paragraph, type TextRun } from '../../src/core';
import { docxFormat, exportDocx, importDocx } from '../../src/io';
import { parseXml, textContent, descendant } from '../../src/io/util/markup';
import { readZip, writeZip } from '../../src/io/util/zip';
import { firstTable, plainParagraphs, PNG_DATA_URI, richDocument } from './fixtures';

const runs = (p: Paragraph) => p.children.filter((c): c is TextRun => c.type === 'text');
const enc = new TextEncoder();

describe('DOCX: scrittura', () => {
  it('produce un pacchetto con le parti e le relazioni attese', async () => {
    const zip = await readZip(await exportDocx(richDocument()));
    const names = [...zip.keys()];
    for (const part of [
      '[Content_Types].xml',
      '_rels/.rels',
      'word/document.xml',
      'word/_rels/document.xml.rels',
      'word/styles.xml',
      'word/numbering.xml',
      'word/settings.xml',
      'docProps/core.xml',
      'word/header1.xml',
      'word/footer1.xml',
      'word/media/image1.png',
    ]) {
      expect(names).toContain(part);
    }
    const text = (p: string) => new TextDecoder().decode(zip.get(p));
    expect(text('[Content_Types].xml')).toContain('<Default Extension="png" ContentType="image/png"/>');
    const rels = text('word/_rels/document.xml.rels');
    expect(rels).toContain('Target="https://example.com/a?b=1&amp;c=2" TargetMode="External"');
    expect(rels).toContain('Target="media/image1.png"');
    // Ogni parte XML deve essere ben formata: il parser non deve perdere il nodo radice.
    for (const n of names.filter((n) => n.endsWith('.xml') || n.endsWith('.rels'))) {
      expect(parseXml(text(n)).children.length).toBeGreaterThan(0);
    }
    const doc = text('word/document.xml');
    // Ordine dello schema in w:pPr: spacing prima di ind, ind prima di jc.
    expect(doc).toMatch(/<w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="auto"\/><w:ind w:left="360" w:hanging="180"\/><w:jc w:val="both"\/>/);
    expect(doc).toContain('<w:fldSimple w:instr=" PAGE ">');
    expect(doc).toContain('<w:gridSpan w:val="2"/>');
    expect(doc).toContain('<w:vMerge w:val="restart"/>');
    expect(doc).toContain('<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>');
  });

  it('salva i metadati in docProps/core.xml', async () => {
    const zip = await readZip(await exportDocx(richDocument()));
    const core = parseXml(new TextDecoder().decode(zip.get('docProps/core.xml')));
    expect(textContent(descendant(core, 'dc:title')!)).toBe('Prova');
    expect(textContent(descendant(core, 'dc:creator')!)).toBe('Nicola');
    expect(textContent(descendant(core, 'dcterms:created')!)).toBe('2026-01-02T03:04:05Z');
  });
});

describe('DOCX: andata e ritorno', () => {
  it('conserva pedice e apice, anche col testo normale accanto', async () => {
    const { createDocument, paragraph, textRun } = await import('../../src/core');
    const original = createDocument([
      paragraph([textRun('H'), textRun('2', { verticalAlign: 'subscript' }), textRun('O e x'), textRun('2', { verticalAlign: 'superscript' }), textRun(' fine', { verticalAlign: 'baseline' })]),
    ]);
    const doc = await docxFormat.import!(await docxFormat.export!(original));
    const rs = runs(plainParagraphs(doc)[0]!);
    expect(rs.map((r) => [r.text, r.format.verticalAlign ?? 'baseline'])).toEqual([
      ['H', 'baseline'],
      ['2', 'subscript'],
      ['O e x', 'baseline'],
      ['2', 'superscript'],
      [' fine', 'baseline'],
    ]);
  });

  it('conserva testo, stili, formattazione, elenchi, tabelle, immagini, campi e sezioni', async () => {
    const original = richDocument();
    const doc = await docxFormat.import!(await docxFormat.export!(original));

    expect(plainParagraphs(doc).map(paragraphText)).toEqual(plainParagraphs(original).map(paragraphText));
    expect(doc.metadata).toMatchObject({ title: 'Prova', author: 'Nicola', subject: 'Test', keywords: ['uno', 'due'] });

    const ps = plainParagraphs(doc);
    expect(ps[0]!.styleId).toBe('Title');
    expect(ps[1]!.styleId).toBe('Heading1');
    expect(doc.styles.styles.Heading1!.name).toBe('Titolo 1');
    expect(resolveCharFormat(doc.styles, ps[1]!)).toMatchObject({ bold: true, fontSize: 16, color: '#2f5496', fontFamily: 'Calibri' });

    expect(ps[2]!.format).toEqual({ align: 'justify', indentLeft: 18, indentFirstLine: -9, spaceBefore: 6, spaceAfter: 12, lineSpacing: { rule: 'multiple', value: 1.5 } });
    const formats = Object.fromEntries(runs(ps[2]!).map((r) => [r.text, r.format]));
    expect(formats['grassetto']).toEqual({ bold: true });
    expect(formats['corsivo sottolineato']).toEqual({ italic: true, underline: 'single' });
    expect(formats[' rosso 14']).toEqual({ color: '#ff0000', fontSize: 14, fontFamily: 'Times New Roman' });
    expect(formats['evidenziato']).toEqual({ highlight: '#ffff00' });
    expect(formats['2']).toEqual({ verticalAlign: 'superscript' });
    expect(formats['collegamento']).toEqual({ link: 'https://example.com/a?b=1&c=2' });
    expect(formats[' e barrato']).toEqual({ strike: true });

    const listId = ps[3]!.format.list!.listId;
    expect(ps[4]!.format.list).toEqual({ listId, level: 1 });
    expect(doc.lists[listId]!.levels[0]).toMatchObject({ format: 'bullet', text: '•', indentLeft: 36, hanging: 18 });
    expect(doc.lists[listId]!.levels[1]).toMatchObject({ format: 'lowerLetter', text: '%2)' });

    const image = ps[6]!.children[1]!;
    expect(image).toMatchObject({ type: 'image', width: 72, height: 36, alt: 'Logo' });
    expect(doc.media[(image as { src: string }).src]).toBe(PNG_DATA_URI);

    const t = firstTable(doc);
    expect(t.columnWidths).toEqual([150, 150, 150]);
    expect(t.format.align).toBe('center');
    expect(t.format.borders?.insideV).toEqual({ style: 'single', width: 0.5, color: '#000000' });
    expect(t.rows[0]!.isHeader).toBe(true);
    expect(t.rows[0]!.cells[0]).toMatchObject({ colSpan: 2, format: { shading: '#d9e2f3' } });
    expect(t.rows[0]!.cells[1]).toMatchObject({ vMerge: 'restart', format: { verticalAlign: 'center' } });
    expect(t.rows[1]!.cells[2]!.vMerge).toBe('continue');

    const fields = ps.at(-2)!.children.filter((c) => c.type === 'field');
    expect(fields.map((f) => f.type === 'field' && f.kind)).toEqual(['page', 'numPages']);

    expect(doc.sections).toHaveLength(2);
    const [s1, s2] = doc.sections;
    expect(paragraphText(s1!.header!.default![0] as Paragraph)).toBe('Intestazione');
    expect((s1!.footer!.default![0] as Paragraph).children[1]).toMatchObject({ type: 'field', kind: 'page' });
    // I twip arrotondano a 1/20 di punto.
    const round = (o: object) => JSON.parse(JSON.stringify(o), (_k, v) => (typeof v === 'number' ? Math.round(v) : v));
    expect(round(s1!.page)).toEqual(round(original.sections[0]!.page));
    expect(s2!.page.orientation).toBe('landscape');
    expect(s2!.breakType).toBe('nextPage');
    expect(s2!.header).toEqual(s1!.header);
  });
});

/** DOCX scritto a mano come lo produrrebbe Word in italiano. */
async function wordLikeDocx(): Promise<Uint8Array> {
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const document = `<?xml version="1.0" encoding="UTF-8"?>
<w:document ${W}><w:body>
  <w:p><w:pPr><w:pStyle w:val="Titolo1"/></w:pPr><w:r><w:t>Intro</w:t></w:r></w:p>
  <w:p><w:pPr><w:pStyle w:val="Paragrafoelenco"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr></w:pPr><w:r><w:t>voce</w:t></w:r></w:p>
  <w:p>
    <w:r><w:t xml:space="preserve">Pagina </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE \\* MERGEFORMAT </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>7</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:t xml:space="preserve"> vedi </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>HYPERLINK "https://nik.it"</w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:rStyle w:val="Collegamentoipertestuale"/></w:rPr><w:t>sito</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:ins w:id="1" w:author="x"><w:r><w:t xml:space="preserve"> aggiunto</w:t></w:r></w:ins>
    <w:del w:id="2" w:author="x"><w:r><w:delText>tolto</w:delText></w:r></w:del>
    <w:hyperlink r:id="rId9"><w:r><w:t xml:space="preserve"> rel</w:t></w:r></w:hyperlink>
  </w:p>
  <w:sdt><w:sdtContent><w:p><w:r><w:rPr><w:b/><w:b w:val="0"/></w:rPr><w:t>in controllo</w:t></w:r></w:p></w:sdtContent></w:sdt>
  <w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:num="2" w:space="720"/></w:sectPr></w:pPr><w:r><w:t>fine sezione 1</w:t></w:r></w:p>
  <w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/></w:tblGrid>
    <w:tr><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc></w:tr>
    <w:tr><w:tc><w:p/></w:tc><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc></w:tr>
  </w:tbl>
  <w:sectPr><w:titlePg/><w:pgNumType w:start="3"/></w:sectPr>
</w:body></w:document>`;
  const styles = `<?xml version="1.0" encoding="UTF-8"?>
<w:styles ${W}>
  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
  <w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normale"><w:name w:val="Normal"/><w:qFormat/></w:style>
  <w:style w:type="paragraph" w:styleId="Titolo1"><w:name w:val="heading 1"/><w:basedOn w:val="Normale"/><w:next w:val="Normale"/><w:pPr><w:keepNext/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/><w:sz w:val="32"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Paragrafoelenco"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normale"/><w:pPr><w:ind w:left="720"/></w:pPr></w:style>
  <w:style w:type="character" w:styleId="Collegamentoipertestuale"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>
  <w:style w:type="numbering" w:styleId="Nessunelenco"><w:name w:val="No List"/></w:style>
</w:styles>`;
  const numbering = `<?xml version="1.0" encoding="UTF-8"?>
<w:numbering ${W}>
  <w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val=""/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr></w:lvl></w:abstractNum>
  <w:num w:numId="5"><w:abstractNumId w:val="2"/></w:num>
</w:numbering>`;
  const theme = `<?xml version="1.0"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`;
  const rels = `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
  <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="mailto:a@b.it" TargetMode="External"/>
</Relationships>`;
  const pkgRels = `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="/word/document.xml"/></Relationships>`;
  return writeZip([
    ['_rels/.rels', enc.encode(pkgRels)],
    ['word/document.xml', enc.encode(document)],
    ['word/_rels/document.xml.rels', enc.encode(rels)],
    ['word/styles.xml', enc.encode(styles)],
    ['word/numbering.xml', enc.encode(numbering)],
    ['word/theme/theme1.xml', enc.encode(theme)],
  ]);
}

describe('DOCX: lettura di file di Word', () => {
  it('riconosce gli stili predefiniti localizzati e i caratteri del tema', async () => {
    const doc = await importDocx(await wordLikeDocx());
    const ps = plainParagraphs(doc);
    expect(ps[0]!.styleId).toBe('Heading1');
    expect(doc.styles.defaultParagraphStyle).toBe('Normal');
    expect(doc.styles.defaults.char).toMatchObject({ fontFamily: 'Aptos', fontSize: 11 });
    expect(doc.styles.defaults.para.lineSpacing).toEqual({ rule: 'multiple', value: 1.08 });
    expect(resolveCharFormat(doc.styles, ps[0]!)).toMatchObject({ fontFamily: 'Calibri Light', fontSize: 16 });
    expect(doc.styles.styles.Heading1).toMatchObject({ basedOn: 'Normal', next: 'Normal', paraFormat: { keepWithNext: true, outlineLevel: 1 } });
    expect(doc.styles.styles.Nessunelenco).toBeUndefined();
  });

  it('legge elenchi con punti elenco Symbol', async () => {
    const doc = await importDocx(await wordLikeDocx());
    const voce = plainParagraphs(doc)[1]!;
    expect(voce.styleId).toBe('Paragrafoelenco');
    expect(voce.format.list).toEqual({ listId: 'num5', level: 0 });
    expect(doc.lists.num5!.levels[0]).toMatchObject({ format: 'bullet', text: '•', indentLeft: 36, hanging: 18 });
    expect(doc.lists.num5!.levels[0]!.charFormat).toBeUndefined();
  });

  it('legge campi complessi, collegamenti e revisioni', async () => {
    const doc = await importDocx(await wordLikeDocx());
    const p = plainParagraphs(doc)[2]!;
    // Le revisioni restano in attesa: il testo eliminato c'è, segnato come tale.
    expect(paragraphText(p)).toBe('Pagina ￼ vedi sito aggiuntotolto rel');
    expect(runs(p).find((r) => r.text === ' aggiunto')!.revision).toEqual({ kind: 'insert', author: 'x' });
    expect(runs(p).find((r) => r.text === 'tolto')!.revision).toEqual({ kind: 'delete', author: 'x' });
    expect(p.children[1]).toMatchObject({ type: 'field', kind: 'page' });
    const sito = runs(p).find((r) => r.text === 'sito')!;
    expect(sito.format.link).toBe('https://nik.it');
    expect(sito.styleId).toBe('Hyperlink');
    expect(runs(p).find((r) => r.text === ' rel')!.format.link).toBe('mailto:a@b.it');
    const sdt = plainParagraphs(doc)[3]!;
    expect(paragraphText(sdt)).toBe('in controllo');
    expect(runs(sdt)[0]!.format.bold).toBe(true);
  });

  it('divide le sezioni e legge le celle unite', async () => {
    const doc = await importDocx(await wordLikeDocx());
    expect(doc.sections).toHaveLength(2);
    const [s1, s2] = doc.sections;
    expect(s1!.blocks.map((b) => (b.type === 'paragraph' ? paragraphText(b) : 'tabella'))).toEqual(['Intro', 'voce', expect.any(String), 'in controllo', 'fine sezione 1']);
    expect(s1!.page).toMatchObject({ width: 612, height: 792, orientation: 'portrait', columns: { count: 2, gap: 36 } });
    expect(s1!.page.margins).toMatchObject({ top: 72, left: 72, header: 36 });
    expect(s2!.titlePage).toBe(true);
    expect(s2!.pageNumberStart).toBe(3);
    const t = firstTable(doc);
    expect(t.columnWidths).toEqual([100, 150]);
    expect(t.rows[0]!.cells[1]!.vMerge).toBe('restart');
    expect(t.rows[1]!.cells[1]!.vMerge).toBe('continue');
    expect(t.rows[1]!.cells[0]!.blocks).toHaveLength(1);
  });

  it('rifiuta un archivio che non contiene un documento', async () => {
    const zip = await writeZip([['a.txt', enc.encode('x')]]);
    await expect(importDocx(zip)).rejects.toThrow(/documento Word/);
  });
});

describe('DOCX: file prodotti da altri programmi', () => {
  const load = async (name: string) => {
    // Il progetto non include i tipi di Node: import dinamico non tipizzato.
    const fsModule = 'node:fs/promises';
    const { readFile } = await import(/* @vite-ignore */ fsModule);
    return importDocx(new Uint8Array(await readFile(new URL(`./files/${name}`, import.meta.url))));
  };

  it('legge un documento salvato da LibreOffice', async () => {
    const doc = await load('libreoffice.docx');
    const ps = plainParagraphs(doc);
    expect(ps.map(paragraphText)).toEqual(plainParagraphs(richDocument()).map(paragraphText));
    expect(ps[0]!.styleId).toBe('Title');
    expect(ps[1]!.styleId).toBe('Heading1');
    expect(runs(ps[2]!).find((r) => r.text === 'grassetto')!.format).toEqual({ bold: true });
    expect(ps[4]!.format.list?.level).toBe(1);
    expect(Object.values(doc.media)[0]).toBe(PNG_DATA_URI);
    expect(doc.sections).toHaveLength(2);
    expect(doc.sections[1]!.page.orientation).toBe('landscape');
    expect(paragraphText(doc.sections[0]!.header!.default![0] as Paragraph)).toBe('Intestazione');
  });

  it('legge un documento prodotto da pandoc', async () => {
    const doc = await load('pandoc.docx');
    const ps = plainParagraphs(doc);
    expect(ps.map(paragraphText)).toEqual(['Titolo', 'Testo grassetto e corsivo con link.', 'uno', 'due', 'annidato', 'A', 'B', '1', '2']);
    expect(ps[0]!.styleId).toBe('Heading1');
    expect(runs(ps[1]!).find((r) => r.text === 'link')!.format.link).toBe('https://pandoc.org');
    expect(ps[4]!.format.list?.level).toBe(1);
    expect(firstTable(doc).rows).toHaveLength(2);
  });
});
