import { describe, expect, it } from 'vitest';
import { Editor, caret, createDocument, createState, paragraph, paragraphText, type Paragraph, type TextRun } from '../../src/core';
import { exportDocx, importDocx } from '../../src/io';
import { readZip, writeZip } from '../../src/io/util/zip';
import { findToc, tocAnchor, tocModule } from '../../src/toc';

const enc = new TextEncoder();
const h = (level: number, text: string) => paragraph(text, {}, `Heading${level}`);

function withToc() {
  const doc = createDocument([paragraph(''), h(1, 'Capitolo'), paragraph('testo'), h(2, 'Paragrafo')]);
  const editor = new Editor({ ...createState(doc), selection: caret(doc.sections[0]!.blocks[0]!.id, 0) });
  editor.use(tocModule);
  editor.execute('toc.insert');
  return editor.state.doc;
}

const bodyOf = (doc: { sections: readonly { blocks: readonly unknown[] }[] }) => doc.sections[0]!.blocks as readonly Paragraph[];

describe('sommario in DOCX', () => {
  it('scrive il campo TOC, gli stili di Word e i segnalibri dei titoli', async () => {
    const doc = withToc();
    const zip = await readZip(await exportDocx(doc));
    const xml = new TextDecoder().decode(zip.get('word/document.xml'));
    const styles = new TextDecoder().decode(zip.get('word/styles.xml'));

    expect(xml).toMatch(/<w:fldChar w:fldCharType="begin"\/><\/w:r><w:r><w:instrText xml:space="preserve"> TOC \\o &quot;1-3&quot; \\h \\z \\u <\/w:instrText>/);
    expect(xml.match(/fldCharType="begin"/g)).toHaveLength(1);
    expect(xml.match(/fldCharType="end"/g)).toHaveLength(1);
    // Il campo si chiude nell'ultima voce, prima del primo titolo.
    expect(xml.indexOf('fldCharType="end"')).toBeLessThan(xml.indexOf('w:bookmarkStart'));

    const heading = bodyOf(doc).find((p) => paragraphText(p) === 'Capitolo')!;
    const name = tocAnchor(heading.id).slice(1);
    expect(xml).toContain(`<w:bookmarkStart w:id="0" w:name="${name}"/>`);
    expect(xml).toContain(`<w:hyperlink w:anchor="${name}">`);
    expect(styles).toContain('<w:name w:val="toc 1"/>');
    expect(styles).toContain('<w:name w:val="TOC Heading"/>');
  });

  it('rilegge il proprio sommario e lo ricollega ai titoli aggiornandolo', async () => {
    const back = await importDocx(await exportDocx(withToc()));
    const at = findToc(back)!;
    expect(at).toBeDefined();
    expect(bodyOf(back).slice(at.start, at.end).map(paragraphText)).toEqual(['Sommario', 'Capitolo\t1', 'Paragrafo\t1']);

    const editor = new Editor(createState(back));
    editor.use(tocModule);
    editor.execute('toc.update');
    const doc = editor.state.doc;
    const heading = bodyOf(doc).find((p) => paragraphText(p) === 'Capitolo')!;
    const entry = bodyOf(doc)[findToc(doc)!.start + 1]!;
    expect((entry.children[0] as TextRun).format.link).toBe(tocAnchor(heading.id));
  });

  it('riconosce il sommario prodotto da Word (campo TOC, PAGEREF, stili localizzati)', async () => {
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
    const entry = (style: string, text: string, anchor: string, page: string, first = false, last = false) => `
      <w:p><w:pPr><w:pStyle w:val="${style}"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9628"/></w:tabs></w:pPr>
        ${first ? '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' : ''}
        <w:hyperlink w:anchor="${anchor}" w:history="1">
          <w:r><w:t>${text}</w:t></w:r><w:r><w:tab/></w:r>
          <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGEREF ${anchor} \\h </w:instrText></w:r>
          <w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${page}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>
        </w:hyperlink>
        ${last ? '<w:r><w:fldChar w:fldCharType="end"/></w:r>' : ''}
      </w:p>`;
    const document = `<?xml version="1.0" encoding="UTF-8"?>
<w:document ${W}><w:body>
  <w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>
    <w:p><w:pPr><w:pStyle w:val="Titolosommario"/></w:pPr><w:r><w:t>Sommario</w:t></w:r></w:p>
    ${entry('Sommario1', 'Primo', '_Toc1', '1', true)}
    ${entry('Sommario2', 'Secondo', '_Toc2', '2', false, true)}
  </w:sdtContent></w:sdt>
  <w:p><w:pPr><w:pStyle w:val="Titolo1"/></w:pPr><w:bookmarkStart w:id="0" w:name="_Toc1"/><w:r><w:t>Primo</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>
  <w:p><w:r><w:br w:type="page"/></w:r></w:p>
  <w:p><w:pPr><w:pStyle w:val="Titolo2"/></w:pPr><w:r><w:t>Secondo</w:t></w:r></w:p>
  <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1134" w:bottom="1134" w:left="1134" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>
</w:body></w:document>`;
    const styles = `<?xml version="1.0" encoding="UTF-8"?>
<w:styles ${W}>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normale"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Titolo1"><w:name w:val="heading 1"/><w:basedOn w:val="Normale"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="Titolo2"><w:name w:val="heading 2"/><w:basedOn w:val="Normale"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="Titolosommario"><w:name w:val="TOC Heading"/><w:basedOn w:val="Titolo1"/><w:pPr><w:outlineLvl w:val="9"/></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="Sommario1"><w:name w:val="toc 1"/><w:basedOn w:val="Normale"/></w:style>
  <w:style w:type="paragraph" w:styleId="Sommario2"><w:name w:val="toc 2"/><w:basedOn w:val="Normale"/><w:pPr><w:ind w:left="220"/></w:pPr></w:style>
</w:styles>`;
    const rels = `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;
    const pkgRels = `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;
    const data = await writeZip([
      ['_rels/.rels', enc.encode(pkgRels)],
      ['word/document.xml', enc.encode(document)],
      ['word/_rels/document.xml.rels', enc.encode(rels)],
      ['word/styles.xml', enc.encode(styles)],
    ]);

    const doc = await importDocx(data);
    const at = findToc(doc)!;
    expect(at).toMatchObject({ start: 0, end: 3 });
    expect(bodyOf(doc).slice(0, 3).map(paragraphText)).toEqual(['Sommario', 'Primo\t1', 'Secondo\t2']);

    const editor = new Editor(createState(doc));
    editor.use(tocModule);
    editor.execute('toc.update');
    const toc = bodyOf(editor.state.doc).slice(0, 3);
    // Il titolo del sommario (basato su Titolo 1) non diventa una voce; gli stili di Word restano.
    expect(toc.map(paragraphText)).toEqual(['Sommario', 'Primo\t1', 'Secondo\t2']);
    expect(toc.map((p) => p.styleId)).toEqual(['Titolosommario', 'Sommario1', 'Sommario2']);
  });
});
