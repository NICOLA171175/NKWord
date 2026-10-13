import { describe, expect, it } from 'vitest';
import { caret, createDocument, createState, firstParagraph, footnoteNumbers, paragraph, textRun, type DocumentModel } from '../../src/core';
import { footnoteRefs, footnoteText, insertFootnote } from '../../src/footnotes';
import { documentToHtml, exportDocx, importDocx } from '../../src/io';
import { parseXml } from '../../src/io/util/markup';
import { readZip, writeZip } from '../../src/io/util/zip';

const enc = new TextEncoder();
const dec = new TextDecoder();

function docWithNotes(): DocumentModel {
  const p = paragraph([textRun('Primo e secondo')]);
  let state = { ...createState(createDocument([p])), selection: caret(p.id, 5) };
  state = insertFootnote(state, 'Nota uno')!.state;
  state = { ...state, selection: caret(p.id, 16) };
  return insertFootnote(state, 'Nota due\nsecondo paragrafo')!.state.doc;
}

describe('DOCX: note a piè di pagina', () => {
  it('scrive footnotes.xml con separatori, relazione, tipo di contenuto e impostazioni', async () => {
    const zip = await readZip(await exportDocx(docWithNotes()));
    const text = (n: string) => dec.decode(zip.get(n));
    const notes = text('word/footnotes.xml');
    expect(parseXml(notes).children.length).toBeGreaterThan(0);
    expect(notes).toContain('<w:footnote w:type="separator" w:id="-1">');
    expect(notes).toContain('<w:footnote w:type="continuationSeparator" w:id="0">');
    expect(notes).toMatch(/<w:footnote w:id="1"><w:p>.*<w:footnoteRef\/>.*Nota uno.*<\/w:footnote>/);
    expect(notes).toContain('<w:pStyle w:val="FootnoteText"/>');
    expect(text('word/_rels/document.xml.rels')).toContain('relationships/footnotes" Target="footnotes.xml"');
    expect(text('[Content_Types].xml')).toContain('<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>');
    expect(text('word/settings.xml')).toContain('<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>');
    const document = text('word/document.xml');
    expect(document).toContain('<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>');
    expect(document).toContain('<w:footnoteReference w:id="2"/>');
    expect(text('word/styles.xml')).toContain('<w:name w:val="footnote text"/>');
  });

  it('senza note non scrive footnotes.xml', async () => {
    const zip = await readZip(await exportDocx(createDocument()));
    expect(zip.has('word/footnotes.xml')).toBe(false);
    expect(dec.decode(zip.get('word/settings.xml'))).not.toContain('footnotePr');
  });

  it('rilegge le note salvate (andata e ritorno)', async () => {
    const back = await importDocx(await exportDocx(docWithNotes()));
    const refs = footnoteRefs(back);
    expect(refs).toHaveLength(2);
    expect(refs.map((r) => footnoteText(back, r.noteId))).toEqual(['Nota uno', 'Nota due\nsecondo paragrafo']);
    expect(refs.map((r) => r.position.offset)).toEqual([5, 16]);
    expect([...footnoteNumbers(back).values()]).toEqual([1, 2]);
    // Il richiamo dentro la nota punta alla nota stessa.
    const note = back.footnotes![refs[0]!.noteId]!;
    const first = note.blocks[0]!;
    expect(first.type === 'paragraph' && first.children[0]).toMatchObject({ type: 'footnoteRef', noteId: refs[0]!.noteId });
  });

  it('legge un file di Word: ignora i separatori e le note non richiamate', async () => {
    const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const files: [string, Uint8Array][] = [
      ['[Content_Types].xml', enc.encode('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>')],
      ['_rels/.rels', enc.encode(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${r}/officeDocument" Target="word/document.xml"/></Relationships>`)],
      ['word/_rels/document.xml.rels', enc.encode(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="${r}/footnotes" Target="footnotes.xml"/></Relationships>`)],
      [
        'word/document.xml',
        enc.encode(
          `<?xml version="1.0"?><w:document xmlns:w="${w}"><w:body><w:p><w:r><w:t>Testo</w:t></w:r>` +
            `<w:r><w:rPr><w:rStyle w:val="Rimandonotaapidipagina"/><w:sz w:val="28"/></w:rPr><w:footnoteReference w:id="3"/></w:r></w:p></w:body></w:document>`,
        ),
      ],
      [
        'word/footnotes.xml',
        enc.encode(
          `<?xml version="1.0"?><w:footnotes xmlns:w="${w}">` +
            `<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>` +
            `<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>` +
            `<w:footnote w:id="2"><w:p><w:r><w:t>orfana</w:t></w:r></w:p></w:footnote>` +
            `<w:footnote w:id="3"><w:p><w:pPr><w:pStyle w:val="Testonotaapidipagina"/></w:pPr><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Fonte: Istat.</w:t></w:r></w:p></w:footnote>` +
            `</w:footnotes>`,
        ),
      ],
    ];
    const doc = await importDocx(await writeZip(files));
    const p = firstParagraph(doc);
    expect(p.children.map((c) => c.type)).toEqual(['text', 'footnoteRef']);
    const ref = p.children[1]!;
    if (ref.type !== 'footnoteRef') throw new Error('richiamo mancante');
    expect(ref.format).toEqual({ fontSize: 14 });
    expect(Object.keys(doc.footnotes!)).toEqual([ref.noteId]);
    expect(footnoteText(doc, ref.noteId)).toBe('Fonte: Istat.');
    const first = doc.footnotes![ref.noteId]!.blocks[0]!;
    expect(first.type === 'paragraph' && first.styleId).toBe('Testonotaapidipagina');
  });

  it('in HTML: numero in apice nel testo e note in fondo', () => {
    const html = documentToHtml(docWithNotes());
    expect(html).toContain('Primo<sup>1</sup>');
    expect(html).toMatch(/<div class="footnotes">\n<hr>\n<div class="footnote">\n<p[^>]*><sup>1<\/sup> Nota uno<\/p>/);
  });
});
