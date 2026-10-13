import { describe, expect, it } from 'vitest';
import { createDocument, Editor, createState, paragraph, textRun, type Paragraph, type TextRun } from '../../src/core';
import { commentThreads, createCommentsModule } from '../../src/comments';
import { exportDocx, importDocx } from '../../src/io';
import { parseXml } from '../../src/io/util/markup';
import { readZip, writeZip } from '../../src/io/util/zip';

const enc = new TextEncoder();
const dec = new TextDecoder();

function commentedDoc() {
  const a = paragraph([textRun('Primo '), textRun('paragrafo', { link: 'https://example.com' })]);
  const b = paragraph([textRun('secondo paragrafo')]);
  const editor = new Editor({ ...createState(createDocument([a, b])), selection: { anchor: { paragraphId: a.id, offset: 6 }, focus: { paragraphId: b.id, offset: 7 } } });
  editor.use(createCommentsModule({ author: 'Nicola Rossi', now: () => new Date('2026-10-13T08:00:00Z') }));
  editor.execute('comment.insert', { text: 'Riga uno\nRiga due & <tre>' });
  const id = Object.keys(editor.state.doc.comments!)[0]!;
  editor.execute('comment.reply', { parentId: id, text: 'Va bene', author: 'Anna Bianchi' });
  editor.execute('comment.resolve', { id, resolved: true });
  return editor.state.doc;
}

describe('commenti in DOCX', () => {
  it('scrive comments.xml, commentsExtended.xml e i marcatori nel corpo', async () => {
    const zip = await readZip(await exportDocx(commentedDoc()));
    const text = (p: string) => dec.decode(zip.get(p));
    expect(text('[Content_Types].xml')).toContain('/word/comments.xml');
    expect(text('word/_rels/document.xml.rels')).toContain('Target="comments.xml"');
    const comments = text('word/comments.xml');
    expect(comments).toContain('w:author="Nicola Rossi"');
    expect(comments).toContain('w:initials="NR"');
    expect(comments).toContain('Riga due &amp; &lt;tre&gt;');
    expect(comments).toContain('w:author="Anna Bianchi"');
    const ext = text('word/commentsExtended.xml');
    expect(ext).toContain('w15:paraId="00000001" w15:done="1"');
    expect(ext).toContain('w15:paraId="00000002" w15:paraIdParent="00000001"');
    const body = text('word/document.xml');
    // L'inizio sta dentro il collegamento, prima del run commentato; la fine nel secondo paragrafo.
    expect(body).toMatch(/<w:hyperlink [^>]*><w:commentRangeStart w:id="0"\/><w:commentRangeStart w:id="1"\/><w:r>/);
    expect(body).toMatch(/secondo<\/w:t><\/w:r><w:commentRangeEnd w:id="0"\/><w:r><w:commentReference w:id="0"\/><\/w:r>/);
    for (const part of ['word/comments.xml', 'word/commentsExtended.xml', 'word/document.xml']) {
      expect(parseXml(text(part)).children.length).toBeGreaterThan(0);
    }
  });

  it('senza commenti non aggiunge parti', async () => {
    const zip = await readZip(await exportDocx(createDocument([paragraph([textRun('x')])])));
    expect(zip.has('word/comments.xml')).toBe(false);
  });

  it('andata e ritorno conserva commenti, risposte e stato', async () => {
    const doc = await importDocx(await exportDocx(commentedDoc()));
    const [t] = commentThreads(doc);
    expect(commentThreads(doc)).toHaveLength(1);
    expect(t!.comment).toMatchObject({ author: 'Nicola Rossi', initials: 'NR', text: 'Riga uno\nRiga due & <tre>', resolved: true, date: '2026-10-13T08:00:00Z' });
    expect(t!.replies.map((r) => [r.author, r.text])).toEqual([['Anna Bianchi', 'Va bene']]);
    const [a, b] = doc.sections[0]!.blocks as Paragraph[];
    expect(t!.range).toEqual({ from: { paragraphId: a!.id, offset: 6 }, to: { paragraphId: b!.id, offset: 7 } });
    const linked = a!.children[1] as TextRun;
    expect(linked.format).toMatchObject({ link: 'https://example.com', commentIds: t!.comment.id });
  });

  it('legge un file come lo scrive Word (risposte con intervallo proprio)', async () => {
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const files: [string, string][] = [
      ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'],
      ['_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'],
      ['word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/><Relationship Id="rId6" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/></Relationships>'],
      ['word/document.xml', `<w:document ${W}><w:body><w:commentRangeStart w:id="7"/><w:p><w:commentRangeStart w:id="9"/><w:r><w:t>Ciao </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>mondo</w:t></w:r><w:commentRangeEnd w:id="7"/><w:commentRangeEnd w:id="9"/><w:r><w:rPr><w:rStyle w:val="Rimandocommento"/></w:rPr><w:commentReference w:id="7"/></w:r><w:r><w:t> fine</w:t></w:r></w:p></w:body></w:document>`],
      ['word/comments.xml', `<w:comments ${W} xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:comment w:id="7" w:author="Mario" w:date="2026-01-02T03:04:00Z" w:initials="M"><w:p w14:paraId="1A2B3C4D"><w:r><w:annotationRef/></w:r><w:r><w:t>Controlla</w:t></w:r></w:p></w:comment><w:comment w:id="9" w:author="Lucia" w:date="2026-01-03T00:00:00Z"><w:p w14:paraId="0F0F0F0F"><w:r><w:t>Fatto</w:t></w:r></w:p></w:comment></w:comments>`],
      ['word/commentsExtended.xml', '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:commentEx w15:paraId="1A2B3C4D" w15:done="0"/><w15:commentEx w15:paraId="0F0F0F0F" w15:paraIdParent="1A2B3C4D" w15:done="0"/></w15:commentsEx>'],
    ];
    const doc = await importDocx(await writeZip(files.map(([n, x]) => [n, enc.encode(x)])));
    const threads = commentThreads(doc);
    expect(threads).toHaveLength(1);
    expect(threads[0]!.comment).toMatchObject({ author: 'Mario', initials: 'M', text: 'Controlla' });
    expect(threads[0]!.comment.resolved).toBeUndefined();
    expect(threads[0]!.replies.map((r) => r.text)).toEqual(['Fatto']);
    const p = doc.sections[0]!.blocks[0] as Paragraph;
    const id = threads[0]!.comment.id;
    expect(p.children.map((c) => [(c as TextRun).text, (c as TextRun).format.commentIds])).toEqual([
      ['Ciao ', id],
      ['mondo', id],
      [' fine', undefined],
    ]);
  });
});
