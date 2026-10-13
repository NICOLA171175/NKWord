import { describe, expect, it } from 'vitest';
import { createDocument, paragraph, paragraphText, textRun, type Block } from '../../src/core';
import { DefaultClipboardService, sliceDocument } from '../../src/io';
import { firstTable, plainParagraphs, richDocument } from './fixtures';

const docOf = (blocks: Block[]) => createDocument(blocks);

describe('frammenti', () => {
  it('taglia il testo dentro un paragrafo conservando la formattazione', () => {
    const p = paragraph([textRun('ab'), textRun('cd', { bold: true }), { type: 'tab' }, textRun('ef')]);
    const doc = docOf([p]);
    const frag = sliceDocument(doc, { paragraphId: p.id, offset: 5 }, { paragraphId: p.id, offset: 1 });
    const [only] = plainParagraphs(frag);
    expect(paragraphText(only!)).toBe('bcd\t');
    expect(only!.children[1]).toEqual(textRun('cd', { bold: true }));
  });

  it('copre più paragrafi e tiene intere le tabelle attraversate', () => {
    const doc = richDocument();
    const ps = plainParagraphs(doc);
    const from = { paragraphId: ps[5]!.id, offset: 8 }; // "Secondo punto" → "punto"
    const to = { paragraphId: ps.at(-2)!.id, offset: 6 }; // "Pagina"
    const frag = sliceDocument(doc, from, to);
    expect(frag.sections).toHaveLength(1);
    const blocks = frag.sections[0]!.blocks;
    expect(blocks.map((b) => b.type)).toEqual(['paragraph', 'paragraph', 'table', 'paragraph']);
    expect(paragraphText(blocks[0] as never)).toBe('punto');
    expect(paragraphText(blocks[3] as never)).toBe('Pagina');
    expect(firstTable(frag)).toBe(firstTable(doc));
  });

  it('appiattisce una tabella selezionata solo in parte', () => {
    const doc = richDocument();
    const ps = plainParagraphs(doc);
    const a2 = ps.find((p) => paragraphText(p) === 'A2')!;
    const frag = sliceDocument(doc, { paragraphId: ps[6]!.id, offset: 0 }, { paragraphId: a2.id, offset: 2 });
    expect(frag.sections[0]!.blocks.every((b) => b.type === 'paragraph')).toBe(true);
    expect(plainParagraphs(frag).map(paragraphText)).toEqual(['Immagine: ￼', 'A1+B1', 'C1', 'A2']);
  });
});

describe('ClipboardService', () => {
  const clipboard = new DefaultClipboardService();

  it('serializza una selezione in HTML e testo', () => {
    const p1 = paragraph([textRun('Ciao '), textRun('mondo', { bold: true })]);
    const p2 = paragraph('seconda riga');
    const doc = docOf([p1, p2]);
    const out = clipboard.serialize(doc, { paragraphId: p1.id, offset: 5 }, { paragraphId: p2.id, offset: 7 });
    expect(out.text).toBe('mondo\nseconda');
    expect(out.html).toContain('<b>mondo</b>');
    expect(out.html).not.toContain('<html');
  });

  it('preferisce l\'HTML al testo in incolla', () => {
    const frag = clipboard.parse({ html: '<html><body><!--StartFragment--><p><i>ciao</i></p><!--EndFragment--></body></html>', text: 'ciao' });
    const [p] = plainParagraphs(frag);
    expect(p!.children[0]).toEqual(textRun('ciao', { italic: true }));
  });

  it('usa il testo semplice se manca l\'HTML', () => {
    const frag = clipboard.parse({ text: 'uno\ndue' });
    expect(plainParagraphs(frag).map(paragraphText)).toEqual(['uno', 'due']);
  });

  it('copia e incolla fra istanze conserva la struttura', () => {
    const doc = richDocument();
    const ps = plainParagraphs(doc);
    const out = clipboard.serialize(doc, { paragraphId: ps[0]!.id, offset: 0 }, { paragraphId: ps.at(-1)!.id, offset: 5 });
    const frag = clipboard.parse(out);
    expect(plainParagraphs(frag).map(paragraphText).slice(0, 3)).toEqual(['Titolo del documento', 'Capitolo uno', paragraphText(ps[2]!)]);
    expect(plainParagraphs(frag)[1]!.styleId).toBe('Heading1');
    expect(firstTable(frag).rows[0]!.cells[0]!.colSpan).toBe(2);
  });
});
