import { describe, expect, it } from 'vitest';
import {
  Editor,
  createState,
  defaultPageSetup,
  paragraph,
  paragraphLength,
  textRun,
  type Paragraph,
  type Section,
} from '../../src/core';
import { createLayoutModule } from '../../src/layout';
import { doc, engine, lineText, measurer, pageTexts, plain } from './helpers';

const page = defaultPageSetup();
const contentWidth = page.width - page.margins.left - page.margins.right;

/** Pagina stretta: 20 caratteri da 11 pt (5,5 pt) per riga. */
const narrow: Section['page'] = {
  ...page,
  width: 110 + 20,
  height: 200,
  margins: { top: 20, bottom: 20, left: 10, right: 10, header: 5, footer: 5 },
};

describe('righe', () => {
  it('posiziona un paragrafo corto in cima alla pagina, al margine', () => {
    const p = plain('Ciao mondo');
    const layout = engine.layout(doc([p]));
    expect(layout.pages).toHaveLength(1);
    const [line] = layout.pages[0]!.lines;
    expect(line).toMatchObject({ paragraphId: p.id, start: 0, end: 10, x: page.margins.left, y: page.margins.top });
    expect(line!.height).toBeCloseTo(11);
    expect(line!.baseline).toBeCloseTo(8.8);
    expect(lineText(line!)).toBe('Ciao mondo');
    expect(layout.pages[0]!.contentRect.width).toBeCloseTo(contentWidth);
  });

  it('va a capo fra le parole e copre tutto il paragrafo', () => {
    const p = plain('uno due tre quattro cinque sei sette otto nove dieci');
    const layout = engine.layout(doc([p], { page: narrow }));
    const lines = layout.pages.flatMap((pg) => pg.lines);
    expect(lines.map(lineText)).toEqual(['uno due tre quattro ', 'cinque sei sette otto ', 'nove dieci']);
    expect(lines[0]!.start).toBe(0);
    for (let i = 1; i < lines.length; i++) expect(lines[i]!.start).toBe(lines[i - 1]!.end);
    expect(lines[lines.length - 1]!.end).toBe(paragraphLength(p));
  });

  it('spezza una parola più lunga della riga', () => {
    const p = plain('x'.repeat(45));
    const lines = engine.layout(doc([p], { page: narrow })).pages[0]!.lines;
    expect(lines.map((l) => l.end - l.start)).toEqual([20, 20, 5]);
  });

  it('allinea al centro, a destra e giustifica', () => {
    const center = plain('abcd', { align: 'center' });
    const right = plain('abcd', { align: 'right' });
    const just = plain('aa bb cc dd ee ff gg hh ii', { align: 'justify' });
    const pages = engine.layout(doc([center, right, just], { page: narrow })).pages;
    const lines = pages[0]!.lines;
    expect(lines[0]!.runs[0]!.x).toBeCloseTo((110 - 22) / 2);
    expect(lines[1]!.runs[0]!.x).toBeCloseTo(110 - 22);
    // La prima riga giustificata arriva al margine destro, l'ultima no.
    const first = lines[2]!;
    const last = first.runs[first.runs.length - 1]!;
    expect(last.x + measurer.measure(last.text.trimEnd(), last.format).width).toBeCloseTo(110);
    const lastLine = lines[lines.length - 1]!;
    expect(lastLine.runs[0]!.x).toBe(0);
  });

  it('applica rientri e interlinea esatta', () => {
    const p = plain('testo', { indentLeft: 20, indentFirstLine: 10, lineSpacing: { rule: 'exact', value: 30 } });
    const [line] = engine.layout(doc([p])).pages[0]!.lines;
    expect(line!.x).toBeCloseTo(page.margins.left + 30);
    expect(line!.height).toBe(30);
  });

  it('usa le tabulazioni predefinite e quelle a destra', () => {
    const p = paragraph([textRun('a'), { type: 'tab' }, textRun('b')], { spaceAfter: 0 });
    const [line] = engine.layout(doc([p])).pages[0]!.lines;
    expect(line!.runs.find((r) => r.text === 'b')!.x).toBeCloseTo(36);

    const q = paragraph([textRun('a'), { type: 'tab' }, textRun('fine')], {
      tabs: [{ position: 200, align: 'right' }],
    });
    const [l2] = engine.layout(doc([q])).pages[0]!.lines;
    const fine = l2!.runs.find((r) => r.text === 'fine')!;
    expect(fine.x + fine.width).toBeCloseTo(200);
  });

  it("un'interruzione di riga crea una nuova riga nello stesso paragrafo", () => {
    const p = paragraph([textRun('ab'), { type: 'break', kind: 'line' }, textRun('cd')], { spaceAfter: 0 });
    const lines = engine.layout(doc([p])).pages[0]!.lines;
    expect(lines.map((l) => [l.start, l.end])).toEqual([[0, 3], [3, 5]]);
  });
});

describe('pagine', () => {
  const many = (n: number) => Array.from({ length: n }, (_, i) => plain(`riga ${i}`));

  it('distribuisce i paragrafi su più pagine entro l’area del contenuto', () => {
    // Area alta 160 pt, righe da 11 pt: 14 righe per pagina.
    const layout = engine.layout(doc(many(30), { page: narrow }));
    expect(layout.pages.map((p) => p.lines.length)).toEqual([14, 14, 2]);
    expect(layout.pages.map((p) => p.number)).toEqual([1, 2, 3]);
    for (const p of layout.pages) {
      for (const l of p.lines) expect(l.y + l.height).toBeLessThanOrEqual(p.contentRect.y + p.contentRect.height + 0.01);
    }
  });

  it('interruzione di pagina e "interruzione di pagina prima"', () => {
    const a = paragraph([textRun('a'), { type: 'break', kind: 'page' }, textRun('b')]);
    const c = plain('c', { pageBreakBefore: true });
    const layout = engine.layout(doc([a, c]));
    expect(pageTexts(layout)).toEqual([['a'], ['b'], ['c']]);
  });

  it('mantiene un titolo con il paragrafo successivo', () => {
    const filler = many(13);
    const heading = plain('Titolo', { keepWithNext: true });
    const body = plain('corpo');
    const layout = engine.layout(doc([...filler, heading, body], { page: narrow }));
    expect(pageTexts(layout)[1]).toEqual(['Titolo', 'corpo']);
  });

  it('evita le righe orfane e vedove', () => {
    const words = (n: number) => Array.from({ length: n }, () => 'parola').join(' ');
    // 13 righe occupate, il paragrafo da 2 righe non lascia la prima da sola.
    const orphan = engine.layout(doc([...many(13), plain(words(4))], { page: narrow }));
    expect(orphan.pages[0]!.lines).toHaveLength(13);
    // 10 righe occupate + paragrafo di 5 righe: 3 + 2 invece di 4 + 1.
    const widow = engine.layout(doc([...many(10), plain(words(13))], { page: narrow }));
    expect(widow.pages[0]!.lines).toHaveLength(13);
    expect(widow.pages[1]!.lines).toHaveLength(2);
    const off = engine.layout(doc([...many(10), plain(words(13), { widowControl: false })], { page: narrow }));
    expect(off.pages[1]!.lines).toHaveLength(1);
  });

  it('intestazione, piè di pagina e campi numero di pagina', () => {
    const footer = paragraph([
      textRun('Pagina '),
      { type: 'field', kind: 'page', format: {} },
      textRun(' di '),
      { type: 'field', kind: 'numPages', format: {} },
    ], { spaceAfter: 0 });
    const layout = engine.layout(
      doc(many(20), {
        page: narrow,
        header: { default: [plain('Intestazione')], first: [plain('Prima')] },
        footer: { default: [footer], first: [footer] },
        titlePage: true,
      }),
    );
    expect(layout.pages.map((p) => p.header.map(lineText))).toEqual([['Prima'], ['Intestazione']]);
    expect(layout.pages.map((p) => lineText(p.footer[0]!))).toEqual(['Pagina 1 di 2', 'Pagina 2 di 2']);
    const f = layout.pages[0]!.footer[0]!;
    expect(f.y + f.height).toBeCloseTo(narrow.height - narrow.margins.footer);
  });

  it('sezioni: pagina successiva, continua e numerazione', () => {
    const d = doc([plain('a')]);
    const s2: Section = { id: 's2', page: narrow, blocks: [plain('b')], pageNumberStart: 10 };
    const s3: Section = { id: 's3', page: narrow, blocks: [plain('c')], breakType: 'continuous' };
    const layout = engine.layout({ ...d, sections: [d.sections[0]!, s2, s3] });
    expect(pageTexts(layout)).toEqual([['a'], ['b', 'c']]);
    expect(layout.pages.map((p) => p.number)).toEqual([1, 10]);
    expect(layout.pages[1]!.sectionId).toBe('s2');
  });

  it('pagina dispari: aggiunge una pagina bianca', () => {
    const d = doc([plain('a')]);
    const s2: Section = { id: 's2', page, blocks: [plain('b')], breakType: 'oddPage' };
    const layout = engine.layout({ ...d, sections: [d.sections[0]!, s2] });
    expect(pageTexts(layout)).toEqual([['a'], [], ['b']]);
  });

  it('colonne', () => {
    const cols: Section['page'] = { ...narrow, width: 250, columns: { count: 2, gap: 10 } };
    const layout = engine.layout(doc(many(30), { page: cols }));
    const p = layout.pages[0]!;
    expect(p.lines).toHaveLength(28);
    expect(p.lines[14]!.x).toBeCloseTo(10 + 110 + 10);
    expect(p.lines[14]!.y).toBeCloseTo(p.lines[0]!.y);
  });
});

describe('tabelle', () => {
  const cellPara = (t: string) => plain(t);
  const border = { style: 'single' as const, width: 1, color: '#000000' };

  function makeTable(rows: string[][], extra: Partial<import('../../src/core').Table> = {}) {
    return {
      type: 'table' as const,
      id: 't',
      columnWidths: rows[0]!.map(() => 100),
      rows: rows.map((r, i) => ({
        id: `r${i}`,
        cells: r.map((text, j) => ({ id: `c${i}${j}`, blocks: [cellPara(text)], format: {} })),
      })),
      format: { borders: { top: border, bottom: border, left: border, right: border, insideH: border, insideV: border } },
      ...extra,
    };
  }

  it('posiziona le celle in griglia con bordi', () => {
    const t = makeTable([['a', 'b'], ['c', 'd']]);
    const layout = engine.layout(doc([t, plain('dopo')]));
    const p = layout.pages[0]!;
    const byText = Object.fromEntries(p.lines.map((l) => [lineText(l), l]));
    expect(byText.a!.x).toBeCloseTo(page.margins.left + 5.4);
    expect(byText.b!.x).toBeCloseTo(page.margins.left + 100 + 5.4);
    expect(byText.c!.y).toBeCloseTo(byText.a!.y + 11);
    expect(byText.dopo!.y).toBeCloseTo(byText.a!.y + 22);
    expect(p.decorations.filter((d) => d.kind === 'border').length).toBe(16);
  });

  it('unisce celle in orizzontale e in verticale', () => {
    const t = makeTable([['a', 'b'], ['c', 'd']]);
    const merged = {
      ...t,
      rows: [
        { ...t.rows[0]!, cells: [{ ...t.rows[0]!.cells[0]!, colSpan: 2 }] },
        t.rows[1]!,
      ],
    };
    const p = engine.layout(doc([merged])).pages[0]!;
    expect(p.lines.map(lineText)).toEqual(['a', 'c', 'd']);

    const v = makeTable([['alto', 'x'], ['', 'y']]);
    const vm = {
      ...v,
      rows: [
        { ...v.rows[0]!, cells: [{ ...v.rows[0]!.cells[0]!, vMerge: 'restart' as const, blocks: [plain('1'), plain('2'), plain('3')] }, v.rows[0]!.cells[1]!] },
        { ...v.rows[1]!, cells: [{ ...v.rows[1]!.cells[0]!, vMerge: 'continue' as const }, v.rows[1]!.cells[1]!] },
      ],
    };
    const after = plain('dopo');
    const q = engine.layout(doc([vm, after])).pages[0]!;
    const dopo = q.lines.find((l) => lineText(l) === 'dopo')!;
    const one = q.lines.find((l) => lineText(l) === '1')!;
    // La cella unita (3 righe) allunga la seconda riga della tabella.
    expect(dopo.y - one.y).toBeCloseTo(33);
  });

  it('ripete le righe di intestazione sulla pagina successiva', () => {
    const rows = [['Titolo'], ...Array.from({ length: 20 }, (_, i) => [`r${i}`])];
    const t = makeTable(rows);
    const withHeader = { ...t, rows: t.rows.map((r, i) => (i === 0 ? { ...r, isHeader: true } : r)) };
    const layout = engine.layout(doc([withHeader], { page: narrow }));
    expect(layout.pages.length).toBe(2);
    expect(lineText(layout.pages[1]!.lines[0]!)).toBe('Titolo');
  });
});

describe('elenchi', () => {
  it('numera i paragrafi e i livelli', () => {
    const list = {
      id: 'L',
      levels: [
        { format: 'decimal' as const, text: '%1.', start: 1, indentLeft: 36, hanging: 18 },
        { format: 'lowerLetter' as const, text: '%2)', start: 1, indentLeft: 72, hanging: 18 },
      ],
    };
    const item = (t: string, level = 0) => plain(t, { list: { listId: 'L', level } });
    const d = doc([item('uno'), item('sotto', 1), item('sotto2', 1), item('due'), item('sotto3', 1)]);
    const layout = engine.layout({ ...d, lists: { L: list } });
    const lines = layout.pages[0]!.lines;
    expect(lines.map((l) => l.runs[0]!.text)).toEqual(['1.', 'a)', 'b)', '2.', 'a)']);
    expect(lines[0]!.x).toBeCloseTo(page.margins.left + 18);
    expect(lines[0]!.runs[1]!.x).toBeCloseTo(18);
  });
});

describe('cursore e selezione', () => {
  const p1 = plain('uno due tre quattro cinque sei sette otto nove dieci');
  const p2 = plain('fine');
  const layout = engine.layout(doc([p1, p2], { page: narrow }));

  it('caretRect e hitTest sono coerenti', () => {
    for (let off = 0; off <= paragraphLength(p1); off++) {
      const c = engine.caretRect(layout, { paragraphId: p1.id, offset: off })!;
      expect(c).not.toBeNull();
      const hit = engine.hitTest(layout, c.page, c.rect.x, c.rect.y + c.rect.height / 2)!;
      // Alla fine di una riga spezzata il clic torna prima dello spazio finale.
      expect([off, off - 1]).toContain(hit.offset);
    }
  });

  it("l'offset dell'a capo automatico va all'inizio della riga successiva", () => {
    const c = engine.caretRect(layout, { paragraphId: p1.id, offset: 20 })!;
    expect(c.rect.y).toBeCloseTo(narrow.margins.top + 11);
    expect(c.rect.x).toBeCloseTo(narrow.margins.left);
  });

  it('hitTest fuori dal testo prende la riga più vicina', () => {
    expect(engine.hitTest(layout, 0, 0, 0)).toEqual({ paragraphId: p1.id, offset: 0 });
    expect(engine.hitTest(layout, 0, 500, 1000)).toEqual({ paragraphId: p2.id, offset: 4 });
  });

  it('selectionRects su più righe e paragrafi', () => {
    const rects = engine.selectionRects(layout, { paragraphId: p2.id, offset: 2 }, { paragraphId: p1.id, offset: 4 });
    expect(rects).toHaveLength(4);
    expect(rects[0]!.rect.x).toBeCloseTo(narrow.margins.left + 3 * 5.5 + 2.75);
    expect(rects[3]!.rect.width).toBeCloseTo(2 * 5.5);
    expect(engine.selectionRects(layout, { paragraphId: p2.id, offset: 1 }, { paragraphId: p2.id, offset: 1 })).toEqual([]);
  });

  it('trova il cursore in una cella e in un paragrafo vuoto', () => {
    const empty: Paragraph = plain('');
    const l = engine.layout(doc([empty]));
    const c = engine.caretRect(l, { paragraphId: empty.id, offset: 0 })!;
    expect(c.rect).toMatchObject({ x: page.margins.left, y: page.margins.top });
    expect(engine.hitTest(l, 0, 300, 300)).toEqual({ paragraphId: empty.id, offset: 0 });
  });
});

describe('modulo', () => {
  it('pubblica editor.services.layout e lo rimuove alla disattivazione', () => {
    const editor = new Editor(createState());
    const dispose = editor.use(createLayoutModule({ measurer }));
    expect(editor.services.layout).toBeDefined();
    expect(editor.services.layout!.layout(editor.state.doc).pages).toHaveLength(1);
    dispose();
    expect(editor.services.layout).toBeUndefined();
  });

  it('riusa le righe dei paragrafi invariati', () => {
    const p = plain('testo');
    const d = doc([p]);
    const a = engine.layout(d);
    const b = engine.layout({ ...d });
    expect(b.pages[0]!.lines[0]!.runs).toBe(a.pages[0]!.lines[0]!.runs);
  });
});
