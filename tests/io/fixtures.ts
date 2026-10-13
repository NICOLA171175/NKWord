import {
  createDocument,
  paragraph,
  section,
  table,
  tableCell,
  textRun,
  type Block,
  type DocumentModel,
  type Paragraph,
  type Table,
} from '../../src/core';

/** PNG 1×1 trasparente. */
export const PNG_DATA_URI =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=';

export function plainParagraphs(doc: DocumentModel): Paragraph[] {
  const out: Paragraph[] = [];
  const walk = (blocks: readonly Block[]) => {
    for (const b of blocks) {
      if (b.type === 'paragraph') out.push(b);
      else for (const r of b.rows) for (const c of r.cells) walk(c.blocks);
    }
  };
  for (const s of doc.sections) walk(s.blocks);
  return out;
}

export function firstTable(doc: DocumentModel): Table {
  for (const s of doc.sections) for (const b of s.blocks) if (b.type === 'table') return b;
  throw new Error('nessuna tabella');
}

/** Documento con quasi tutto ciò che il modello sa rappresentare. */
export function richDocument(): DocumentModel {
  const base = createDocument();
  const t = table(2, 3, 450);
  const merged: Table = {
    ...t,
    format: {
      align: 'center',
      borders: {
        top: { style: 'single', width: 0.5, color: '#000000' },
        bottom: { style: 'single', width: 0.5, color: '#000000' },
        left: { style: 'single', width: 0.5, color: '#000000' },
        right: { style: 'single', width: 0.5, color: '#000000' },
        insideH: { style: 'single', width: 0.5, color: '#000000' },
        insideV: { style: 'single', width: 0.5, color: '#000000' },
      },
    },
    rows: [
      {
        ...t.rows[0]!,
        isHeader: true,
        cells: [
          { ...tableCell([paragraph('A1+B1')]), colSpan: 2, format: { shading: '#d9e2f3' } },
          { ...tableCell([paragraph('C1')]), vMerge: 'restart', format: { verticalAlign: 'center' } },
        ],
      },
      {
        ...t.rows[1]!,
        cells: [tableCell([paragraph('A2')]), tableCell([paragraph('B2')]), { ...tableCell(), vMerge: 'continue' }],
      },
    ],
  };
  const body: Block[] = [
    paragraph('Titolo del documento', {}, 'Title'),
    paragraph('Capitolo uno', {}, 'Heading1'),
    paragraph(
      [
        textRun('Normale, '),
        textRun('grassetto', { bold: true }),
        textRun(', '),
        textRun('corsivo sottolineato', { italic: true, underline: 'single' }),
        textRun(' rosso 14', { color: '#ff0000', fontSize: 14, fontFamily: 'Times New Roman' }),
        { type: 'tab' },
        textRun('evidenziato', { highlight: '#ffff00' }),
        textRun('2', { verticalAlign: 'superscript' }),
        { type: 'break', kind: 'line' },
        textRun('collegamento', { link: 'https://example.com/a?b=1&c=2' }),
        textRun(' e barrato', { strike: true }),
      ],
      { align: 'justify', indentLeft: 18, indentFirstLine: -9, spaceBefore: 6, spaceAfter: 12, lineSpacing: { rule: 'multiple', value: 1.5 } },
    ),
    paragraph('Primo punto', { list: { listId: 'L1', level: 0 } }),
    paragraph('Sottopunto', { list: { listId: 'L1', level: 1 } }),
    paragraph('Secondo punto', { list: { listId: 'L1', level: 0 } }),
    paragraph([textRun('Immagine: '), { type: 'image', src: 'img1', width: 72, height: 36, alt: 'Logo' }]),
    merged,
    paragraph([textRun('Pagina '), { type: 'field', kind: 'page', format: {} }, textRun(' di '), { type: 'field', kind: 'numPages', format: {} }]),
  ];
  const second = section([paragraph('Seconda sezione', { align: 'center' })]);
  return {
    ...base,
    sections: [
      {
        ...base.sections[0]!,
        blocks: body,
        header: { default: [paragraph('Intestazione')] },
        footer: { default: [paragraph([textRun('Pag. '), { type: 'field', kind: 'page', format: {} }])] },
      },
      {
        ...second,
        breakType: 'nextPage',
        page: { ...second.page, width: second.page.height, height: second.page.width, orientation: 'landscape' },
      },
    ],
    lists: {
      L1: {
        id: 'L1',
        levels: [
          { format: 'bullet', text: '•', start: 1, indentLeft: 36, hanging: 18 },
          { format: 'lowerLetter', text: '%2)', start: 1, indentLeft: 72, hanging: 18 },
        ],
      },
    },
    media: { img1: PNG_DATA_URI },
    metadata: { title: 'Prova', author: 'Nicola', subject: 'Test', keywords: ['uno', 'due'], created: '2026-01-02T03:04:05Z' },
  };
}
