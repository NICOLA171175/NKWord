import { describe, expect, it } from 'vitest';
import { caret, createDocument, createState, paragraph } from '../../src/core';
import { insertFootnote } from '../../src/footnotes';
import { createApproxMeasurer, createLayoutEngine } from '../../src/layout';
import { encodeWinAnsi, layoutToPdf } from '../../src/pdf';

const measurer = createApproxMeasurer();
const hex = (s: string) => Array.from(encodeWinAnsi(s), (b) => b.toString(16).padStart(2, '0')).join('');

describe('PDF: note a piè di pagina', () => {
  it('disegna il testo delle note in fondo alla pagina', async () => {
    const p = paragraph('Testo');
    const doc = insertFootnote({ ...createState(createDocument([p])), selection: caret(p.id, 5) }, 'Fonte')!.state.doc;
    const bytes = await layoutToPdf(createLayoutEngine(measurer).layout(doc), { measurer });
    const pdf = Array.from(bytes, (c) => String.fromCharCode(c)).join('');
    expect(pdf).toContain(hex('Fonte'));
  });
});
