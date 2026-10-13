import { describe, expect, it } from 'vitest';
import { paragraphText } from '../../src/core';
import { documentToText, textFormat, textToDocument } from '../../src/io';
import { plainParagraphs, richDocument } from './fixtures';

describe('formato testo', () => {
  it('importa righe come paragrafi e tab come inline', async () => {
    const doc = await textFormat.import!(new TextEncoder().encode('﻿uno\r\ndue\tcolonna\n\ntre'));
    const ps = plainParagraphs(doc);
    expect(ps.map(paragraphText)).toEqual(['uno', 'due\tcolonna', '', 'tre']);
    expect(ps[1]!.children[1]).toEqual({ type: 'tab' });
    expect(ps[2]!.children).toEqual([]);
  });

  it('esporta paragrafi, tabelle e interruzioni', () => {
    const text = documentToText(richDocument());
    const lines = text.split('\n');
    expect(lines[0]).toBe('Titolo del documento');
    expect(lines).toContain('A1+B1\tC1');
    expect(lines).toContain('A2\tB2');
    expect(lines).toContain('Immagine: [Logo]');
    expect(lines).toContain('collegamento e barrato');
  });

  it('esporta con fine riga Windows', async () => {
    const out = await textFormat.export!(textToDocument('a\nb'));
    expect(out).toBe('a\r\nb');
  });
});
