import { describe, expect, it } from 'vitest';
import { PAGE_SIZES, defaultPageSetup } from '../../src/core';
import { applyPageSetup } from '../../src/structures';

// Controllo dei comandi (Layout > Dimensioni): il formato carta rispetta l'orientamento della pagina.
describe('Layout > Dimensioni', () => {
  it('su una pagina orizzontale il nuovo formato resta orizzontale', () => {
    const landscape = applyPageSetup(defaultPageSetup(), { orientation: 'landscape' });
    const a5 = applyPageSetup(landscape, { ...PAGE_SIZES.A5 });
    expect(a5.orientation).toBe('landscape');
    expect(a5.width).toBeCloseTo(PAGE_SIZES.A5.height);
    expect(a5.height).toBeCloseTo(PAGE_SIZES.A5.width);
  });

  it('su una pagina verticale resta verticale', () => {
    const letter = applyPageSetup(defaultPageSetup(), { ...PAGE_SIZES.Letter });
    expect(letter).toMatchObject({ orientation: 'portrait', ...PAGE_SIZES.Letter });
  });
});
