import { describe, expect, it } from 'vitest';
import { PAGE_SIZES, cmToPt, emuToPt, ptToCm, ptToEmu, ptToPx, ptToTwips, pxToPt, twipsToPt } from '../../src/core';

describe('units', () => {
  it('converte twip ed EMU in andata e ritorno', () => {
    expect(ptToTwips(12)).toBe(240);
    expect(twipsToPt(240)).toBe(12);
    expect(emuToPt(ptToEmu(36))).toBe(36);
  });
  it('converte cm e pixel', () => {
    expect(ptToCm(cmToPt(2.5))).toBeCloseTo(2.5);
    expect(ptToPx(72)).toBe(96);
    expect(pxToPt(ptToPx(10, 1.5), 1.5)).toBeCloseTo(10);
  });
  it('A4 misura circa 595x842 pt', () => {
    expect(PAGE_SIZES.A4.width).toBeCloseTo(595.28, 1);
    expect(PAGE_SIZES.A4.height).toBeCloseTo(841.89, 1);
  });
});
