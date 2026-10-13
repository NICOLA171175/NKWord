// @vitest-environment happy-dom
import { describe, expect, it } from 'vitest';
import { buttonSize, iconSvg } from '../../src/ui/icons';
import { RIBBON_ITEMS } from '../../src/ui/ribbon-items';

describe('icone della barra multifunzione', () => {
  it('ogni voce con icona ha la sua icona SVG nel set', () => {
    const missing = RIBBON_ITEMS.filter((i) => i.icon && !iconSvg(i)).map((i) => i.label);
    expect(missing).toEqual([]);
  });

  it('distingue le varianti dello stesso comando dagli argomenti', () => {
    const left = RIBBON_ITEMS.find((i) => i.label === 'Allinea a sinistra')!;
    const right = RIBBON_ITEMS.find((i) => i.label === 'Allinea a destra')!;
    expect(iconSvg(left)!.innerHTML).not.toBe(iconSvg(right)!.innerHTML);
  });

  it('mostra solo l\'icona nei gruppi Carattere e Paragrafo, pulsanti grandi altrove', () => {
    expect(buttonSize(RIBBON_ITEMS.find((i) => i.label === 'Grassetto')!)).toBe('icon');
    expect(buttonSize(RIBBON_ITEMS.find((i) => i.label === 'Nuovo commento')!)).toBe('large');
    expect(buttonSize(RIBBON_ITEMS.find((i) => i.label === 'Trova')!)).toBe('small');
  });
});
