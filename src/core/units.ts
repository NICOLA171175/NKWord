/**
 * Unità di misura.
 *
 * Convenzione: nel modello del documento tutte le lunghezze sono in punti
 * tipografici (pt, 1/72 di pollice). Le altre unità servono solo ai confini:
 * twip ed EMU per DOCX, pixel per il rendering, mm/cm/pollici per la UI.
 */

export type Points = number;

export const POINTS_PER_INCH = 72;
export const TWIPS_PER_POINT = 20;
export const EMU_PER_POINT = 12700;
export const MM_PER_INCH = 25.4;
/** Pixel CSS per pollice (fisso per specifica CSS). */
export const CSS_PX_PER_INCH = 96;

export const ptToTwips = (pt: Points): number => Math.round(pt * TWIPS_PER_POINT);
export const twipsToPt = (twips: number): Points => twips / TWIPS_PER_POINT;

export const ptToEmu = (pt: Points): number => Math.round(pt * EMU_PER_POINT);
export const emuToPt = (emu: number): Points => emu / EMU_PER_POINT;

/** Mezzi punti: unità delle dimensioni dei caratteri in DOCX (w:sz). */
export const ptToHalfPoints = (pt: Points): number => Math.round(pt * 2);
export const halfPointsToPt = (hp: number): Points => hp / 2;

export const ptToInch = (pt: Points): number => pt / POINTS_PER_INCH;
export const inchToPt = (inch: number): Points => inch * POINTS_PER_INCH;

export const ptToMm = (pt: Points): number => (pt / POINTS_PER_INCH) * MM_PER_INCH;
export const mmToPt = (mm: number): Points => (mm / MM_PER_INCH) * POINTS_PER_INCH;

export const ptToCm = (pt: Points): number => ptToMm(pt) / 10;
export const cmToPt = (cm: number): Points => mmToPt(cm * 10);

/** `zoom` = 1 significa 100%. */
export const ptToPx = (pt: Points, zoom = 1): number =>
  (pt / POINTS_PER_INCH) * CSS_PX_PER_INCH * zoom;
export const pxToPt = (px: number, zoom = 1): Points =>
  (px / zoom / CSS_PX_PER_INCH) * POINTS_PER_INCH;

/** Formati pagina comuni, in punti (verticale). */
export const PAGE_SIZES = {
  A4: { width: mmToPt(210), height: mmToPt(297) },
  A5: { width: mmToPt(148), height: mmToPt(210) },
  A3: { width: mmToPt(297), height: mmToPt(420) },
  Letter: { width: inchToPt(8.5), height: inchToPt(11) },
  Legal: { width: inchToPt(8.5), height: inchToPt(14) },
} as const satisfies Record<string, { width: Points; height: Points }>;

export type PageSizeName = keyof typeof PAGE_SIZES;
