/**
 * EMU / point / pixel conversion helpers shared across OOXML formats.
 *
 * DrawingML measures in EMUs (English Metric Units): 914400 EMU = 1 inch.
 * Standard assumption for screen rendering: 96 DPI.
 *
 * @packageDocumentation
 * @public
 */

/** Standard DPI for screen rendering */
const STANDARD_DPI = 96;

/**
 * EMUs per inch (1 inch = 914400 EMUs)
 *
 * @public
 */
export const EMUS_PER_INCH = 914400;

/**
 * Points per inch (1 inch = 72 points)
 *
 * @public
 */
export const POINTS_PER_INCH = 72;

/**
 * Pixels per inch at standard DPI
 *
 * @public
 */
export const PIXELS_PER_INCH = STANDARD_DPI;

/**
 * Convert EMUs to pixels (at 96 DPI)
 *
 * 1 inch = 914400 EMUs = 96 pixels
 * Returns 0 for null/undefined/NaN inputs.
 *
 * @public
 */
export function emuToPixels(emu: number | undefined | null): number {
  if (emu == null || isNaN(emu)) return 0;
  return Math.round((emu * PIXELS_PER_INCH) / EMUS_PER_INCH);
}

/**
 * Convert pixels to EMUs.
 * EMU coordinates in OOXML are integer-typed (xs:long); rounding here keeps
 * floating-point drift (e.g. 52 px → 495299.99999999994) out of the document.
 *
 * @public
 */
export function pixelsToEmu(px: number): number {
  return Math.round((px / PIXELS_PER_INCH) * EMUS_PER_INCH);
}

/**
 * Convert points to pixels (at 96 DPI)
 *
 * 1 inch = 72 points = 96 pixels
 * → 1 point = 96/72 pixels = 4/3 pixels
 *
 * @public
 */
export function pointsToPixels(points: number): number {
  return (points / POINTS_PER_INCH) * PIXELS_PER_INCH;
}
