// @betteroffice/docx declaration utils/units.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Unit Conversion Utilities - Convert OOXML units to CSS/pixels
 *
 * OOXML uses various unit systems that need conversion for rendering:
 * - Twips: 1/20 of a point (1440 twips = 1 inch)
 * - EMUs (English Metric Units): 914400 EMUs = 1 inch
 * - Half-points: 1/2 of a point (144 half-points = 1 inch)
 * - Points: 72 points = 1 inch
 * - Eighths of a point: 1/8 of a point (576 eighths = 1 inch)
 *
 * Standard assumption: 96 DPI (pixels per inch)
 * @packageDocumentation
 * @public
 */
export declare const PIXELS_PER_INCH = 96;
export declare function emuToPixels(emu: number | undefined | null): number;
export declare function pixelsToEmu(px: number): number;
export declare function pointsToPixels(points: number): number;
/** Twips per inch (1 inch = 1440 twips) */
export declare const TWIPS_PER_INCH = 1440;
/**
 * Automatic paragraph spacing applied for `w:beforeAutospacing` /
 * `w:afterAutospacing` (HTML-origin "auto" spacing). Word renders these as
 * ~14px before/after regardless of any explicit `w:before`/`w:after` value.
 */
export declare const AUTO_PARAGRAPH_SPACING_PX = 14;
/**
 * Convert twips to pixels (at 96 DPI)
 *
 * 1 inch = 1440 twips = 96 pixels
 * → 1 twip = 96/1440 pixels = 1/15 pixels
 */
export declare function twipsToPixels(twips: number): number;
/**
 * Convert pixels to twips
 */
export declare function pixelsToTwips(px: number): number;
/**
 * Convert EMUs to twips
 */
export declare function emuToTwips(emu: number): number;
/**
 * Convert twips to EMUs
 */
export declare function twipsToEmu(twips: number): number;
/**
 * Convert half-points to pixels (at 96 DPI)
 *
 * Half-points are commonly used for font sizes in OOXML (w:sz).
 */
export declare function halfPointsToPixels(halfPoints: number): number;
/**
 * Convert half-points to points
 */
export declare function halfPointsToPoints(halfPoints: number): number;
/**
 * Convert points to half-points
 */
export declare function pointsToHalfPoints(points: number): number;
/**
 * Convert eighths of a point to pixels (at 96 DPI)
 *
 * Eighths of a point are used for border widths in OOXML.
 */
export declare function eighthsToPixels(eighths: number): number;
/**
 * Round a pixel value to avoid sub-pixel rendering issues
 */
export declare function roundPixels(px: number, decimalPlaces?: number): number;
/**
 * Clamp a value between min and max
 */
export declare function clamp(value: number, min: number, max: number): number;
/**
 * Format a pixel value as CSS string
 */
export declare function formatPx(px: number): string;
