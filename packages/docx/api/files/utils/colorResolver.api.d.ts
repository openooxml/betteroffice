// @betteroffice/docx declaration utils/colorResolver.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Color Resolver - Convert OOXML colors to CSS
 *
 * The generic color machinery (theme slots, tint/shade math, blending, the
 * picker matrix) is re-exported here. This module keeps the WordprocessingML-
 * specific pieces: w:highlight named colors, w:shd shading semantics, and
 * ColorValue construction helpers.
 *
 * OOXML Color References:
 * - w:color/@w:val - RGB hex or "auto"
 * - w:color/@w:themeColor - Theme color slot
 * - w:color/@w:themeTint - Tint modifier (0-255, hex)
 * - w:color/@w:themeShade - Shade modifier (0-255, hex)
 */
import type { ColorValue, Theme, ThemeColorSlot } from '../types/document';
export { blendColors, darkenColor, generateThemeTintShadeMatrix, getContrastingColor, getThemeTintShadeHex, lightenColor, resolveColor, resolveColorToHex, } from './drawingmlColor';
export type { ThemeMatrixCell } from './drawingmlColor';
/**
 * Resolve a highlight color name to CSS
 *
 * @param highlight - Highlight color name (e.g., "yellow", "cyan")
 * @returns CSS color string or empty string for "none"
 */
export declare function resolveHighlightColor(highlight: string | undefined): string;
/**
 * Resolve a shading fill or pattern color to CSS
 *
 * @param color - ColorValue for fill
 * @param theme - Theme for resolving theme colors
 * @returns CSS color string
 */
export declare function resolveShadingColor(color: ColorValue | undefined | null, theme: Theme | null | undefined): string;
/**
 * Check if a color is effectively black
 *
 * @param color - ColorValue object
 * @param theme - Theme for resolving theme colors
 * @returns True if color resolves to black or very dark
 */
export declare function isBlack(color: ColorValue | undefined | null, theme: Theme | null | undefined): boolean;
/**
 * Check if a color is effectively white
 *
 * @param color - ColorValue object
 * @param theme - Theme for resolving theme colors
 * @returns True if color resolves to white or very light
 */
export declare function isWhite(color: ColorValue | undefined | null, theme: Theme | null | undefined): boolean;
/**
 * Parse a color string (various formats) to ColorValue
 *
 * @param colorString - Color string like "FF0000", "auto", or theme color name
 * @returns ColorValue object
 */
export declare function parseColorString(colorString: string | undefined): ColorValue | undefined;
/**
 * Create a ColorValue from theme color reference
 *
 * @param themeColor - Theme color slot name
 * @param tint - Optional tint modifier
 * @param shade - Optional shade modifier
 * @returns ColorValue object
 */
export declare function createThemeColor(themeColor: ThemeColorSlot, tint?: number, shade?: number): ColorValue;
/**
 * Create a ColorValue from RGB hex
 *
 * @param hex - 6-character hex color (no #)
 * @returns ColorValue object
 */
export declare function createRgbColor(hex: string): ColorValue;
/**
 * Ensure a hex color string has a '#' prefix.
 */
export declare function ensureHexPrefix(hex: string): string;
/**
 * Resolve a highlight color value to a CSS-ready string.
 * Tries OOXML named highlight first, then ensures hex prefix.
 */
export declare function resolveHighlightToCss(value: string): string;
/**
 * Check if two colors are equal
 *
 * @param color1 - First color
 * @param color2 - Second color
 * @param theme - Theme for resolving
 * @returns True if colors resolve to the same value
 */
export declare function colorsEqual(color1: ColorValue | undefined | null, color2: ColorValue | undefined | null, theme: Theme | null | undefined): boolean;
