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

import { hexToRgb, resolveColor, resolveThemeColorSlot } from './drawingmlColor';
import type { ColorValue, Theme, ThemeColorSlot } from '../types/document';

export {
  blendColors,
  darkenColor,
  generateThemeTintShadeMatrix,
  getContrastingColor,
  getThemeTintShadeHex,
  lightenColor,
  resolveColor,
  resolveColorToHex,
} from './drawingmlColor';
export type { ThemeMatrixCell } from './drawingmlColor';

/**
 * Highlight color mapping to hex values
 * These are the W3C standard colors for Word highlighting
 */
const HIGHLIGHT_COLORS: Record<string, string> = {
  black: '000000',
  blue: '0000FF',
  cyan: '00FFFF',
  darkBlue: '00008B',
  darkCyan: '008B8B',
  darkGray: 'A9A9A9',
  darkGreen: '006400',
  darkMagenta: '8B008B',
  darkRed: '8B0000',
  darkYellow: '808000',
  green: '00FF00',
  lightGray: 'D3D3D3',
  magenta: 'FF00FF',
  red: 'FF0000',
  white: 'FFFFFF',
  yellow: 'FFFF00',
  none: '',
};

/**
 * Resolve a highlight color name to CSS
 *
 * @param highlight - Highlight color name (e.g., "yellow", "cyan")
 * @returns CSS color string or empty string for "none"
 */
export function resolveHighlightColor(highlight: string | undefined): string {
  if (!highlight || highlight === 'none') {
    return '';
  }

  const hex = HIGHLIGHT_COLORS[highlight];
  return hex ? `#${hex}` : '';
}

/**
 * Resolve a shading fill or pattern color to CSS
 *
 * @param color - ColorValue for fill
 * @param theme - Theme for resolving theme colors
 * @returns CSS color string
 */
export function resolveShadingColor(
  color: ColorValue | undefined | null,
  theme: Theme | null | undefined
): string {
  if (!color) return '';

  // For shading, "auto" typically means transparent
  if (color.auto) {
    return 'transparent';
  }

  return resolveColor(color, theme);
}

/**
 * Check if a color is effectively black
 *
 * @param color - ColorValue object
 * @param theme - Theme for resolving theme colors
 * @returns True if color resolves to black or very dark
 */
export function isBlack(
  color: ColorValue | undefined | null,
  theme: Theme | null | undefined
): boolean {
  if (!color) return false;
  if (color.auto) return true;

  const resolved = resolveColor(color, theme);
  const hex = resolved.replace(/^#/, '').toLowerCase();

  // Check if it's black or very dark
  const rgb = hexToRgb(hex);
  const luminance = (rgb.r + rgb.g + rgb.b) / 3;

  return luminance < 20;
}

/**
 * Check if a color is effectively white
 *
 * @param color - ColorValue object
 * @param theme - Theme for resolving theme colors
 * @returns True if color resolves to white or very light
 */
export function isWhite(
  color: ColorValue | undefined | null,
  theme: Theme | null | undefined
): boolean {
  if (!color) return false;

  const resolved = resolveColor(color, theme);
  const hex = resolved.replace(/^#/, '').toLowerCase();

  // Check if it's white or very light
  const rgb = hexToRgb(hex);
  const luminance = (rgb.r + rgb.g + rgb.b) / 3;

  return luminance > 235;
}

/**
 * Parse a color string (various formats) to ColorValue
 *
 * @param colorString - Color string like "FF0000", "auto", or theme color name
 * @returns ColorValue object
 */
export function parseColorString(colorString: string | undefined): ColorValue | undefined {
  if (!colorString) return undefined;

  const normalized = colorString.trim();

  if (normalized.toLowerCase() === 'auto') {
    return { auto: true };
  }

  // Check if it's a theme color name
  const themeSlot = resolveThemeColorSlot(normalized);
  if (themeSlot) {
    return { themeColor: themeSlot };
  }

  // Assume it's an RGB hex value
  // Remove # if present and normalize to 6 chars
  const hex = normalized.replace(/^#/, '').toUpperCase();

  // Validate hex format
  if (/^[0-9A-F]{6}$/i.test(hex)) {
    return { rgb: hex };
  }

  // 3-character shorthand
  if (/^[0-9A-F]{3}$/i.test(hex)) {
    const expanded = hex
      .split('')
      .map((c) => c + c)
      .join('');
    return { rgb: expanded };
  }

  // Unknown format, return as RGB anyway
  return { rgb: hex.padStart(6, '0').slice(0, 6) };
}

/**
 * Create a ColorValue from theme color reference
 *
 * @param themeColor - Theme color slot name
 * @param tint - Optional tint modifier
 * @param shade - Optional shade modifier
 * @returns ColorValue object
 */
export function createThemeColor(
  themeColor: ThemeColorSlot,
  tint?: number,
  shade?: number
): ColorValue {
  const result: ColorValue = { themeColor };

  if (tint !== undefined && tint > 0 && tint < 1) {
    result.themeTint = Math.round(tint * 255)
      .toString(16)
      .toUpperCase()
      .padStart(2, '0');
  }

  if (shade !== undefined && shade > 0 && shade < 1) {
    result.themeShade = Math.round(shade * 255)
      .toString(16)
      .toUpperCase()
      .padStart(2, '0');
  }

  return result;
}

/**
 * Create a ColorValue from RGB hex
 *
 * @param hex - 6-character hex color (no #)
 * @returns ColorValue object
 */
export function createRgbColor(hex: string): ColorValue {
  return { rgb: hex.replace(/^#/, '').toUpperCase() };
}

// ============================================================================
// HEX UTILITIES
// ============================================================================

/**
 * Ensure a hex color string has a '#' prefix.
 */
export function ensureHexPrefix(hex: string): string {
  return hex.startsWith('#') ? hex : `#${hex}`;
}

/**
 * Resolve a highlight color value to a CSS-ready string.
 * Tries OOXML named highlight first, then ensures hex prefix.
 */
export function resolveHighlightToCss(value: string): string {
  return resolveHighlightColor(value) || ensureHexPrefix(value);
}

/**
 * Check if two colors are equal
 *
 * @param color1 - First color
 * @param color2 - Second color
 * @param theme - Theme for resolving
 * @returns True if colors resolve to the same value
 */
export function colorsEqual(
  color1: ColorValue | undefined | null,
  color2: ColorValue | undefined | null,
  theme: Theme | null | undefined
): boolean {
  if (!color1 && !color2) return true;
  if (!color1 || !color2) return false;

  const resolved1 = resolveColor(color1, theme).toUpperCase();
  const resolved2 = resolveColor(color2, theme).toUpperCase();

  return resolved1 === resolved2;
}
