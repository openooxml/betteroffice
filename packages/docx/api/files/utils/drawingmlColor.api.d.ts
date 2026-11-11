// @betteroffice/docx declaration utils/drawingmlColor.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Color primitives and math shared across OOXML formats.
 *
 * Handles theme color slot references (accent1, dk1, ...), RGB hex values,
 * tint/shade modifiers, and the derived color helpers (lighten/darken/blend/
 * contrast) built on top of them.
 *
 * Tint/Shade Calculations (ECMA-376 §17.3.2.41):
 * - Tint makes a color lighter (per-channel interpolation toward white)
 * - Shade makes a color darker (per-channel interpolation toward black)
 * - The wire value is a hex byte (00-FF) meaning "how much color to keep"
 */
import type { ColorValue, Theme, ThemeColorScheme, ThemeColorSlot } from '../types/drawingml';
/**
 * Default theme colors (Office 2016 default theme).
 *
 * Single source of truth for the fallback palette used when a document has
 * no (or a malformed) theme part.
 *
 * @public
 */
export declare const DEFAULT_THEME_COLORS: ThemeColorScheme;
/**
 * Parse RGB hex color to component values
 *
 * @param hex - 6-character hex color (no #)
 * @returns RGB object with r, g, b values 0-255
 *
 * @public
 */
export declare function hexToRgb(hex: string): {
    r: number;
    g: number;
    b: number;
};
/**
 * Convert RGB values to hex color
 *
 * @param r - Red 0-255
 * @param g - Green 0-255
 * @param b - Blue 0-255
 * @returns 6-character hex color (no #)
 *
 * @public
 */
export declare function rgbToHex(r: number, g: number, b: number): string;
/**
 * Apply tint to a color (make lighter by blending with white)
 *
 * @param hex - 6-character hex color (no #)
 * @param tint - Tint value 0-1: how much of the original color to keep.
 *   1 = no change, 0 = fully white. Per ECMA-376 §17.3.2.41.
 * @returns Modified hex color
 *
 * @public
 */
export declare function applyTint(hex: string, tint: number): string;
/**
 * Apply shade to a color (make darker by blending with black)
 *
 * @param hex - 6-character hex color (no #)
 * @param shade - Shade value 0-1: how much of the original color to keep.
 *   1 = no change, 0 = fully black. Per ECMA-376 §17.3.2.41.
 * @returns Modified hex color
 *
 * @public
 */
export declare function applyShade(hex: string, shade: number): string;
/**
 * Get a theme color by slot name, falling back to the Office 2016 defaults
 * when the theme is missing or lacks the slot.
 *
 * @param theme - Theme object (or null/undefined for defaults)
 * @param slot - Color slot name (aliases like background1/text2 resolve)
 * @returns Hex color (6 characters, no #)
 *
 * @public
 */
export declare function getThemeColorValue(theme: Theme | null | undefined, slot: ThemeColorSlot): string;
/**
 * Resolve a theme color name to a standard slot
 *
 * @param colorName - Theme color name (could be alias)
 * @returns Standard ThemeColorSlot or null if unknown
 *
 * @public
 */
export declare function resolveThemeColorSlot(colorName: string): ThemeColorSlot | null;
/**
 * Resolve a ColorValue to a CSS color string
 *
 * @param color - ColorValue object with rgb, themeColor, tint/shade, or auto
 * @param theme - Theme for resolving theme colors
 * @param defaultColor - Default color if auto or undefined (default: black)
 * @returns CSS color string (e.g., "#FF0000" or "inherit")
 *
 * @public
 */
export declare function resolveColor(color: ColorValue | undefined | null, theme: Theme | null | undefined, defaultColor?: string): string;
/**
 * Resolve any ColorValue (text, fill/shading, border, underline) to a 6-char
 * uppercase hex string — or `undefined` if transparent/unset/unresolvable.
 *
 * Shared display-side resolver. Prefer this over reading `.rgb` directly so
 * that `themeColor` + `themeTint`/`themeShade` are honored consistently across
 * all render paths.
 *
 * When a themed color is present but `theme` is null/undefined, falls back to
 * `color.rgb` if Word wrote one for compat; otherwise returns `undefined`.
 *
 * @returns 6-char uppercase hex without `#`, or `undefined`.
 *
 * @public
 */
export declare function resolveColorToHex(color: ColorValue | undefined | null, theme: Theme | null | undefined): string | undefined;
/**
 * Get contrasting text color for a background
 *
 * @param backgroundColor - Background ColorValue
 * @param theme - Theme for resolving theme colors
 * @returns Black or white hex color for best contrast
 *
 * @public
 */
export declare function getContrastingColor(backgroundColor: ColorValue | undefined | null, theme: Theme | null | undefined): string;
/**
 * Darken a color by a percentage
 *
 * @param color - ColorValue to darken
 * @param theme - Theme for resolving
 * @param percent - Percentage to darken (0-100)
 * @returns CSS color string
 *
 * @public
 */
export declare function darkenColor(color: ColorValue | undefined | null, theme: Theme | null | undefined, percent: number): string;
/**
 * Lighten a color by a percentage
 *
 * @param color - ColorValue to lighten
 * @param theme - Theme for resolving
 * @param percent - Percentage to lighten (0-100)
 * @returns CSS color string
 *
 * @public
 */
export declare function lightenColor(color: ColorValue | undefined | null, theme: Theme | null | undefined, percent: number): string;
/**
 * Blend two colors together
 *
 * @param color1 - First color
 * @param color2 - Second color
 * @param ratio - Blend ratio (0 = all color1, 1 = all color2)
 * @param theme - Theme for resolving
 * @returns CSS color string
 *
 * @public
 */
export declare function blendColors(color1: ColorValue | undefined | null, color2: ColorValue | undefined | null, ratio: number, theme: Theme | null | undefined): string;
/**
 * Theme color matrix cell
 *
 * @public
 */
export interface ThemeMatrixCell {
    /** Resolved hex color (6 chars, no #) */
    hex: string;
    /** Theme color slot */
    themeSlot: ThemeColorSlot;
    /** Tint hex modifier if applicable (e.g., "CC") */
    tint?: string;
    /** Shade hex modifier if applicable (e.g., "BF") */
    shade?: string;
    /** Human-readable label (e.g., "Accent 1, Lighter 60%") */
    label: string;
}
/**
 * Compute a single tinted or shaded hex color from a base color.
 *
 * @param baseHex - 6-character hex color (no #)
 * @param type - 'tint' to lighten, 'shade' to darken
 * @param fraction - Amount (0-1). For tint: 0=no change, 1=white. For shade: 0=black, 1=no change.
 * @returns 6-character hex color (no #)
 *
 * @public
 */
export declare function getThemeTintShadeHex(baseHex: string, type: 'tint' | 'shade', fraction: number): string;
/**
 * Generate the 10×6 theme color matrix for an advanced color picker.
 *
 * Columns: lt1, dk1, lt2, dk2, accent1-6 (matches Word's order)
 * Rows: base, 80% tint, 60% tint, 40% tint, 25% shade, 50% shade
 *
 * @param colorScheme - Theme color scheme (falls back to Office 2016 defaults)
 * @returns 6 rows × 10 columns of ThemeMatrixCell
 *
 * @public
 */
export declare function generateThemeTintShadeMatrix(colorScheme?: ThemeColorScheme | null): ThemeMatrixCell[][];
