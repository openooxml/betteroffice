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

import type { Theme, ThemeColorScheme } from './theme';

/**
 * Theme color slots from a theme's clrScheme (plus the WordprocessingML
 * alias slots background1/text1/background2/text2).
 *
 * @public
 */
export type ThemeColorSlot =
  | 'dk1'
  | 'lt1'
  | 'dk2'
  | 'lt2'
  | 'accent1'
  | 'accent2'
  | 'accent3'
  | 'accent4'
  | 'accent5'
  | 'accent6'
  | 'hlink'
  | 'folHlink'
  | 'background1'
  | 'text1'
  | 'background2'
  | 'text2';

/**
 * ECMA-376 color reference — either a direct RGB hex, a theme slot
 * reference (with optional tint/shade), or `auto` for context-dependent
 * defaults (usually black for text on light backgrounds). When both
 * `rgb` and `themeColor` are set, the theme wins on Word import and the
 * `rgb` acts as a fallback for renderers without theme support.
 *
 * See ECMA-376 §17.18.39 (`ST_ThemeColor`).
 *
 * @public
 */
export interface ColorValue {
  /** RGB hex value without # (e.g., "FF0000") */
  rgb?: string;
  /** Theme color slot reference */
  themeColor?: ThemeColorSlot;
  /** Tint modifier (0-255 as hex string, e.g., "80") - makes color lighter */
  themeTint?: string;
  /** Shade modifier (0-255 as hex string) - makes color darker */
  themeShade?: string;
  /** Auto color - context-dependent (usually black for text) */
  auto?: boolean;
}

/**
 * Default theme colors (Office 2016 default theme).
 *
 * Single source of truth for the fallback palette used when a document has
 * no (or a malformed) theme part.
 *
 * @public
 */
export const DEFAULT_THEME_COLORS: ThemeColorScheme = {
  dk1: '000000', // Black
  lt1: 'FFFFFF', // White
  dk2: '44546A', // Dark blue-gray
  lt2: 'E7E6E6', // Light gray
  accent1: '4472C4', // Blue
  accent2: 'ED7D31', // Orange
  accent3: 'A5A5A5', // Gray
  accent4: 'FFC000', // Gold
  accent5: '5B9BD5', // Light blue
  accent6: '70AD47', // Green
  hlink: '0563C1', // Hyperlink blue
  folHlink: '954F72', // Followed hyperlink purple
};

/**
 * Map alternative theme color names to standard slots
 * OOXML uses different names in different contexts
 */
const THEME_COLOR_ALIASES: Record<string, ThemeColorSlot> = {
  // Standard names
  dk1: 'dk1',
  lt1: 'lt1',
  dk2: 'dk2',
  lt2: 'lt2',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
  hlink: 'hlink',
  folHlink: 'folHlink',
  // Alternative names used in some OOXML contexts
  dark1: 'dk1',
  light1: 'lt1',
  dark2: 'dk2',
  light2: 'lt2',
  hyperlink: 'hlink',
  followedHyperlink: 'folHlink',
  // Background/text names (map to dk1/lt1)
  background1: 'lt1',
  text1: 'dk1',
  background2: 'lt2',
  text2: 'dk2',
  tx1: 'dk1',
  tx2: 'dk2',
  bg1: 'lt1',
  bg2: 'lt2',
};

/**
 * Parse a hex color modifier value (tint or shade)
 * OOXML stores tint/shade as hex string (00-FF) representing 0-255
 *
 * @param hexValue - Hex string like "80" or "FF"
 * @returns Decimal value 0-1
 */
function parseModifierValue(hexValue: string | undefined): number {
  if (!hexValue) return 1;

  const parsed = parseInt(hexValue, 16);
  if (isNaN(parsed)) return 1;

  // Value is 0-255, convert to 0-1
  return parsed / 255;
}

/**
 * Parse RGB hex color to component values
 *
 * @param hex - 6-character hex color (no #)
 * @returns RGB object with r, g, b values 0-255
 *
 * @public
 */
export function hexToRgb(hex: string): { r: number; g: number; b: number } {
  // Ensure 6 characters
  const normalized = hex.padStart(6, '0').slice(0, 6);

  const r = parseInt(normalized.slice(0, 2), 16);
  const g = parseInt(normalized.slice(2, 4), 16);
  const b = parseInt(normalized.slice(4, 6), 16);

  return {
    r: isNaN(r) ? 0 : r,
    g: isNaN(g) ? 0 : g,
    b: isNaN(b) ? 0 : b,
  };
}

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
export function rgbToHex(r: number, g: number, b: number): string {
  const toHex = (n: number) =>
    Math.max(0, Math.min(255, Math.round(n)))
      .toString(16)
      .padStart(2, '0');

  return `${toHex(r)}${toHex(g)}${toHex(b)}`.toUpperCase();
}

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
export function applyTint(hex: string, tint: number): string {
  if (tint >= 1) return hex;
  if (tint <= 0) return 'FFFFFF';

  // OOXML per-channel linear interpolation toward white:
  // new_channel = channel * t + 255 * (1 - t)
  const rgb = hexToRgb(hex);
  return rgbToHex(
    Math.min(255, Math.max(0, Math.round(rgb.r * tint + 255 * (1 - tint)))),
    Math.min(255, Math.max(0, Math.round(rgb.g * tint + 255 * (1 - tint)))),
    Math.min(255, Math.max(0, Math.round(rgb.b * tint + 255 * (1 - tint))))
  );
}

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
export function applyShade(hex: string, shade: number): string {
  if (shade >= 1) return hex;
  if (shade <= 0) return '000000';

  // OOXML per-channel linear interpolation toward black:
  // new_channel = channel * s
  const rgb = hexToRgb(hex);
  return rgbToHex(
    Math.min(255, Math.max(0, Math.round(rgb.r * shade))),
    Math.min(255, Math.max(0, Math.round(rgb.g * shade))),
    Math.min(255, Math.max(0, Math.round(rgb.b * shade)))
  );
}

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
export function getThemeColorValue(theme: Theme | null | undefined, slot: ThemeColorSlot): string {
  // Map alias slots to actual color scheme keys
  const schemeKey = THEME_COLOR_ALIASES[slot] ?? slot;

  // Define the actual keys that exist on ThemeColorScheme
  const schemeKeys = [
    'dk1',
    'lt1',
    'dk2',
    'lt2',
    'accent1',
    'accent2',
    'accent3',
    'accent4',
    'accent5',
    'accent6',
    'hlink',
    'folHlink',
  ] as const;
  type SchemeKey = (typeof schemeKeys)[number];

  const isSchemeKey = (key: string): key is SchemeKey => schemeKeys.includes(key as SchemeKey);

  if (!theme?.colorScheme) {
    if (isSchemeKey(schemeKey)) {
      return DEFAULT_THEME_COLORS[schemeKey] ?? '000000';
    }
    return '000000';
  }

  if (isSchemeKey(schemeKey)) {
    return theme.colorScheme[schemeKey] ?? DEFAULT_THEME_COLORS[schemeKey] ?? '000000';
  }

  return '000000';
}

/**
 * Resolve a theme color name to a standard slot
 *
 * @param colorName - Theme color name (could be alias)
 * @returns Standard ThemeColorSlot or null if unknown
 *
 * @public
 */
export function resolveThemeColorSlot(colorName: string): ThemeColorSlot | null {
  if (!colorName) return null;

  const normalized = colorName.toLowerCase();
  const slot = THEME_COLOR_ALIASES[colorName] ?? THEME_COLOR_ALIASES[normalized];

  return slot ?? null;
}

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
export function resolveColor(
  color: ColorValue | undefined | null,
  theme: Theme | null | undefined,
  defaultColor: string = '000000'
): string {
  if (!color) {
    return `#${defaultColor}`;
  }

  // Handle "auto" color
  if (color.auto) {
    // Auto typically means black for text, but can be context-dependent
    return `#${defaultColor}`;
  }

  let hexColor: string;

  // Check for theme color first
  if (color.themeColor) {
    const slot = resolveThemeColorSlot(color.themeColor);
    if (slot) {
      hexColor = getThemeColorValue(theme, slot);
    } else {
      // Unknown theme color, use RGB if available or default
      hexColor = color.rgb ?? defaultColor;
    }

    // Apply tint/shade modifiers
    if (color.themeTint) {
      const tintValue = parseModifierValue(color.themeTint);
      hexColor = applyTint(hexColor, tintValue);
    } else if (color.themeShade) {
      const shadeValue = parseModifierValue(color.themeShade);
      hexColor = applyShade(hexColor, shadeValue);
    }
  } else if (color.rgb) {
    // "auto" in OOXML means automatic color (typically black)
    hexColor = color.rgb === 'auto' ? defaultColor : color.rgb;
  } else {
    // No color specified
    hexColor = defaultColor;
  }

  // Ensure proper format
  return `#${hexColor.toUpperCase().replace(/^#/, '')}`;
}

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
export function resolveColorToHex(
  color: ColorValue | undefined | null,
  theme: Theme | null | undefined
): string | undefined {
  if (!color || color.auto) return undefined;

  if (color.themeColor && theme) {
    // resolveColor always returns `#XXXXXX`; drop the `#`.
    return resolveColor(color, theme).slice(1);
  }

  if (color.rgb && color.rgb !== 'auto') {
    return color.rgb.toUpperCase().replace(/^#/, '');
  }

  return undefined;
}

/**
 * Get contrasting text color for a background
 *
 * @param backgroundColor - Background ColorValue
 * @param theme - Theme for resolving theme colors
 * @returns Black or white hex color for best contrast
 *
 * @public
 */
export function getContrastingColor(
  backgroundColor: ColorValue | undefined | null,
  theme: Theme | null | undefined
): string {
  if (!backgroundColor) return '#000000';

  const bgResolved = resolveColor(backgroundColor, theme);
  const bgHex = bgResolved.replace(/^#/, '');
  const bgRgb = hexToRgb(bgHex);

  // Calculate relative luminance using sRGB formula
  const luminance = (0.299 * bgRgb.r + 0.587 * bgRgb.g + 0.114 * bgRgb.b) / 255;

  // Return black for light backgrounds, white for dark
  return luminance > 0.5 ? '#000000' : '#FFFFFF';
}

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
export function darkenColor(
  color: ColorValue | undefined | null,
  theme: Theme | null | undefined,
  percent: number
): string {
  const resolved = resolveColor(color, theme);
  const hex = resolved.replace(/^#/, '');
  // percent=80 means darken 80% → keep 20% of original
  const shade = 1 - percent / 100;
  return `#${applyShade(hex, shade)}`;
}

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
export function lightenColor(
  color: ColorValue | undefined | null,
  theme: Theme | null | undefined,
  percent: number
): string {
  const resolved = resolveColor(color, theme);
  const hex = resolved.replace(/^#/, '');
  // percent=80 means lighten 80% → keep 20% of original
  const tint = 1 - percent / 100;
  return `#${applyTint(hex, tint)}`;
}

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
export function blendColors(
  color1: ColorValue | undefined | null,
  color2: ColorValue | undefined | null,
  ratio: number,
  theme: Theme | null | undefined
): string {
  const resolved1 = resolveColor(color1, theme).replace(/^#/, '');
  const resolved2 = resolveColor(color2, theme).replace(/^#/, '');

  const rgb1 = hexToRgb(resolved1);
  const rgb2 = hexToRgb(resolved2);

  const blended = {
    r: Math.round(rgb1.r * (1 - ratio) + rgb2.r * ratio),
    g: Math.round(rgb1.g * (1 - ratio) + rgb2.g * ratio),
    b: Math.round(rgb1.b * (1 - ratio) + rgb2.b * ratio),
  };

  return `#${rgbToHex(blended.r, blended.g, blended.b)}`;
}

// ============================================================================
// THEME COLOR MATRIX FOR ADVANCED COLOR PICKER
// ============================================================================

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
 * Theme color column order matching Word's color picker:
 * Background 1 (lt1), Text 1 (dk1), Background 2 (lt2), Text 2 (dk2), Accent 1-6
 */
const THEME_MATRIX_COLUMNS: Array<{ slot: ThemeColorSlot; name: string }> = [
  { slot: 'lt1', name: 'Background 1' },
  { slot: 'dk1', name: 'Text 1' },
  { slot: 'lt2', name: 'Background 2' },
  { slot: 'dk2', name: 'Text 2' },
  { slot: 'accent1', name: 'Accent 1' },
  { slot: 'accent2', name: 'Accent 2' },
  { slot: 'accent3', name: 'Accent 3' },
  { slot: 'accent4', name: 'Accent 4' },
  { slot: 'accent5', name: 'Accent 5' },
  { slot: 'accent6', name: 'Accent 6' },
];

/**
 * Tint/shade row definitions matching Word's picker.
 * Row 0 = base, rows 1-3 = tints (lighter), rows 4-5 = shades (darker).
 */
const THEME_MATRIX_ROWS: Array<{
  type: 'base' | 'tint' | 'shade';
  value: number; // fraction 0-1
  hexValue: string; // OOXML hex modifier
  labelSuffix: string;
}> = [
  { type: 'base', value: 0, hexValue: '', labelSuffix: '' },
  { type: 'tint', value: 0.8, hexValue: 'CC', labelSuffix: ', Lighter 80%' },
  { type: 'tint', value: 0.6, hexValue: '99', labelSuffix: ', Lighter 60%' },
  { type: 'tint', value: 0.4, hexValue: '66', labelSuffix: ', Lighter 40%' },
  { type: 'shade', value: 0.75, hexValue: 'BF', labelSuffix: ', Darker 25%' },
  { type: 'shade', value: 0.5, hexValue: '80', labelSuffix: ', Darker 50%' },
];

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
export function getThemeTintShadeHex(
  baseHex: string,
  type: 'tint' | 'shade',
  fraction: number
): string {
  if (type === 'tint') {
    // fraction is "how much to lighten" (0 = no change, 1 = fully white)
    // applyTint wants "how much to keep" (1 = no change, 0 = fully white) → invert
    return applyTint(baseHex, 1 - fraction);
  }
  // fraction is "how much to keep" (1 = no change, 0 = fully black) — matches applyShade
  return applyShade(baseHex, fraction);
}

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
export function generateThemeTintShadeMatrix(
  colorScheme?: ThemeColorScheme | null
): ThemeMatrixCell[][] {
  const scheme = colorScheme ?? DEFAULT_THEME_COLORS;

  return THEME_MATRIX_ROWS.map((row) => {
    return THEME_MATRIX_COLUMNS.map((col) => {
      const baseHex =
        scheme[col.slot as keyof ThemeColorScheme] ??
        DEFAULT_THEME_COLORS[col.slot as keyof ThemeColorScheme] ??
        '000000';

      let hex: string;
      if (row.type === 'base') {
        hex = baseHex.toUpperCase();
      } else if (row.type === 'tint') {
        // row.value is "how much to lighten" (0.8 = 80% lighter)
        // applyTint wants "how much to keep" → invert
        hex = applyTint(baseHex, 1 - row.value);
      } else {
        // row.value for shade is "how much to keep" (0.75 = keep 75% = darken 25%)
        hex = applyShade(baseHex, row.value);
      }

      const cell: ThemeMatrixCell = {
        hex,
        themeSlot: col.slot,
        label: `${col.name}${row.labelSuffix}`,
      };

      if (row.type === 'tint' && row.hexValue) {
        cell.tint = row.hexValue;
      } else if (row.type === 'shade' && row.hexValue) {
        cell.shade = row.hexValue;
      }

      return cell;
    });
  });
}
