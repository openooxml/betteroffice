/**
 * Shared DrawingML parsing utilities — colors, fills, and outlines from
 * a:-namespace shape properties (a:solidFill, a:gradFill, a:ln, ...).
 */

import { getThemeColorValue, type ColorValue } from './color';
import type { ShapeFill, ShapeOutline } from './shape';
import { findByFullName, getAttribute, getChildElements, type XmlLike } from './xml';

// ============================================================================
// COLOR PARSING
// ============================================================================

/**
 * Map OOXML scheme names to standard theme color slots.
 * Used when parsing a:schemeClr elements in DrawingML.
 */
const SCHEME_TO_THEME_COLOR: Record<string, ColorValue['themeColor']> = {
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
  dk1: 'dk1',
  lt1: 'lt1',
  dk2: 'dk2',
  lt2: 'lt2',
  tx1: 'text1',
  tx2: 'text2',
  bg1: 'background1',
  bg2: 'background2',
  hlink: 'hlink',
  folHlink: 'folHlink',
};

/**
 * Common preset color names to RGB hex values.
 */
const PRESET_COLORS: Record<string, string> = {
  black: '000000',
  white: 'FFFFFF',
  red: 'FF0000',
  green: '00FF00',
  blue: '0000FF',
  yellow: 'FFFF00',
  cyan: '00FFFF',
  magenta: 'FF00FF',
};

/**
 * Apply color modifiers (shade, tint) from child elements of a color element.
 * Converts DrawingML 100000ths-scale values to hex (0-FF) for OOXML compatibility.
 */
function applyColorModifiers(color: ColorValue, element: XmlLike): ColorValue {
  const children = getChildElements(element);

  const shade = children.find((el) => el.name === 'a:shade');
  if (shade) {
    const val = getAttribute(shade, null, 'val');
    if (val) {
      color.themeShade = Math.round((parseInt(val, 10) / 100000) * 255)
        .toString(16)
        .padStart(2, '0')
        .toUpperCase();
    }
  }

  const tint = children.find((el) => el.name === 'a:tint');
  if (tint) {
    const val = getAttribute(tint, null, 'val');
    if (val) {
      color.themeTint = Math.round((parseInt(val, 10) / 100000) * 255)
        .toString(16)
        .padStart(2, '0')
        .toUpperCase();
    }
  }

  return color;
}

/**
 * Parse a color value from a DrawingML element.
 * Handles: a:srgbClr, a:schemeClr, a:sysClr, a:prstClr
 * Applies shade/tint modifiers when present.
 *
 * @public
 */
export function parseColorElement(element: XmlLike | null): ColorValue | undefined {
  if (!element) return undefined;

  const children = getChildElements(element);

  // sRGB color: a:srgbClr[@val]
  const srgbClr = children.find((el) => el.name === 'a:srgbClr');
  if (srgbClr) {
    const val = getAttribute(srgbClr, null, 'val');
    if (val) {
      return applyColorModifiers({ rgb: val }, srgbClr);
    }
  }

  // Scheme color (theme): a:schemeClr[@val]
  const schemeClr = children.find((el) => el.name === 'a:schemeClr');
  if (schemeClr) {
    const val = getAttribute(schemeClr, null, 'val');
    if (val) {
      const color: ColorValue = {
        themeColor: SCHEME_TO_THEME_COLOR[val] ?? 'dk1',
      };
      return applyColorModifiers(color, schemeClr);
    }
  }

  // System color: a:sysClr[@lastClr]
  const sysClr = children.find((el) => el.name === 'a:sysClr');
  if (sysClr) {
    const lastClr = getAttribute(sysClr, null, 'lastClr');
    return { rgb: lastClr ?? '000000' };
  }

  // Preset color: a:prstClr[@val]
  const prstClr = children.find((el) => el.name === 'a:prstClr');
  if (prstClr) {
    const val = getAttribute(prstClr, null, 'val');
    if (val && PRESET_COLORS[val]) {
      return { rgb: PRESET_COLORS[val] };
    }
  }

  return undefined;
}

// ============================================================================
// FILL & OUTLINE PARSING
// ============================================================================

/**
 * Parse fill from shape properties (a:solidFill, a:noFill, a:gradFill).
 *
 * @public
 */
export function parseFill(spPr: XmlLike | null): ShapeFill | undefined {
  if (!spPr) return undefined;

  const children = getChildElements(spPr);

  if (children.find((el) => el.name === 'a:noFill')) {
    return { type: 'none' };
  }

  const solidFill = children.find((el) => el.name === 'a:solidFill');
  if (solidFill) {
    return { type: 'solid', color: parseColorElement(solidFill) };
  }

  if (children.find((el) => el.name === 'a:gradFill')) {
    return { type: 'gradient' };
  }

  return undefined;
}

/**
 * Parse outline from shape properties (a:ln).
 *
 * @public
 */
export function parseOutline(spPr: XmlLike | null): ShapeOutline | undefined {
  const ln = spPr ? findByFullName(spPr, 'a:ln') : null;
  if (!ln) return undefined;

  const children = getChildElements(ln);

  if (children.find((el) => el.name === 'a:noFill')) {
    return undefined;
  }

  const outline: ShapeOutline = {};

  const w = getAttribute(ln, null, 'w');
  if (w) outline.width = parseInt(w, 10);

  const solidFill = children.find((el) => el.name === 'a:solidFill');
  if (solidFill) outline.color = parseColorElement(solidFill);

  const prstDash = children.find((el) => el.name === 'a:prstDash');
  if (prstDash) {
    const val = getAttribute(prstDash, null, 'val');
    if (val) outline.style = val as ShapeOutline['style'];
  }

  return outline;
}

// ============================================================================
// COLOR RESOLUTION (for shapes/text boxes without theme context)
// ============================================================================

/**
 * Resolve a ColorValue to a CSS hex string using the default theme colors
 * (Office 2016 palette — see {@link DEFAULT_THEME_COLORS}).
 * For use when no Theme object is available (e.g., shape/text box parsing).
 *
 * @public
 */
export function resolveColorValueToHex(color: ColorValue | undefined): string | undefined {
  if (!color) return undefined;

  if (color.rgb) return `#${color.rgb}`;

  if (color.themeColor) {
    return `#${getThemeColorValue(null, color.themeColor)}`;
  }

  return undefined;
}
