/**
 * Theme parsing — colors and fonts from a DrawingML theme part
 * (e.g. word/theme/theme1.xml, ppt/theme/theme1.xml).
 *
 * Extracts color scheme (dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink)
 * and font scheme (majorFont, minorFont) from the theme.
 *
 * OOXML Reference:
 * - Uses DrawingML namespace (a:)
 * - Colors can be srgbClr, sysClr, or schemeClr
 */

import { DEFAULT_THEME_COLORS } from './color';
import {
  findChild,
  findChildren,
  getAttribute,
  getChildElements,
  getLocalName,
  type XmlLike,
} from './xml';

/**
 * Theme color scheme (a:clrScheme)
 *
 * @public
 */
export interface ThemeColorScheme {
  /** Dark 1 color (usually black) */
  dk1?: string;
  /** Light 1 color (usually white) */
  lt1?: string;
  /** Dark 2 color */
  dk2?: string;
  /** Light 2 color */
  lt2?: string;
  /** Accent colors 1-6 */
  accent1?: string;
  accent2?: string;
  accent3?: string;
  accent4?: string;
  accent5?: string;
  accent6?: string;
  /** Hyperlink color */
  hlink?: string;
  /** Followed hyperlink color */
  folHlink?: string;
}

/**
 * Theme font (with script variants)
 *
 * @public
 */
export interface ThemeFont {
  /** Latin font */
  latin?: string;
  /** East Asian font */
  ea?: string;
  /** Complex script font */
  cs?: string;
  /** Script-specific fonts */
  fonts?: Record<string, string>;
}

/**
 * Theme font scheme (a:fontScheme)
 *
 * @public
 */
export interface ThemeFontScheme {
  /** Major font (headings) */
  majorFont?: ThemeFont;
  /** Minor font (body text) */
  minorFont?: ThemeFont;
}

/**
 * Theme (from a theme1.xml part)
 *
 * @public
 */
export interface Theme {
  /** Theme name */
  name?: string;
  /** Color scheme */
  colorScheme?: ThemeColorScheme;
  /** Font scheme */
  fontScheme?: ThemeFontScheme;
  /** Format scheme (fills, lines, effects) - simplified */
  formatScheme?: {
    name?: string;
  };
}

/**
 * Default font scheme
 */
const DEFAULT_FONTS: ThemeFontScheme = {
  majorFont: {
    latin: 'Calibri Light',
    ea: '',
    cs: '',
    fonts: {},
  },
  minorFont: {
    latin: 'Calibri',
    ea: '',
    cs: '',
    fonts: {},
  },
};

/**
 * Default theme when no theme part exists
 */
const DEFAULT_THEME: Theme = {
  name: 'Office Theme',
  colorScheme: DEFAULT_THEME_COLORS,
  fontScheme: DEFAULT_FONTS,
};

/**
 * Color slot names in theme
 */
const COLOR_SLOTS = [
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

/**
 * Parse a color element (srgbClr, sysClr, or schemeClr)
 *
 * @param element - Color child element
 * @returns Hex color value (6 characters, no #)
 */
function parseThemeColorElement(element: XmlLike | null): string | null {
  if (!element) return null;

  const localName = getLocalName(element.name || '');

  switch (localName) {
    case 'srgbClr': {
      // Direct RGB color: <a:srgbClr val="4472C4"/>
      const val = getAttribute(element, 'a', 'val') ?? getAttribute(element, null, 'val');
      return val ?? null;
    }

    case 'sysClr': {
      // System color with fallback: <a:sysClr val="windowText" lastClr="000000"/>
      // Use lastClr as the fallback since we can't access actual system colors
      const lastClr =
        getAttribute(element, 'a', 'lastClr') ?? getAttribute(element, null, 'lastClr');
      if (lastClr) return lastClr;

      // Fallback based on common system color names
      const val = getAttribute(element, 'a', 'val') ?? getAttribute(element, null, 'val');
      switch (val) {
        case 'windowText':
        case 'menuText':
        case 'captionText':
        case 'btnText':
          return '000000';
        case 'window':
        case 'menu':
        case 'btnFace':
        case 'btnHighlight':
          return 'FFFFFF';
        case 'highlight':
          return '0078D7';
        case 'highlightText':
          return 'FFFFFF';
        case 'grayText':
          return '808080';
        default:
          return null;
      }
    }

    case 'schemeClr': {
      // Reference to another scheme color - rare in color scheme itself
      // Usually found in fill/line definitions with modifiers
      // For the color scheme, we just need the val
      const val = getAttribute(element, 'a', 'val') ?? getAttribute(element, null, 'val');
      // This is a reference, not a final color - return null for now
      // The actual resolution would need the full color scheme
      return val === 'phClr' ? null : null;
    }

    default:
      return null;
  }
}

/**
 * Parse the color scheme from a:clrScheme element
 *
 * @param clrScheme - The a:clrScheme element
 * @returns ThemeColorScheme with resolved hex colors
 */
function parseColorScheme(clrScheme: XmlLike | null): ThemeColorScheme {
  const result: ThemeColorScheme = { ...DEFAULT_THEME_COLORS };

  if (!clrScheme) return result;

  // Each color slot has a child element with the slot name
  for (const slot of COLOR_SLOTS) {
    // Find the slot element (e.g., a:dk1, a:accent1)
    const slotElement = findChild(clrScheme, 'a', slot);

    if (slotElement) {
      // The actual color is a child (srgbClr, sysClr, or schemeClr)
      const children = getChildElements(slotElement);
      if (children.length > 0) {
        const color = parseThemeColorElement(children[0]);
        if (color) {
          result[slot] = color;
        }
      }
    }
  }

  return result;
}

/**
 * Parse a font definition (majorFont or minorFont)
 *
 * @param fontElement - The a:majorFont or a:minorFont element
 * @returns ThemeFont with font family names
 */
function parseThemeFonts(fontElement: XmlLike | null): ThemeFont {
  const result: ThemeFont = {
    latin: '',
    ea: '',
    cs: '',
    fonts: {},
  };

  if (!fontElement) return result;

  // Parse main font elements
  const latinEl = findChild(fontElement, 'a', 'latin');
  if (latinEl) {
    result.latin =
      getAttribute(latinEl, 'a', 'typeface') ?? getAttribute(latinEl, null, 'typeface') ?? '';
  }

  const eaEl = findChild(fontElement, 'a', 'ea');
  if (eaEl) {
    result.ea = getAttribute(eaEl, 'a', 'typeface') ?? getAttribute(eaEl, null, 'typeface') ?? '';
  }

  const csEl = findChild(fontElement, 'a', 'cs');
  if (csEl) {
    result.cs = getAttribute(csEl, 'a', 'typeface') ?? getAttribute(csEl, null, 'typeface') ?? '';
  }

  // Parse script-specific fonts (a:font elements with script attribute)
  const fontElements = findChildren(fontElement, 'a', 'font');
  for (const font of fontElements) {
    const script = getAttribute(font, 'a', 'script') ?? getAttribute(font, null, 'script');
    const typeface = getAttribute(font, 'a', 'typeface') ?? getAttribute(font, null, 'typeface');

    if (script && typeface) {
      result.fonts = result.fonts || {};
      result.fonts[script] = typeface;
    }
  }

  return result;
}

/**
 * Parse the font scheme from a:fontScheme element
 *
 * @param fontScheme - The a:fontScheme element
 * @returns ThemeFontScheme with major and minor fonts
 */
function parseFontScheme(fontScheme: XmlLike | null): ThemeFontScheme {
  const result: ThemeFontScheme = { ...DEFAULT_FONTS };

  if (!fontScheme) return result;

  const majorFontEl = findChild(fontScheme, 'a', 'majorFont');
  if (majorFontEl) {
    result.majorFont = parseThemeFonts(majorFontEl);
  }

  const minorFontEl = findChild(fontScheme, 'a', 'minorFont');
  if (minorFontEl) {
    result.minorFont = parseThemeFonts(minorFontEl);
  }

  return result;
}

/**
 * Parse a theme part from its root `a:theme` element.
 *
 * The host is responsible for XML parsing; pass the parsed root element
 * (any {@link XmlLike}-compatible tree).
 *
 * @param root - Root a:theme element, or null/undefined if not present
 * @returns Parsed Theme object with colors and fonts
 *
 * @public
 */
export function parseTheme(root: XmlLike | null | undefined): Theme {
  // Return defaults if no theme element
  if (!root) {
    return { ...DEFAULT_THEME };
  }

  try {
    // Get theme name from root element
    const themeName =
      getAttribute(root, 'a', 'name') ?? getAttribute(root, null, 'name') ?? 'Office Theme';

    // Find a:themeElements which contains clrScheme and fontScheme
    const themeElements = findChild(root, 'a', 'themeElements');

    // Parse color scheme
    const clrScheme = findChild(themeElements, 'a', 'clrScheme');
    const colorScheme = parseColorScheme(clrScheme);

    // Parse font scheme
    const fontSchemeEl = findChild(themeElements, 'a', 'fontScheme');
    const fontScheme = parseFontScheme(fontSchemeEl);

    return {
      name: themeName,
      colorScheme,
      fontScheme,
    };
  } catch (error) {
    console.warn('Failed to parse theme:', error);
    return { ...DEFAULT_THEME };
  }
}

/**
 * Map a `w:themeFontLang` EastAsian language tag to the DrawingML `script`
 * attribute used by the theme's `<a:font script="…">` entries.
 */
function eastAsiaLangToScript(lang: string): string | null {
  const l = lang.toLowerCase();
  if (l.startsWith('ja')) return 'Jpan';
  if (l.startsWith('ko')) return 'Hang';
  if (l.startsWith('zh')) {
    // Traditional Chinese locales use the Hant typeface; everything else Hans.
    if (l.includes('hant') || /-(tw|hk|mo)\b/.test(l)) return 'Hant';
    return 'Hans';
  }
  return null;
}

/**
 * Map a `w:themeFontLang` bidi/complex-script language tag to a DrawingML
 * `script` attribute.
 */
function bidiLangToScript(lang: string): string | null {
  const l = lang.toLowerCase();
  if (l.startsWith('ar')) return 'Arab';
  if (l.startsWith('he') || l.startsWith('iw')) return 'Hebr';
  if (l.startsWith('th')) return 'Thai';
  if (l.startsWith('hi') || l.startsWith('mr') || l.startsWith('ne')) return 'Deva';
  return null;
}

/**
 * Resolve empty EastAsian/complex-script theme font slots using the document's
 * `w:themeFontLang`.
 *
 * Office's default theme leaves `<a:ea typeface="">`/`<a:cs typeface="">` empty
 * and instead lists the real typeface per script (`<a:font script="Jpan" …>`).
 * Word selects among those by the user/document language declared in
 * `w:themeFontLang`. We mirror that here so a run that references
 * `minorEastAsia` in a Japanese document resolves to `ＭＳ 明朝` rather than an
 * empty string — an empty font breaks line-break measurement (canvas falls back
 * to its 10px default) while the painter renders CJK glyphs at full width, so
 * text overflows the right margin.
 *
 * Mutates the theme in place. This is safe for round-tripping because the
 * theme part is preserved verbatim on save; the parsed `Theme` is only
 * consumed for layout/rendering.
 *
 * @public
 */
export function applyThemeFontLang(
  theme: Theme | null | undefined,
  themeFontLang: { eastAsia?: string; bidi?: string } | undefined
): void {
  if (!theme?.fontScheme || !themeFontLang) return;

  const eaScript = themeFontLang.eastAsia ? eastAsiaLangToScript(themeFontLang.eastAsia) : null;
  const csScript = themeFontLang.bidi ? bidiLangToScript(themeFontLang.bidi) : null;

  for (const font of [theme.fontScheme.majorFont, theme.fontScheme.minorFont]) {
    if (!font) continue;
    if (!font.ea && eaScript && font.fonts?.[eaScript]) {
      font.ea = font.fonts[eaScript];
    }
    if (!font.cs && csScript && font.fonts?.[csScript]) {
      font.cs = font.fonts[csScript];
    }
  }
}

/**
 * Get a color from the theme by slot name
 *
 * @param theme - Parsed theme
 * @param slot - Color slot name (dk1, lt1, accent1, etc.)
 * @returns Hex color value (6 characters, no #)
 *
 * @public
 */
export function getThemeColor(
  theme: Theme | null | undefined,
  slot: keyof ThemeColorScheme
): string {
  if (!theme?.colorScheme) {
    return DEFAULT_THEME_COLORS[slot] ?? '000000';
  }

  return theme.colorScheme[slot] ?? DEFAULT_THEME_COLORS[slot] ?? '000000';
}

/**
 * Get the major font (heading font) from theme
 *
 * @param theme - Parsed theme
 * @param script - Optional script code (defaults to latin)
 * @returns Font family name
 *
 * @public
 */
export function getMajorFont(theme: Theme | null | undefined, script: string = 'latin'): string {
  if (!theme?.fontScheme?.majorFont) {
    return DEFAULT_FONTS.majorFont?.latin ?? 'Calibri Light';
  }

  const majorFont = theme.fontScheme.majorFont;

  if (script === 'latin') return majorFont.latin || 'Calibri Light';
  if (script === 'ea') return majorFont.ea || '';
  if (script === 'cs') return majorFont.cs || '';

  // Check script-specific fonts
  if (majorFont.fonts?.[script]) {
    return majorFont.fonts[script];
  }

  // Default to latin
  return majorFont.latin || 'Calibri Light';
}

/**
 * Get the minor font (body font) from theme
 *
 * @param theme - Parsed theme
 * @param script - Optional script code (defaults to latin)
 * @returns Font family name
 *
 * @public
 */
export function getMinorFont(theme: Theme | null | undefined, script: string = 'latin'): string {
  if (!theme?.fontScheme?.minorFont) {
    return DEFAULT_FONTS.minorFont?.latin ?? 'Calibri';
  }

  const minorFont = theme.fontScheme.minorFont;

  if (script === 'latin') return minorFont.latin || 'Calibri';
  if (script === 'ea') return minorFont.ea || '';
  if (script === 'cs') return minorFont.cs || '';

  // Check script-specific fonts
  if (minorFont.fonts?.[script]) {
    return minorFont.fonts[script];
  }

  // Default to latin
  return minorFont.latin || 'Calibri';
}

/**
 * Resolve a theme font reference to an actual font name
 *
 * Theme font references are like: majorAscii, majorHAnsi, minorAscii, minorHAnsi, etc.
 *
 * @param theme - Parsed theme
 * @param themeRef - Theme font reference
 * @returns Font family name
 *
 * @public
 */
export function resolveThemeFontRef(theme: Theme | null | undefined, themeRef: string): string {
  if (!themeRef) return 'Calibri';

  // Parse the reference: major/minor + script type
  const isMajor = themeRef.toLowerCase().includes('major');
  const isMinor = themeRef.toLowerCase().includes('minor');

  // Determine script from reference
  let script = 'latin';
  const lowerRef = themeRef.toLowerCase();

  if (lowerRef.includes('eastasia')) {
    script = 'ea';
  } else if (lowerRef.includes('bidi') || lowerRef.includes('cs')) {
    script = 'cs';
  }
  // ascii and hAnsi both map to latin

  if (isMajor) {
    return getMajorFont(theme, script);
  } else if (isMinor) {
    return getMinorFont(theme, script);
  }

  // Default to minor latin
  return getMinorFont(theme, 'latin');
}

/**
 * Get all font families from the theme for preloading
 *
 * @param theme - Parsed theme
 * @returns Array of unique font family names
 *
 * @public
 */
export function getThemeFonts(theme: Theme | null | undefined): string[] {
  const fonts = new Set<string>();

  if (theme?.fontScheme) {
    const { majorFont, minorFont } = theme.fontScheme;

    // Add main fonts
    if (majorFont?.latin) fonts.add(majorFont.latin);
    if (majorFont?.ea) fonts.add(majorFont.ea);
    if (majorFont?.cs) fonts.add(majorFont.cs);

    if (minorFont?.latin) fonts.add(minorFont.latin);
    if (minorFont?.ea) fonts.add(minorFont.ea);
    if (minorFont?.cs) fonts.add(minorFont.cs);

    // Add script-specific fonts
    if (majorFont?.fonts) {
      for (const font of Object.values(majorFont.fonts)) {
        if (font) fonts.add(font);
      }
    }

    if (minorFont?.fonts) {
      for (const font of Object.values(minorFont.fonts)) {
        if (font) fonts.add(font);
      }
    }
  }

  // Remove empty strings
  fonts.delete('');

  return Array.from(fonts);
}

/**
 * Get the default theme (Office 2016 theme)
 *
 * @returns Default Theme object
 *
 * @public
 */
export function getDefaultTheme(): Theme {
  return { ...DEFAULT_THEME };
}
