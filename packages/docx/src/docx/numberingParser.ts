/**
 * Numbering/List Parser for DOCX
 *
 * Parses numbering.xml to extract:
 * - Abstract numbering definitions (templates with levels)
 * - Numbering instances (concrete references with optional overrides)
 *
 * OOXML Structure:
 * - w:abstractNum - Template definitions with 9 levels (0-8)
 * - w:num - Instances that reference abstractNum and can override levels
 * - w:lvl - Level definition with start, format, text pattern, etc.
 */

import type {
  NumberingDefinitions,
  AbstractNumbering,
  NumberingInstance,
  ListLevel,
  ListRendering,
  NumberFormat,
} from '../types/document';
import { parseDocumentWithRust } from './rustParseFacade';
import { rezipContainer } from './wasm';

export type NumberingMap = {
  definitions: NumberingDefinitions;
  getLevel: (numId: number, ilvl: number) => ListLevel | null;
  getAbstract: (abstractNumId: number) => AbstractNumbering | null;
  getInstance: (numId: number) => NumberingInstance | null;
  hasNumbering: (numId: number) => boolean;
};

const UTF8 = new TextEncoder();

/** Rust-backed adapter for the published raw-numbering.xml leaf API. */
export function parseNumbering(numberingXml: string | null): NumberingMap {
  if (!numberingXml) return createNumberingMap({ abstractNums: [], nums: [] });
  const packageBytes = rezipContainer({
    '[Content_Types].xml': UTF8.encode(
      '<?xml version="1.0" encoding="UTF-8"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' +
        '</Types>'
    ),
    'word/document.xml': UTF8.encode(
      '<?xml version="1.0" encoding="UTF-8"?>' +
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:body><w:p/></w:body></w:document>'
    ),
    'word/_rels/document.xml.rels': UTF8.encode(
      '<?xml version="1.0" encoding="UTF-8"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>' +
        '</Relationships>'
    ),
    'word/numbering.xml': UTF8.encode(numberingXml),
  });
  const buffer = packageBytes.buffer.slice(
    packageBytes.byteOffset,
    packageBytes.byteOffset + packageBytes.byteLength
  ) as ArrayBuffer;
  const definitions = parseDocumentWithRust(buffer, {
    parseHeadersFooters: false,
    parseNotes: false,
    detectVariables: false,
  }).document.package.numbering ?? { abstractNums: [], nums: [] };
  return createNumberingMap(definitions);
}

const numberingMapCache = new WeakMap<NumberingDefinitions, NumberingMap>();

export function getCachedNumberingMap(definitions: NumberingDefinitions): NumberingMap {
  let map = numberingMapCache.get(definitions);
  if (!map) {
    map = createNumberingMap(definitions);
    numberingMapCache.set(definitions, map);
  }
  return map;
}

export function createNumberingMap(definitions: NumberingDefinitions): NumberingMap {
  // Build lookup maps for efficient access
  const abstractMap = new Map<number, AbstractNumbering>();
  for (const abs of definitions.abstractNums) {
    abstractMap.set(abs.abstractNumId, abs);
  }

  const numMap = new Map<number, NumberingInstance>();
  for (const num of definitions.nums) {
    numMap.set(num.numId, num);
  }

  return {
    definitions,

    getLevel(numId: number, ilvl: number): ListLevel | null {
      const num = numMap.get(numId);
      if (!num) return null;

      // Check for level override first
      if (num.levelOverrides) {
        const override = num.levelOverrides.find((o) => o.ilvl === ilvl);
        if (override) {
          if (override.lvl) {
            // Full level redefinition
            return override.lvl;
          }
          // Start override - need to get base level and modify
          const abstractNum = abstractMap.get(num.abstractNumId);
          if (abstractNum) {
            const baseLevel = abstractNum.levels.find((l) => l.ilvl === ilvl);
            if (baseLevel && override.startOverride !== undefined) {
              return {
                ...baseLevel,
                start: override.startOverride,
              };
            }
          }
        }
      }

      // Get from abstract numbering
      let abstractNum = abstractMap.get(num.abstractNumId);
      if (!abstractNum) return null;

      // Follow numStyleLink: when an abstractNum has numStyleLink instead of
      // defining levels directly, find the abstractNum that owns that style
      // (has matching styleLink) and use its levels. Per ECMA-376 §17.9.21/22.
      if (abstractNum.numStyleLink && abstractNum.levels.length === 0) {
        for (const candidate of abstractMap.values()) {
          if (candidate.styleLink === abstractNum.numStyleLink && candidate.levels.length > 0) {
            abstractNum = candidate;
            break;
          }
        }
      }

      return abstractNum.levels.find((l) => l.ilvl === ilvl) ?? null;
    },

    getAbstract(abstractNumId: number): AbstractNumbering | null {
      return abstractMap.get(abstractNumId) ?? null;
    },

    getInstance(numId: number): NumberingInstance | null {
      return numMap.get(numId) ?? null;
    },

    hasNumbering(numId: number): boolean {
      return numMap.has(numId);
    },
  };
}

/**
 * Resolve a paragraph's `numPr` against the numbering definitions into the
 * `ListRendering` the layout pipeline needs (marker template, per-level
 * numFmts, counter key, start override). Returns null when the numPr doesn't
 * name a real level — including `numId === 0`, "no numbering" per ECMA-376.
 *
 * Shared by the parser (document load) and `applyStyle` (style picker), so a
 * style-attached list renders identically in both paths.
 */
export function computeListRendering(
  numPr: { numId?: number; ilvl?: number },
  numbering: NumberingMap
): ListRendering | null {
  const { numId, ilvl = 0 } = numPr;
  if (numId === undefined || numId === 0) return null;

  const level = numbering.getLevel(numId, ilvl);
  if (!level) return null;

  // Collect numFmts for levels 0..ilvl so multi-level templates like
  // "%1.%2." can resolve each %N with its own format (e.g., upperRoman
  // parent + decimal child).
  const levelNumFmts: NumberFormat[] = [];
  for (let i = 0; i <= ilvl; i += 1) {
    const parent = numbering.getLevel(numId, i);
    levelNumFmts.push(parent?.numFmt ?? 'decimal');
  }

  const instance = numbering.getInstance(numId);
  const overrideForLevel = instance?.levelOverrides?.find((o) => o.ilvl === ilvl);

  return {
    level: ilvl,
    numId,
    marker: level.lvlText,
    isBullet: level.numFmt === 'bullet',
    numFmt: level.numFmt,
    markerHidden: level.rPr?.hidden || undefined,
    markerFontFamily: level.rPr?.fontFamily?.ascii || level.rPr?.fontFamily?.hAnsi || undefined,
    // w:sz is in half-points; convert to points for downstream use
    markerFontSize: level.rPr?.fontSize ? level.rPr.fontSize / 2 : undefined,
    markerSuffix: level.suffix,
    levelNumFmts,
    abstractNumId: instance?.abstractNumId,
    startOverride: overrideForLevel?.startOverride,
  };
}

/**
 * Format a number according to the specified format
 *
 * @param num - The number to format
 * @param format - The number format
 * @returns Formatted string
 */
export function formatNumber(num: number, format: NumberFormat): string {
  switch (format) {
    case 'decimal':
      return num.toString();

    case 'decimalZero':
      return padDecimal(num, 2);

    case 'decimalZero3':
      return padDecimal(num, 3);

    case 'decimalZero4':
      return padDecimal(num, 4);

    case 'decimalZero5':
      return padDecimal(num, 5);

    case 'upperRoman':
      return toRoman(num).toUpperCase();

    case 'lowerRoman':
      return toRoman(num).toLowerCase();

    case 'upperLetter':
      return toLetter(num).toUpperCase();

    case 'lowerLetter':
      return toLetter(num).toLowerCase();

    case 'ordinal':
      return toOrdinal(num);

    case 'bullet':
      return '•'; // Default bullet

    case 'none':
      return '';

    case 'decimalEnclosedParen':
      return `(${num})`;

    case 'numberInDash':
      return `-${num}-`;

    default:
      // For CJK and other special formats, fall back to decimal
      return num.toString();
  }
}

/** Zero-pad a counter to `width` digits ("decimalZero" family, §17.18.59). */
export function padDecimal(num: number, width: number): string {
  if (num < 0) return num.toString();
  return num.toString().padStart(width, '0');
}

/**
 * Convert number to Roman numerals
 */
function toRoman(num: number): string {
  if (num <= 0 || num > 3999) return num.toString();

  const romanNumerals: [number, string][] = [
    [1000, 'm'],
    [900, 'cm'],
    [500, 'd'],
    [400, 'cd'],
    [100, 'c'],
    [90, 'xc'],
    [50, 'l'],
    [40, 'xl'],
    [10, 'x'],
    [9, 'ix'],
    [5, 'v'],
    [4, 'iv'],
    [1, 'i'],
  ];

  let result = '';
  let remaining = num;

  for (const [value, numeral] of romanNumerals) {
    while (remaining >= value) {
      result += numeral;
      remaining -= value;
    }
  }

  return result;
}

/**
 * Convert number to letter (a, b, c, ... z, aa, ab, ...)
 */
function toLetter(num: number): string {
  if (num <= 0) return '';

  let result = '';
  let remaining = num;

  while (remaining > 0) {
    remaining--;
    result = String.fromCharCode(97 + (remaining % 26)) + result;
    remaining = Math.floor(remaining / 26);
  }

  return result;
}

/**
 * Convert number to ordinal (1st, 2nd, 3rd, ...)
 */
function toOrdinal(num: number): string {
  const suffix = ['th', 'st', 'nd', 'rd'];
  const v = num % 100;
  return num + (suffix[(v - 20) % 10] || suffix[v] || suffix[0]);
}

/**
 * Render list marker text by replacing placeholders with formatted numbers
 *
 * @param lvlText - The level text pattern (e.g., "%1.", "%1.%2")
 * @param counters - Array of counter values for each level (index 0 = level 0, etc.)
 * @param formats - Array of number formats for each level
 * @returns Rendered marker text
 */
export function renderListMarker(
  lvlText: string,
  counters: number[],
  formats: NumberFormat[]
): string {
  let result = lvlText;

  // Replace %1 through %9 with formatted counter values
  for (let i = 1; i <= 9; i++) {
    const placeholder = `%${i}`;
    if (result.includes(placeholder)) {
      const counterIndex = i - 1;
      const counter = counters[counterIndex] ?? 1;
      const format = formats[counterIndex] ?? 'decimal';
      const formatted = formatNumber(counter, format);
      result = result.replace(placeholder, formatted);
    }
  }

  return result;
}

/**
 * Get the bullet character for a bullet list level
 *
 * @param level - The list level definition
 * @returns The bullet character to display
 */
export function getBulletCharacter(level: ListLevel): string {
  // If lvlText is set and not empty, use it
  if (level.lvlText) {
    return level.lvlText;
  }

  // Check font for common bullet font mappings
  const fontFamily = level.rPr?.fontFamily?.ascii || level.rPr?.fontFamily?.hAnsi;

  if (fontFamily) {
    const fontLower = fontFamily.toLowerCase();

    // Symbol font common bullets
    if (fontLower === 'symbol') {
      return '•'; // Standard bullet
    }

    // Wingdings common bullets
    if (fontLower.includes('wingding')) {
      return '❑'; // Square bullet
    }
  }

  // Default bullet
  return '•';
}

/**
 * Check if a list level is a bullet (not numbered)
 */
export function isBulletLevel(level: ListLevel): boolean {
  return level.numFmt === 'bullet' || level.numFmt === 'none';
}
