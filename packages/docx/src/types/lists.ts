/**
 * Lists & Numbering Types
 *
 * Types for bullet lists, numbered lists, and numbering definitions.
 */

import type { TextFormatting, ParagraphFormatting } from './formatting';

// ============================================================================
// LISTS & NUMBERING
// ============================================================================

/**
 * Number format type
 */
export type NumberFormat =
  | 'decimal'
  | 'upperRoman'
  | 'lowerRoman'
  | 'upperLetter'
  | 'lowerLetter'
  | 'ordinal'
  | 'cardinalText'
  | 'ordinalText'
  | 'hex'
  | 'chicago'
  | 'ideographDigital'
  | 'japaneseCounting'
  | 'aiueo'
  | 'iroha'
  | 'decimalFullWidth'
  | 'decimalHalfWidth'
  | 'japaneseLegal'
  | 'japaneseDigitalTenThousand'
  | 'decimalEnclosedCircle'
  | 'decimalFullWidth2'
  | 'aiueoFullWidth'
  | 'irohaFullWidth'
  | 'decimalZero'
  // Synthetic in-memory formats for Word's `w:numFmt w:val="custom"` with an
  // XSLT-style zero-padded format string ("001, 002, ...", "0001, ...",
  // "00001, ..."). Not OOXML enum values — never serialized (numbering.xml is
  // preserved as-is on save); they exist so the render pipeline can carry the
  // pad width through the existing NumberFormat plumbing.
  | 'decimalZero3'
  | 'decimalZero4'
  | 'decimalZero5'
  | 'bullet'
  | 'ganada'
  | 'chosung'
  | 'decimalEnclosedFullstop'
  | 'decimalEnclosedParen'
  | 'decimalEnclosedCircleChinese'
  | 'ideographEnclosedCircle'
  | 'ideographTraditional'
  | 'ideographZodiac'
  | 'ideographZodiacTraditional'
  | 'taiwaneseCounting'
  | 'ideographLegalTraditional'
  | 'taiwaneseCountingThousand'
  | 'taiwaneseDigital'
  | 'chineseCounting'
  | 'chineseLegalSimplified'
  | 'chineseCountingThousand'
  | 'koreanDigital'
  | 'koreanCounting'
  | 'koreanLegal'
  | 'koreanDigital2'
  | 'vietnameseCounting'
  | 'russianLower'
  | 'russianUpper'
  | 'none'
  | 'numberInDash'
  | 'hebrew1'
  | 'hebrew2'
  | 'arabicAlpha'
  | 'arabicAbjad'
  | 'hindiVowels'
  | 'hindiConsonants'
  | 'hindiNumbers'
  | 'hindiCounting'
  | 'thaiLetters'
  | 'thaiNumbers'
  | 'thaiCounting';

/**
 * Multi-level suffix (what follows the number)
 */
export type LevelSuffix = 'tab' | 'space' | 'nothing';

/**
 * One indentation level of an abstract numbering definition (`w:lvl`).
 * Carries the number format, the marker template (`lvlText` — e.g.
 * `"%1.%2."`), the level's paragraph properties (indent, hanging) and
 * character properties (font, size, color for the marker itself).
 *
 * `ilvl` ranges 0-8 in standard Word documents.
 */
export interface ListLevel {
  /** Level index (0-8) */
  ilvl: number;
  /** Starting number */
  start?: number;
  /** Number format */
  numFmt: NumberFormat;
  /** Level text (e.g., "%1." or "•") */
  lvlText: string;
  /** Justification */
  lvlJc?: 'left' | 'center' | 'right';
  /** Suffix after number */
  suffix?: LevelSuffix;
  /** Paragraph properties for this level */
  pPr?: ParagraphFormatting;
  /** Run properties for the number/bullet */
  rPr?: TextFormatting;
  /** Restart numbering from higher level */
  lvlRestart?: number;
  /** Is legal numbering style */
  isLgl?: boolean;
  /** Legacy settings */
  legacy?: {
    legacy?: boolean;
    legacySpace?: number;
    legacyIndent?: number;
  };
}

/**
 * Abstract numbering definition (`w:abstractNum`) — the reusable template
 * for a list: which `NumberFormat` at each indentation level, what
 * marker text, what paragraph/character formatting. Multiple
 * `NumberingInstance`s (`w:num`) can reference one abstract numbering
 * to share the template while keeping independent counters.
 *
 * See ECMA-376 §17.9.
 */
export interface AbstractNumbering {
  /** Abstract numbering ID */
  abstractNumId: number;
  /** Multi-level type */
  multiLevelType?: 'hybridMultilevel' | 'multilevel' | 'singleLevel';
  /** Numbering style link */
  numStyleLink?: string;
  /** Style link */
  styleLink?: string;
  /** Level definitions */
  levels: ListLevel[];
  /** Name */
  name?: string;
}

/**
 * Numbering instance (w:num)
 */
export interface NumberingInstance {
  /** Numbering ID (referenced by paragraphs) */
  numId: number;
  /** Reference to abstract numbering */
  abstractNumId: number;
  /** Level overrides */
  levelOverrides?: Array<{
    ilvl: number;
    startOverride?: number;
    lvl?: ListLevel;
  }>;
}

/**
 * Computed list marker for one paragraph — what the layout engine and
 * painter need to render the "1.", "a)", "•" prefix. Not part of the
 * wire format; the parser fills this from the `numbering.xml` chain plus
 * the paragraph's `numPr`. Paragraphs without list rendering omit it.
 */
export interface ListRendering {
  /** Computed marker text (e.g., "1.", "a)", "•") */
  marker: string;
  /** List level (0-8) */
  level: number;
  /** Numbering ID */
  numId: number;
  /** Whether this is a bullet or numbered list */
  isBullet: boolean;
  /** Number format type (decimal, lowerRoman, upperRoman, etc.) */
  numFmt?: NumberFormat;
  /** Whether the list marker is hidden (w:vanish on level rPr) */
  markerHidden?: boolean;
  /** Marker font family from numbering level rPr (ascii name) */
  markerFontFamily?: string;
  /** Marker font size from numbering level rPr, in points */
  markerFontSize?: number;
  /**
   * Suffix character placed after the marker before body text (§17.9.25).
   * Default is `tab`; `space` inserts a single space; `nothing` no gap.
   * Drives marker-slot sizing in `getListMarkerInlineWidth`.
   */
  markerSuffix?: LevelSuffix;
  /**
   * NumberFormat for each level from 0..ilvl (inclusive).
   * Used to resolve multi-level templates like "%1.%2." where each %N
   * may need a different format (e.g., upperRoman parent + decimal child).
   */
  levelNumFmts?: NumberFormat[];
  /** abstractNumId the paragraph's numId points to (counters key on this). */
  abstractNumId?: number;
  /**
   * Start value from the numId's lvlOverride for the paragraph's ilvl, if any.
   * Per ECMA-376 §17.9.18, this resets the shared abstractNum counter the
   * first time the numId appears.
   */
  startOverride?: number;
}

/**
 * Top-level numbering data from `numbering.xml` — the set of abstract
 * templates and the per-document `NumberingInstance`s that reference
 * them. Paragraphs reference a `numId` (instance), not an
 * `abstractNumId` directly.
 */
export interface NumberingDefinitions {
  /** Abstract numbering definitions */
  abstractNums: AbstractNumbering[];
  /** Numbering instances */
  nums: NumberingInstance[];
}
