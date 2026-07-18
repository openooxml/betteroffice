/**
 * Text, Paragraph, and Table Formatting Types
 *
 * Properties that control how text, paragraphs, and table structures
 * are formatted in OOXML (w:rPr, w:pPr, w:tblPr, etc.).
 */

import type { ColorValue, BorderSpec, ShadingProperties } from './colors';

// ============================================================================
// TEXT FORMATTING (Run Properties - rPr)
// ============================================================================

/**
 * Underline style options
 */
export type UnderlineStyle =
  | 'none'
  | 'single'
  | 'words'
  | 'double'
  | 'thick'
  | 'dotted'
  | 'dottedHeavy'
  | 'dash'
  | 'dashedHeavy'
  | 'dashLong'
  | 'dashLongHeavy'
  | 'dotDash'
  | 'dashDotHeavy'
  | 'dotDotDash'
  | 'dashDotDotHeavy'
  | 'wave'
  | 'wavyHeavy'
  | 'wavyDouble';

/**
 * Text effect animations
 */
export type TextEffect =
  | 'none'
  | 'blinkBackground'
  | 'lights'
  | 'antsBlack'
  | 'antsRed'
  | 'shimmer'
  | 'sparkle';

/**
 * Emphasis mark type
 */
export type EmphasisMark = 'none' | 'dot' | 'comma' | 'circle' | 'underDot';

/** Word font-binding hint from `w:rFonts/@w:hint`. Undefined means `default`. */
export type RunFontHint = 'default' | 'eastAsia' | 'cs';

/**
 * Modern (Office 2010+ `w14:*`) run text effects — glow, modern shadow,
 * reflection, text fill, and compound text outline. Distances/radii are in
 * CSS px (converted from EMU at parse); directions are degrees clockwise;
 * opacities are 0..1 fractions; colors are resolved `#rrggbb` strings.
 * All members optional and additive so documents without w14 effects
 * round-trip byte-identically.
 */
export interface TextModernEffects {
  /** `w14:glow` — colored halo around the glyph outline. */
  glow?: { color?: string; radius?: number };
  /** `w14:shadow` — parameterized modern shadow. */
  shadow?: { color?: string; blurRadius?: number; distance?: number; direction?: number };
  /** `w14:reflection` — mirrored copy below the text. */
  reflection?: {
    blurRadius?: number;
    startOpacity?: number;
    endOpacity?: number;
    distance?: number;
    direction?: number;
  };
  /** `w14:textFill` — solid or gradient glyph fill replacing the run color. */
  textFill?: {
    kind?: 'none' | 'solid' | 'gradient';
    color?: string;
    angle?: number;
    stops?: Array<{ position?: number; color?: string }>;
  };
  /** `w14:textOutline` — stroked glyph outline. Width in px. */
  textOutline?: { color?: string; width?: number; dash?: string; noFill?: boolean };
}

/**
 * The three independent language slots in `w:lang`. Missing slots inherit
 * through the style/default chain and therefore have no effect.
 */
export interface RunLanguage {
  /** Latin/default language (`w:lang/@w:val`). */
  latin?: string;
  /** East Asian language (`w:lang/@w:eastAsia`). */
  eastAsia?: string;
  /** Complex-script language (`w:lang/@w:bidi`). */
  bidi?: string;
}

/**
 * Character-level formatting (`w:rPr`) — the full set of run properties
 * Word supports: weight, slant, font, size, color, highlight, underline,
 * strikethrough, vertical position, language, complex-script variants,
 * spacing/kerning, emphasis marks, and more.
 *
 * Most fields mirror their ECMA-376 element names (see §17.3.2). Missing
 * keys inherit from the run's paragraph style → linked style → document
 * defaults chain.
 */
export interface TextFormatting {
  // Basic formatting
  /** Bold (w:b) */
  bold?: boolean;
  /** Bold complex script (w:bCs) */
  boldCs?: boolean;
  /** Italic (w:i) */
  italic?: boolean;
  /** Italic complex script (w:iCs) */
  italicCs?: boolean;

  // Underline & strikethrough
  /** Underline style and color (w:u) */
  underline?: {
    style: UnderlineStyle;
    color?: ColorValue;
  };
  /** Strikethrough (w:strike) */
  strike?: boolean;
  /** Double strikethrough (w:dstrike) */
  doubleStrike?: boolean;

  // Vertical alignment
  /** Superscript/subscript (w:vertAlign) */
  vertAlign?: 'baseline' | 'superscript' | 'subscript';

  // Capitalization
  /** Small caps (w:smallCaps) */
  smallCaps?: boolean;
  /** All caps (w:caps) */
  allCaps?: boolean;

  // Visibility
  /** Hidden text (w:vanish) */
  hidden?: boolean;

  // Colors and highlighting
  /** Text color (w:color) */
  color?: ColorValue;
  /** Highlight/background color (w:highlight) */
  highlight?:
    | 'black'
    | 'blue'
    | 'cyan'
    | 'darkBlue'
    | 'darkCyan'
    | 'darkGray'
    | 'darkGreen'
    | 'darkMagenta'
    | 'darkRed'
    | 'darkYellow'
    | 'green'
    | 'lightGray'
    | 'magenta'
    | 'none'
    | 'red'
    | 'white'
    | 'yellow';
  /** Character shading (w:shd) */
  shading?: ShadingProperties;

  // Font properties
  /** Font size in half-points (w:sz) - e.g., 24 = 12pt */
  fontSize?: number;
  /** Font size complex script (w:szCs) */
  fontSizeCs?: number;
  /** Font family (w:rFonts) */
  fontFamily?: {
    ascii?: string;
    hAnsi?: string;
    eastAsia?: string;
    cs?: string;
    /** Theme font reference */
    asciiTheme?:
      | 'majorAscii'
      | 'majorHAnsi'
      | 'majorEastAsia'
      | 'majorBidi'
      | 'minorAscii'
      | 'minorHAnsi'
      | 'minorEastAsia'
      | 'minorBidi';
    hAnsiTheme?: string;
    eastAsiaTheme?: string;
    csTheme?: string;
    /** Script-binding hint (`w:rFonts/@w:hint`). Undefined = `default`. */
    hint?: RunFontHint;
  };

  /** Three-slot run language (`w:lang`). Undefined = inherit/no override. */
  language?: RunLanguage;

  // Spacing and position
  /** Character spacing in twips (w:spacing) */
  spacing?: number;
  /** Raised/lowered text position in half-points (w:position) */
  position?: number;
  /** Horizontal text scale percentage (w:w) */
  scale?: number;
  /** Kerning threshold in half-points (w:kern) */
  kerning?: number;

  // Effects
  /** Text effect animation (w:effect) */
  effect?: TextEffect;
  /** Emphasis mark (w:em) */
  emphasisMark?: EmphasisMark;
  /** Emboss effect (w:emboss) */
  emboss?: boolean;
  /** Imprint/engrave effect (w:imprint) */
  imprint?: boolean;
  /** Outline effect (w:outline) */
  outline?: boolean;
  /** Shadow effect (w:shadow) */
  shadow?: boolean;
  /** Modern w14 text effects (glow/shadow/reflection/textFill/textOutline). */
  modernEffects?: TextModernEffects;

  // Complex script
  /** Right-to-left text (w:rtl) */
  rtl?: boolean;
  /** Complex script formatting (w:cs) */
  cs?: boolean;

  // Style reference
  /** Character style ID (w:rStyle) */
  styleId?: string;
}

// ============================================================================
// PARAGRAPH FORMATTING (Paragraph Properties - pPr)
// ============================================================================

/**
 * Tab stop alignment
 */
export type TabStopAlignment = 'left' | 'center' | 'right' | 'decimal' | 'bar' | 'clear' | 'num';

/**
 * Tab leader character
 */
export type TabLeader = 'none' | 'dot' | 'hyphen' | 'underscore' | 'heavy' | 'middleDot';

/**
 * Tab stop definition
 */
export interface TabStop {
  /** Position in twips from left margin */
  position: number;
  /** Alignment at tab stop */
  alignment: TabStopAlignment;
  /** Leader character */
  leader?: TabLeader;
}

/**
 * Line spacing rule
 */
export type LineSpacingRule = 'auto' | 'exact' | 'atLeast';

/**
 * Paragraph alignment/justification
 */
export type ParagraphAlignment =
  | 'left'
  | 'center'
  | 'right'
  | 'both'
  | 'distribute'
  | 'mediumKashida'
  | 'highKashida'
  | 'lowKashida'
  | 'thaiDistribute';

/**
 * Per-side flags identifying which `<w:spacing>` attrs were inline (not
 * inherited from a style chain). Used to suppress style-only spacing on
 * empty paragraphs per Word's behavior.
 */
export type SpacingExplicit = { before?: boolean; after?: boolean };

/**
 * Paragraph-level formatting (`w:pPr`) — alignment, indentation, spacing
 * (before/after, line height), pagination flags (keepNext, keepLines,
 * pageBreakBefore, widowControl), tabs, borders, shading, numbering
 * reference, style reference, and frame/anchored-text properties.
 *
 * Most fields mirror their ECMA-376 element names (see §17.3.1).
 * Inheritance: direct formatting beats the linked style which beats
 * document defaults.
 */
export interface ParagraphFormatting {
  // Alignment
  /** Paragraph alignment (w:jc) */
  alignment?: ParagraphAlignment;
  /** Text direction (w:bidi) */
  bidi?: boolean;

  // Spacing
  /** Spacing before in twips (w:spacing/@w:before) */
  spaceBefore?: number;
  /** Spacing after in twips (w:spacing/@w:after) */
  spaceAfter?: number;
  /** Line spacing value (w:spacing/@w:line) */
  lineSpacing?: number;
  /** Line spacing rule (w:spacing/@w:lineRule) */
  lineSpacingRule?: LineSpacingRule;
  /** Auto space before (w:spacing/@w:beforeAutospacing) */
  beforeAutospacing?: boolean;
  /** Auto space after (w:spacing/@w:afterAutospacing) */
  afterAutospacing?: boolean;
  /**
   * Per-side flags marking which `<w:spacing>` attrs came from this
   * paragraph's own pPr (vs inherited). Word collapses style-inherited
   * spacing on empty paragraphs but honors the explicit values.
   */
  spacingExplicit?: SpacingExplicit;

  // Indentation
  /** Left indent in twips (w:ind/@w:left) */
  indentLeft?: number;
  /** Right indent in twips (w:ind/@w:right) */
  indentRight?: number;
  /** First line indent in twips - positive for indent, negative for hanging (w:ind/@w:firstLine or @w:hanging) */
  indentFirstLine?: number;
  /** Whether first line is hanging indent */
  hangingIndent?: boolean;

  // Borders
  /** Paragraph borders (w:pBdr) */
  borders?: {
    top?: BorderSpec;
    bottom?: BorderSpec;
    left?: BorderSpec;
    right?: BorderSpec;
    between?: BorderSpec;
    bar?: BorderSpec;
  };

  // Background
  /** Paragraph shading (w:shd) */
  shading?: ShadingProperties;

  // Tab stops
  /** Custom tab stops (w:tabs) */
  tabs?: TabStop[];

  // Page break control
  /** Keep with next paragraph (w:keepNext) */
  keepNext?: boolean;
  /** Keep lines together (w:keepLines) */
  keepLines?: boolean;
  /** Widow/orphan control (w:widowControl) */
  widowControl?: boolean;
  /** Page break before (w:pageBreakBefore) */
  pageBreakBefore?: boolean;
  /** Contextual spacing — suppress space between paragraphs of the same style (w:contextualSpacing) */
  contextualSpacing?: boolean;

  // Numbering/List
  /** Numbering properties (w:numPr) */
  numPr?: {
    /** Numbering definition ID (w:numId) */
    numId?: number;
    /** List level (0-8) (w:ilvl) */
    ilvl?: number;
  };
  /**
   * When `numPr` was resolved from the paragraph STYLE's pPr rather than the
   * paragraph's own `<w:numPr>`, this records the style-sourced value. The
   * serializer omits `numPr` while it still equals this value — writing it as
   * direct formatting would flip Word's indent precedence (a directly
   * referenced level's indents beat the style's; a style-referenced level's
   * do not) and break the document on save/reload. Cleared the moment the
   * user changes the numbering (values diverge).
   */
  numPrFromStyle?: {
    numId?: number;
    ilvl?: number;
  };

  // Outline level (for TOC)
  /** Outline level 0-9 (w:outlineLvl) */
  outlineLevel?: number;

  // Style reference
  /** Paragraph style ID (w:pStyle) */
  styleId?: string;

  // Frame properties
  /** Text frame properties (w:framePr) */
  frame?: {
    width?: number;
    height?: number;
    hAnchor?: 'text' | 'margin' | 'page';
    vAnchor?: 'text' | 'margin' | 'page';
    x?: number;
    y?: number;
    xAlign?: 'left' | 'center' | 'right' | 'inside' | 'outside';
    yAlign?: 'top' | 'center' | 'bottom' | 'inside' | 'outside' | 'inline';
    wrap?: 'around' | 'auto' | 'none' | 'notBeside' | 'through' | 'tight';
  };

  // Suppress
  /** Suppress line numbers (w:suppressLineNumbers) */
  suppressLineNumbers?: boolean;
  /** Suppress auto hyphens (w:suppressAutoHyphens) */
  suppressAutoHyphens?: boolean;

  // Default run properties for this paragraph
  /** Run properties to apply to all runs (w:rPr) */
  runProperties?: TextFormatting;
}

// ============================================================================
// TABLE FORMATTING (w:tblPr, w:trPr, w:tcPr)
// ============================================================================

/**
 * Table width type
 */
export type TableWidthType = 'auto' | 'dxa' | 'nil' | 'pct';

/**
 * Table measurement (width or height)
 */
export interface TableMeasurement {
  /** Value in twips (for dxa) or fifths of a percent (for pct) */
  value: number;
  /** Measurement type */
  type: TableWidthType;
}

/**
 * Table borders
 */
export interface TableBorders {
  top?: BorderSpec;
  bottom?: BorderSpec;
  left?: BorderSpec;
  right?: BorderSpec;
  insideH?: BorderSpec;
  insideV?: BorderSpec;
  /** Logical leading edge (`w:start`); undefined = no authored edge. */
  start?: BorderSpec;
  /** Logical trailing edge (`w:end`); undefined = no authored edge. */
  end?: BorderSpec;
  /** Top-left to bottom-right diagonal (`w:tl2br`). */
  tl2br?: BorderSpec;
  /** Top-right to bottom-left diagonal (`w:tr2bl`). */
  tr2bl?: BorderSpec;
}

/**
 * Cell margins
 */
export interface CellMargins {
  top?: TableMeasurement;
  bottom?: TableMeasurement;
  left?: TableMeasurement;
  right?: TableMeasurement;
  /** Logical leading margin (`w:start`); undefined falls back to physical left/right. */
  start?: TableMeasurement;
  /** Logical trailing margin (`w:end`); undefined falls back to physical left/right. */
  end?: TableMeasurement;
}

/**
 * Table look flags (for table styles)
 */
export interface TableLook {
  firstColumn?: boolean;
  firstRow?: boolean;
  lastColumn?: boolean;
  lastRow?: boolean;
  noHBand?: boolean;
  noVBand?: boolean;
  /** Original hexadecimal `w:tblLook/@w:val`; undefined = derive from flags. */
  value?: string;
}

/** One conditional region selected by `w:tblStylePr`/`w:cnfStyle`. */
export type TableStyleRegion =
  | 'band1Horz'
  | 'band1Vert'
  | 'band2Horz'
  | 'band2Vert'
  | 'firstCol'
  | 'firstRow'
  | 'lastCol'
  | 'lastRow'
  | 'neCell'
  | 'nwCell'
  | 'seCell'
  | 'swCell';

/**
 * Lossless style-cascade provenance carried after resolution. Every member is
 * optional; an absent object means legacy direct-formatting behavior.
 */
export interface TableStyleCascade {
  /** Explicitly selected `w:tblStyle`, if any. */
  selectedStyleId?: string;
  /** Type-default table style used when no explicit style exists. */
  defaultStyleId?: string;
  /** `basedOn` chain, base first. Undefined = not resolved yet. */
  basedOnStyleIds?: string[];
  /** Exact conditional regions applied, in cascade order. */
  appliedRegions?: TableStyleRegion[];
}

/**
 * Floating table properties
 */
export interface FloatingTableProperties {
  /** Horizontal anchor */
  horzAnchor?: 'margin' | 'page' | 'text';
  /** Vertical anchor */
  vertAnchor?: 'margin' | 'page' | 'text';
  /** Horizontal position */
  tblpX?: number;
  tblpXSpec?: 'left' | 'center' | 'right' | 'inside' | 'outside';
  /** Vertical position */
  tblpY?: number;
  tblpYSpec?: 'top' | 'center' | 'bottom' | 'inside' | 'outside' | 'inline';
  /** Distance from surrounding text */
  topFromText?: number;
  bottomFromText?: number;
  leftFromText?: number;
  rightFromText?: number;
}

/**
 * Table formatting properties (w:tblPr)
 */
export interface TableFormatting {
  /** Table width */
  width?: TableMeasurement;
  /** Table justification */
  justification?: 'left' | 'center' | 'right';
  /** Cell spacing */
  cellSpacing?: TableMeasurement;
  /** Table indent from left margin */
  indent?: TableMeasurement;
  /** Table borders */
  borders?: TableBorders;
  /** Default cell margins */
  cellMargins?: CellMargins;
  /** Table layout */
  layout?: 'fixed' | 'autofit';
  /** Table style ID */
  styleId?: string;
  /** Table look (conditional formatting flags) */
  look?: TableLook;
  /** Shading/background */
  shading?: ShadingProperties;
  /** Overlap for floating tables */
  overlap?: 'never' | 'overlap';
  /** Floating table properties */
  floating?: FloatingTableProperties;
  /** Right to left table */
  bidi?: boolean;
  /** Horizontal band size (`w:tblStyleRowBandSize`). Undefined = 1. */
  styleRowBandSize?: number;
  /** Vertical band size (`w:tblStyleColBandSize`). Undefined = 1. */
  styleColBandSize?: number;
  /** Optional resolved style provenance. Undefined = legacy cascade. */
  styleCascade?: TableStyleCascade;
}

/**
 * Table row formatting properties (w:trPr)
 */
export interface TableRowFormatting {
  /** Row height */
  height?: TableMeasurement;
  /** Height rule */
  heightRule?: 'auto' | 'atLeast' | 'exact';
  /** Header row (repeats on each page) */
  header?: boolean;
  /** Allow row to break across pages */
  cantSplit?: boolean;
  /** Row justification */
  justification?: 'left' | 'center' | 'right';
  /** Hidden row */
  hidden?: boolean;
  /** Conditional format style */
  conditionalFormat?: ConditionalFormatStyle;
  /** Leading omitted grid columns (`w:gridBefore`). Undefined = 0. */
  gridBefore?: number;
  /** Trailing omitted grid columns (`w:gridAfter`). Undefined = 0. */
  gridAfter?: number;
  /** Preferred width before cells (`w:wBefore`). */
  widthBefore?: TableMeasurement;
  /** Preferred width after cells (`w:wAfter`). */
  widthAfter?: TableMeasurement;
}

/**
 * Conditional format style
 */
export interface ConditionalFormatStyle {
  /** First row */
  firstRow?: boolean;
  /** Last row */
  lastRow?: boolean;
  /** First column */
  firstColumn?: boolean;
  /** Last column */
  lastColumn?: boolean;
  /** Odd horizontal band */
  oddHBand?: boolean;
  /** Even horizontal band */
  evenHBand?: boolean;
  /** Odd vertical band */
  oddVBand?: boolean;
  /** Even vertical band */
  evenVBand?: boolean;
  /** Northwest corner */
  nwCell?: boolean;
  /** Northeast corner */
  neCell?: boolean;
  /** Southwest corner */
  swCell?: boolean;
  /** Southeast corner */
  seCell?: boolean;
}

/**
 * Table cell formatting properties (w:tcPr)
 */
export interface TableCellFormatting {
  /** Cell width */
  width?: TableMeasurement;
  /** Cell borders */
  borders?: TableBorders;
  /** Cell margins (override table default) */
  margins?: CellMargins;
  /** Cell shading/background */
  shading?: ShadingProperties;
  /** Vertical alignment */
  verticalAlign?: 'top' | 'center' | 'bottom';
  /** Text direction */
  textDirection?: 'lr' | 'lrV' | 'rl' | 'rlV' | 'tb' | 'tbV' | 'tbRl' | 'tbRlV' | 'btLr';
  /** Grid span (horizontal merge) */
  gridSpan?: number;
  /** Vertical merge */
  vMerge?: 'restart' | 'continue';
  /** Fit text to cell width */
  fitText?: boolean;
  /** Wrap text */
  noWrap?: boolean;
  /** Hide cell marker */
  hideMark?: boolean;
  /** Conditional format style */
  conditionalFormat?: ConditionalFormatStyle;
}
