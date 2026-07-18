/**
 * Type vocabulary of the pagination pipeline: FlowBlocks go in, per-block
 * extents come back from measurement, and the paginator emits fragments
 * placed on pages.
 * @packageDocumentation
 * @public
 */

import type { PresetGeometryPathCommand } from '@betteroffice/drawingml';
import type { Chart } from '../../types/document';
import { isWrapNone, wrapsAroundText } from '../../docx/wrapTypes';
import { emuToPixels } from '../../utils/units';

/**
 * Inline SDT widget metadata carried by a painted text run.
 */
export interface InlineSdtWidget {
  kind: 'checkbox';
  /** Stable per-document id derived from the ProseMirror node position. */
  groupId: string;
  /** ProseMirror position of the inline SDT node. */
  pos: number;
  /** Word tag value (`w:tag`). */
  tag?: string;
  /** Word alias value (`w:alias`). */
  alias?: string;
  /** Live checkbox glyph state. */
  checked?: boolean;
}

/**
 * One run's visible slice on a laid-out line: the run to render/measure plus
 * its on-line text. For a boundary text run this is a shallow copy sliced to
 * the line's head/tail char (PM positions shifted to match); tabs, images,
 * line breaks, and fields pass through whole with an empty `text`. `text` is
 * the run's textual contribution to the line (empty for non-text runs), so
 * joining a line's segment texts reconstructs the line's visible characters.
 * Materialized during layout by the Rust engine.
 * @public
 */
export type ResolvedSegment = {
  run: Run;
  text: string;
};

/**
 * A laid-out line's resolved run segments (see {@link ResolvedSegment}). One
 * is carried per line on `ParagraphFragment.resolvedLines`.
 * @public
 */
export type ResolvedLine = {
  segments: ResolvedSegment[];
};

/**
 * Key that identifies a flow block across measure/paginate/paint passes —
 * usually the block's index, sometimes an `${index}-${type}` compound.
 */
export type BlockId = string | number;

/** Script-specific font slots carried until shaping chooses a face. */
export type RunFontSlots = {
  ascii?: string;
  hAnsi?: string;
  eastAsia?: string;
  cs?: string;
  asciiTheme?: string;
  hAnsiTheme?: string;
  eastAsiaTheme?: string;
  csTheme?: string;
  hint?: 'default' | 'eastAsia' | 'cs';
};

/** Three-slot `w:lang` projection; missing slots inherit/no-op. */
export type RunLanguageSlots = {
  latin?: string;
  eastAsia?: string;
  bidi?: string;
};

// =============================================================================
// flow blocks — what the paginator consumes
// =============================================================================

/**
 * Character-level formatting shared by every run flavor (mirrors w:rPr).
 */
export type RunFormatting = {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean | { style?: string; color?: string };
  strike?: boolean;
  color?: string;
  highlight?: string;
  fontFamily?: string;
  /** All authored font slots. Undefined = use legacy `fontFamily`. */
  fontSlots?: RunFontSlots;
  fontSize?: number;
  /** Complex-script font size in px. Undefined = use `fontSize`. */
  fontSizeCs?: number;
  /** Complex-script bold override. Undefined = use `bold`. */
  boldCs?: boolean;
  /** Complex-script italic override. Undefined = use `italic`. */
  italicCs?: boolean;
  /** Force complex-script formatting (`w:cs`). Undefined = false. */
  complexScript?: boolean;
  /** Run languages. Undefined = no language override. */
  language?: RunLanguageSlots;
  letterSpacing?: number;
  superscript?: boolean;
  subscript?: boolean;
  /** Render glyphs as uppercase regardless of source case (OOXML w:caps). */
  allCaps?: boolean;
  /** Render lowercase glyphs as small uppercase (OOXML w:smallCaps). */
  smallCaps?: boolean;
  /**
   * Vertical baseline shift in CSS pixels (positive = up). OOXML w:position is
   * authored in half-points; converted to px during the bridge so the painter
   * can apply it directly via vertical-align without re-doing the math.
   */
  positionPx?: number;
  /**
   * Horizontal text scale as a percentage (100 = normal, 50 = half-width,
   * 200 = double-width). OOXML w:w specifies pct in raw % (e.g. 90 means 90%).
   */
  horizontalScale?: number;
  /**
   * Minimum font size in points at which kerning kicks in (OOXML w:kern, half-
   * points in source). When this run's effective font size is at or above this
   * threshold, the painter enables font-kerning: normal.
   */
  kerningMinPt?: number;
  /** Engraved/imprint effect (OOXML w:imprint, §17.3.2.18). */
  imprint?: boolean;
  /** Embossed/raised effect (OOXML w:emboss, §17.3.2.13). */
  emboss?: boolean;
  /** Drop-shadow effect (OOXML w:shadow, §17.3.2.31). */
  textShadow?: boolean;
  /** Outlined / hollow text (OOXML w:outline, §17.3.2.23). */
  textOutline?: boolean;
  /**
   * CJK emphasis mark (OOXML w:em, §17.3.2.12). Maps to CSS `text-emphasis`
   * with a position and style; the painter handles the variant lookup.
   */
  emphasisMark?: 'dot' | 'comma' | 'circle' | 'underDot';
  /** Hidden run (OOXML w:vanish, §17.3.2.41). Painter skips the run. */
  hidden?: boolean;
  /**
   * Per-run right-to-left direction (OOXML w:rtl, §17.3.2.30). Independent
   * from the paragraph's bidi flag — a single run may flip direction within
   * an LTR paragraph.
   */
  rtl?: boolean;
  /**
   * Legacy text-effect animation (OOXML w:effect, §17.3.2.11). The painter
   * surfaces it as a class hook so host CSS can opt in to animations.
   */
  textEffect?: 'blinkBackground' | 'lights' | 'antsBlack' | 'antsRed' | 'shimmer' | 'sparkle';
  /**
   * Modern w14 text effects (glow/shadow/reflection/textFill/textOutline),
   * threaded losslessly to the display-list contract. Undefined = none.
   */
  modernEffects?: import('../../types/formatting').TextModernEffects;
  /** Hyperlink info if this run is a link */
  hyperlink?: HyperlinkInfo;
  /** Footnote reference ID (if this run contains a footnote reference) */
  footnoteRefId?: number;
  /** Endnote reference ID (if this run contains an endnote reference) */
  endnoteRefId?: number;
  /** Comment IDs if this run is within a comment range */
  commentIds?: number[];
  /** Whether this run is a tracked insertion */
  isInsertion?: boolean;
  /** Whether this run is a tracked deletion */
  isDeletion?: boolean;
  /** Author of the tracked change */
  changeAuthor?: string;
  /** Date of the tracked change */
  changeDate?: string;
  /** Revision ID of the tracked change (for sidebar matching) */
  changeRevisionId?: number;
  /** Logical source-order index, independent of bidi visual paint order. */
  logicalOrder?: number;
  /** Resolved Unicode bidi embedding level. Undefined = resolve downstream. */
  bidiLevel?: number;
};

/**
 * Link target attached to a run.
 */
export type HyperlinkInfo = {
  href: string;
  tooltip?: string;
  /** Safe target frame. Undefined = host default. */
  target?: string;
  /** Authored visited-history bit. Undefined = false/unmanaged. */
  history?: boolean;
  /** Location in the target document. */
  docLocation?: string;
  /**
   * Skip the painter's "Word default" blue + underline fallback when the run
   * has no resolved color/underline. Set by the layout for hyperlinks
   * inside TOC paragraphs, where Word renders entries in the TOCx paragraph
   * color rather than the Hyperlink character style's color.
   */
  noDefaultStyle?: boolean;
};

/**
 * Run carrying literal text.
 */
export type TextRun = RunFormatting & {
  kind: 'text';
  text: string;
  /** Hyperlink information if this run is a link. */
  hyperlink?: HyperlinkInfo;
  /** PM doc offset where the run's text begins. */
  pmStart?: number;
  /** PM doc offset one past the run's final character. */
  pmEnd?: number;
  /** Inline content-control widget metadata when this run is the visible glyph. */
  inlineSdtWidget?: InlineSdtWidget;
};

/** Run for one tab character; measurement resolves its width from the tab grid. */
export type TabRun = RunFormatting & {
  kind: 'tab';
  pmStart?: number;
  pmEnd?: number;
  /** resolved advance in pixels, filled in by measurement */
  width?: number;
  /** Shaped leader glyph metadata. Undefined = no/synthetic legacy leader. */
  leaderGlyphs?: {
    glyph?: string;
    count?: number;
    advance?: number;
    font?: string;
    fontSize?: number;
    color?: string;
  };
};

/**
 * Anchor placement for a floating image: each axis pairs a reference frame
 * with either a fixed offset or a named alignment.
 */
export type ImageRunPosition = {
  useSimplePos?: boolean;
  simplePos?: { x?: number; y?: number };
  relativeHeight?: number;
  behindDoc?: boolean;
  horizontal?: {
    align?: string;
    posOffset?: number;
    relativeTo?: string;
  };
  vertical?: {
    align?: string;
    posOffset?: number;
    relativeTo?: string;
  };
};

export type WrapTextDirection = 'bothSides' | 'left' | 'right' | 'largest';

/**
 * Run holding a picture, inline or anchored.
 */
export type ImageRun = {
  kind: 'image';
  src: string;
  width: number;
  height: number;
  alt?: string;
  /** CSS transform string (rotation, flip) */
  transform?: string;
  /** Position for floating/anchored images */
  position?: ImageRunPosition;
  /** Wrap type from DOCX (inline, square, tight, through, topAndBottom, etc.) */
  wrapType?: string;
  /** Display mode for CSS rendering */
  displayMode?: 'inline' | 'block' | 'float';
  /** CSS float direction */
  cssFloat?: 'left' | 'right' | 'none';
  /** Wrap distances in pixels */
  distTop?: number;
  distBottom?: number;
  distLeft?: number;
  distRight?: number;
  /** wp:srcRect crop fractions in [0, 1]; emit as CSS clip-path inset. */
  cropTop?: number;
  cropRight?: number;
  cropBottom?: number;
  cropLeft?: number;
  /** a:alphaModFix → CSS opacity in [0, 1]. */
  opacity?: number;
  /** Decomposed rotation; undefined = parse legacy `transform`. */
  rotationDeg?: number;
  /** Decomposed horizontal flip; undefined = false. */
  flipH?: boolean;
  /** Decomposed vertical flip; undefined = false. */
  flipV?: boolean;
  /** Post-transform footprint used for flow. Undefined = width/height. */
  rotationBounds?: { width?: number; height?: number; offsetX?: number; offsetY?: number };
  /** Square/tight/through side choice. Undefined = bothSides. */
  wrapText?: WrapTextDirection;
  /** Tight/through contour in authored coordinates. */
  wrapPolygon?: Array<{ x?: number; y?: number }>;
  /** `wp:anchor/@allowOverlap`; undefined = true. */
  allowOverlap?: boolean;
  /** `wp:anchor/@layoutInCell`; undefined = true. */
  layoutInCell?: boolean;
  /** Effect extent in px. Undefined = zero. */
  effectExtent?: { top?: number; right?: number; bottom?: number; left?: number };
  /** Ordered image effects. Undefined = none. */
  effects?: Array<{ kind?: string; amount?: number; colors?: string[] }>;
  /** Picture border metadata. Undefined = none. */
  outline?: CellBorderSpec;
  /** Accessibility decorative flag. Undefined = false. */
  decorative?: boolean;
  /** Full hyperlink metadata for picture links. */
  hyperlink?: HyperlinkInfo;
  /** Whether this picture is itself a tracked insertion (`<w:ins>`). */
  isInsertion?: boolean;
  /** Whether this picture is itself a tracked deletion (`<w:del>`). */
  isDeletion?: boolean;
  /** Author of the tracked change wrapping the picture. */
  changeAuthor?: string;
  /** Date of the tracked change wrapping the picture. */
  changeDate?: string;
  /** Revision id of the tracked change (for sidebar matching). */
  changeRevisionId?: number;
  pmStart?: number;
  pmEnd?: number;
};

/** Run for an explicit w:br — ends the line, not the paragraph. */
export type LineBreakRun = {
  kind: 'lineBreak';
  pmStart?: number;
  pmEnd?: number;
};

/**
 * Run whose text is a field instruction (PAGE, NUMPAGES, ...) resolved by the
 * painter once the page context is known.
 */
export type FieldRun = RunFormatting & {
  kind: 'field';
  fieldType: 'PAGE' | 'NUMPAGES' | 'DATE' | 'TIME' | 'OTHER';
  /**
   * Raw Word field type token (e.g. `TOC`, `PAGEREF`) when `fieldType`
   * collapsed it to a painter category. Carried for a11y announcement only.
   */
  rawType?: string;
  /**
   * Raw field instruction text (`w:instrText`), carried INERT so the display
   * list can announce field identity. Never parsed into behavior or executed.
   */
  instruction?: string;
  /** Fallback text if field can't be resolved */
  fallback?: string;
  pmStart?: number;
  pmEnd?: number;
};

/**
 * Any run a paragraph can contain.
 */
export type Run = TextRun | TabRun | ImageRun | LineBreakRun | FieldRun;

/** Paragraph spacing (w:spacing): above/below plus the w:lineRule line rule. */
export type ParagraphSpacing = {
  before?: number;
  after?: number;
  line?: number;
  lineUnit?: 'px' | 'multiplier';
  lineRule?: 'auto' | 'exact' | 'atLeast';
};

/**
 * Paragraph indents in pixels (w:ind), including first-line and hanging.
 */
export type ParagraphIndent = {
  left?: number;
  right?: number;
  firstLine?: number;
  hanging?: number;
};

/** How text anchors on a tab stop; mirrors w:tab/@w:val (§17.3.1.37). */
export type TabAlignment = 'start' | 'end' | 'center' | 'decimal' | 'bar' | 'clear';

/**
 * One entry of a paragraph's tab-stop grid.
 */
export type TabStop = {
  /** Tab alignment mode */
  val: TabAlignment;
  /** Position in twips from left margin */
  pos: number;
  /** Optional leader character */
  leader?: 'none' | 'dot' | 'hyphen' | 'underscore' | 'heavy' | 'middleDot';
};

/**
 * One edge of a paragraph border (w:pBdr child), pre-converted to pixels.
 */
export type BorderStyle = {
  color?: string; // css color
  style?: string;
  width?: number; // px
  space?: number; // gap between border and text in px (w:space, pt in source)
};

/** Paragraph border edges, incl. the between-paragraphs rule and bar edge. */
export type ParagraphBorders = {
  top?: BorderStyle;
  bottom?: BorderStyle;
  left?: BorderStyle;
  right?: BorderStyle;
  between?: BorderStyle;
  bar?: BorderStyle;
};

/**
 * Numbering reference of a list paragraph (w:numPr): instance id + level.
 */
export type ListNumPr = {
  numId?: number;
  ilvl?: number;
};

/**
 * Everything paragraph-level the measurer and painter need (derived from
 * w:pPr plus resolved list/numbering state).
 */
export type ParagraphAttrs = {
  alignment?: 'left' | 'center' | 'right' | 'justify';
  spacing?: ParagraphSpacing;
  /** See ParagraphFormatting.spacingExplicit. */
  spacingExplicit?: { before?: boolean; after?: boolean };
  indent?: ParagraphIndent;
  keepNext?: boolean;
  keepLines?: boolean;
  pageBreakBefore?: boolean;
  styleId?: string;
  contextualSpacing?: boolean;
  /** Right-to-left paragraph direction */
  bidi?: boolean;
  borders?: ParagraphBorders;
  shading?: string; // CSS background color
  tabs?: TabStop[]; // Custom tab stops
  // List properties
  numPr?: ListNumPr;
  listMarker?: string; // Pre-computed marker text (e.g., "1.", "•", "a)")
  listIsBullet?: boolean;
  listMarkerHidden?: boolean; // w:vanish on numbering level rPr
  listMarkerFontFamily?: string; // from numbering level rPr (w:rFonts)
  listMarkerFontSize?: number; // from numbering level rPr, in points
  listMarkerSuffix?: 'tab' | 'space' | 'nothing'; // §17.9.25 w:suff; default 'tab'
  /**
   * Tracked-change state of the list numbering itself. When a list is applied
   * (or removed) under suggesting mode, the marker paints in the insertion /
   * deletion color so an inserted list item's number reads as part of the
   * suggestion, matching Word. `undefined` = numbering is not a pending change.
   */
  listMarkerRevision?: 'ins' | 'del';
  /** Document-wide `w:defaultTabStop` in twips (§17.6.13). Default 720. */
  defaultTabStopTwips?: number;
  // Default font for empty paragraphs (from style's rPr / pPr/rPr)
  defaultFontSize?: number; // in points
  defaultFontFamily?: string;
  /**
   * Skip the empty-paragraph line-height fallback. Used by HF measurement
   * for the canonical OOXML "trailing empty paragraph after a table" pattern
   * — Word renders that paragraph as a zero-height anchor, not as a full
   * line-height of phantom space. See `normalizeHeaderFooterMeasureBlocks`
   * (#381).
   */
  suppressEmptyParagraphHeight?: boolean;
  /**
   * Tracked-change marker on the paragraph mark itself
   * (`<w:pPr><w:rPr><w:ins/>`). The painter renders a pilcrow and margin
   * change bar; the cache key includes presence so two paragraphs with
   * different revisions don't share a measurement.
   */
  pPrIns?: import('../../types/content/trackedChange').RevisionInfo | null;
  /** Tracked-change marker on the paragraph mark (`<w:pPr><w:rPr><w:del/>`). */
  pPrDel?: import('../../types/content/trackedChange').RevisionInfo | null;
};

/**
 * Flow block for one paragraph: its runs plus resolved attributes.
 */
export type ParagraphBlock = {
  /** Enclosing block-SDT group memberships (outermost first), if any. */
  sdtGroups?: SdtGroup[];
  kind: 'paragraph';
  id: BlockId;
  /** Stable Word `w14:paraId` / PM `paraId`, when available. */
  paraId?: string;
  runs: Run[];
  attrs?: ParagraphAttrs;
  /** ProseMirror start position for this block. */
  pmStart?: number;
  /** ProseMirror end position for this block. */
  pmEnd?: number;
};

/**
 * One edge of a table-cell border, in painter-ready CSS units.
 */
export type CellBorderSpec = {
  width?: number; // pixels
  color?: string; // CSS color
  style?: string; // CSS border-style (solid, dashed, dotted, double)
};

/**
 * Border edges of a single cell.
 */
export type CellBorders = {
  top?: CellBorderSpec;
  right?: CellBorderSpec;
  bottom?: CellBorderSpec;
  left?: CellBorderSpec;
};

/**
 * One cell of a table row; its content is a nested block list.
 */
export type TableCell = {
  id: BlockId;
  blocks: LayoutBlock[];
  colSpan?: number;
  rowSpan?: number;
  width?: number;
  /** Original DOCX cell width value, before unit conversion. */
  widthValue?: number;
  /** Original DOCX cell width type ('auto', 'pct', 'dxa', 'nil'). */
  widthType?: string;
  /** Preferred `tcW` retained as one typed value. Undefined = legacy fields. */
  preferredWidth?: { value?: number; type?: 'auto' | 'pct' | 'dxa' | 'nil' };
  /** Grid column where this cell starts. Undefined = derive sequentially. */
  gridStart?: number;
  /** Min-content width in px. Undefined = not measured. */
  minContentWidth?: number;
  /** Max-content width in px. Undefined = not measured. */
  maxContentWidth?: number;
  verticalAlign?: 'top' | 'center' | 'bottom';
  background?: string;
  borders?: CellBorders;
  /** Per-cell padding in pixels (from w:tcMar or table-level w:tblCellMar) */
  padding?: { top: number; right: number; bottom: number; left: number };
  /**
   * `w:noWrap`: when true, the cell forbids text wrapping inside it. The
   * painter renders this as `white-space: nowrap` on the content container
   * — content stays on one line and the cell expands horizontally.
   */
  noWrap?: boolean;
  /** Tracked cell marker (`<w:cellIns>` / `<w:cellDel>` / `<w:cellMerge>`) — painter colors the top border. */
  trackedMarker?: import('../../types/content/trackedChange').CellMarker;
};

/**
 * One row of a table block.
 */
export type TableRow = {
  id: BlockId;
  cells: TableCell[];
  height?: number;
  heightRule?: 'auto' | 'atLeast' | 'exact';
  isHeader?: boolean;
  /**
   * `w:cantSplit` (§17.4.6): the row may not break across a page boundary.
   * The layout engine keeps such a row whole, moving it wholesale to the next
   * page rather than splitting its content.
   */
  cantSplit?: boolean;
  /** Leading omitted grid columns (`w:gridBefore`). Undefined = 0. */
  gridBefore?: number;
  /** Trailing omitted grid columns (`w:gridAfter`). Undefined = 0. */
  gridAfter?: number;
  /** Preferred width before cells. Undefined = 0/auto. */
  widthBefore?: { value?: number; type?: 'auto' | 'pct' | 'dxa' | 'nil' };
  /** Preferred width after cells. Undefined = 0/auto. */
  widthAfter?: { value?: number; type?: 'auto' | 'pct' | 'dxa' | 'nil' };
  /** Tracked row ins / del (`<w:trPr><w:ins/>` / `<w:del/>`). */
  trackedIns?: import('../../types/content/trackedChange').RevisionInfo;
  /** see trackedIns */ trackedDel?: import('../../types/content/trackedChange').RevisionInfo;
};

/**
 * Placement of a floating table (w:tblpPr), converted to pixels: anchors,
 * fixed or named x/y, and the wrap distances around it.
 */
export type FloatingTablePosition = {
  horzAnchor?: 'margin' | 'page' | 'text';
  tblpX?: number;
  tblpXSpec?: 'left' | 'center' | 'right' | 'inside' | 'outside';
  vertAnchor?: 'margin' | 'page' | 'text';
  tblpY?: number;
  tblpYSpec?: 'top' | 'center' | 'bottom' | 'inside' | 'outside' | 'inline';
  topFromText?: number;
  rightFromText?: number;
  bottomFromText?: number;
  leftFromText?: number;
};

/**
 * Flow block for a table: rows plus grid/justification/float properties.
 */
export type TableBlock = {
  /** Enclosing block-SDT group memberships (outermost first), if any. */
  sdtGroups?: SdtGroup[];
  kind: 'table';
  id: BlockId;
  rows: TableRow[];
  columnWidths?: number[];
  /** Authored `w:tblGrid` values, before algorithmic adjustment. */
  gridWidths?: number[];
  /** Table width value (twips for dxa, 50ths of percent for pct). */
  width?: number;
  /** Table width type ('auto', 'pct', 'dxa', 'nil'). */
  widthType?: string;
  /** Preferred `tblW` retained as one typed value. Undefined = legacy fields. */
  preferredWidth?: { value?: number; type?: 'auto' | 'pct' | 'dxa' | 'nil' };
  /** `w:tblLayout`; undefined preserves the legacy resolver. */
  layoutMode?: 'fixed' | 'autofit';
  /** Intrinsic sizing policy. Undefined = legacy fill-available-width. */
  widthAlgorithm?: 'legacy' | 'fixed' | 'autofit';
  /** Resolved table-style provenance/conditionals. */
  styleCascade?: {
    selectedStyleId?: string;
    defaultStyleId?: string;
    basedOnStyleIds?: string[];
    appliedRegions?: string[];
    rowBandSize?: number;
    columnBandSize?: number;
  };
  /** Resolved table background. Undefined = transparent. */
  background?: string;
  /** Table horizontal alignment */
  justification?: 'left' | 'center' | 'right';
  /** Visual RTL column order (`w:bidiVisual`): painter renders logical column 0 rightmost. */
  bidi?: boolean;
  /** Table indent from left margin (in pixels, from w:tblInd) */
  indent?: number;
  /** Floating table properties (pixel values). */
  floating?: FloatingTablePosition;
  pmStart?: number;
  pmEnd?: number;
};

/**
 * Flow block for an image promoted out of the text flow (anchored/floating).
 */
export type ImageBlock = {
  /** Enclosing block-SDT group memberships (outermost first), if any. */
  sdtGroups?: SdtGroup[];
  kind: 'image';
  id: BlockId;
  src: string;
  width: number;
  height: number;
  alt?: string;
  /** CSS transform string (rotation, flip) */
  transform?: string;
  opacity?: number;
  rotationDeg?: number;
  flipH?: boolean;
  flipV?: boolean;
  rotationBounds?: { width?: number; height?: number; offsetX?: number; offsetY?: number };
  anchor?: {
    isAnchored?: boolean;
    offsetH?: number;
    offsetV?: number;
    behindDoc?: boolean;
    position?: ImageRunPosition;
    relativeHeight?: number;
    allowOverlap?: boolean;
    layoutInCell?: boolean;
    wrapType?: string;
    wrapText?: WrapTextDirection;
    wrapPolygon?: Array<{ x?: number; y?: number }>;
  };
  /** Hyperlink URL for clickable image */
  hlinkHref?: string;
  hlinkTitle?: string;
  decorative?: boolean;
  crop?: { top?: number; right?: number; bottom?: number; left?: number };
  effects?: Array<{ kind?: string; amount?: number; colors?: string[] }>;
  outline?: CellBorderSpec;
  pmStart?: number;
  pmEnd?: number;
};

/** Fill data carried by a DrawingML shape block. */
export type ShapeBlockFill = {
  type: 'none' | 'solid' | 'gradient' | 'pattern' | 'picture';
  color?: string;
  gradientType?: string;
  gradientAngle?: number;
  gradientStops?: Array<{ position: number; color: string }>;
  patternPreset?: string;
  foregroundColor?: string;
  backgroundColor?: string;
  pictureRelId?: string;
  /** Resolved SAFE embedded picture source (`blob:`/`data:` only; never external). */
  pictureSrc?: string;
  /** Picture source crop fractions (`a:srcRect`, 0..1; negative = outset). */
  pictureSrcRect?: { left?: number; top?: number; right?: number; bottom?: number };
  /** Picture fill mode. Undefined = 'stretch'. */
  pictureFillMode?: 'stretch' | 'tile';
  /** Tile parameters when `pictureFillMode` = 'tile'. Offsets px, scales fractions. */
  pictureTile?: {
    offsetX?: number;
    offsetY?: number;
    scaleX?: number;
    scaleY?: number;
    alignment?: string;
    flip?: 'none' | 'x' | 'y' | 'xy';
  };
  /** Stretch target rect fractions of the shape box (`a:fillRect`, 0..1). */
  pictureStretchRect?: { left?: number; top?: number; right?: number; bottom?: number };
  /** Picture alpha 0..1 (`a:alphaModFix`). Undefined = opaque. */
  pictureOpacity?: number;
  themeRefIndex?: number;
};

/** Stroke data carried by a DrawingML shape block. */
export type ShapeBlockStroke = {
  color?: string;
  /** Width in CSS pixels. */
  width?: number;
  /** DrawingML/CSS dash style token. */
  dash?: string;
  compound?: string;
  alignment?: string;
  cap?: string;
  join?: string;
  miterLimit?: number;
  customDash?: number[];
  headEnd?: { type?: string; width?: string; length?: string };
  tailEnd?: { type?: string; width?: string; length?: string };
};

/** Rotation/flip transform carried by a DrawingML shape block. */
export type ShapeBlockTransform = {
  /** Rotation in degrees. */
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
};

/**
 * Flow block for a basic DrawingML autoshape/connector.
 *
 * `geometryPath` is normalized to the shape-local `[0, 1]` box; consumers scale
 * it by `width`/`height` and place the corresponding `ShapeFragment` at `x/y`.
 */
export type ShapeBlock = {
  /** Enclosing block-SDT group memberships (outermost first), if any. */
  sdtGroups?: SdtGroup[];
  kind: 'shape';
  id: BlockId;
  /** Preset geometry name or source shape label; `geometryPath` is authoritative. */
  shapeType: string;
  geometryPath: PresetGeometryPathCommand[];
  fill?: ShapeBlockFill;
  stroke?: ShapeBlockStroke;
  transform?: ShapeBlockTransform;
  width: number;
  height: number;
  /** Optional placed coordinates; page fragments are the authoritative placement. */
  x?: number;
  y?: number;
  /** Optional inner paragraphs for future text-bearing shape rendering. */
  innerText?: ParagraphBlock[];
  /** Pre-measured inner paragraph measures for display-list shape text. */
  innerMeasures?: ParagraphExtent[];
  /** Child shapes positioned relative to this shape's top-left corner. */
  children?: ShapeBlock[];
  /** General heterogeneous scene. Undefined = legacy shape-only children. */
  scene?: import('../../types/content/shape').DrawingScene;
  /** Ordered 2-D effects. Undefined = none. */
  effects?: import('../../types/content/shape').ShapeEffect[];
  /** Full text-body/autofit settings. */
  textBodyProperties?: import('../../types/content/shape').ShapeTextBodyProperties;
  /** Anchor/wrap metadata. Undefined = in-flow. */
  position?: ImageRunPosition;
  wrapType?: string;
  wrapText?: WrapTextDirection;
  relativeHeight?: number;
  behindDoc?: boolean;
  decorative?: boolean;
  title?: string;
  description?: string;
  /** PM/doc positions for selection mapping. */
  docStart?: number;
  docEnd?: number;
  pmStart?: number;
  pmEnd?: number;
};

/**
 * Flow block for a basic DrawingML chart.
 *
 * v1 measurement is the drawing extent bbox; display-list emission renders the
 * common chart families from the normalized chart model.
 */
export type ChartBlock = {
  /** Enclosing block-SDT group memberships (outermost first), if any. */
  sdtGroups?: SdtGroup[];
  kind: 'chart';
  id: BlockId;
  chart: Chart;
  width: number;
  height: number;
  position?: ImageRunPosition;
  wrapType?: string;
  wrapText?: WrapTextDirection;
  relativeHeight?: number;
  behindDoc?: boolean;
  docStart?: number;
  docEnd?: number;
  pmStart?: number;
  pmEnd?: number;
};

/**
 * Flow block marking a section boundary and the page geometry the next
 * section requests (applied by the Rust engine's section handling;
 * ECMA-376 §17.6.22).
 */
export type SectionBreakBlock = {
  /** Enclosing block-SDT group memberships (outermost first), if any. */
  sdtGroups?: SdtGroup[];
  kind: 'sectionBreak';
  id: BlockId;
  type?: 'continuous' | 'nextPage' | 'evenPage' | 'oddPage' | 'nextColumn';
  /** Stable section identity. Undefined = derive from break order. */
  sectionId?: string;
  /** Zero-based section index. Undefined = derive during pagination. */
  sectionIndex?: number;
  pageSize?: { w: number; h: number };
  orientation?: 'portrait' | 'landscape';
  margins?: PageMargins;
  columns?: ColumnLayout;
  /** Effective header/footer references for the section. */
  headerFooterRefs?: PageHeaderFooterRefs;
  /** Section-relative page-number settings. */
  pageNumbering?: PageNumberingContract;
  /** Page-border contract retained for per-page selection. */
  pageBorders?: PageBorderContract;
  /** Effective watermark metadata. Undefined = none. */
  watermark?: import('../../types/content/watermark').Watermark;
  /** Vertical page alignment. Undefined = top. */
  verticalAlign?: 'top' | 'center' | 'both' | 'bottom';
  /** Per-section note settings. Undefined = document defaults. */
  noteSettings?: NoteSettingsContract;
};

/**
 * Flow block for a hard page break (w:br type="page").
 */
export type PageBreakBlock = {
  /** Enclosing block-SDT group memberships (outermost first), if any. */
  sdtGroups?: SdtGroup[];
  kind: 'pageBreak';
  id: BlockId;
  pmStart?: number;
  pmEnd?: number;
};

/**
 * Flow block for a hard column break (w:br type="column").
 */
export type ColumnBreakBlock = {
  /** Enclosing block-SDT group memberships (outermost first), if any. */
  sdtGroups?: SdtGroup[];
  kind: 'columnBreak';
  id: BlockId;
  pmStart?: number;
  pmEnd?: number;
};

/** Default internal margins for text boxes (OOXML defaults in pixels) */
export const DEFAULT_TEXTBOX_MARGINS = { top: 4, bottom: 4, left: 7, right: 7 };

/** Default text box width in pixels when no width is specified */
export const DEFAULT_TEXTBOX_WIDTH = 200;

/**
 * Text box block — positioned container with paragraph content.
 */
export type TextBoxBlock = {
  /** Enclosing block-SDT group memberships (outermost first), if any. */
  sdtGroups?: SdtGroup[];
  kind: 'textBox';
  id: BlockId;
  /** Width in pixels */
  width: number;
  /** Height in pixels (may be auto-calculated) */
  height?: number;
  /** Fill/background color */
  fillColor?: string;
  /** Border width in pixels */
  outlineWidth?: number;
  /** Border color */
  outlineColor?: string;
  /** Border style */
  outlineStyle?: string;
  /** Internal padding */
  margins?: { top: number; bottom: number; left: number; right: number };
  /** Paragraph blocks inside the text box */
  content: ParagraphBlock[];
  /** Display mode copied from the ProseMirror text box node */
  displayMode?: 'inline' | 'float' | 'block';
  /** CSS float direction copied from the ProseMirror text box node */
  cssFloat?: 'left' | 'right' | 'none';
  /** OOXML wrap type for anchored text boxes */
  wrapType?: string;
  /** OOXML wrapText direction */
  wrapText?: WrapTextDirection;
  /** Anchor target used during DOCX import/export */
  anchorTarget?: 'followingBlock';
  /** Position for floating/anchored text boxes */
  position?: ImageRunPosition;
  /** Wrap distances in pixels */
  distTop?: number;
  distBottom?: number;
  distLeft?: number;
  distRight?: number;
  pmStart?: number;
  pmEnd?: number;
};

/**
 * Identity of a block-level Structured Document Tag (content control) enclosing
 * a run of flow blocks. Block SDTs flatten into their child flow blocks for
 * pagination; each child carries its group(s) so the painter draws the boundary.
 */
export interface SdtGroup {
  id: string; // stable per-document id (derived from the PM node position)
  sdtType: string; // control type (richText, dropDownList, ...)
  tag?: string; // w:tag
  alias?: string; // w:alias
  lock?: string; // lock mode
  checked?: boolean; // live checkbox glyph state
  bound?: boolean; // data-bound (w:dataBinding): suppress the editable widget
  repeatingItem?: boolean; // w15:repeatingSectionItem: show add/remove affordances
  /** Authored numeric `w:id`; undefined = address by `id`/PM position. */
  controlId?: number;
  /** Stable PM position for untagged controls. */
  pos?: number;
  /** Typed overlay state. Undefined = derive from content. */
  controlState?: import('../../types/content/sdt').SdtControlState;
  /** Typed list/date metadata used by overlays. */
  properties?: import('../../types/content/sdt').SdtProperties;
}

/**
 * Union of every block kind the layout engine knows about.
 *
 * Two sites must stay in sync with this type:
 * - `createMeasureBlock` in `layout/measure/rustMeasureSource.ts` (the sole
 *   measurer — it ends in `assertExhaustiveLayoutBlock(block, '<site>')`,
 *   so adding a variant here without updating it is a typecheck error)
 * - the Rust engine's block enum (`crates/docx-layout/src/types.rs`), which
 *   deserializes the same contract
 */
export type LayoutBlock =
  | ParagraphBlock
  | TableBlock
  | ImageBlock
  | ShapeBlock
  | ChartBlock
  | TextBoxBlock
  | SectionBreakBlock
  | PageBreakBlock
  | ColumnBreakBlock;

/**
 * Exhaustiveness guard for `LayoutBlock`-shaped switches. Call from the
 * `default` arm with the still-typed value; TypeScript will refuse to
 * compile if any variant of `LayoutBlock` was missed. The thrown error
 * names the calling site so runtime failures (e.g. an old adapter
 * compiled against a newer core) point future debuggers at the contract.
 */
export function assertExhaustiveLayoutBlock(block: never, site: string): never {
  const kind = (block as { kind?: string }).kind ?? '<unknown>';
  throw new Error(
    `${site}: unhandled LayoutBlock kind "${kind}". ` +
      `Add the case alongside the other LayoutBlock switches (see types.ts).`
  );
}

// =============================================================================
// extents — per-block measurement results
// =============================================================================

/**
 * One typeset line of a measured paragraph: which run/char span it covers
 * and its vertical metrics.
 */
export type TypesetRow = {
  /** Starting run index (inclusive). */
  headRun: number;
  /** Starting character index within headRun. */
  headChar: number;
  /** Ending run index (inclusive). */
  tailRun: number;
  /** Ending character index within tailRun (exclusive). */
  tailChar: number;
  /** Total width of the line in pixels. */
  width: number;
  /** Ascent (height above baseline) in pixels. */
  ascent: number;
  /** Descent (height below baseline) in pixels. */
  descent: number;
  /** Total line height in pixels. */
  lineHeight: number;
  /** Left offset from floating images (pixels from content left edge). */
  leftOffset?: number;
  /** Right offset from floating images (pixels from content right edge). */
  rightOffset?: number;
  /** Optional split segments for centered floating exclusions. */
  segments?: TypesetRowSegment[];
  /**
   * Vertical space inserted before this line to skip past floats that leave
   * no usable horizontal width at the natural line Y. Painters render this
   * as marginTop on the line element; measurement adds it to totalHeight.
   */
  floatSkipBefore?: number;
  /** Exact per-run advances in visual paint order. Undefined = legacy estimation. */
  runAdvances?: TypesetRunAdvance[];
  /** Exact shaped cluster advances. Undefined = legacy estimation. */
  clusterAdvances?: TypesetClusterAdvance[];
  /** Bidi slices with separate visual/logical order. */
  bidiSlices?: TypesetBidiSlice[];
};

export type TypesetRunAdvance = {
  runIndex?: number;
  startChar?: number;
  endChar?: number;
  advance?: number;
  logicalOrder?: number;
};

export type TypesetClusterAdvance = {
  runIndex?: number;
  startChar?: number;
  endChar?: number;
  advance?: number;
  xOffset?: number;
  bidiLevel?: number;
  logicalOrder?: number;
};

export type TypesetBidiSlice = {
  runIndex?: number;
  startChar?: number;
  endChar?: number;
  advance?: number;
  bidiLevel?: number;
  visualOrder?: number;
  logicalOrder?: number;
};

export type TypesetRowSegment = {
  headRun: number;
  headChar: number;
  tailRun: number;
  tailChar: number;
  leftOffset: number;
  availableWidth: number;
  width: number;
};

/**
 * Extent of a paragraph: its typeset lines and their combined height.
 */
export type ParagraphExtent = {
  kind: 'paragraph';
  lines: TypesetRow[];
  totalHeight: number;
};

/**
 * Extent of an image block (its display box).
 */
export type ImageExtent = {
  kind: 'image';
  width: number;
  height: number;
};

/**
 * Extent of a DrawingML shape block (v1 measurement is its bounding box).
 */
export type ShapeExtent = {
  kind: 'shape';
  width: number;
  height: number;
  /** Pre-measured inner paragraph measures for text-bearing shapes. */
  innerMeasures?: ParagraphExtent[];
};

/**
 * Extent of a chart block (its drawing bbox).
 */
export type ChartExtent = {
  kind: 'chart';
  width: number;
  height: number;
};

/**
 * Extent of one table cell, with the extents of its nested blocks.
 */
export type TableCellExtent = {
  blocks: BlockExtent[];
  width: number;
  height: number;
  colSpan?: number;
  rowSpan?: number;
};

/**
 * Extent of one table row.
 */
export type TableRowExtent = {
  cells: TableCellExtent[];
  height: number;
};

/**
 * Extent of a whole table: per-row extents plus the resolved column grid.
 */
export type TableExtent = {
  kind: 'table';
  rows: TableRowExtent[];
  columnWidths: number[];
  totalWidth: number;
  totalHeight: number;
};

/** Extent placeholder for a section break — occupies no space itself. */
export type SectionBreakExtent = {
  kind: 'sectionBreak';
};

/** Extent placeholder for a page break — zero-size by definition. */
export type PageBreakExtent = {
  kind: 'pageBreak';
};

/** Extent placeholder for a column break — zero-size by definition. */
export type ColumnBreakExtent = {
  kind: 'columnBreak';
};

/**
 * Extent of a text box, with its inner paragraphs pre-measured.
 */
export type TextBoxExtent = {
  kind: 'textBox';
  width: number;
  height: number;
  /** Pre-measured inner paragraph measures (avoids re-measuring during render) */
  innerMeasures: ParagraphExtent[];
};

/**
 * Extent of any flow block.
 */
export type BlockExtent =
  | ParagraphExtent
  | ImageExtent
  | ShapeExtent
  | ChartExtent
  | TableExtent
  | TextBoxExtent
  | SectionBreakExtent
  | PageBreakExtent
  | ColumnBreakExtent;

// =============================================================================
// fragments — block slices placed on a page
// =============================================================================

/**
 * Fields every fragment carries: which block it came from, where it sits on
 * its page, and the PM span it maps to.
 */
export type FragmentBase = {
  /** Block ID this fragment belongs to. */
  blockId: BlockId;
  /** X position on page (relative to page left). */
  x: number;
  /** Y position on page (relative to page top). */
  y: number;
  /** Width of the fragment. */
  width: number;
  /** ProseMirror start position (for click mapping). */
  pmStart?: number;
  /** ProseMirror end position (for click mapping). */
  pmEnd?: number;
};

/**
 * The slice of a paragraph shown on one page — a half-open line range
 * `[fromLine, toLine)` of the measure, so a page-split paragraph yields one
 * fragment per page.
 */
export type ParagraphFragment = FragmentBase & {
  kind: 'paragraph';
  /** First line index (inclusive) from the measure. */
  fromLine: number;
  /** Last line index (exclusive) from the measure. */
  toLine: number;
  /** Height of this fragment. */
  height: number;
  /** True if this continues from a previous page. */
  carriedFromPrev?: boolean;
  /** True if this continues onto the next page. */
  carriedToNext?: boolean;
  /** Per-line resolved run segments for `[fromLine, toLine)`, aligned so `resolvedLines[k]` is line `fromLine + k` (see {@link ResolvedLine}); materialized during layout, omitted from golden serialization. */
  resolvedLines?: ResolvedLine[];
};

/**
 * The slice of a table shown on one page — a half-open row range plus
 * optional top/bottom clips when a single row breaks mid-content.
 */
export type TableFragment = FragmentBase & {
  kind: 'table';
  /** First row index (inclusive). */
  rowStart: number;
  /** Last row index (exclusive). */
  rowEnd: number;
  /** Height of this fragment. */
  height: number;
  /** True if this is a floating table. */
  isFloating?: boolean;
  /** True if this continues from a previous page. */
  carriedFromPrev?: boolean;
  /** True if this continues onto the next page. */
  carriedToNext?: boolean;
  /** Number of header rows prepended to this continuation fragment (0 or undefined for first fragment). */
  headerRowCount?: number;
  /**
   * Pixels to skip from the top of `rowStart`. Non-zero when this fragment's
   * first row is the continuation of a row that broke across a page boundary
   * (Word's "allow row to break across pages"). The painter renders the row
   * shifted up by this amount so the already-shown top slice is clipped.
   */
  clipTop?: number;
  /**
   * Visible height (px) measured from the top of the LAST row (`rowEnd - 1`).
   * Set when that row breaks mid-content onto the next page; `undefined`
   * means the last row is fully visible. When `rowStart === rowEnd - 1`, the
   * visible band of that single row is `[clipTop, clipBottom)`.
   */
  clipBottom?: number;
};

/**
 * An image placed on a page.
 */
export type ImageFragment = FragmentBase & {
  kind: 'image';
  /** Height of the image. */
  height: number;
  /** True if this is an anchored/floating image. */
  isAnchored?: boolean;
  /** Z-index for layering. */
  zIndex?: number;
  /** Original content frame when `width`/`height` are transformed bbox dimensions. */
  contentFrame?: { x?: number; y?: number; width?: number; height?: number };
};

/**
 * A DrawingML shape placed on a page.
 */
export type ShapeFragment = FragmentBase & {
  kind: 'shape';
  /** Height of the shape bbox. */
  height: number;
  /** PM/doc positions for display-list consumers that use docStart/docEnd names. */
  docStart?: number;
  docEnd?: number;
  /** True when positioned outside normal flow. */
  isAnchored?: boolean;
  /** Z-index for layering. */
  zIndex?: number;
};

/**
 * A DrawingML chart placed on a page.
 */
export type ChartFragment = FragmentBase & {
  kind: 'chart';
  /** Height of the chart bbox. */
  height: number;
  docStart?: number;
  docEnd?: number;
  /** True when positioned outside normal flow. Undefined = false. */
  isAnchored?: boolean;
  /** Stable z-order. Undefined = source order. */
  zIndex?: number;
};

/**
 * A text box placed on a page.
 */
export type TextBoxFragment = FragmentBase & {
  kind: 'textBox';
  /** Height of the text box. */
  height: number;
  /** True when positioned outside normal document flow. */
  isFloating?: boolean;
  /** Stack order hint for anchored text boxes. */
  zIndex?: number;
};

/**
 * Any placed fragment.
 */
export type Fragment =
  | ParagraphFragment
  | TableFragment
  | ImageFragment
  | ShapeFragment
  | ChartFragment
  | TextBoxFragment;

// =============================================================================
// pages and layout — the paginator's output
// =============================================================================

/**
 * Page margins in pixels (w:pgMar), including header/footer distances.
 */
export type PageMargins = {
  top: number;
  right: number;
  bottom: number;
  left: number;
  /** w:header — offset of the header band from the sheet's top edge. */
  header?: number;
  /** w:footer — offset of the footer band from the sheet's bottom edge. */
  footer?: number;
  /** Binding gutter in px. Undefined = 0. */
  gutter?: number;
};

/** Effective header/footer relationship ids for one section/page. */
export type PageHeaderFooterRefs = {
  headerDefault?: string;
  headerFirst?: string;
  headerEven?: string;
  footerDefault?: string;
  footerFirst?: string;
  footerEven?: string;
};

/** Section-relative logical page label contract. */
export type PageNumberingContract = {
  start?: number;
  format?: string;
  chapterStyle?: number;
  chapterSeparator?: 'colon' | 'emDash' | 'enDash' | 'hyphen' | 'period';
};

/** Page-border data retained until per-page section selection. */
export type PageBorderContract = {
  top?: import('../../types/colors').BorderSpec;
  right?: import('../../types/colors').BorderSpec;
  bottom?: import('../../types/colors').BorderSpec;
  left?: import('../../types/colors').BorderSpec;
  display?: 'allPages' | 'firstPage' | 'notFirstPage';
  offsetFrom?: 'page' | 'text';
  zOrder?: 'front' | 'back';
};

/** Cascaded note settings used by section-aware pagination. */
export type NoteSettingsContract = {
  footnote?: import('../../types/content/headerFooter').FootnoteProperties;
  endnote?: import('../../types/content/headerFooter').EndnoteProperties;
  footnoteColumns?: number;
};

/** One measured, non-fragmented note in a page note region. */
export type NoteLayoutItem = {
  kind?: 'footnote' | 'endnote';
  id?: number;
  displayLabel?: string;
  blocks?: LayoutBlock[];
  measures?: BlockExtent[];
  height?: number;
  anchorDocStart?: number;
  anchorDocEnd?: number;
  customMarkFollows?: boolean;
};

/** Footnote/endnote area attached to a page/section boundary. */
export type NoteAreaContract = {
  kind?: 'footnote' | 'endnote';
  placement?: 'pageBottom' | 'beneathText' | 'sectEnd' | 'docEnd';
  y?: number;
  height?: number;
  columns?: number;
  separator?: NoteLayoutItem;
  notes?: NoteLayoutItem[];
  sectionId?: string;
};

/**
 * One laid-out page: its fragments, geometry, and header/footer/footnote
 * bookkeeping.
 */
export type Page = {
  /** Page number (1-indexed). */
  number: number;
  /** Fragments positioned on this page. */
  fragments: Fragment[];
  /** Page margins. */
  margins: PageMargins;
  /** Page size (width, height). */
  size: { w: number; h: number };
  /** Page orientation. */
  orientation?: 'portrait' | 'landscape';
  /** Section index this page belongs to. */
  sectionIndex?: number;
  /** Stable section identity. Undefined = derive from `sectionIndex`. */
  sectionId?: string;
  /** Zero-based physical page ordinal within its section. */
  sectionPageIndex?: number;
  /** One-based logical page number after section restart. */
  sectionPageNumber?: number;
  /** Fully formatted PAGE-field label. Undefined = decimal `number`. */
  pageLabel?: string;
  /** Effective page numbering contract. */
  pageNumbering?: PageNumberingContract;
  /** Header/footer references for this page. */
  headerFooterRefs?: PageHeaderFooterRefs;
  /** Effective per-page header distance in px. Undefined = margins.header/default. */
  headerDistance?: number;
  /** Effective per-page footer distance in px. Undefined = margins.footer/default. */
  footerDistance?: number;
  /** Effective section page borders. Undefined = none. */
  pageBorders?: PageBorderContract;
  /** Effective header-owned watermark. Undefined = none. */
  watermark?: import('../../types/content/watermark').Watermark;
  /** Vertical section alignment. Undefined = top. */
  verticalAlign?: 'top' | 'center' | 'both' | 'bottom';
  /** Footnote IDs that appear on this page (for rendering). */
  footnoteIds?: number[];
  /** Height reserved for the footnote area at page bottom (pixels). */
  footnoteReservedHeight?: number;
  /** Footnote-area columns (`w15:footnoteColumns`); absent/1 = single column. */
  footnoteColumns?: number;
  /** Typed note regions. Undefined = use legacy footnote id/reservation fields. */
  noteAreas?: NoteAreaContract[];
  /** Column layout for this page (if multi-column). */
  columns?: ColumnLayout;
};

/**
 * Multi-column setup for a section (w:cols).
 */
export type ColumnLayout = {
  count: number;
  gap: number;
  equalWidth?: boolean;
  /** Draw vertical separator line between columns (w:sep). */
  separator?: boolean;
  /** Authored unequal column geometry. Undefined = derive equal widths. */
  columns?: Array<{ width?: number; space?: number }>;
};

/**
 * Laid-out content of one header or footer variant.
 */
export type HeaderFooterLayout = {
  height: number;
  fragments: Fragment[];
};

/**
 * The paginator's complete result — everything the painter needs.
 */
export type Layout = {
  /** Serialization contract version. Undefined reads as legacy version 0. */
  contractVersion?: number;
  /** Default page size for the document. */
  pageSize: { w: number; h: number };
  /** All rendered pages with positioned fragments. */
  pages: Page[];
  /** Column configuration (if multi-column). */
  columns?: ColumnLayout;
  /** Header layouts by type (default, first, even). */
  headers?: Record<string, HeaderFooterLayout>;
  /** Footer layouts by type (default, first, even). */
  footers?: Record<string, HeaderFooterLayout>;
  /** Gap between pages in pixels (for rendering). */
  pageGap?: number;
  /**
   * Resume bookmarks for incremental pagination (derived data — omitted from
   * golden serialization). See `checkpoints.ts`.
   */
  checkpoints?: import('./checkpoints').LayoutCheckpoint[];
};

// =============================================================================
// layout options
// =============================================================================

/**
 * Measured header/footer heights keyed by variant.
 */
export type HeaderFooterContentHeights = Partial<
  Record<'default' | 'first' | 'even' | 'odd', number>
>;

/**
 * A footnote's blocks and measures, resolved ahead of pagination so page
 * bottoms can reserve the right space.
 */
export type FootnoteContent = {
  /** Footnote ID. */
  id: number;
  /** Display number (e.g. 1, 2, 3). */
  displayNumber: number;
  /** FlowBlocks for rendering the footnote content. */
  blocks: LayoutBlock[];
  /** Measurements for the blocks. */
  measures: BlockExtent[];
  /** Total height in pixels. */
  height: number;
  /** Stable note kind. Undefined = footnote. */
  noteKind?: 'footnote' | 'endnote';
  /** Fully formatted number/custom mark. Undefined = decimal displayNumber. */
  displayLabel?: string;
  /** Body-reference document range. */
  anchor?: { docStart?: number; docEnd?: number };
  /** `customMarkFollows`; undefined = false. */
  customMarkFollows?: boolean;
};

/**
 * Inputs that configure a pagination run.
 */
export type LayoutOptions = {
  /** Serialization contract version. Undefined reads as legacy version 0. */
  contractVersion?: number;
  /** Initial page size. */
  pageSize: { w: number; h: number };
  /** Initial page margins. */
  margins: PageMargins;
  /** Body-level (final section) page size, used after the last explicit section break. */
  finalPageSize?: { w: number; h: number };
  /** Body-level (final section) margins, used after the last explicit section break. */
  finalMargins?: PageMargins;
  /** Column configuration. */
  columns?: ColumnLayout;
  /** Gap between rendered pages (for UI). */
  pageGap?: number;
  /** Default line height multiplier. */
  defaultLineHeight?: number;
  /** Header content heights by variant. */
  headerContentHeights?: HeaderFooterContentHeights;
  /** Footer content heights by variant. */
  footerContentHeights?: HeaderFooterContentHeights;
  /** Whether section has different first page header/footer. */
  titlePage?: boolean;
  /** Whether section has different even/odd headers/footers. */
  evenAndOddHeaders?: boolean;
  /** Per-page footnote reserved heights (pageNumber → height in pixels). */
  footnoteReservedHeights?: Map<number, number>;
  /** Section break type for the body-level (final) section (for section transition logic). */
  bodyBreakType?: 'continuous' | 'nextPage' | 'evenPage' | 'oddPage';
  /** Effective section states, indexed by section. Undefined = legacy globals. */
  sections?: Array<{
    sectionId?: string;
    pageSize?: { w?: number; h?: number };
    margins?: Partial<PageMargins>;
    columns?: ColumnLayout;
    headerFooterRefs?: PageHeaderFooterRefs;
    pageNumbering?: PageNumberingContract;
    pageBorders?: PageBorderContract;
    watermark?: import('../../types/content/watermark').Watermark;
    noteSettings?: NoteSettingsContract;
  }>;
};

// =============================================================================
// misc helper types
// =============================================================================

/**
 * What a pointer position resolved to: the page, the fragment under it (if
 * any), and the position local to that fragment.
 */
export type HitTestResult = {
  /** Page index (0-based). */
  pageIndex: number;
  /** Fragment that was hit, if any. */
  fragment?: Fragment;
  /** Local X coordinate within the fragment. */
  localX?: number;
  /** Local Y coordinate within the fragment. */
  localY?: number;
};

/**
 * A location expressed in block/run/character terms, with the matching PM
 * position when known.
 */
export type DocumentPosition = {
  /** Block index. */
  blockIndex: number;
  /** Run index within the block (for paragraphs). */
  runIndex?: number;
  /** Character offset within the run. */
  charOffset?: number;
  /** ProseMirror position. */
  pmPos?: number;
};

// =============================================================================
// relocated from the deleted DOM painter (RETIREMENT G1)
// =============================================================================
//
// DOM-free geometry + render-model types/helpers that a RETAINED consumer
// (`editor/computeLayout.ts`, the Rust `paginationSource` seam host) depends on.
// They were previously defined under the DOM painter directory, deleted when
// the canvas renderer became the sole renderer; they live here in the kept
// wasm-contract types module. See
// openspec/changes/rust-canvas-engine/RETIREMENT.md.

/**
 * Page geometry needed to translate OOXML `relativeFrom` anchors into painter
 * coordinates. All values are in CSS pixels.
 */
export interface PageGeometry {
  pageWidth: number;
  pageHeight: number;
  marginLeft: number;
  marginTop: number;
  contentWidth: number;
  contentHeight: number;
}

/**
 * Derive {@link PageGeometry} (px) from a laid-out page's size + margins. Pure
 * arithmetic — shared by the measure pipeline (float-band reservation), the
 * painter, and the layout compute pass so a reserved zone lines up with where
 * an anchored object is placed.
 */
export function pageGeometryFromPage(page: Pick<Page, 'size' | 'margins'>): PageGeometry {
  return {
    pageWidth: page.size.w,
    pageHeight: page.size.h,
    marginLeft: page.margins.left,
    marginTop: page.margins.top,
    contentWidth: page.size.w - page.margins.left - page.margins.right,
    contentHeight: page.size.h - page.margins.top - page.margins.bottom,
  };
}

/**
 * Minimal anchored object shape needed to resolve OOXML positioning against
 * page geometry. Shared by the retained measure pipeline and the DOM painter's
 * compatibility re-export until paint/ is deleted.
 */
export interface AnchoredObjectPositionInput {
  width: number;
  height: number;
  position?: {
    horizontal?: { relativeTo?: string; posOffset?: number; align?: string };
    vertical?: { relativeTo?: string; posOffset?: number; align?: string };
  };
  cssFloat?: 'left' | 'right' | 'none';
}

/**
 * Content-area top Y (px) of an anchored object, resolving its OOXML vertical
 * anchor (`page` / `margin` / `topMargin` / `bottomMargin` / `paragraph`, with
 * `align` or `posOffset`). Exposed so the measure pipeline can reserve a
 * `topAndBottom` band at the exact Y the painter will place the object.
 */
export function resolveAnchoredObjectVerticalTop(
  object: AnchoredObjectPositionInput,
  fragmentY: number,
  geometry?: PageGeometry
): number {
  return resolveVerticalAnchor(object, fragmentY, geometry);
}

function resolveVerticalAnchor(
  object: AnchoredObjectPositionInput,
  fragmentY: number,
  geometry: PageGeometry | undefined
): number {
  const vertical = object.position?.vertical;
  if (!vertical) return fragmentY;

  const band = verticalAnchorBand(vertical.relativeTo, fragmentY, geometry);
  if (vertical.align === 'top') {
    return band.base;
  }
  if (vertical.align === 'center') {
    return band.size ? band.base + (band.size - object.height) / 2 : fragmentY;
  }
  if (vertical.align === 'bottom') {
    return band.size ? band.base + band.size - object.height : fragmentY;
  }
  if (vertical.posOffset !== undefined) {
    return band.base + emuToPixels(vertical.posOffset);
  }

  return vertical.relativeTo === 'paragraph' || vertical.relativeTo === 'line'
    ? fragmentY
    : band.base;
}

function verticalAnchorBand(
  relativeTo: string | undefined,
  fragmentY: number,
  geometry: PageGeometry | undefined
): { base: number; size: number } {
  const pageHeight = geometry?.pageHeight ?? 0;
  const marginTop = geometry?.marginTop ?? 0;
  const contentHeight = geometry?.contentHeight ?? 0;

  switch (relativeTo) {
    case 'paragraph':
    case 'line':
      return { base: fragmentY, size: 0 };
    case 'page':
      return { base: -marginTop, size: pageHeight };
    case 'topMargin':
      return { base: -marginTop, size: marginTop };
    case 'bottomMargin':
      return { base: contentHeight, size: marginTop };
    case 'margin':
    case 'insideMargin':
    case 'outsideMargin':
    default:
      return { base: 0, size: contentHeight };
  }
}

/**
 * Check if a floating image should create text wrapping exclusion zones.
 * wrapNone images (`behind` / `inFront`) are positioned floats but do not
 * shrink line widths; text paints over or under them.
 */
export function isTextWrappingFloatingImageRun(run: ImageRun): boolean {
  if (isWrapNone(run.wrapType) || run.wrapType === 'topAndBottom') return false;
  if (wrapsAroundText(run.wrapType)) return true;
  return run.displayMode === 'float' && run.cssFloat !== 'none';
}

/**
 * Header/footer content for rendering.
 */
export interface HeaderFooterContent {
  /** Flow blocks for the header/footer content. */
  blocks: LayoutBlock[];
  /** Measurements for the blocks. */
  measures: BlockExtent[];
  /** Total height of the content (in-flow stack incl. floating blocks). */
  height: number;
  /**
   * In-flow band height: the height of strictly in-flow content
   * (paragraphs, tables, inline images/text boxes), EXCLUDING anchored /
   * floating objects. This is what grows the header/footer band and pushes
   * the body margin, mirroring Word: a page/margin-anchored shape (e.g. a
   * full-page letterhead in a header) is positioned independently and does
   * NOT push body text down. Use this — not `height`/`visualBottom` — for
   * margin extension. Falls back to `height` when undefined.
   */
  flowHeight?: number;
  /** Top-most visual extent relative to the nominal flow origin. */
  visualTop?: number;
  /** Bottom-most visual extent relative to the nominal flow origin. */
  visualBottom?: number;
}

/**
 * A single footnote item ready for rendering at page bottom.
 */
export interface FootnoteRenderItem {
  /** Display number (e.g. "1", "2") */
  displayNumber: string;
  /** Plain text content */
  text: string;
  /** Measured body-pipeline content used for WYSIWYG painting. */
  content?: FootnoteContent;
}
