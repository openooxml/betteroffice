// renderer-agnostic display list: the single contract between layout output and
// every rendering backend (web canvas, DOM mirror, native, PDF). Derived from a
// computed Layout — never built inside the paginator — so any pagination engine
// (TS or Rust) feeds any backend. All coordinates are page-local px, same unit
// space as Layout. v0 carries shaped-by-browser text runs (string + CSS font);
// phase 2 (docx-text) adds glyph-run primitives without breaking consumers.

export interface DisplayList {
  /** Contract version. Undefined reads as legacy version 0. */
  contractVersion?: number;
  pages: DisplayPage[];
}

/**
 * Modern (`w14:*`) run text-effect payload carried on text/glyph primitives.
 * Same shape as the document model's `TextModernEffects` — px distances,
 * degree directions, 0..1 opacities, resolved `#rrggbb` colors.
 */
export type DisplayTextModernEffects = import('../../types/formatting').TextModernEffects;

export interface DisplayPage {
  pageIndex: number;
  width: number;
  height: number;
  /** Authored body content box (page-local px), derived from page margins. */
  contentBounds?: DisplayBounds;
  /** Authored body column boxes in reading order (page-local px). */
  columnBounds?: DisplayBounds[];
  /** Stable section identity for this page. Undefined = legacy/global section. */
  sectionId?: string;
  /** Zero-based section ordinal. */
  sectionIndex?: number;
  /** Zero-based page ordinal within the section. */
  sectionPageIndex?: number;
  /** Effective page number within the section, after any restart. */
  sectionPageNumber?: number;
  /** Formatted PAGE label (for example, "vii"). */
  pageLabel?: string;
  primitives: DisplayPrimitive[]; // body content, paint order
  /** Resolved page background. Undefined = transparent/host white. */
  background?: string;
  pageBorders?: PageBorderPrimitive[];
  header?: HfRegion;
  footer?: HfRegion;
  /** Footnote/endnote regions in page coordinates. Undefined = none. */
  noteAreas?: NoteRegion[];
}

/** A page-local rectangle emitted as display-list metadata. */
export interface DisplayBounds {
  x: number;
  y: number;
  width: number;
  height: number;
}

/** Page note region; primitives retain normal doc/a11y metadata. */
export interface NoteRegion {
  kind?: 'footnote' | 'endnote';
  sectionId?: string;
  y?: number;
  height?: number;
  columns?: number;
  separatorPrimitives?: DisplayPrimitive[];
  primitives?: DisplayPrimitive[];
  noteIds?: number[];
  /**
   * Per-note backlink metadata (W17): the body-doc PM range of each note's
   * reference mark plus its formatted label, so the a11y mirror can wire
   * note ↔ reference associations (doc-noteref / doc-backlink). Parallel to
   * `noteIds`; undefined = legacy region without backlink data.
   */
  notes?: NoteRegionNote[];
}

/** Backlink metadata for one note in a NoteRegion. */
export interface NoteRegionNote {
  id?: number;
  /** Body-doc PM range of the reference mark anchoring this note. */
  anchorDocStart?: number;
  anchorDocEnd?: number;
  /** Formatted reference label (display number / custom mark). File-derived. */
  label?: string;
}

// header/footer band. Primitives are in page coordinates like body primitives;
// doc positions inside refer to the HF ProseMirror doc identified by rId, NOT
// the body doc — hit-testing and the mirror must scope by region (the painted
// DOM analogue of .layout-page-header / .layout-page-footer).
export interface HfRegion {
  rId: string;
  kind: 'header' | 'footer';
  y: number;
  height: number;
  primitives: DisplayPrimitive[]; // paint order
}

export type DisplayPrimitive =
  | TextRunPrimitive
  | GlyphRunPrimitive
  | RectPrimitive
  | LinePrimitive
  | ImagePrimitive
  | ShapePrimitive
  | DecorationPrimitive;

// grid position of the table cell a primitive paints inside. row/col are
// 0-based anchor-grid coordinates with vmerge/colspan resolved (a vertically
// merged cell keeps its anchor row); spans are >= 1. `continuation` marks the
// synthetic slice of a vmerge cell re-painted on a continuation page — the
// data-vmerge-continuation analogue (not selectable, doc positions stripped).
// Repeated header rows re-painted on later pages are literal re-paints and
// are NOT flagged, matching the DOM painter.
export interface TableCellRef {
  row: number;
  col: number;
  rowSpan: number;
  colSpan: number;
  continuation?: boolean;
  /** Stable semantic cell identity. */
  cellId?: string;
  /** Cell participates as a row/column header. */
  isHeader?: boolean;
  /** Header cell was repeated on a continuation page. */
  repeatedHeader?: boolean;
  /** Word `w:noWrap` cell behavior. */
  noWrap?: boolean;
  /** IDs of header cells that label this data cell. */
  headerIds?: string[];
  /** Collapsed-border ownership retained for exact fragment replay. */
  ownsTopBorder?: boolean;
  ownsRightBorder?: boolean;
  ownsBottomBorder?: boolean;
  ownsLeftBorder?: boolean;
}

/** Table-fragment semantics shared by every primitive emitted for that fragment. */
export interface DisplayTableMetadata {
  tableId?: string;
  rowStart?: number;
  rowEnd?: number;
  rowCount?: number;
  columnCount?: number;
  headerRowCount?: number;
  caption?: string;
  description?: string;
  parentTableId?: string;
}

/** Resolved review presentation for the primitive's `commentIds`. */
export interface DisplayCommentMetadata {
  status?: 'active' | 'resolved';
  authorId?: string;
  paletteIndex?: number;
  color?: string;
  selected?: boolean;
  /** Reviewer display name (file-derived, attacker-controlled). */
  authorName?: string;
  /** Comment date string, passed through verbatim. */
  date?: string;
  /** Plain-text comment body excerpt (builder-capped, file-derived). */
  text?: string;
  /** Number of replies in the thread. */
  replyCount?: number;
  /** Reply summaries in thread order (builder-capped). */
  replies?: DisplayCommentReply[];
}

/** One reply summary inside DisplayCommentMetadata. */
export interface DisplayCommentReply {
  authorName?: string;
  date?: string;
  text?: string;
}

/**
 * Inert field identity for a11y announcement. `instruction` is the raw Word
 * field instruction, carried for announcement/inspection ONLY — it is never
 * parsed into behavior or executed by any consumer (see the repo security guidelines security).
 */
export interface DisplayFieldMetadata {
  /** Painter-resolved category (PAGE, NUMPAGES, DATE, TIME, OTHER). */
  category?: string;
  /** Raw Word field type token (e.g. TOC, PAGEREF, REF). */
  type?: string;
  /** Raw field instruction text (announce-only, builder-capped). */
  instruction?: string;
}

/** Footnote/endnote reference identity on the body reference-mark primitive. */
export interface DisplayNoteRef {
  kind?: 'footnote' | 'endnote';
  id?: number;
}

/** Scoped clip/group membership emitted by the Rust compiler. */
export interface DisplayClipGroupMetadata {
  id?: string;
  clip?: { x?: number; y?: number; w?: number; h?: number };
  opacity?: number;
}

export interface StructuralRevision {
  scope: 'pmark' | 'table' | 'row' | 'cell';
  author: string;
  date?: string;
  revisionId: string;
  kind: 'ins' | 'del' | 'merge';
  rowIndex?: number;
  colIndex?: number;
}

export interface SdtAttrs {
  groupId: string;
  sdtType: string;
  depth?: number;
  tag?: string;
  alias?: string;
  lock?: string;
  checked?: boolean;
  bound?: boolean;
  repeatingItem?: boolean;
}

export interface InlineSdtWidgetAttrs {
  kind: 'checkbox';
  /** Rich control kind. Undefined = legacy checkbox. */
  controlKind?: 'checkbox' | 'dropDownList' | 'comboBox' | 'date' | 'picture' | 'repeatingSection';
  groupId: string;
  pos: number;
  tag?: string;
  alias?: string;
  checked?: boolean;
  controlId?: number;
  value?: string;
  selectedIndex?: number;
  listItems?: Array<{ displayText?: string; value?: string }>;
  dateFormat?: string;
  dateLanguage?: string;
  locked?: boolean;
}

export interface ChartA11yAttrs {
  label: string;
}

// attributes shared by primitives that map back to document content; replaces
// the painted-DOM dataset contract (data-doc-start/end, data-block-id, ...)
export interface DocAttrs {
  docStart?: number;
  docEnd?: number;
  blockId?: number;
  // raw string block id, present when the id is NOT numeric (the live
  // pipeline's compound `block-N` keys). Group by blockKey ?? String(blockId)
  // — exactly one of the two is set whenever the primitive has block identity.
  blockKey?: string;
  // Exact PM range of the paragraph fragment that owns this primitive. Run
  // ranges start inside the paragraph (normally pmStart+1), so the mirror
  // cannot reconstruct the painter wrapper's data-doc-start from docStart.
  fragmentDocStart?: number;
  fragmentDocEnd?: number;
  // stable Word `w14:paraId` / PM `paraId` of the enclosing paragraph, when the
  // source carries one. The a11y mirror stamps it as `data-para-id` on the
  // paragraph wrapper so `scrollToParaId`-style lookups resolve against the
  // mirror. Absent when the paragraph has no paraId (matches the DOM painter).
  paraId?: string;
  // measured line window [fromLine, toLine) of the paragraph fragment this
  // primitive belongs to — the display-list analogue of the painter's
  // data-from-line / data-to-line. Present only on paragraph fragments the a11y
  // mirror surfaces as a paragraph wrapper (body, header/footer, text box);
  // table-cell paragraphs omit it (rendered as ARIA cells, which have no
  // fragment node). The mirror stamps them on the paragraph wrapper.
  fromLine?: number;
  toLine?: number;
  /**
   * Table cell the primitive paints inside, when any. The a11y mirror builds
   * real ARIA table semantics (role table/row/cell, aria-row/colindex) from
   * this — without it a table block reads as a flat run of text.
   */
  cell?: TableCellRef;
  /** Sanitized hyperlink target when this primitive paints linked content. */
  href?: string;
  /** Optional hyperlink tooltip/title for popup UI. */
  tooltip?: string;
  /** ScreenTip exposed as DOM `title`; undefined falls back to `tooltip`. */
  linkTitle?: string;
  /** Safe hyperlink target frame. */
  linkTarget?: string;
  /** Link history bit. Undefined = host-unmanaged. */
  linkHistory?: boolean;
  /** Location within the target document. */
  linkDocLocation?: string;
  commentIds?: string[];
  /**
   * Inert field identity when this primitive paints a field result. The a11y
   * mirror announces it (data-field-* + aria-description); nothing executes it.
   */
  field?: DisplayFieldMetadata;
  /**
   * Footnote/endnote reference identity when this primitive is the body
   * reference mark — the mirror renders it as a doc-noteref link to the
   * note's mirror element (W17 backlinks).
   */
  noteRef?: DisplayNoteRef;
  revision?: { author: string; date: string; revisionId: string; kind: 'ins' | 'del' };
  /** Synthetic numbering glyph for the first line of a list paragraph. */
  listMarker?: boolean;
  /** Pending suggestion paint for a synthetic numbering glyph. */
  listMarkerRevision?: 'ins' | 'del';
  structuralRevision?: StructuralRevision;
  /** Innermost block-level content-control identity for this primitive. */
  sdt?: SdtAttrs;
  /** Full outer-to-inner content-control ancestry. */
  sdtPath?: SdtAttrs[];
  /** Inline content-control widget metadata when this text run is its glyph. */
  inlineSdtWidget?: InlineSdtWidgetAttrs;
  /** Accessibility summary for primitives that compose one chart block. */
  chart?: ChartA11yAttrs;
  /** Logical/source order, independent of visual bidi paint order. */
  logicalOrder?: number;
  /** Resolved bidi embedding level. */
  bidiLevel?: number;
  /** BCP-47 language for a11y/font fallback. */
  lang?: string;
  /** Decorative content is omitted from the accessibility mirror. */
  decorative?: boolean;
  /** Authored accessibility name. */
  ariaLabel?: string;
  /** Authored accessibility description. */
  ariaDescription?: string;
  /** Hidden authored object. Undefined = visible. */
  hiddenObject?: boolean;
  /** Stable logical scene/group identity. */
  groupId?: string;
  /** Review presentation for comment-associated primitives. */
  comment?: DisplayCommentMetadata;
  /**
   * Scoped clip/group membership. Undefined = ungrouped. Kept as metadata
   * until the canvas/mirror batch activates `ClipGroupPrimitive` in the union.
   */
  clipGroup?: DisplayClipGroupMetadata;
  /** Enclosing table-fragment identity and accessibility semantics. */
  table?: DisplayTableMetadata;
}

export interface TextRunPrimitive extends DocAttrs {
  kind: 'text';
  text: string;
  x: number; // pen origin
  baselineY: number;
  width: number; // measured advance of the whole run
  font: string; // CSS font shorthand (v0, browser-shaped); phase 2 adds fontId+glyphs
  color: string;
  letterSpacing?: number;
  // extra advance (px) added after each U+0020 space cluster — the canvas
  // backend replays it as `ctx.wordSpacing`. Set only on justified lines
  // (`jc=both/distribute`); `width` already includes the stretch so the mirror
  // geometry matches. Absent = 0 (no stretch).
  wordSpacing?: number;
  rtl?: boolean;
  /** 0..1 alpha applied only to this run. Absent = opaque. */
  opacity?: number;
  /** Degrees clockwise around the primitive's geometry center. Absent = 0. */
  rotationDeg?: number;
  /** Horizontal text scale as a percentage (100 = normal). Absent = 100. */
  horizontalScale?: number;
  /** Render glyphs uppercase without changing the accessible source text. */
  allCaps?: boolean;
  /** Render lowercase glyphs as small uppercase. */
  smallCaps?: boolean;
  /** Hidden run shown in editing view as dimmed/dotted, matching the DOM painter. */
  hidden?: boolean;
  /** CSS-like shadow effect ported from run formatting. */
  textShadow?: 'shadow' | 'emboss' | 'imprint';
  /** Outlined / hollow text. */
  textOutline?: boolean;
  /** CJK emphasis mark. */
  emphasisMark?: 'dot' | 'comma' | 'circle' | 'underDot';
  /** Legacy text-effect class hook for the a11y mirror. */
  textEffect?: 'blinkBackground' | 'lights' | 'antsBlack' | 'antsRed' | 'shimmer' | 'sparkle';
  /** Modern w14 text effects (glow/shadow/reflection/textFill/textOutline). */
  modernEffects?: DisplayTextModernEffects;
  /** Shaped tab/TOC leader metadata. Undefined = ordinary text. */
  leaderGlyphs?: Omit<LeaderGlyphPrimitive, 'kind'>;
}

// one shaped, positioned glyph inside a GlyphRun. `id` is a glyph index into
// the font identified by the run's `fontId` (NOT a Unicode codepoint); `x`/`y`
// are the pen origin and baseline in page-local px; `cluster` is the byte index
// into the run's source `text` this glyph came from (glyph↔char map for
// hit-testing and the a11y mirror).
export interface PositionedGlyph {
  id: number;
  x: number;
  y: number;
  cluster: number;
  // pen advance for this glyph in page-local px (shaped x_advance plus any
  // justification word-spacing for a U+0020 cluster). `x + advance` is the next
  // glyph's origin; the trailing glyph's `x + advance` closes the run's true
  // right extent, so hit-testing and the a11y mirror read the real run width off
  // the glyphs instead of estimating a uniform trailing advance. Optional for
  // back-compat: pre-shaper display lists omit it and the mirror falls back to
  // the estimate.
  advance?: number;
  /** Logical cluster order before bidi reordering. */
  logicalOrder?: number;
  /** Bidi embedding level for this cluster. */
  bidiLevel?: number;
}

// phase-2 text primitive: a run shaped by the Rust engine into font glyphs
// rather than a browser-shaped CSS string (TextRunPrimitive). The canvas
// backend paints each glyph as a Path2D outline fetched from the font bytes
// (fontId → FontStore), scaled by `size`/upem; the a11y mirror renders `text`
// as a real text node at the run's geometry. `color` is "#rrggbb"; `size` is
// px. Coexists with TextRunPrimitive so v0 (browser-shaped) and phase-2
// (glyph-shaped) runs can share one display list.
export interface GlyphRunPrimitive extends DocAttrs {
  kind: 'glyphRun';
  fontId: number;
  size: number; // px
  color: string; // "#rrggbb"
  text: string; // source text — the a11y mirror renders this as a real text node
  glyphs: PositionedGlyph[];
  // resolved CSS font shorthand of the face this run was shaped with (same
  // recipe as TextRunPrimitive.font). The canvas backend's fillText safety net
  // uses it when glyph outlines are unavailable, so the fallback keeps the
  // measured family/weight/style instead of degrading to generic sans-serif.
  // Absent (pre-contract emissions) = `${size}px sans-serif`.
  fallbackFont?: string;
  // extra advance (px) after each U+0020 space cluster on justified lines
  // (jc=both/distribute); the shaper baked the same stretch into the glyph pen
  // origins, so paint and mirror geometry agree. Absent = 0.
  wordSpacing?: number;
  rtl?: boolean;
  /** 0..1 alpha applied only to this run. Absent = opaque. */
  opacity?: number;
  /** Degrees clockwise around the primitive's geometry center. Absent = 0. */
  rotationDeg?: number;
  /** Horizontal text scale as a percentage (100 = normal). Absent = 100. */
  horizontalScale?: number;
  /** Render glyphs uppercase without changing the accessible source text. */
  allCaps?: boolean;
  /** Render lowercase glyphs as small uppercase. */
  smallCaps?: boolean;
  /** Hidden run shown in editing view as dimmed/dotted, matching the DOM painter. */
  hidden?: boolean;
  /** CSS-like shadow effect ported from run formatting. */
  textShadow?: 'shadow' | 'emboss' | 'imprint';
  /** Outlined / hollow text. */
  textOutline?: boolean;
  /** CJK emphasis mark. */
  emphasisMark?: 'dot' | 'comma' | 'circle' | 'underDot';
  /** Legacy text-effect class hook for the a11y mirror. */
  textEffect?: 'blinkBackground' | 'lights' | 'antsBlack' | 'antsRed' | 'shimmer' | 'sparkle';
  /** Modern w14 text effects (glow/shadow/reflection/textFill/textOutline). */
  modernEffects?: DisplayTextModernEffects;
  /** Shaped tab/TOC leader metadata. Undefined = ordinary glyph run. */
  leaderGlyphs?: Omit<LeaderGlyphPrimitive, 'kind'>;
}

export interface RectPrimitive extends DocAttrs {
  kind: 'rect';
  x: number;
  y: number;
  w: number;
  h: number;
  fill: string;
  /** Alpha multiplier. Undefined = 1. */
  opacity?: number;
}

/** Word/CSS line recipe retained instead of flattening to a solid stroke. */
export type DisplayBorderStyle =
  | 'solid'
  | 'double'
  | 'dotted'
  | 'dashed'
  | 'dashDot'
  | 'dashDotDot'
  | 'triple'
  | 'thinThick'
  | 'thickThin'
  | 'wave'
  | 'doubleWave'
  | 'groove'
  | 'ridge'
  | 'inset'
  | 'outset';

// Table/cell border lines additionally carry ownership metadata through the
// inherited DocAttrs: `cell` names the owning grid cell (with its border
// ownership flags) and `table` the enclosing fragment, so consumers such as
// displayListTables can associate border lines exactly instead of relying on
// the geometric fallback. Absent on pre-contract emissions.
export interface LinePrimitive extends DocAttrs {
  kind: 'line';
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  strokeWidth: number;
  color: string;
  dash?: number[];
  role?: 'border' | 'table-border' | 'table-cut' | 'separator';
  /** Full border recipe. Undefined = infer solid/dash from `dash`. */
  borderStyle?: DisplayBorderStyle;
  /** Secondary color for 3-D/multi-stroke recipes. */
  secondaryColor?: string;
  /** Alpha multiplier. Undefined = 1. */
  opacity?: number;
  /** Owner of the retained Word border recipe. */
  borderOwner?: 'cell' | 'fragment' | 'paragraph' | 'textBox';
}

export interface PageBorderSide {
  width: number;
  color: string;
  style: 'solid' | 'double' | 'dotted' | 'dashed' | 'groove' | 'ridge' | 'inset' | 'outset';
}

export interface PageBorderPrimitive {
  kind: 'pageBorder';
  x: number;
  y: number;
  w: number;
  h: number;
  zOrder?: 'front' | 'back';
  top?: PageBorderSide;
  right?: PageBorderSide;
  bottom?: PageBorderSide;
  left?: PageBorderSide;
}

export interface ImagePrimitive extends DocAttrs {
  kind: 'image';
  relId: string;
  x: number;
  y: number;
  w: number;
  h: number;
  rotationDeg?: number;
  /** 0..1 alpha applied only to this image. Absent = opaque. */
  opacity?: number;
  /** Flip about the content-frame center. Undefined = false. */
  flipH?: boolean;
  /** Flip about the content-frame center. Undefined = false. */
  flipV?: boolean;
  /** Original image frame when w/h are a transformed layout bbox. */
  contentFrame?: { x?: number; y?: number; w?: number; h?: number };
  /** Canvas/CSS filter string for renderer-owned effects, e.g. watermark washout. */
  filter?: string;
  /** True for decorative/non-content images such as picture watermarks. */
  decorative?: boolean;
  crop?: { top: number; right: number; bottom: number; left: number }; // fractions 0..1
  // alternative text (wp:docPr descr → ImageRun/ImageBlock alt); file-derived
  // attacker-controlled data, capped at 2048 chars by the builder
  altText?: string;
  /** Ordered image effects. Undefined = use legacy `filter` only. */
  effects?: Array<{
    kind?: string;
    amount?: number;
    threshold?: number;
    colors?: string[];
    rawName?: string;
  }>;
  /** Picture border recipe. Undefined = none. */
  border?: {
    width?: number;
    color?: string;
    style?: string;
    dash?: number[];
  };
}

export type ShapePathCommand =
  | { type: 'move'; x: number; y: number }
  | { type: 'line'; x: number; y: number }
  | { type: 'quad'; cpx: number; cpy: number; x: number; y: number }
  | {
      type: 'cubic';
      cp1x: number;
      cp1y: number;
      cp2x: number;
      cp2y: number;
      x: number;
      y: number;
    }
  | { type: 'close' };

export interface ShapePrimitive extends DocAttrs {
  kind: 'shape';
  x: number;
  y: number;
  w: number;
  h: number;
  /** Page-local path commands scaled from the shape's normalized geometry. */
  geometryPath: ShapePathCommand[];
  fill?: string;
  stroke?: {
    color: string;
    width: number;
    /** DrawingML/CSS dash style token. */
    dash?: string;
  };
  transform?: {
    rotation?: number;
    flipH?: boolean;
    flipV?: boolean;
  };
  /** True when the shape carries no accessible inner text. */
  decorative?: boolean;
  /** Alpha multiplier. Undefined = 1. */
  opacity?: number;
  /** Lossless paint server; undefined = legacy `fill`. */
  fillPaint?: {
    kind?: 'none' | 'solid' | 'gradient' | 'pattern' | 'picture' | 'theme';
    color?: string;
    angle?: number;
    gradientType?: string;
    stops?: Array<{ position?: number; color?: string }>;
    patternPreset?: string;
    foregroundColor?: string;
    backgroundColor?: string;
    pictureRelId?: string;
    /**
     * Resolved SAFE embedded source for a picture fill (`blob:`/`data:` only —
     * external-mode blip relationships are never resolved, so opening a file
     * cannot trigger a network fetch). The canvas image resolver rejects any
     * other scheme as defense in depth. Absent = unresolvable, paint the
     * legacy `fill`/`color` instead.
     */
    pictureSrc?: string;
    /** Picture source crop fractions (`a:srcRect`, 0..1; negative = outset). */
    pictureSrcRect?: { left?: number; top?: number; right?: number; bottom?: number };
    /** Picture fill mode. Absent = 'stretch' (`a:stretch`). */
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
    /** Picture alpha 0..1 (`a:alphaModFix`). Absent = opaque. */
    pictureOpacity?: number;
    themeRefIndex?: number;
  };
  /** Lossless DrawingML stroke details beyond the legacy `stroke` triple. */
  strokePaint?: {
    color?: string;
    width?: number;
    dash?: string;
    customDash?: number[];
    compound?: 'single' | 'double' | 'thickThin' | 'thinThick' | 'triple';
    alignment?: 'center' | 'inset';
    cap?: 'flat' | 'round' | 'square';
    join?: 'bevel' | 'miter' | 'round';
    miterLimit?: number;
    headEnd?: { type?: string; width?: string; length?: string };
    tailEnd?: { type?: string; width?: string; length?: string };
  };
  /** Ordered 2-D effects. Undefined = none. */
  effects?: Array<{
    kind?: string;
    color?: string;
    opacity?: number;
    blurRadius?: number;
    distance?: number;
    direction?: number;
    size?: number;
    rawName?: string;
  }>;
  /** Effect extents in page px. Undefined = zero. */
  effectExtent?: { top?: number; right?: number; bottom?: number; left?: number };
  /** Versioned heterogeneous DrawingML group/canvas scene. */
  drawingScene?: import('../../types/content/shape').DrawingScene;
  /** Full DrawingML `a:bodyPr` projection for shape text. */
  textBodyProperties?: import('../../types/content/shape').ShapeTextBodyProperties;
}

export interface DecorationPrimitive extends DocAttrs {
  kind: 'decoration';
  deco: 'underline' | 'strike' | 'highlight' | 'comment-range' | 'spell';
  x: number;
  y: number;
  w: number;
  h: number;
  color: string;
  // dashed rule instead of a solid one — set for the tracked-change insertion
  // underline (the painter's `border-bottom: 2px dashed`). Absent = solid.
  dashed?: boolean;
  // dotted rule instead of a solid one — set for hidden-run dotted underline.
  dotted?: boolean;
  /** Exact highlight font-box/source slice. Undefined = legacy rect. */
  highlightSlice?: {
    sourceStart?: number;
    sourceEnd?: number;
    ascent?: number;
    descent?: number;
    includesTrailingWhitespace?: boolean;
  };
  /** Full border/decoration recipe. Undefined = solid/dashed/dotted flags. */
  style?: DisplayBorderStyle;
}

/** Scoped opacity/transform/clip group. Empty/missing members are no-ops. */
export interface ClipGroupPrimitive extends DocAttrs {
  /** Primitive tag. Undefined until Batch F activates the union arm. */
  kind?: 'clipGroup';
  clip?: { x?: number; y?: number; w?: number; h?: number };
  opacity?: number;
  transform?: {
    translateX?: number;
    translateY?: number;
    rotationDeg?: number;
    scaleX?: number;
    scaleY?: number;
    flipH?: boolean;
    flipV?: boolean;
  };
  primitives?: DisplayPrimitive[];
}

/** Shaped tab/TOC leader glyphs; absent metrics mean no painted leader. */
export interface LeaderGlyphPrimitive extends DocAttrs {
  /** Primitive tag. Undefined when embedded as run metadata. */
  kind?: 'leaderGlyphs';
  glyph?: string;
  count?: number;
  x?: number;
  baselineY?: number;
  advance?: number;
  width?: number;
  font?: string;
  fontId?: number;
  size?: number;
  color?: string;
  rtl?: boolean;
}
