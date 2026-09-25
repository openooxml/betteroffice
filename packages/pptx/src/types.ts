export interface TextStyle {
  bold?: boolean;
  italic?: boolean;
  fontSizePt?: number;
  color?: string;
  fontFamily?: string;
  underline?: string;
  spacingPt?: number;
  baselinePct?: number;
}

export type TextStylePatch = TextStyle;

export interface TextStyleSnapshot {
  bold: boolean | null;
  italic: boolean | null;
  fontSizePt: number | null;
  color: string | null;
  fontFamily: string | null;
  underline: string | null;
  spacingPt?: number | null;
  baselinePct?: number | null;
  /** `a:rPr@cap`: how the run is cased when drawn, never in the stored text. */
  caps?: 'none' | 'small' | 'all' | null;
}

export interface TextRunSnapshot {
  text: string;
  style: TextStyleSnapshot;
}

/** OOXML `a:pPr@algn` token. */
export type ParagraphAlignment = 'l' | 'ctr' | 'r' | 'just';

export interface ParagraphSnapshot {
  id: string;
  alignment: string | null;
  level: number;
  bulletJson: string | null;
  runs: TextRunSnapshot[];
}

export interface StorySnapshot {
  id: string;
  length: number;
  paragraphs: ParagraphSnapshot[];
}

export interface PptxTextSearchOptions {
  /** Defaults to false. */
  caseSensitive?: boolean;
  /** Maximum matches; unlimited by default. */
  limit?: number;
}

/** Zero-based slide index; story-local UTF-16 offsets. */
export interface PptxTextMatch {
  slideIndex: number;
  slideId: string;
  shapeId: string;
  storyId: string;
  start: number;
  end: number;
  text: string;
}

export type ShapeKind = 'shape' | 'picture' | 'graphicFrame' | 'group';

export interface ColorValue {
  rgb?: string;
  themeColor?: string;
  themeTint?: string;
  themeShade?: string;
  auto?: boolean;
  luminanceModulation?: number;
  luminanceOffset?: number;
  saturationModulation?: number;
  alpha?: number;
}

export interface ShapeGradient {
  type: string;
  angle?: number;
  stops: Array<{ position: number; color: ColorValue }>;
}

/** An authored fill as the deck stores it. */
export interface ShapeFill {
  type: string;
  color?: ColorValue;
  gradient?: ShapeGradient;
}

export interface ShapeLineEnd {
  type: string;
  width: string | null;
  length: string | null;
}

/** An authored outline as the deck stores it; `width` is EMU. */
export interface ShapeOutline {
  width?: number;
  color?: ColorValue;
  gradient?: ShapeGradient;
  style?: string;
  cap?: string;
  join?: string;
  headEnd?: ShapeLineEnd;
  tailEnd?: ShapeLineEnd;
}

export type BlipEffect =
  | { type: 'biLevel'; threshold: number }
  | { type: 'grayscale' }
  | { type: 'luminance'; brightness: number; contrast: number }
  | { type: 'duotone'; shadow: ColorValue | null; highlight: ColorValue | null }
  | { type: 'colorChange'; from: ColorValue | null; to: ColorValue | null; useAlpha?: boolean };

export interface ShapeSnapshot {
  id: string;
  sourceId: number;
  kind: ShapeKind;
  name: string;
  x: number;
  y: number;
  width: number;
  height: number;
  rotationDeg: number;
  flipH: boolean;
  flipV: boolean;
  /** Hides this shape and its descendants; omitted when false. */
  hidden?: boolean;
  geometry: string;
  adjustValues: Record<string, number>;
  placeholder: unknown | null;
  fill: ShapeFill | null;
  resolvedFillColor: string | null;
  outline: ShapeOutline | null;
  resolvedOutlineColor: string | null;
  mediaPartPath: string | null;
  /** Image data added to this session, retained across saves. */
  pendingMedia?: { contentType: string; base64: string } | null;
  blipEffects?: BlipEffect[];
  graphic: unknown | null;
  textStories: StorySnapshot[];
  children: ShapeSnapshot[];
}

export interface SlideSnapshot {
  id: string;
  sourcePartPath: string | null;
  layoutPartPath: string | null;
  name: string | null;
  /** Speaker notes as plain text; absent when the slide has none. */
  notes?: string;
  shapes: ShapeSnapshot[];
}

export interface DeckSnapshot {
  widthEmu: number;
  heightEmu: number;
  slides: SlideSnapshot[];
  commentFlavor?: CommentFlavor;
  comments?: CommentSnapshot[];
}

/** Legacy comments or modern threads. */
export type CommentFlavor = 'legacy' | 'modern';

export interface CommentSnapshot {
  id: string;
  slideId: string;
  author: string;
  initials: string;
  text: string;
  created: string | null;
  xEmu: number;
  yEmu: number;
  /** Set on a reply; names the thread root. Modern decks only. */
  parentId: string | null;
  resolved: boolean;
}

export interface CommentReceipt {
  commentId: string;
  slideId: string;
  parentId: string | null;
  resolved: boolean;
}

export interface SlideReceipt {
  slideId: string;
  fromIndex: number | null;
  toIndex: number | null;
}

export interface ShapeReceipt {
  slideId: string;
  shapeId: string;
  index: number;
}

export interface ShapeZOrderReceipt {
  slideId: string;
  shapeId: string;
  fromIndex: number;
  toIndex: number;
}

export interface ShapeRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface TransformReceipt {
  slideId: string;
  shapeId: string;
  before: ShapeRect;
  after: ShapeRect;
}

export interface TextReceipt {
  storyId: string;
  start: number;
  end: number;
  text: string;
}

export interface ShapeDraft {
  name: string;
  rect: ShapeRect;
  text: string;
  style: TextStyle;
}

export interface PresetShapeDraft {
  name: string;
  geometry: string;
  rect: ShapeRect;
  fill?: string | null;
}

export interface PictureDraft {
  name: string;
  rect: ShapeRect;
  /** The image's MIME type, e.g. `image/png`. */
  contentType: string;
  /** The image bytes, base64-encoded. */
  mediaBase64: string;
}

export interface ShapeStroke {
  color?: string;
  widthPt?: number;
}

export interface ShapeFillReceipt {
  slideId: string;
  shapeId: string;
  before: string | null;
  after: string | null;
}

export interface ShapeStrokeReceipt {
  slideId: string;
  shapeId: string;
  before: ShapeStroke | null;
  after: ShapeStroke | null;
}

export interface ShapeAdjustReceipt {
  slideId: string;
  shapeId: string;
  before: Record<string, number>;
  after: Record<string, number>;
}

export interface HistoryResult {
  applied: boolean;
  snapshot: DeckSnapshot;
}

/** Renderer stage latencies of one profiled slide layout, in ms. */
export interface LayoutProfile {
  scopeMs: number;
  layoutMs: number;
  serializeMs: number;
}

export interface ProfiledLayout {
  layout: SlideDisplayList;
  profile: LayoutProfile;
}

/** Boundary stage latencies of one profiled edit, in ms. */
export interface EditProfile {
  parseMs: number;
  applyMs: number;
  serializeMs: number;
}

export interface HistoryProfile {
  undoMs: number;
  snapshotMs: number;
  serializeMs: number;
}

export interface Profiled<T, P = EditProfile> {
  receipt: T;
  profile: P;
}

export interface PptxFontFace {
  family: string;
  bold?: boolean;
  italic?: boolean;
  bytes: Uint8Array;
}

export type GeometryPathCommand =
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

export type Paint =
  | { kind: 'solid'; color: string }
  | {
      kind: 'gradient';
      gradientType: 'linear' | 'radial' | 'rectangular' | 'path';
      angleDeg?: number;
      stops: Array<{ position: number; color: string }>;
    };

export interface StrokeEnd {
  kind: string;
  width: number;
  length: number;
}

export interface Stroke {
  /** Solid colour or first gradient stop. */
  color: string;
  width: number;
  dashed?: boolean;
  paint?: Paint;
  join?: 'round' | 'bevel' | 'miter';
  headEnd?: StrokeEnd;
  tailEnd?: StrokeEnd;
}

/** An `a:outerShdw`: a blurred copy of the shape's own path, offset and tinted. */
export interface Shadow {
  paths?: Array<{ path: GeometryPathCommand[]; fill: boolean; stroke?: Stroke }>;
  color: string;
  blur?: number;
  dx?: number;
  dy?: number;
  scaleX?: number;
  scaleY?: number;
}

export interface PrimitiveTransform {
  rotationDeg?: number;
  flipH?: boolean;
  flipV?: boolean;
}

interface PrimitiveBase {
  objectId: number;
  shapeId?: string;
  x: number;
  y: number;
  w: number;
  h: number;
  transform?: PrimitiveTransform;
}

export interface ShapePrimitive extends PrimitiveBase {
  kind: 'shape';
  name: string;
  geometry: string;
  path: GeometryPathCommand[];
  /** An unsupported outline or clip was replaced by a rectangle. */
  geometryFallback?: boolean;
  clip?: GeometryPathCommand[];
  evenOdd?: boolean;
  adjustValues?: Record<string, number>;
  fill?: Paint;
  stroke?: Stroke;
  shadow?: Shadow;
}

/** An `a:blip` colour transform, colours already resolved to `#rrggbbaa`. */
export type ImageEffect =
  | { kind: 'biLevel'; threshold: number }
  | { kind: 'grayscale' }
  | { kind: 'luminance'; brightness: number; contrast: number }
  | { kind: 'duotone'; shadow: string; highlight: string }
  | { kind: 'colorChange'; from: string; to: string; useAlpha?: boolean };
export interface ImageCrop {
  left?: number;
  top?: number;
  right?: number;
  bottom?: number;
}

export interface ImagePrimitive extends PrimitiveBase {
  kind: 'image';
  name: string;
  assetId?: string;
  /** Applied to the bitmap in order before it is drawn. */
  effects?: ImageEffect[];
  /** Fraction of the source discarded per edge, from `a:srcRect`. */
  crop?: ImageCrop;
  /** Outline the picture is masked to, when its `spPr` gives it one. */
  path?: GeometryPathCommand[];
  /** The authored mask is unsupported and uses a rectangle fallback. */
  geometryFallback?: boolean;
  stroke?: Stroke;
  shadow?: Shadow;
}

export interface CaretStop {
  position: number;
  x: number;
}

export interface PositionedGlyph {
  glyphId: number;
  cluster: number;
  x: number;
  advance: number;
  xOffset: number;
  yOffset: number;
}

export interface PositionedTextRun {
  text: string;
  start: number;
  end: number;
  x: number;
  width: number;
  fontId: number;
  fontFamily: string;
  fontSizePx: number;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  color: string;
  letterSpacingPx?: number;
  baselineOffsetPx?: number;
  glyphs: PositionedGlyph[];
}

export interface PositionedTextLine {
  x: number;
  y: number;
  width: number;
  height: number;
  baseline: number;
  start: number;
  end: number;
  runs: PositionedTextRun[];
  caretStops: CaretStop[];
}

export interface TextBoxPrimitive extends PrimitiveBase {
  kind: 'textBox';
  storyId?: string;
  anchor: 'top' | 'center' | 'bottom';
  paragraphs: Array<{
    align?: 'left' | 'center' | 'right' | 'justify';
    level: number;
    runs: Array<{
      text: string;
      fontFamily: string;
      fontSizePt: number;
      bold?: boolean;
      italic?: boolean;
      underline?: boolean;
      color: string;
    }>;
  }>;
  lines: PositionedTextLine[];
  overflow?: boolean;
}

export interface PlaceholderPrimitive extends PrimitiveBase {
  kind: 'placeholder';
  name: string;
  label?: string;
}

/** A plotted chart. Its parts paint clipped to the chart rectangle, and
 *  `label` is the screen-reader summary of the whole chart. */
export interface ChartPrimitive extends PrimitiveBase {
  kind: 'chart';
  name: string;
  label: string;
  primitives: SlidePrimitive[];
}

/** A laid-out table. Its cells paint clipped to the table rectangle, and
 *  `label` is the screen-reader summary of the whole table. */
export interface TablePrimitive extends PrimitiveBase {
  kind: 'table';
  name: string;
  label: string;
  primitives: SlidePrimitive[];
}

export type SlidePrimitive =
  | ShapePrimitive
  | ImagePrimitive
  | TextBoxPrimitive
  | PlaceholderPrimitive
  | ChartPrimitive
  | TablePrimitive;

export interface SlideDisplayList {
  contractVersion: number;
  width: number;
  height: number;
  background?: Paint;
  primitives: SlidePrimitive[];
}

export type HitTestResult =
  | { kind: 'shape'; shapeId: string }
  | { kind: 'text'; shapeId: string; storyId: string; position: number };

export interface UpdateEvent {
  origin: 'local' | 'remote';
  update: Uint8Array;
}
