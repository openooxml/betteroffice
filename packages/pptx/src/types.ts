export interface TextStyle {
  bold?: boolean;
  italic?: boolean;
  fontSizePt?: number;
  color?: string;
  fontFamily?: string;
  underline?: string;
}

export type TextStylePatch = TextStyle;

export interface TextStyleSnapshot {
  bold: boolean | null;
  italic: boolean | null;
  fontSizePt: number | null;
  color: string | null;
  fontFamily: string | null;
  underline: string | null;
}

export interface TextRunSnapshot {
  text: string;
  style: TextStyleSnapshot;
}

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

export type ShapeKind = 'shape' | 'picture' | 'graphicFrame' | 'group';

export interface ColorValue {
  rgb?: string;
  themeColor?: string;
  themeTint?: string;
  themeShade?: string;
  auto?: boolean;
}

export interface ShapeFill {
  type: string;
  color?: ColorValue;
}

export interface ShapeOutline {
  width?: number;
  color?: ColorValue;
  style?: string;
  cap?: string;
  join?: string;
}

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
  geometry: string;
  adjustValues: Record<string, number>;
  placeholder: unknown | null;
  fill: ShapeFill | null;
  outline: ShapeOutline | null;
  mediaPartPath: string | null;
  graphic: unknown | null;
  textStories: StorySnapshot[];
  children: ShapeSnapshot[];
}

export interface SlideSnapshot {
  id: string;
  sourcePartPath: string | null;
  layoutPartPath: string | null;
  name: string | null;
  shapes: ShapeSnapshot[];
}

export interface DeckSnapshot {
  widthEmu: number;
  heightEmu: number;
  slides: SlideSnapshot[];
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

export interface Stroke {
  color: string;
  width: number;
  dashed?: boolean;
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
  adjustValues?: Record<string, number>;
  fill?: Paint;
  stroke?: Stroke;
}

export interface ImagePrimitive extends PrimitiveBase {
  kind: 'image';
  name: string;
  assetId?: string;
  stroke?: Stroke;
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

export type SlidePrimitive =
  | ShapePrimitive
  | ImagePrimitive
  | TextBoxPrimitive
  | PlaceholderPrimitive;

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
