/**
 * DrawingML shapes (`wps:wsp`) and text boxes — preset shape types,
 * fill, outline, shape text body, transform.
 *
 * This module binds the shared ShapeTextBody to docx paragraphs and keeps the
 * docx-level Shape/TextBox wrappers.
 */

import type {
  PresetGeometryPathCommand,
  ShapeFill,
  ShapeOutline,
  ShapeTextBody as ShapeTextBodyBase,
  ShapeType,
} from '@betteroffice/drawingml';
import type { ImageSize, ImagePosition, ImageWrap, ImageTransform } from './image';
import type { Paragraph } from './paragraph';
import type { BlockContent } from './section';
import type { Image } from './image';
import type { Chart } from './chart';
import type { ColorValue, ThemeColorSlot } from '../colors';

export type { ShapeFill, ShapeOutline, ShapeType } from '@betteroffice/drawingml';

/**
 * Text body inside a shape (content typed as docx paragraphs)
 */
export type ShapeTextBody = ShapeTextBodyBase<Paragraph>;

/** Detailed DrawingML paint server retained alongside the legacy shared fill. */
export interface ShapeFillPaint {
  kind?: 'none' | 'solid' | 'gradient' | 'pattern' | 'picture' | 'theme';
  color?: ColorValue;
  themeColor?: ThemeColorSlot;
  gradientType?: 'linear' | 'path' | 'radial' | 'rectangular';
  angle?: number;
  stops?: Array<{ position?: number; color?: ColorValue }>;
  pathShape?: 'circle' | 'rect' | 'shape';
  focusRect?: { left?: number; top?: number; right?: number; bottom?: number };
  rotateWithShape?: boolean;
  patternPreset?: string;
  foregroundColor?: ColorValue;
  backgroundColor?: ColorValue;
  picture?: Image;
  themeRefIndex?: number;
  /** Picture-fill source crop fractions (`a:srcRect`, 0..1; negative = outset). */
  srcRect?: { left?: number; top?: number; right?: number; bottom?: number };
  /** Picture-fill mode: `a:stretch` (default) or `a:tile`. */
  fillMode?: 'stretch' | 'tile';
  /** Picture-fill tile parameters (`a:tile`). Offsets px, scales fractions. */
  tile?: {
    offsetX?: number;
    offsetY?: number;
    scaleX?: number;
    scaleY?: number;
    alignment?: string;
    flip?: 'none' | 'x' | 'y' | 'xy';
  };
  /** Picture-fill stretch target rect fractions (`a:stretch/a:fillRect`, 0..1). */
  stretchRect?: { left?: number; top?: number; right?: number; bottom?: number };
  /** Picture-fill alpha 0..1 (`a:alphaModFix` on the blip). */
  pictureOpacity?: number;
}

/** Detailed DrawingML line contract, including arrows and compound/custom dash. */
export interface ShapeStrokePaint {
  fill?: ShapeFillPaint;
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
  themeRefIndex?: number;
}

/** Common 2-D shape effect. Missing parameters mean the OOXML default. */
export interface ShapeEffect {
  kind?: 'shadow' | 'glow' | 'reflection' | 'softEdge' | 'blur' | 'unknown';
  color?: ColorValue;
  opacity?: number;
  blurRadius?: number;
  distance?: number;
  direction?: number;
  size?: number;
  rawName?: string;
}

/** Full `a:bodyPr` projection used by text-body/autofit layout. */
export interface ShapeTextBodyProperties {
  vertical?:
    | 'horizontal'
    | 'vertical'
    | 'vertical270'
    | 'wordArtVertical'
    | 'eastAsianVertical'
    | 'mongolianVertical';
  rotation?: number;
  upright?: boolean;
  anchor?: 'top' | 'middle' | 'bottom' | 'distributed' | 'justified';
  anchorCenter?: boolean;
  columns?: number;
  columnSpacing?: number;
  wrap?: 'square' | 'none';
  horizontalOverflow?: 'clip' | 'overflow';
  verticalOverflow?: 'clip' | 'ellipsis' | 'overflow';
  margins?: { top?: number; bottom?: number; left?: number; right?: number };
  autoFit?: 'none' | 'normal' | 'shape';
  fontScale?: number;
  lineSpacingReduction?: number;
  fromWordArt?: boolean;
  presetTextWarp?: string;
}

/** Heterogeneous DrawingML group/canvas scene node. */
export interface DrawingSceneNode {
  kind?: 'shape' | 'group' | 'canvas' | 'picture' | 'chart' | 'graphicFrame' | 'contentPart';
  id?: string;
  name?: string;
  shape?: Shape;
  image?: Image;
  chart?: Chart;
  children?: DrawingSceneNode[];
  transform?: {
    offsetX?: number;
    offsetY?: number;
    width?: number;
    height?: number;
    childOffsetX?: number;
    childOffsetY?: number;
    childWidth?: number;
    childHeight?: number;
    rotation?: number;
    flipH?: boolean;
    flipV?: boolean;
  };
  fill?: ShapeFillPaint;
  effects?: ShapeEffect[];
  relationshipId?: string;
}

/** Versioned general DrawingML scene. Undefined version reads as legacy v0. */
export interface DrawingScene {
  version?: number;
  root?: DrawingSceneNode;
  title?: string;
  description?: string;
  decorative?: boolean;
  hidden?: boolean;
}

/**
 * Shape/drawing object (wps:wsp)
 */
export interface Shape {
  type: 'shape';
  /** Shape type preset */
  shapeType: ShapeType;
  /** Unique ID */
  id?: string;
  /** Name */
  name?: string;
  /** Size in EMUs */
  size: ImageSize;
  /** Offset in EMUs relative to the containing drawing/group, if any. */
  offset?: { x: number; y: number };
  /** Position for floating shapes */
  position?: ImagePosition;
  /** Wrap settings */
  wrap?: ImageWrap;
  /** Fill */
  fill?: ShapeFill;
  /** Outline/stroke */
  outline?: ShapeOutline;
  /** Transform */
  transform?: ImageTransform;
  /** Text content inside the shape */
  textBody?: ShapeTextBody;
  /** Normalized custom or resolved shape geometry path */
  geometryPath?: PresetGeometryPathCommand[];
  /** Nested shapes when this shape is used as a drawing/diagram group. */
  children?: Shape[];
  /** Custom geometry points */
  customGeometry?: string;
  /** General heterogeneous group/canvas scene. Undefined = legacy shape tree. */
  scene?: DrawingScene;
  /** Lossless fill details. Undefined = use `fill`. */
  fillPaint?: ShapeFillPaint;
  /** Lossless stroke details. Undefined = use `outline`. */
  strokePaint?: ShapeStrokePaint;
  /** Ordered 2-D effects. Undefined = none. */
  effects?: ShapeEffect[];
  /** Effect ink/wrap extents in EMUs. Undefined = zero. */
  effectExtent?: { top?: number; right?: number; bottom?: number; left?: number };
  /** Full text-body properties. Undefined = use `textBody` defaults. */
  textBodyProperties?: ShapeTextBodyProperties;
  /** Authored accessibility title. Undefined = use name/text. */
  title?: string;
  /** Authored accessibility description. */
  description?: string;
  /** Decorative flag. Undefined = infer false. */
  decorative?: boolean;
  /** Hidden drawing flag. Undefined = visible. */
  hidden?: boolean;
  /** Stable z-order. Undefined = source order. */
  relativeHeight?: number;
}

/**
 * Text box (floating text container)
 */
export interface TextBox {
  type: 'textBox';
  /** Unique ID */
  id?: string;
  /** Size */
  size: ImageSize;
  /** Position */
  position?: ImagePosition;
  /** Wrap settings */
  wrap?: ImageWrap;
  /** Fill */
  fill?: ShapeFill;
  /** Outline */
  outline?: ShapeOutline;
  /** Text content */
  content: Paragraph[];
  /** Internal margins */
  margins?: {
    top?: number;
    bottom?: number;
    left?: number;
    right?: number;
  };
  /** Full block grammar. Undefined = use legacy paragraph-only `content`. */
  blockContent?: BlockContent[];
  /** Full `a:bodyPr` projection. Undefined = legacy margins/top anchor. */
  bodyProperties?: ShapeTextBodyProperties;
}
