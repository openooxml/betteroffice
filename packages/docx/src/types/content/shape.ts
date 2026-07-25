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
} from '../drawingml';
import type {
  DrawingScene,
  ShapeEffect,
  ShapeFillPaint,
  ShapeStrokePaint,
  ShapeTextBodyProperties,
} from '../../layout/drawing';
import type { ImageSize, ImagePosition, ImageWrap, ImageTransform } from './image';
import type { Paragraph } from './paragraph';
import type { BlockContent } from './section';

export type { ShapeFill, ShapeOutline, ShapeType } from '../drawingml';
export type {
  DrawingScene,
  DrawingSceneNode,
  ShapeEffect,
  ShapeFillPaint,
  ShapeStrokePaint,
  ShapeTextBodyProperties,
} from '../../layout/drawing';

/**
 * Text body inside a shape (content typed as docx paragraphs)
 */
export type ShapeTextBody = ShapeTextBodyBase<Paragraph>;

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
