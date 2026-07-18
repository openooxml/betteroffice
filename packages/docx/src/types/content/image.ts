/**
 * Embedded images (`w:drawing` → `pic:pic`): size, wrap, position,
 * transform, padding, crop.
 */

import type { WrapType } from '../../docx/wrapTypes';
import type { ShapeOutline } from './shape';

/** One point in a `wp:wrapPolygon`, in the drawing's authored coordinates. */
export interface ImageWrapPoint {
  x?: number;
  y?: number;
}

/**
 * Image size specification
 */
export interface ImageSize {
  /** Width in EMUs (English Metric Units) */
  width: number;
  /** Height in EMUs */
  height: number;
}

/**
 * Image wrap type for floating images
 */
export interface ImageWrap {
  type: WrapType;
  /** Wrap text direction */
  wrapText?: 'bothSides' | 'left' | 'right' | 'largest';
  /** Distance from text */
  distT?: number;
  distB?: number;
  distL?: number;
  distR?: number;
  /** Tight/through contour. Undefined = rectangular wrap. */
  polygon?: ImageWrapPoint[];
}

/**
 * Position for floating images
 */
export interface ImagePosition {
  /** Use `wp:simplePos` instead of positionH/V. Undefined = false. */
  useSimplePos?: boolean;
  /** `wp:simplePos` coordinates in EMUs. */
  simplePos?: { x?: number; y?: number };
  /** Stable z-order (`wp:anchor/@relativeHeight`). Undefined = source order. */
  relativeHeight?: number;
  /** Behind-text plane (`wp:anchor/@behindDoc`). Undefined = false. */
  behindDoc?: boolean;
  /** Hidden drawing. Undefined = visible. */
  hidden?: boolean;
  /** Locked drawing metadata. Undefined = editable. */
  locked?: boolean;
  /** Horizontal positioning */
  horizontal: {
    relativeTo:
      | 'character'
      | 'column'
      | 'insideMargin'
      | 'leftMargin'
      | 'margin'
      | 'outsideMargin'
      | 'page'
      | 'rightMargin';
    alignment?: 'left' | 'right' | 'center' | 'inside' | 'outside';
    posOffset?: number;
  };
  /** Vertical positioning */
  vertical: {
    relativeTo:
      | 'insideMargin'
      | 'line'
      | 'margin'
      | 'outsideMargin'
      | 'page'
      | 'paragraph'
      | 'topMargin'
      | 'bottomMargin';
    alignment?: 'top' | 'bottom' | 'center' | 'inside' | 'outside';
    posOffset?: number;
  };
}

/**
 * Image transformation
 */
export interface ImageTransform {
  /** Rotation in degrees */
  rotation?: number;
  /** Flip horizontal */
  flipH?: boolean;
  /** Flip vertical */
  flipV?: boolean;
}

/** Post-transform layout footprint, kept separate from the content frame. */
export interface ImageRotationBounds {
  width?: number;
  height?: number;
  offsetX?: number;
  offsetY?: number;
}

/** One ordered DrawingML blip effect. Unknown effects stay named and inert. */
export interface ImageEffect {
  kind?:
    | 'brightnessContrast'
    | 'saturation'
    | 'grayscale'
    | 'biLevel'
    | 'colorChange'
    | 'duotone'
    | 'alpha'
    | 'blur'
    | 'unknown';
  amount?: number;
  threshold?: number;
  colors?: string[];
  rawName?: string;
}

/** Safe hyperlink metadata attached to picture non-visual properties. */
export interface ImageHyperlink {
  href?: string;
  tooltip?: string;
  target?: string;
  history?: boolean;
  docLocation?: string;
}

/**
 * Image padding/margins
 */
export interface ImagePadding {
  top?: number;
  bottom?: number;
  left?: number;
  right?: number;
}

/**
 * Image crop, expressed as fractions of the source image to trim from each
 * edge. OOXML's `<a:srcRect l="10000" t="0" r="5000" b="0"/>` uses units of
 * 1/100000 (so 10000 → 0.1 → 10% trimmed from the left). We store the
 * normalised fraction so both the renderer and the saver can read it
 * directly without re-parsing units.
 */
export interface ImageCrop {
  left?: number;
  top?: number;
  right?: number;
  bottom?: number;
}

/**
 * Embedded image (`w:drawing` with an inline or anchored picture). Carries
 * the relationship-id pointer to the binary in `word/media/`, its
 * resolved data URL (`src`), display dimensions, optional crop /
 * transform / wrap behaviors, and anchor positioning for floating
 * images.
 *
 * See ECMA-376 §20.4 (DrawingML wordprocessingDrawing).
 */
export interface Image {
  type: 'image';
  /** Unique ID */
  id?: string;
  /** Relationship ID for the image data */
  rId: string;
  /** Resolved image data (base64 or blob URL) */
  src?: string;
  /** Image MIME type */
  mimeType?: string;
  /** Original filename */
  filename?: string;
  /** Alt text for accessibility */
  alt?: string;
  /** Title/description */
  title?: string;
  /** Image size */
  size: ImageSize;
  /** Original size before any transforms */
  originalSize?: ImageSize;
  /** Wrap settings */
  wrap: ImageWrap;
  /** Position for floating images */
  position?: ImagePosition;
  /** Image transformations */
  transform?: ImageTransform;
  /** Padding around image */
  padding?: ImagePadding;
  /** Source-image crop (fractional, OOXML `a:srcRect`). */
  crop?: ImageCrop;
  /** Opacity in [0, 1] (OOXML `a:alphaModFix amt`). Undefined = fully opaque. */
  opacity?: number;
  /** Whether this is a decorative image */
  decorative?: boolean;
  /**
   * `wp:anchor layoutInCell` — when true (default), an anchored image inside
   * a table cell is constrained to the cell. When false, the image escapes
   * the cell into the page area. Round-tripped on save.
   */
  layoutInCell?: boolean;
  /**
   * `wp:anchor allowOverlap` — when true (default), anchored objects may
   * overlap; when false, Word repositions them to avoid collisions. We
   * don't currently reposition; we round-trip the flag so saving preserves
   * the author's intent.
   */
  allowOverlap?: boolean;
  /** Hyperlink URL for clickable image */
  hlinkHref?: string;
  /** Image outline/border */
  outline?: ShapeOutline;
  /** Image effects */
  effects?: {
    brightness?: number;
    contrast?: number;
    saturation?: number;
    /** Ordered blip-effect list. Undefined = legacy scalar effects only. */
    ordered?: ImageEffect[];
  };
  /** Rotated layout bbox. Undefined = use untransformed `size`. */
  rotationBounds?: ImageRotationBounds;
  /** Complete picture hyperlink metadata. Undefined = use `hlinkHref`. */
  hyperlink?: ImageHyperlink;
  /** Effect extents in EMUs. Undefined = no extra wrap/ink extent. */
  effectExtent?: ImagePadding;
}
