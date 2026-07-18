/**
 * Color & Styling Primitives
 *
 * Basic types used throughout OOXML for colors, borders, and shading.
 * The color primitives (ColorValue, ThemeColorSlot) are shared across OOXML
 * formats; the WordprocessingML border/shading types stay here.
 */

import type { ColorValue } from '@betteroffice/drawingml';

export type { ColorValue, ThemeColorSlot } from '@betteroffice/drawingml';

/**
 * One side of a border — style, color, width, spacing. Used by paragraph
 * borders, table borders (per-cell or whole-table), and page borders.
 * `size` is in eighths of a point (Word's wire format); `space` is in
 * points.
 *
 * See ECMA-376 §17.18.2 (`ST_Border`).
 */
export interface BorderSpec {
  /** Border style */
  style:
    | 'none'
    | 'single'
    | 'double'
    | 'dotted'
    | 'dashed'
    | 'thick'
    | 'triple'
    | 'thinThickSmallGap'
    | 'thickThinSmallGap'
    | 'thinThickMediumGap'
    | 'thickThinMediumGap'
    | 'thinThickLargeGap'
    | 'thickThinLargeGap'
    | 'wave'
    | 'doubleWave'
    | 'dashSmallGap'
    | 'dashDotStroked'
    | 'threeDEmboss'
    | 'threeDEngrave'
    | 'outset'
    | 'inset'
    | 'nil';
  /** Color of the border */
  color?: ColorValue;
  /** Width in eighths of a point (1/8 pt) */
  size?: number;
  /** Spacing from text in points */
  space?: number;
  /** Shadow effect */
  shadow?: boolean;
  /** Frame effect */
  frame?: boolean;
  /**
   * Relationship id for relationship-backed border art. Undefined means a
   * normal line border; renderers must not fetch external relationships.
   */
  relationshipId?: string;
  /** Optional authored art/corner token retained when `style` is approximated. */
  artName?: string;
}

/**
 * Cell/paragraph/run shading — Word's combined "fill + pattern overlay"
 * model. `fill` is the solid background; `color` is the pattern overlay
 * drawn on top; `pattern` selects the pattern type (defaults to
 * `'clear'` = solid `fill`, no pattern).
 *
 * See ECMA-376 §17.4.32 (`CT_Shd`).
 */
export interface ShadingProperties {
  /** Pattern fill color */
  color?: ColorValue;
  /** Background fill color */
  fill?: ColorValue;
  /** Shading pattern type */
  pattern?:
    | 'clear'
    | 'solid'
    | 'horzStripe'
    | 'vertStripe'
    | 'reverseDiagStripe'
    | 'diagStripe'
    | 'horzCross'
    | 'diagCross'
    | 'thinHorzStripe'
    | 'thinVertStripe'
    | 'thinReverseDiagStripe'
    | 'thinDiagStripe'
    | 'thinHorzCross'
    | 'thinDiagCross'
    | 'pct5'
    | 'pct10'
    | 'pct12'
    | 'pct15'
    | 'pct20'
    | 'pct25'
    | 'pct30'
    | 'pct35'
    | 'pct37'
    | 'pct40'
    | 'pct45'
    | 'pct50'
    | 'pct55'
    | 'pct60'
    | 'pct62'
    | 'pct65'
    | 'pct70'
    | 'pct75'
    | 'pct80'
    | 'pct85'
    | 'pct87'
    | 'pct90'
    | 'pct95'
    | 'nil';
}
