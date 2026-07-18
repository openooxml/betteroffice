/**
 * Border Conversion
 *
 * Shared OOXML BorderSpec → layout/pagination BorderStyle conversion, used by
 * paragraph borders, table cell borders, and header/footer borders.
 */

import type { CellBorders, BorderStyle } from '../../pagination/types';
import type { Theme } from '../../../types/document';
import { resolveColor } from '../../../utils/colorResolver';
import { pointsToPixels } from '../../../utils/units';

/**
 * Convert border width from eighths of a point to pixels.
 * OOXML stores border widths in eighths of a point.
 */
function borderWidthToPixels(eighthsOfPoint: number): number {
  // 1 point = 1.333 pixels at 96 DPI
  // eighths of a point: divide by 8 first
  return Math.max(1, Math.round((eighthsOfPoint / 8) * 1.333));
}

// OOXML border style → CSS border-style mapping
const OOXML_TO_CSS_BORDER: Record<string, string> = {
  single: 'solid',
  double: 'double',
  dotted: 'dotted',
  dashed: 'dashed',
  thick: 'solid',
  dashSmallGap: 'dashed',
  dotDash: 'dashed',
  dotDotDash: 'dotted',
  triple: 'double',
  wave: 'solid',
  doubleWave: 'double',
  threeDEmboss: 'ridge',
  threeDEngrave: 'groove',
  outset: 'outset',
  inset: 'inset',
};

/**
 * Convert an OOXML BorderSpec to a layout/pagination BorderStyle.
 * Shared by paragraph borders, cell borders, and header/footer borders.
 */
export function convertBorderSpecToLayout(
  border: {
    style?: string;
    size?: number;
    space?: number;
    color?: { rgb?: string; themeColor?: string; themeTint?: string; themeShade?: string };
  },
  theme?: Theme | null
): BorderStyle | undefined {
  if (!border || !border.style || border.style === 'none' || border.style === 'nil') {
    return undefined;
  }
  const result: BorderStyle = {
    style: OOXML_TO_CSS_BORDER[border.style] || 'solid',
    width: borderWidthToPixels(border.size ?? 0),
    color: border.color
      ? resolveColor(border.color as Parameters<typeof resolveColor>[0], theme)
      : '#000000',
  };
  if (border.space !== undefined) {
    result.space = pointsToPixels(border.space);
  }
  return result;
}

/**
 * Extract cell borders from ProseMirror attributes.
 * Borders are full BorderSpec objects with style/size/color.
 */
export function extractCellBorders(
  attrs: Record<string, unknown>,
  theme?: Theme | null
): CellBorders | undefined {
  const borders = attrs.borders as Record<
    string,
    {
      style?: string;
      size?: number;
      color?: { rgb?: string; themeColor?: string; themeTint?: string; themeShade?: string };
    }
  > | null;

  if (!borders) {
    return undefined;
  }

  const result: CellBorders = {};
  const sides = ['top', 'bottom', 'left', 'right'] as const;

  for (const side of sides) {
    const border = borders[side];
    const converted = border ? convertBorderSpecToLayout(border, theme) : undefined;
    result[side] = converted ?? { width: 0, style: 'none' };
  }

  return Object.keys(result).length > 0 ? result : undefined;
}
