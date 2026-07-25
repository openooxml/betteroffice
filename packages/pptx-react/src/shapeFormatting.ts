import type { ShapeSnapshot } from '@betteroffice/pptx';
import type { ShapeFormatting } from './components/Toolbar';

const EMU_PER_POINT = 12_700;

export function shapeFormattingFromShape(shape: ShapeSnapshot | null): ShapeFormatting {
  const fillColor =
    shape?.fill?.type === 'none' ? null : normalizeResolvedColor(shape?.resolvedFillColor);
  const strokeColor = normalizeResolvedColor(shape?.resolvedOutlineColor);
  return {
    geometry: shape?.geometry,
    fillColor,
    strokeColor,
    strokeWidthPt:
      strokeColor && shape?.outline?.width !== undefined
        ? shape.outline.width / EMU_PER_POINT
        : strokeColor
          ? 1
          : null,
    adjustments: shape?.adjustValues ?? {},
  };
}

function normalizeResolvedColor(value: string | null | undefined): string | null {
  if (!value) return null;
  const rgb = value.replace(/^#/, '');
  return /^[0-9a-f]{6}$/i.test(rgb) ? `#${rgb}` : null;
}
