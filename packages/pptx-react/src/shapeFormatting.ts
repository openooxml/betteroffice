import type { ColorValue, ShapeSnapshot } from '@betteroffice/pptx';
import type { ShapeFormatting } from './components/Toolbar';

const EMU_PER_POINT = 12_700;
const DEFAULT_THEME_COLORS: Record<string, string> = {
  dk1: '000000',
  text1: '000000',
  lt1: 'FFFFFF',
  background1: 'FFFFFF',
  dk2: '44546A',
  text2: '44546A',
  lt2: 'E7E6E6',
  background2: 'E7E6E6',
  accent1: '4472C4',
  accent2: 'ED7D31',
  accent3: 'A5A5A5',
  accent4: 'FFC000',
  accent5: '5B9BD5',
  accent6: '70AD47',
  hlink: '0563C1',
  folHlink: '954F72',
};

export function shapeFormattingFromShape(shape: ShapeSnapshot | null): ShapeFormatting {
  const fillColor =
    shape?.fill?.type === 'none' ? null : colorValueToHex(shape?.fill?.color) ?? null;
  const strokeColor = colorValueToHex(shape?.outline?.color) ?? null;
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

function colorValueToHex(color: ColorValue | undefined): string | null {
  const value = color?.rgb ?? (color?.themeColor ? DEFAULT_THEME_COLORS[color.themeColor] : null);
  if (!value) return null;
  const rgb = value.replace(/^#/, '');
  return /^[0-9a-f]{6}$/i.test(rgb) ? `#${rgb}` : null;
}
