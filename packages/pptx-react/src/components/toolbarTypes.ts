import type { ParagraphAlignment } from '@betteroffice/pptx';
import type { TranslationKey } from '@betteroffice/pptx-i18n';

export const SHAPE_PRESETS = [
  { geometry: 'rect', labelKey: 'toolbar.shapes.rect' },
  { geometry: 'roundRect', labelKey: 'toolbar.shapes.roundRect' },
  { geometry: 'ellipse', labelKey: 'toolbar.shapes.ellipse' },
  { geometry: 'triangle', labelKey: 'toolbar.shapes.triangle' },
  { geometry: 'rtTriangle', labelKey: 'toolbar.shapes.rtTriangle' },
  { geometry: 'diamond', labelKey: 'toolbar.shapes.diamond' },
  { geometry: 'parallelogram', labelKey: 'toolbar.shapes.parallelogram' },
  { geometry: 'trapezoid', labelKey: 'toolbar.shapes.trapezoid' },
  { geometry: 'pentagon', labelKey: 'toolbar.shapes.pentagon' },
  { geometry: 'hexagon', labelKey: 'toolbar.shapes.hexagon' },
  { geometry: 'octagon', labelKey: 'toolbar.shapes.octagon' },
  { geometry: 'star5', labelKey: 'toolbar.shapes.star5' },
  { geometry: 'rightArrow', labelKey: 'toolbar.shapes.rightArrow' },
  { geometry: 'leftArrow', labelKey: 'toolbar.shapes.leftArrow' },
  { geometry: 'upArrow', labelKey: 'toolbar.shapes.upArrow' },
  { geometry: 'downArrow', labelKey: 'toolbar.shapes.downArrow' },
  { geometry: 'chevron', labelKey: 'toolbar.shapes.chevron' },
] as const satisfies ReadonlyArray<{ geometry: string; labelKey: TranslationKey }>;

export type PptxShapePreset = (typeof SHAPE_PRESETS)[number]['geometry'];
export type PptxEditorTool = 'select' | 'textBox' | `shape:${PptxShapePreset}`;
export type PptxZoom = number | 'fit';

export interface SelectionFormatting {
  fontFamily?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  textColor?: string;
  align?: ParagraphAlignment;
}

export type FormattingAction =
  | 'bold'
  | 'italic'
  | 'underline'
  | { type: 'fontFamily'; value: string }
  | { type: 'fontSize'; value: number }
  | { type: 'textColor'; value: string }
  | { type: 'align'; value: ParagraphAlignment };

export interface ShapeFormatting {
  geometry?: string;
  fillColor?: string | null;
  strokeColor?: string | null;
  strokeWidthPt?: number | null;
  adjustments?: Record<string, number>;
}

export type ShapeZOrder = 'front' | 'forward' | 'backward' | 'back';

export type ShapeFormattingAction =
  | { type: 'fillColor'; value: string | null }
  | { type: 'strokeColor'; value: string | null }
  | { type: 'strokeWidth'; value: number | null }
  | { type: 'adjust'; name: string; value: number }
  | { type: 'zOrder'; value: ShapeZOrder };

export interface SlideLayoutOption {
  partPath: string | null;
  label?: string;
}

export function shapePresetFromTool(tool: string): PptxShapePreset | null {
  if (!tool.startsWith('shape:')) return null;
  const geometry = tool.slice('shape:'.length);
  return SHAPE_PRESETS.some((preset) => preset.geometry === geometry)
    ? (geometry as PptxShapePreset)
    : null;
}

/** The adjustment a single-value control edits: `adj`, else `adj1`, else the first by name. */
export function primaryAdjustment(
  adjustments: Record<string, number> | undefined
): [string, number] | null {
  if (!adjustments) return null;
  if (adjustments.adj !== undefined) return ['adj', adjustments.adj];
  if (adjustments.adj1 !== undefined) return ['adj1', adjustments.adj1];
  const name = Object.keys(adjustments).sort()[0];
  return name ? [name, adjustments[name]] : null;
}

/** The largest fraction the toolbar offers for an adjustment: half the shape for corner radii. */
export function adjustmentLimit(geometry: string | null | undefined, name: string): number {
  return geometry === 'roundRect' && name === 'adj' ? 0.5 : 1;
}
