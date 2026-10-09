import { canvasPointToModel, modelPointToCanvas } from '@betteroffice/vsdx';
import type { ModelPoint, PageDisplayList, ShapeSnapshot } from '@betteroffice/vsdx';
import { numericCellValue } from './components/ribbon/commands';

export const GRID_SPACING_IN = 0.25;
export const SNAP_THRESHOLD_CSS = 6;
export const GUIDE_STROKE = '#e81123';

export interface SnapTargets { x: number[]; y: number[]; }

export interface SnapContext {
  zoom: number;
  gridSpacing?: number;
  snapEnabled: boolean;
  suppressed: boolean;
  isRotate: boolean;
  targets: SnapTargets;
}

export interface SnapResult { point: ModelPoint; guides: SnapTargets; }

/** Screen-pixel threshold expressed in model inches at the given zoom. */
export const snapThresholdModel = (zoom: number): number => {
  const safe = Number.isFinite(zoom) && zoom > 0 ? zoom : 1;
  return SNAP_THRESHOLD_CSS / (96 * safe);
};

interface ShapeBox { x0: number; y0: number; x1: number; y1: number; }

const shapeBox = (shape: ShapeSnapshot): ShapeBox | null => {
  try {
    const width = numericCellValue(shape, 'Width');
    const height = numericCellValue(shape, 'Height');
    const pinX = numericCellValue(shape, 'PinX');
    const pinY = numericCellValue(shape, 'PinY');
    const locPinX = numericCellValue(shape, 'LocPinX', width / 2);
    const locPinY = numericCellValue(shape, 'LocPinY', height / 2);
    if (![width, height, pinX, pinY, locPinX, locPinY].every(Number.isFinite)) return null;
    const x0 = pinX - locPinX;
    const y0 = pinY - locPinY;
    return { x0, y0, x1: x0 + width, y1: y0 + height };
  } catch {
    return null;
  }
};

/** Page-space edges and centres of top-level shapes except the dragged one. */
export const collectSnapTargets = (shapes: readonly ShapeSnapshot[], excludeId: string): SnapTargets => {
  const x: number[] = [];
  const y: number[] = [];
  for (const shape of shapes) {
    if (shape.id === excludeId) continue;
    const box = shapeBox(shape);
    if (!box) continue;
    x.push(box.x0, (box.x0 + box.x1) / 2, box.x1);
    y.push(box.y0, (box.y0 + box.y1) / 2, box.y1);
  }
  return { x, y };
};

interface AxisSnap { shift: number; guide: number | null; }

const nearestGrid = (line: number, spacing: number): number => Math.round(line / spacing) * spacing;

const nearestWithin = (line: number, candidates: readonly number[], threshold: number): number | null => {
  let best: number | null = null;
  let bestDistance = threshold + 1e-12;
  for (const candidate of candidates) {
    const distance = Math.abs(line - candidate);
    if (distance <= threshold && distance < bestDistance) {
      best = candidate;
      bestDistance = distance;
    }
  }
  return best;
};

const snapAxis = (lines: readonly number[], targets: readonly number[], spacing: number, threshold: number): AxisSnap => {
  let shape: { distance: number; shift: number; guide: number } | null = null;
  let grid: { distance: number; shift: number } | null = null;
  for (const line of lines) {
    const hit = nearestWithin(line, targets, threshold);
    if (hit !== null) {
      const distance = Math.abs(line - hit);
      if (!shape || distance < shape.distance) shape = { distance, shift: hit - line, guide: hit };
    }
    if (spacing > 0) {
      const crossed = nearestGrid(line, spacing);
      const distance = Math.abs(line - crossed);
      if (distance <= threshold && (!grid || distance < grid.distance)) grid = { distance, shift: crossed - line };
    }
  }
  if (shape && (!grid || shape.distance <= grid.distance)) return { shift: shape.shift, guide: shape.guide };
  if (grid) return { shift: grid.shift, guide: null };
  return { shift: 0, guide: null };
};

const dragBox = (pin: ModelPoint, locPin: ModelPoint | undefined, size: { width: number; height: number }): ShapeBox => {
  const lx = locPin?.x ?? size.width / 2;
  const ly = locPin?.y ?? size.height / 2;
  const x0 = pin.x - lx;
  const y0 = pin.y - ly;
  return { x0, y0, x1: x0 + size.width, y1: y0 + size.height };
};

/** Snaps a drag release point to grid and shape targets; preview and commit share it. */
export const snapRelease = (
  start: { model: ModelPoint; pin: ModelPoint; locPin?: ModelPoint; size: { width: number; height: number } },
  release: ModelPoint,
  context: SnapContext,
): SnapResult => {
  const idle: SnapResult = { point: release, guides: { x: [], y: [] } };
  if (!context.snapEnabled || context.suppressed || context.isRotate) return idle;
  const threshold = snapThresholdModel(context.zoom);
  if (!(threshold > 0)) return idle;
  const spacing = context.gridSpacing ?? GRID_SPACING_IN;
  const box = dragBox(start.pin, start.locPin, start.size);
  const dx = release.x - start.model.x;
  const dy = release.y - start.model.y;
  const linesX = [box.x0 + dx, (box.x0 + box.x1) / 2 + dx, box.x1 + dx];
  const linesY = [box.y0 + dy, (box.y0 + box.y1) / 2 + dy, box.y1 + dy];
  const snappedX = snapAxis(linesX, context.targets.x, spacing, threshold);
  const snappedY = snapAxis(linesY, context.targets.y, spacing, threshold);
  const guides: SnapTargets = { x: snappedX.guide !== null ? [snappedX.guide] : [], y: snappedY.guide !== null ? [snappedY.guide] : [] };
  if (snappedX.shift === 0 && snappedY.shift === 0) return { point: release, guides };
  return { point: { x: release.x + snappedX.shift, y: release.y + snappedY.shift }, guides };
};

/** Page model bounds derived from the canvas corners. */
export const pageModelBounds = (frame: PageDisplayList): { minX: number; minY: number; maxX: number; maxY: number } => {
  const corners = [
    canvasPointToModel(frame.paintTransform, 0, 0),
    canvasPointToModel(frame.paintTransform, frame.width, 0),
    canvasPointToModel(frame.paintTransform, 0, frame.height),
    canvasPointToModel(frame.paintTransform, frame.width, frame.height),
  ];
  const xs = corners.map((point) => point.x);
  const ys = corners.map((point) => point.y);
  return { minX: Math.min(...xs), minY: Math.min(...ys), maxX: Math.max(...xs), maxY: Math.max(...ys) };
};

/** Paints the page grid on the overlay canvas, under selection and guides. */
export const paintGrid = (
  context: CanvasRenderingContext2D,
  frame: PageDisplayList,
  dpr: number,
  scale: number,
  spacingIn = GRID_SPACING_IN,
): void => {
  const zoom = Number.isFinite(scale) && scale > 0 ? scale : 1;
  if (!(spacingIn > 0)) return;
  if (spacingIn * 96 * zoom < 6) return;
  const bounds = pageModelBounds(frame);
  context.save();
  try {
    context.setTransform(dpr * zoom, 0, 0, dpr * zoom, 0, 0);
    context.lineWidth = 1 / zoom;
    const minor = 'rgba(110, 130, 150, 0.28)';
    const major = 'rgba(110, 130, 150, 0.55)';
    for (let gx = Math.ceil(bounds.minX / spacingIn) * spacingIn; gx <= bounds.maxX + 1e-9; gx += spacingIn) {
      const majorLine = Math.abs(gx - Math.round(gx)) < 1e-9;
      const a = modelPointToCanvas(frame.paintTransform, gx, bounds.minY);
      const b = modelPointToCanvas(frame.paintTransform, gx, bounds.maxY);
      context.strokeStyle = majorLine ? major : minor;
      context.beginPath();
      context.moveTo(a.x, a.y);
      context.lineTo(b.x, b.y);
      context.stroke();
    }
    for (let gy = Math.ceil(bounds.minY / spacingIn) * spacingIn; gy <= bounds.maxY + 1e-9; gy += spacingIn) {
      const majorLine = Math.abs(gy - Math.round(gy)) < 1e-9;
      const a = modelPointToCanvas(frame.paintTransform, bounds.minX, gy);
      const b = modelPointToCanvas(frame.paintTransform, bounds.maxX, gy);
      context.strokeStyle = majorLine ? major : minor;
      context.beginPath();
      context.moveTo(a.x, a.y);
      context.lineTo(b.x, b.y);
      context.stroke();
    }
  } finally {
    context.restore();
  }
};

/** Paints thin alignment guides for shape-to-shape snaps on the overlay canvas. */
export const paintSmartGuides = (
  context: CanvasRenderingContext2D,
  frame: PageDisplayList,
  dpr: number,
  scale: number,
  guides: SnapTargets,
): void => {
  if (!guides.x.length && !guides.y.length) return;
  const zoom = Number.isFinite(scale) && scale > 0 ? scale : 1;
  const bounds = pageModelBounds(frame);
  context.save();
  try {
    context.setTransform(dpr * zoom, 0, 0, dpr * zoom, 0, 0);
    context.strokeStyle = GUIDE_STROKE;
    context.lineWidth = 1 / zoom;
    for (const gx of guides.x) {
      const a = modelPointToCanvas(frame.paintTransform, gx, bounds.minY);
      const b = modelPointToCanvas(frame.paintTransform, gx, bounds.maxY);
      context.beginPath();
      context.moveTo(a.x, a.y);
      context.lineTo(b.x, b.y);
      context.stroke();
    }
    for (const gy of guides.y) {
      const a = modelPointToCanvas(frame.paintTransform, bounds.minX, gy);
      const b = modelPointToCanvas(frame.paintTransform, bounds.maxX, gy);
      context.beginPath();
      context.moveTo(a.x, a.y);
      context.lineTo(b.x, b.y);
      context.stroke();
    }
  } finally {
    context.restore();
  }
};

