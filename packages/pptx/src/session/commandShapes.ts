import type { ShapeRect, ShapeSnapshot, SlideSnapshot } from '../types';

export const MIN_SHAPE_SIZE_EMU = 76_200;

export function findShape(shapes: ShapeSnapshot[], shapeId: string): ShapeSnapshot | null {
  for (const shape of shapes) {
    if (shape.id === shapeId) return shape;
    const child = findShape(shape.children, shapeId);
    if (child) return child;
  }
  return null;
}

export function findTopLevelShape(slide: SlideSnapshot, shapeId: string): ShapeSnapshot | null {
  for (const shape of slide.shapes) {
    if (shape.id === shapeId || findShape(shape.children, shapeId)) return shape;
  }
  return null;
}

export function effectiveShapeRect(shape: ShapeSnapshot): ShapeRect | null {
  if (shape.width > 0 && shape.height > 0) {
    return { x: shape.x, y: shape.y, width: shape.width, height: shape.height };
  }
  const inherited = shape.inherited;
  if (!inherited || inherited.width <= 0 || inherited.height <= 0) return null;
  return {
    x: inherited.x,
    y: inherited.y,
    width: inherited.width,
    height: inherited.height,
  };
}

export function canMoveShape(shape: ShapeSnapshot): boolean {
  return effectiveShapeRect(shape) !== null;
}

export function canResizeShape(shape: ShapeSnapshot): boolean {
  const degrees = shape.inherited?.rotationDeg ?? shape.rotationDeg;
  const rotation = ((degrees % 360) + 360) % 360;
  const flipH = shape.inherited?.flipH ?? shape.flipH;
  const flipV = shape.inherited?.flipV ?? shape.flipV;
  return canMoveShape(shape) && rotation === 0 && !flipH && !flipV;
}
