import type {
  DeckSnapshot,
  ShapeSnapshot,
  SlideDisplayList,
  SlideSnapshot,
  TextBoxPrimitive,
} from '@betteroffice/pptx';

export interface SlidePoint {
  x: number;
  y: number;
}

export interface FrameBounds {
  x: number;
  y: number;
  width: number;
  height: number;
}

interface ClientRect {
  left: number;
  top: number;
  width: number;
  height: number;
}

export function slidePoint(
  rect: ClientRect,
  frame: Pick<SlideDisplayList, 'width' | 'height'>,
  clientX: number,
  clientY: number
): SlidePoint | null {
  if (
    !Number.isFinite(rect.width) ||
    !Number.isFinite(rect.height) ||
    rect.width <= 0 ||
    rect.height <= 0
  ) {
    return null;
  }
  const x = ((clientX - rect.left) * frame.width) / rect.width;
  const y = ((clientY - rect.top) * frame.height) / rect.height;
  return Number.isFinite(x) && Number.isFinite(y) ? { x, y } : null;
}

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

export function canMoveShape(shape: ShapeSnapshot): boolean {
  return shape.width > 0 && shape.height > 0;
}

export function frameBoundsForShape(
  deck: DeckSnapshot,
  frame: SlideDisplayList,
  shape: ShapeSnapshot
): FrameBounds | null {
  const shapeIds = new Set<string>();
  collectShapeIds(shape, shapeIds);
  const primitives = frame.primitives.filter(
    (primitive) => primitive.shapeId && shapeIds.has(primitive.shapeId)
  );
  if (primitives.length === 0) {
    if (
      shape.width <= 0 ||
      shape.height <= 0 ||
      deck.widthEmu <= 0 ||
      deck.heightEmu <= 0
    ) {
      return null;
    }
    return {
      x: (shape.x * frame.width) / deck.widthEmu,
      y: (shape.y * frame.height) / deck.heightEmu,
      width: (shape.width * frame.width) / deck.widthEmu,
      height: (shape.height * frame.height) / deck.heightEmu,
    };
  }
  const bounds = primitives.map((primitive) => {
    const angle = ((primitive.transform?.rotationDeg ?? 0) * Math.PI) / 180;
    const width = Math.abs(primitive.w * Math.cos(angle)) + Math.abs(primitive.h * Math.sin(angle));
    const height = Math.abs(primitive.w * Math.sin(angle)) + Math.abs(primitive.h * Math.cos(angle));
    return {
      x: primitive.x + (primitive.w - width) / 2,
      y: primitive.y + (primitive.h - height) / 2,
      width,
      height,
    };
  });
  const left = Math.min(...bounds.map((bound) => bound.x));
  const top = Math.min(...bounds.map((bound) => bound.y));
  const right = Math.max(...bounds.map((bound) => bound.x + bound.width));
  const bottom = Math.max(...bounds.map((bound) => bound.y + bound.height));
  return { x: left, y: top, width: right - left, height: bottom - top };
}

export function textPositionAtPoint(
  frame: SlideDisplayList,
  shapeId: string,
  storyId: string,
  point: SlidePoint
): number | null {
  const textBox = frame.primitives.find(
    (primitive): primitive is TextBoxPrimitive =>
      primitive.kind === 'textBox' &&
      primitive.shapeId === shapeId &&
      primitive.storyId === storyId
  );
  if (!textBox || textBox.lines.length === 0) return null;
  const line = textBox.lines.reduce((nearest, candidate) =>
    lineDistance(candidate.y, candidate.height, point.y) <
    lineDistance(nearest.y, nearest.height, point.y)
      ? candidate
      : nearest
  );
  const firstCaret = line.caretStops[0];
  if (!firstCaret) return line.start;
  const caret = line.caretStops.reduce(
    (nearest, candidate) =>
      Math.abs(candidate.x - point.x) < Math.abs(nearest.x - point.x) ? candidate : nearest,
    firstCaret
  );
  return caret.position;
}

export function movedShapePosition(
  deck: DeckSnapshot,
  frame: SlideDisplayList,
  shape: ShapeSnapshot,
  delta: SlidePoint
): Pick<ShapeSnapshot, 'x' | 'y'> {
  return {
    x: shape.x + Math.round((delta.x * deck.widthEmu) / frame.width),
    y: shape.y + Math.round((delta.y * deck.heightEmu) / frame.height),
  };
}

export function passedDragThreshold(
  startX: number,
  startY: number,
  clientX: number,
  clientY: number,
  threshold = 4
): boolean {
  return Math.hypot(clientX - startX, clientY - startY) >= threshold;
}

function collectShapeIds(shape: ShapeSnapshot, ids: Set<string>): void {
  ids.add(shape.id);
  for (const child of shape.children) collectShapeIds(child, ids);
}

function lineDistance(y: number, height: number, pointY: number): number {
  if (pointY < y) return y - pointY;
  if (pointY > y + height) return pointY - y - height;
  return 0;
}
