import type {
  DeckSnapshot,
  ShapeSnapshot,
  SlideDisplayList,
  SlidePrimitive,
  SlideSnapshot,
  TextBoxPrimitive,
} from '@betteroffice/pptx';

export interface SlidePoint {
  x: number;
  y: number;
}

export interface TextPointLocation {
  position: number;
  lineIndex: number;
}

export type HoverTarget = 'text' | 'shape';

export type ResizeHandle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w';

export const RESIZE_HANDLES: readonly ResizeHandle[] = [
  'nw',
  'n',
  'ne',
  'e',
  'se',
  's',
  'sw',
  'w',
];

const BORDER_BAND_SCREEN_PX = 6;
export const MIN_SHAPE_SIZE_EMU = 76_200;

type PrimitiveGeometry = Pick<SlidePrimitive, 'x' | 'y' | 'w' | 'h' | 'transform'>;

/** `f32::to_radians` folds PI/180 to one f32 constant before multiplying. */
const RADIANS_PER_DEGREE = Math.fround(Math.fround(Math.PI) / 180);

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

export function indexShapes(shapes: ShapeSnapshot[]): Map<string, ShapeSnapshot> {
  const index = new Map<string, ShapeSnapshot>();
  const pending = [...shapes];
  while (pending.length > 0) {
    const shape = pending.pop()!;
    index.set(shape.id, shape);
    pending.push(...shape.children);
  }
  return index;
}

export function findTopLevelShape(slide: SlideSnapshot, shapeId: string): ShapeSnapshot | null {
  for (const shape of slide.shapes) {
    if (shape.id === shapeId || findShape(shape.children, shapeId)) return shape;
  }
  return null;
}

export function shapeTargetExists(
  slide: SlideSnapshot | undefined,
  target: { slideId: string; shapeId: string }
): boolean {
  return Boolean(
    slide && slide.id === target.slideId && findShape(slide.shapes, target.shapeId)
  );
}

export function gestureOwnsPointer<T extends { pointerId: number }>(
  gesture: T | null,
  pointerId: number
): gesture is T {
  return gesture?.pointerId === pointerId;
}

export function canMoveShape(shape: ShapeSnapshot): boolean {
  return shape.width > 0 && shape.height > 0;
}

export function canResizeShape(shape: ShapeSnapshot): boolean {
  const rotation = ((shape.rotationDeg % 360) + 360) % 360;
  return canMoveShape(shape) && rotation === 0 && !shape.flipH && !shape.flipV;
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
  return textLocationAtPoint(frame, shapeId, storyId, point)?.position ?? null;
}

export function textLocationAtPoint(
  frame: SlideDisplayList,
  shapeId: string,
  storyId: string,
  point: SlidePoint
): TextPointLocation | null {
  const textBox = frame.primitives.find(
    (primitive): primitive is TextBoxPrimitive =>
      primitive.kind === 'textBox' &&
      primitive.shapeId === shapeId &&
      primitive.storyId === storyId
  );
  if (!textBox || textBox.lines.length === 0) return null;
  const local = localPoint(textBox, point);
  let lineIndex = 0;
  for (let index = 1; index < textBox.lines.length; index += 1) {
    const candidate = textBox.lines[index];
    const nearest = textBox.lines[lineIndex];
    if (
      lineDistance(candidate.y, candidate.height, local.y) <
      lineDistance(nearest.y, nearest.height, local.y)
    ) {
      lineIndex = index;
    }
  }
  const line = textBox.lines[lineIndex];
  const firstCaret = line.caretStops[0];
  if (!firstCaret) return { position: line.start, lineIndex };
  const caret = line.caretStops.reduce(
    (nearest, candidate) =>
      Math.abs(candidate.x - local.x) < Math.abs(nearest.x - local.x) ? candidate : nearest,
    firstCaret
  );
  return { position: caret.position, lineIndex };
}

/** Mirrors the engine's `hitTest`, which resolves text only where a caret can
 *  land — an empty story reads as a plain shape in both. */
export function hoverTargetAtPoint(
  frame: SlideDisplayList,
  point: SlidePoint
): HoverTarget | null {
  // wasm-bindgen narrows hitTest's arguments to f32 before the engine sees them
  const at = { x: Math.fround(point.x), y: Math.fround(point.y) };
  for (let index = frame.primitives.length - 1; index >= 0; index -= 1) {
    const primitive = frame.primitives[index];
    if (!primitive.shapeId) continue;
    const local = containedPoint(primitive, at);
    if (!local) continue;
    return primitive.kind === 'textBox' && primitive.storyId && hasCaretNear(primitive, local.y)
      ? 'text'
      : 'shape';
  }
  return null;
}

/** Undoes the rotate-then-flip the primitive paints with. Mirrors the engine's
 *  `HitRegion::local_point`; the display list is f32, so the arithmetic is too —
 *  an f64 sum puts the far edges up to 1e-5 off, which whole pixels fall into. */
export function localPoint(primitive: PrimitiveGeometry, point: SlidePoint): SlidePoint {
  const transform = primitive.transform;
  if (!transform || (!transform.rotationDeg && !transform.flipH && !transform.flipV)) {
    return point;
  }
  const f32 = Math.fround;
  const centerX = f32(f32(primitive.x) + f32(f32(primitive.w) / 2));
  const centerY = f32(f32(primitive.y) + f32(f32(primitive.h) / 2));
  const angle = f32(f32(transform.rotationDeg ?? 0) * RADIANS_PER_DEGREE);
  const cos = f32(Math.cos(angle));
  const sin = f32(Math.sin(angle));
  const dx = f32(point.x - centerX);
  const dy = f32(point.y - centerY);
  let localX = f32(f32(dx * cos) + f32(dy * sin));
  let localY = f32(f32(dy * cos) - f32(dx * sin));
  if (transform.flipH) localX = -localX;
  if (transform.flipV) localY = -localY;
  return { x: f32(centerX + localX), y: f32(centerY + localY) };
}

/** The far edge of a primitive, as the engine's f32 arithmetic computes it. */
export function farEdge(origin: number, extent: number): number {
  return Math.fround(Math.fround(origin) + Math.fround(extent));
}

/** The local point when it lands inside the primitive, else null. */
function containedPoint(primitive: SlidePrimitive, point: SlidePoint): SlidePoint | null {
  const local = localPoint(primitive, point);
  return local.x >= Math.fround(primitive.x) &&
    local.x <= farEdge(primitive.x, primitive.w) &&
    local.y >= Math.fround(primitive.y) &&
    local.y <= farEdge(primitive.y, primitive.h)
    ? local
    : null;
}

function hasCaretNear(textBox: TextBoxPrimitive, y: number): boolean {
  if (textBox.lines.length === 0) return false;
  const nearest = textBox.lines.reduce((best, candidate) =>
    lineDistance(candidate.y, candidate.height, y) < lineDistance(best.y, best.height, y)
      ? candidate
      : best
  );
  return nearest.caretStops.length > 0;
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

/** Where a handle's grip sits on the selection box. */
export function handleAnchor(bounds: FrameBounds, handle: ResizeHandle): SlidePoint {
  const x = handle.includes('w')
    ? bounds.x
    : handle.includes('e')
      ? bounds.x + bounds.width
      : bounds.x + bounds.width / 2;
  const y = handle.startsWith('n')
    ? bounds.y
    : handle.startsWith('s')
      ? bounds.y + bounds.height
      : bounds.y + bounds.height / 2;
  return { x, y };
}

export function resizeCursor(handle: ResizeHandle): string {
  if (handle === 'nw' || handle === 'se') return 'nwse-resize';
  if (handle === 'ne' || handle === 'sw') return 'nesw-resize';
  return handle === 'n' || handle === 's' ? 'ns-resize' : 'ew-resize';
}

/** The selection box a drag of `delta` on `handle` produces. Edges opposite
 *  the handle stay put, and the box never collapses past `minimum`. */
export function resizedBounds(
  bounds: FrameBounds,
  handle: ResizeHandle,
  delta: SlidePoint,
  minimum: number
): FrameBounds {
  let { x, y, width, height } = bounds;
  if (handle.includes('w')) {
    const dx = Math.min(delta.x, width - minimum);
    x += dx;
    width -= dx;
  } else if (handle.includes('e')) {
    width = Math.max(minimum, width + delta.x);
  }
  if (handle.startsWith('n')) {
    const dy = Math.min(delta.y, height - minimum);
    y += dy;
    height -= dy;
  } else if (handle.startsWith('s')) {
    height = Math.max(minimum, height + delta.y);
  }
  return { x, y, width, height };
}

/** The delta a resize commits. The release event's own position wins, so a
 *  move whose render has not landed cannot commit a stale size; `tracked` is
 *  the fallback for a release the canvas cannot resolve to a slide point. */
export function resizeCommitDelta(
  start: SlidePoint,
  tracked: SlidePoint,
  release: SlidePoint | null
): SlidePoint {
  return release ? { x: release.x - start.x, y: release.y - start.y } : tracked;
}

/** The same resize in EMU, ready for `moveShape` and `resizeShape`. */
export function resizedShapeBox(
  deck: DeckSnapshot,
  frame: SlideDisplayList,
  shape: ShapeSnapshot,
  handle: ResizeHandle,
  delta: SlidePoint
): { x: number; y: number; width: number; height: number } | null {
  if (!canResizeShape(shape)) return null;
  const emu = {
    x: Math.round((delta.x * deck.widthEmu) / frame.width),
    y: Math.round((delta.y * deck.heightEmu) / frame.height),
  };
  const box = resizedBounds(
    { x: shape.x, y: shape.y, width: shape.width, height: shape.height },
    handle,
    emu,
    MIN_SHAPE_SIZE_EMU
  );
  return { x: box.x, y: box.y, width: box.width, height: box.height };
}

export function resizedShapeBounds(
  deck: DeckSnapshot,
  frame: SlideDisplayList,
  shape: ShapeSnapshot,
  handle: ResizeHandle,
  delta: SlidePoint
): FrameBounds | null {
  const box = resizedShapeBox(deck, frame, shape, handle, delta);
  if (!box || deck.widthEmu <= 0 || deck.heightEmu <= 0) return null;
  return {
    x: (box.x * frame.width) / deck.widthEmu,
    y: (box.y * frame.height) / deck.heightEmu,
    width: (box.width * frame.width) / deck.widthEmu,
    height: (box.height * frame.height) / deck.heightEmu,
  };
}

/** What a pointer gesture acts on. This is the engine's hit with one addition:
 *  the band along a text box's edge reads as the shape, so the border grabs the
 *  box while its middle types — the split PowerPoint draws. `hitTest` has no
 *  such band, and `hoverTargetAtPoint` stays an exact mirror of it. */
export function pointerTargetAtPoint(
  frame: SlideDisplayList,
  point: SlidePoint,
  scale = 1
): HoverTarget | null {
  const target = hoverTargetAtPoint(frame, point);
  if (target !== 'text') return target;
  const at = { x: Math.fround(point.x), y: Math.fround(point.y) };
  for (let index = frame.primitives.length - 1; index >= 0; index -= 1) {
    const primitive = frame.primitives[index];
    if (!primitive.shapeId) continue;
    const local = containedPoint(primitive, at);
    if (!local) continue;
    return withinBorderBand(primitive, local, scale) ? 'shape' : 'text';
  }
  return target;
}

function withinBorderBand(
  primitive: PrimitiveGeometry,
  point: SlidePoint,
  scale: number
): boolean {
  const normalizedScale = Number.isFinite(scale) && scale > 0 ? scale : 1;
  const band = Math.min(
    BORDER_BAND_SCREEN_PX / normalizedScale,
    primitive.w / 3,
    primitive.h / 3
  );
  return (
    point.x - primitive.x <= band ||
    primitive.x + primitive.w - point.x <= band ||
    point.y - primitive.y <= band ||
    primitive.y + primitive.h - point.y <= band
  );
}
