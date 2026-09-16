import { canvasPointToModel, modelPointToCanvas } from '@betteroffice/vsdx';
import type { Affine, ModelPoint, PageDisplayList, PagePrimitive, TextBoxPrimitive } from '@betteroffice/vsdx';
export type ResizeHandle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w';
export const RESIZE_HANDLES: readonly ResizeHandle[] = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];
export type RotateHandle = 'rotate';
interface FrameBounds { x: number; y: number; width: number; height: number; }
export interface DragStart { canvas: ModelPoint; model: ModelPoint; resize: boolean; handle?: ResizeHandle; rotate?: boolean; pin: ModelPoint; locPin?: ModelPoint; locPinAtSize?: (width: number, height: number) => ModelPoint; size: { width: number; height: number }; parentTransforms?: readonly Affine[]; angle?: number; flipX?: boolean; flipY?: boolean; pointerId?: number; startX?: number; startY?: number; thresholdPassed?: boolean; }
const MIN_SHAPE_INCHES = 0.01;
/** Fluent 2 colorNeutralStrokeAccessible. */
export const SELECTION_STROKE = '#616161';
export const SELECTION_HANDLE_FILL = '#ffffff';
export const SELECTION_HANDLE_CSS = 7;
export const SELECTION_ROTATE_RADIUS_CSS = 5;
export const SELECTION_ROTATE_OFFSET_CSS = 18;
export const HANDLE_HIT_TOLERANCE_CSS = 6;
export const ROTATION_SNAP_STEP = Math.PI / 12;
export const CANVAS_KEYBOARD_DPI = 96;
export const CANVAS_KEYBOARD_NUDGE_MULTIPLIER = 10;
export type CanvasKeyboardIntent = { kind: 'undo' } | { kind: 'redo' } | { kind: 'delete' } | { kind: 'escape' } | { kind: 'nudge'; dx: number; dy: number };
export interface CanvasKeyboardEventLike { key: string; ctrlKey?: boolean; metaKey?: boolean; shiftKey?: boolean; altKey?: boolean; target?: unknown; }
/** One screen pixel in model inches at the given zoom. */
export const keyboardNudgeStep = (zoom: number): number => {
  const safe = Number.isFinite(zoom) && zoom > 0 ? zoom : 1;
  return 1 / (CANVAS_KEYBOARD_DPI * safe);
};
export const isEditableKeyboardTarget = (target: unknown): boolean => {
  if (!target || typeof target !== 'object') return false;
  const element = target as { tagName?: unknown; isContentEditable?: unknown; contentEditable?: unknown; closest?: unknown };
  const tag = typeof element.tagName === 'string' ? element.tagName.toUpperCase() : '';
  if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT') return true;
  if (element.isContentEditable === true) return true;
  if (typeof element.contentEditable === 'string' && element.contentEditable.toLowerCase() === 'true') return true;
  if (typeof element.closest === 'function') {
    try {
      const editable = (element.closest as (selector: string) => unknown)('[contenteditable="true"], input, textarea, select');
      if (editable) return true;
    } catch { void 0; }
  }
  return false;
};
/** Pure key-to-intent mapping for the editor canvas. Y is up, so ArrowUp yields +dy. */
export const canvasKeyboardIntent = (event: CanvasKeyboardEventLike, zoom: number): CanvasKeyboardIntent | null => {
  if (isEditableKeyboardTarget(event.target)) return null;
  const ctrl = Boolean(event.ctrlKey);
  const meta = Boolean(event.metaKey);
  const shift = Boolean(event.shiftKey);
  const alt = Boolean(event.altKey);
  const mod = ctrl || meta;
  const key = event.key;
  if (key === 'Escape') return mod || alt ? null : { kind: 'escape' };
  if (mod && !alt) {
    const lower = key.toLowerCase();
    if (lower === 'z' && !shift) return { kind: 'undo' };
    if (lower === 'y' && !shift) return { kind: 'redo' };
    if (lower === 'z' && shift) return { kind: 'redo' };
    return null;
  }
  if (mod || alt) return null;
  if (key === 'Delete' || key === 'Backspace') return { kind: 'delete' };
  const step = keyboardNudgeStep(zoom) * (shift ? CANVAS_KEYBOARD_NUDGE_MULTIPLIER : 1);
  if (key === 'ArrowLeft') return { kind: 'nudge', dx: -step, dy: 0 };
  if (key === 'ArrowRight') return { kind: 'nudge', dx: step, dy: 0 };
  if (key === 'ArrowUp') return { kind: 'nudge', dx: 0, dy: step };
  if (key === 'ArrowDown') return { kind: 'nudge', dx: 0, dy: -step };
  return null;
};
export const passedDragThreshold = (startX: number, startY: number, clientX: number, clientY: number, threshold = 4): boolean => Math.hypot(clientX - startX, clientY - startY) >= threshold;
export const resizeCursor = (handle: ResizeHandle): string => {
  if (handle === 'nw' || handle === 'se') return 'nwse-resize';
  if (handle === 'ne' || handle === 'sw') return 'nesw-resize';
  return handle === 'n' || handle === 's' ? 'ns-resize' : 'ew-resize';
};
export const resizedBounds = (bounds: FrameBounds, handle: ResizeHandle, delta: ModelPoint, minimum: number): FrameBounds => {
  let { x, y, width, height } = bounds;
  if (handle.includes('w')) { const dx = Math.min(delta.x, width - minimum); x += dx; width -= dx; }
  else if (handle.includes('e')) width = Math.max(minimum, width + delta.x);
  if (handle.startsWith('n')) { const dy = Math.min(delta.y, height - minimum); y += dy; height -= dy; }
  else if (handle.startsWith('s')) height = Math.max(minimum, height + delta.y);
  return { x, y, width, height };
};
const yDownHandle = (handle: ResizeHandle): ResizeHandle => {
  if (handle === 'n') return 's';
  if (handle === 's') return 'n';
  if (handle === 'nw') return 'sw';
  if (handle === 'sw') return 'nw';
  if (handle === 'ne') return 'se';
  if (handle === 'se') return 'ne';
  return handle;
};
export const snapRotationAngle = (angle: number, snap: boolean): number => snap ? Math.round(angle / ROTATION_SNAP_STEP) * ROTATION_SNAP_STEP : angle;
/** The engine refuses a LocPin it cannot evaluate; the stored value is what the renderer falls back to. */
const probedLocPin = (start: DragStart, size: { width: number; height: number }): ModelPoint | undefined => {
  if (!(size.width > 0) || !(size.height > 0)) return undefined;
  try {
    const probed = start.locPinAtSize?.(size.width, size.height);
    return probed && Number.isFinite(probed.x) && Number.isFinite(probed.y) ? probed : undefined;
  } catch { return undefined; }
};
const locPinInches = (start: DragStart, size = start.size): ModelPoint => probedLocPin(start, size) ?? {
  x: start.locPin?.x ?? size.width / 2,
  y: start.locPin?.y ?? size.height / 2,
};
export const resolveDragGeometry = (start: DragStart, release: ModelPoint): { x: number; y: number; width: number; height: number } => {
  const toParent = (point: ModelPoint) => (start.parentTransforms ?? []).reduce((local, transform) => canvasPointToModel(transform, local.x, local.y), point);
  const origin = toParent(start.model);
  const end = toParent(release);
  const deltaX = end.x - origin.x;
  const deltaY = end.y - origin.y;
  if (start.rotate) return { x: start.pin.x, y: start.pin.y, width: start.size.width, height: start.size.height };
  if (start.handle) {
    const cos = Math.cos(start.angle ?? 0), sin = Math.sin(start.angle ?? 0);
    const flipSignX = start.flipX ? -1 : 1;
    const flipSignY = start.flipY ? -1 : 1;
    const localX = (cos * deltaX + sin * deltaY) * flipSignX;
    const localY = (-sin * deltaX + cos * deltaY) * flipSignY;
    const locPin = locPinInches(start);
    const visualHandle: ResizeHandle = start.handle;
    const localHandle: ResizeHandle = ((): ResizeHandle => {
      let mapped: ResizeHandle = visualHandle;
      if (start.flipX) {
        if (mapped.includes('w')) mapped = mapped.replace('w', 'e') as ResizeHandle;
        else if (mapped.includes('e')) mapped = mapped.replace('e', 'w') as ResizeHandle;
      }
      if (start.flipY) {
        if (mapped.startsWith('n')) mapped = (`s${mapped.slice(1)}`) as ResizeHandle;
        else if (mapped.startsWith('s')) mapped = (`n${mapped.slice(1)}`) as ResizeHandle;
      }
      return mapped;
    })();
    const boxX = start.pin.x - locPin.x;
    const boxY = start.pin.y - locPin.y;
    const box: FrameBounds = { x: boxX, y: boxY, width: start.size.width, height: start.size.height };
    const next = resizedBounds(box, yDownHandle(localHandle), { x: localX, y: localY }, MIN_SHAPE_INCHES);
    const { x: newLocPinX, y: newLocPinY } = locPinInches(start, next);
    const shiftX = (next.x + newLocPinX) - (boxX + locPin.x);
    const shiftY = (next.y + newLocPinY) - (boxY + locPin.y);
    const unflippedX = shiftX * flipSignX;
    const unflippedY = shiftY * flipSignY;
    return { x: start.pin.x + cos * unflippedX - sin * unflippedY, y: start.pin.y + sin * unflippedX + cos * unflippedY, width: next.width, height: next.height };
  }
  if (start.resize) {
    const cos = Math.cos(start.angle ?? 0), sin = Math.sin(start.angle ?? 0);
    const widthDelta = (cos * deltaX + sin * deltaY) * (start.flipX ? -1 : 1);
    const heightDelta = (-sin * deltaX + cos * deltaY) * (start.flipY ? -1 : 1);
    return { x: start.pin.x, y: start.pin.y, width: Math.max(MIN_SHAPE_INCHES, start.size.width + widthDelta), height: Math.max(MIN_SHAPE_INCHES, start.size.height + heightDelta) };
  }
  return { x: start.pin.x + deltaX, y: start.pin.y + deltaY, width: start.size.width, height: start.size.height };
};
export const resolveRotationAngle = (start: DragStart, release: ModelPoint, snap = false): number => {
  const toParent = (point: ModelPoint) => (start.parentTransforms ?? []).reduce((local, transform) => canvasPointToModel(transform, local.x, local.y), point);
  const origin = toParent(start.model);
  const end = toParent(release);
  const startVecX = origin.x - start.pin.x;
  const startVecY = origin.y - start.pin.y;
  const endVecX = end.x - start.pin.x;
  const endVecY = end.y - start.pin.y;
  if ((startVecX === 0 && startVecY === 0) || (endVecX === 0 && endVecY === 0)) return start.angle ?? 0;
  let delta = Math.atan2(endVecY, endVecX) - Math.atan2(startVecY, startVecX);
  while (delta > Math.PI) delta -= Math.PI * 2;
  while (delta <= -Math.PI) delta += Math.PI * 2;
  return snapRotationAngle((start.angle ?? 0) + delta, snap);
};
const applyForward = (transform: Affine, point: ModelPoint): ModelPoint => ({ x: transform.a * point.x + transform.c * point.y + transform.e, y: transform.b * point.x + transform.d * point.y + transform.f });
/** Expresses a page-axis nudge as the drag that would cover the same screen distance. */
export const resolveNudgeGeometry = (start: DragStart, dx: number, dy: number): { x: number; y: number; width: number; height: number } => resolveDragGeometry({ ...start, canvas: { x: 0, y: 0 }, model: { x: 0, y: 0 } }, { x: dx, y: dy });
export const previewOutline = (start: DragStart, release: ModelPoint, paintTransform: Affine, snap = false): ModelPoint[] => {
  const geometry = resolveDragGeometry(start, release);
  const angle = start.rotate ? resolveRotationAngle(start, release, snap) : (start.angle ?? 0);
  const cos = Math.cos(angle), sin = Math.sin(angle);
  const locPin = locPinInches(start, geometry);
  const effFx = geometry.width > 0 ? locPin.x / geometry.width : 0.5;
  const effFy = geometry.height > 0 ? locPin.y / geometry.height : 0.5;
  const fvx = start.flipX ? 1 - effFx : effFx;
  const fvy = start.flipY ? 1 - effFy : effFy;
  const centreOffsetX = (0.5 - fvx) * geometry.width;
  const centreOffsetY = (0.5 - fvy) * geometry.height;
  const centre = { x: geometry.x + cos * centreOffsetX - sin * centreOffsetY, y: geometry.y + sin * centreOffsetX + cos * centreOffsetY };
  const halfWidth = geometry.width / 2, halfHeight = geometry.height / 2;
  const offsets: ReadonlyArray<readonly [number, number]> = [[-halfWidth, -halfHeight], [halfWidth, -halfHeight], [halfWidth, halfHeight], [-halfWidth, halfHeight]];
  return offsets.map(([offsetX, offsetY]) => {
    const parent = { x: centre.x + offsetX * cos - offsetY * sin, y: centre.y + offsetX * sin + offsetY * cos };
    const page = (start.parentTransforms ?? []).reduceRight((point, transform) => applyForward(transform, point), parent);
    return modelPointToCanvas(paintTransform, page.x, page.y);
  });
};
export const paintDragPreview = (context: CanvasRenderingContext2D, corners: readonly ModelPoint[], dpr: number, scale: number): void => {
  if (!corners.length) return;
  context.save();
  try {
    context.setTransform(dpr * scale, 0, 0, dpr * scale, 0, 0);
    context.strokeStyle = '#0f6cbd';
    context.lineWidth = 1;
    context.setLineDash([4, 4]);
    context.beginPath();
    context.moveTo(corners[0].x, corners[0].y);
    for (let index = 1; index < corners.length; index += 1) context.lineTo(corners[index].x, corners[index].y);
    context.closePath();
    context.stroke();
  } finally { context.restore(); }
};
export const selectionHandlePositions = (corners: readonly ModelPoint[]): { handles: Record<ResizeHandle, ModelPoint>; topCenter: ModelPoint } => {
  const mid = (a: ModelPoint, b: ModelPoint): ModelPoint => ({ x: (a.x + b.x) / 2, y: (a.y + b.y) / 2 });
  const south = mid(corners[0], corners[1]);
  const east = mid(corners[1], corners[2]);
  const north = mid(corners[2], corners[3]);
  const west = mid(corners[3], corners[0]);
  return {
    handles: { sw: corners[0], se: corners[1], ne: corners[2], nw: corners[3], s: south, e: east, n: north, w: west },
    topCenter: north,
  };
};
export const rotationGripPosition = (corners: readonly ModelPoint[], zoom: number): ModelPoint => {
  const { topCenter } = selectionHandlePositions(corners);
  const centre = { x: (corners[0].x + corners[2].x) / 2, y: (corners[0].y + corners[2].y) / 2 };
  const outX = topCenter.x - centre.x;
  const outY = topCenter.y - centre.y;
  const length = Math.hypot(outX, outY) || 1;
  const offset = SELECTION_ROTATE_OFFSET_CSS / Math.max(zoom, 1e-6);
  return { x: topCenter.x + (outX / length) * offset, y: topCenter.y + (outY / length) * offset };
};
export const paintSelectionFrame = (context: CanvasRenderingContext2D, corners: readonly ModelPoint[], dpr: number, scale: number, resizeHandles: readonly ResizeHandle[] = RESIZE_HANDLES): void => {
  if (corners.length < 4) return;
  const zoom = Number.isFinite(scale) && scale > 0 ? scale : 1;
  const handleRadius = SELECTION_HANDLE_CSS / zoom / 2;
  const gripRadius = SELECTION_ROTATE_RADIUS_CSS / zoom;
  const grip = rotationGripPosition(corners, zoom);
  const { handles, topCenter } = selectionHandlePositions(corners);
  context.save();
  try {
    context.setTransform(dpr * zoom, 0, 0, dpr * zoom, 0, 0);
    context.strokeStyle = SELECTION_STROKE;
    context.lineWidth = 1 / zoom;
    context.setLineDash([]);
    context.beginPath();
    context.moveTo(corners[0].x, corners[0].y);
    for (let index = 1; index < corners.length; index += 1) context.lineTo(corners[index].x, corners[index].y);
    context.closePath();
    context.stroke();
    context.beginPath();
    context.moveTo(topCenter.x, topCenter.y);
    context.lineTo(grip.x, grip.y);
    context.stroke();
    for (const key of resizeHandles) {
      const anchor = handles[key];
      context.beginPath();
      context.arc(anchor.x, anchor.y, handleRadius, 0, Math.PI * 2);
      context.fillStyle = SELECTION_HANDLE_FILL;
      context.fill();
      context.stroke();
    }
    context.beginPath();
    context.arc(grip.x, grip.y, gripRadius, 0, Math.PI * 2);
    context.fillStyle = SELECTION_HANDLE_FILL;
    context.fill();
    context.stroke();
  } finally { context.restore(); }
};
const pointInQuad = (point: ModelPoint, corners: readonly ModelPoint[]): boolean => {
  let sign = 0;
  for (let index = 0; index < 4; index += 1) {
    const a = corners[index];
    const b = corners[(index + 1) % 4];
    const cross = (b.x - a.x) * (point.y - a.y) - (b.y - a.y) * (point.x - a.x);
    if (cross !== 0) {
      const current = cross > 0 ? 1 : -1;
      if (sign === 0) sign = current;
      else if (sign !== current) return false;
    }
  }
  return true;
};
export const hitTestSelection = (point: ModelPoint, corners: readonly ModelPoint[], zoom: number, toleranceCss = HANDLE_HIT_TOLERANCE_CSS): ResizeHandle | RotateHandle | null => {
  if (corners.length < 4) return null;
  const safeZoom = Number.isFinite(zoom) && zoom > 0 ? zoom : 1;
  const tolerance = toleranceCss / safeZoom;
  const grip = rotationGripPosition(corners, safeZoom);
  if (Math.hypot(point.x - grip.x, point.y - grip.y) <= Math.max(tolerance, SELECTION_ROTATE_RADIUS_CSS / safeZoom)) return 'rotate';
  const { handles } = selectionHandlePositions(corners);
  const radius = Math.max(tolerance, SELECTION_HANDLE_CSS / safeZoom / 2);
  const ranked = RESIZE_HANDLES.map((key) => ({ key, distance: Math.hypot(point.x - handles[key].x, point.y - handles[key].y) })).filter((entry) => entry.distance <= radius).sort((left, right) => left.distance - right.distance);
  if (!ranked.length) return null;
  if (ranked.length > 1 && pointInQuad(point, corners)) {
    const edgeA = Math.hypot(corners[1].x - corners[0].x, corners[1].y - corners[0].y);
    const edgeB = Math.hypot(corners[2].x - corners[1].x, corners[2].y - corners[1].y);
    const tiny = Math.max(edgeA, edgeB) < SELECTION_HANDLE_CSS / safeZoom;
    if (tiny && (ranked[1].distance - ranked[0].distance) < 1.5 / safeZoom) return null;
  }
  return ranked[0].key;
};

export interface TextEditFont { family: string; sizePx: number; bold: boolean; italic: boolean; color: string; }
export interface TextEditOverlay { width: number; height: number; matrix: Affine; font: TextEditFont; }
interface TextEditTarget { x: number; y: number; width: number; height: number; transform: Affine; primitive: PagePrimitive; }

const IDENTITY_AFFINE: Affine = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
const DEFAULT_TEXT_SIZE_IN = 10 / 72;

export const isPrintableEntryKey = (event: CanvasKeyboardEventLike): boolean =>
  event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey && !isEditableKeyboardTarget(event.target);

const composeAffine = (outer: Affine, inner: Affine): Affine => ({
  a: outer.a * inner.a + outer.c * inner.b,
  b: outer.b * inner.a + outer.d * inner.b,
  c: outer.a * inner.c + outer.c * inner.d,
  d: outer.b * inner.c + outer.d * inner.d,
  e: outer.a * inner.e + outer.c * inner.f + outer.e,
  f: outer.b * inner.e + outer.d * inner.f + outer.f,
});

const primitiveTransform = (primitive: PagePrimitive): Affine => ('transform' in primitive ? primitive.transform : undefined) ?? IDENTITY_AFFINE;

const localBounds = (primitive: PagePrimitive, depth = 0): FrameBounds | null => {
  if (primitive.kind === 'textBox' || primitive.kind === 'image' || primitive.kind === 'placeholder') return primitive;
  const points: ModelPoint[] = [];
  if (primitive.kind === 'shape') {
    for (const command of primitive.path) {
      const record = command as unknown as Record<string, unknown>;
      for (const [xKey, yKey] of [['x', 'y'], ['cpx', 'cpy'], ['cp1x', 'cp1y'], ['cp2x', 'cp2y']] as const) {
        const x = record[xKey], y = record[yKey];
        if (typeof x === 'number' && typeof y === 'number' && Number.isFinite(x) && Number.isFinite(y)) points.push({ x, y });
      }
    }
  } else if (primitive.kind === 'group' && depth < 256) {
    for (const child of primitive.primitives) {
      const bounds = localBounds(child, depth + 1);
      if (!bounds) continue;
      const transform = primitiveTransform(child);
      for (const [x, y] of [[bounds.x, bounds.y], [bounds.x + bounds.width, bounds.y], [bounds.x, bounds.y + bounds.height], [bounds.x + bounds.width, bounds.y + bounds.height]] as const) {
        points.push({ x: transform.a * x + transform.c * y + transform.e, y: transform.b * x + transform.d * y + transform.f });
      }
    }
  }
  if (!points.length) return null;
  const xs = points.map((point) => point.x), ys = points.map((point) => point.y);
  const x = Math.min(...xs), y = Math.min(...ys);
  return { x, y, width: Math.max(...xs) - x, height: Math.max(...ys) - y };
};

const textEditTargets = (primitives: readonly PagePrimitive[], id: string, parent: Affine, depth = 0): TextEditTarget[] => {
  if (depth >= 256) return [];
  const targets: TextEditTarget[] = [];
  for (const primitive of primitives) {
    const transform = composeAffine(parent, primitiveTransform(primitive));
    if (primitive.id === id) {
      const bounds = localBounds(primitive);
      if (bounds) targets.push({ ...bounds, transform, primitive });
    }
    if (primitive.kind === 'group') targets.push(...textEditTargets(primitive.primitives, id, transform, depth + 1));
  }
  return targets;
};

export const withoutTextBox = (primitives: readonly PagePrimitive[], id: string, depth = 0): PagePrimitive[] => {
  if (depth >= 256) return [...primitives];
  const kept: PagePrimitive[] = [];
  for (const primitive of primitives) {
    if (primitive.kind === 'textBox' && primitive.id === id) continue;
    kept.push(primitive.kind === 'group' ? { ...primitive, primitives: withoutTextBox(primitive.primitives, id, depth + 1) } : primitive);
  }
  return kept;
};

export const textEditOverlay = (frame: PageDisplayList, primitiveId: string, zoom: number): TextEditOverlay | null => {
  const targets = textEditTargets(frame.primitives, primitiveId, IDENTITY_AFFINE);
  const target = targets.find((candidate) => candidate.primitive.kind === 'textBox') ?? targets[0];
  if (!target) return null;
  const safeZoom = Number.isFinite(zoom) && zoom > 0 ? zoom : 1;
  const toCanvas = composeAffine({ a: safeZoom, b: 0, c: 0, d: safeZoom, e: 0, f: 0 }, composeAffine(frame.paintTransform, target.transform));
  const placed = composeAffine(toCanvas, { a: 1, b: 0, c: 0, d: -1, e: target.x, f: target.y + target.height });
  const scale = Math.sqrt(Math.abs(placed.a * placed.d - placed.b * placed.c));
  if (!Number.isFinite(scale) || scale <= 0) return null;
  const box = target.primitive.kind === 'textBox' ? (target.primitive as TextBoxPrimitive) : null;
  const runs = box?.paragraphs.flatMap((paragraph) => paragraph.runs) ?? [];
  const run = runs.find((candidate) => candidate.text.length > 0) ?? runs[0];
  return {
    width: Math.max(1, target.width * scale),
    height: Math.max(1, target.height * scale),
    matrix: { a: placed.a / scale, b: placed.b / scale, c: placed.c / scale, d: placed.d / scale, e: placed.e, f: placed.f },
    font: {
      family: run?.family || 'Calibri',
      sizePx: Math.max(1, (run?.sizeIn ?? DEFAULT_TEXT_SIZE_IN) * scale),
      bold: run?.bold ?? false,
      italic: run?.italic ?? false,
      color: run?.color || '#000000',
    },
  };
};
