import { expect, test } from 'bun:test';
import { canvasPointToModel, modelPointToCanvas } from '@betteroffice/vsdx';
import type { Affine, ModelPoint, PageDisplayList, TextBoxPrimitive } from '@betteroffice/vsdx';
import { RESIZE_HANDLES, SELECTION_STROKE, canvasKeyboardIntent, hitTestSelection, isEditableKeyboardTarget, isPrintableEntryKey, keyboardNudgeStep, paintSelectionFrame, paintDragPreview, passedDragThreshold, previewOutline, resizedBounds, resizeCursor, resolveDragGeometry, resolveNudgeGeometry, resolveRotationAngle, rotationGripPosition, selectionHandlePositions, textEditOverlay, withoutTextBox } from './interactions';
const pagePaintTransform = { a: 96, b: 0, c: 0, d: -96, e: 0, f: 1056 };
const identity = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
test('passedDragThreshold needs four css pixels by default', () => {
  expect(passedDragThreshold(10, 10, 12, 12)).toBe(false);
  expect(passedDragThreshold(10, 10, 14, 10)).toBe(true);
  expect(passedDragThreshold(10, 10, 13, 10)).toBe(false);
  expect(passedDragThreshold(10, 10, 16, 10, 8)).toBe(false);
  expect(passedDragThreshold(10, 10, 18, 10, 8)).toBe(true);
});
test('previewOutline draws the moved box in canvas coordinates', () => {
  const start = { canvas: { x: 0, y: 0 }, model: { x: 3, y: 3 }, resize: false, pin: { x: 5, y: 2 }, size: { width: 2, height: 1 } };
  expect(previewOutline(start, { x: 4, y: 4 }, pagePaintTransform)).toEqual([{ x: 480, y: 816 }, { x: 672, y: 816 }, { x: 672, y: 720 }, { x: 480, y: 720 }]);
});
test('previewOutline rotates the box with the shape angle', () => {
  const start = { canvas: { x: 0, y: 0 }, model: { x: 0, y: 0 }, resize: false, pin: { x: 2, y: 3 }, size: { width: 4, height: 5 }, angle: Math.PI / 2 };
  const corners = previewOutline(start, { x: 0, y: 0 }, identity);
  expect(corners[0].x).toBeCloseTo(4.5, 10); expect(corners[0].y).toBeCloseTo(1, 10);
  expect(corners[1].x).toBeCloseTo(4.5, 10); expect(corners[1].y).toBeCloseTo(5, 10);
  expect(corners[2].x).toBeCloseTo(-0.5, 10); expect(corners[2].y).toBeCloseTo(5, 10);
  expect(corners[3].x).toBeCloseTo(-0.5, 10); expect(corners[3].y).toBeCloseTo(1, 10);
});
test('previewOutline keeps a centred flip on the same visual frame', () => {
  const base = { canvas: { x: 0, y: 0 }, model: { x: 3, y: 3 }, resize: false, pin: { x: 5, y: 2 }, size: { width: 2, height: 1 } };
  const plain = previewOutline(base, { x: 3, y: 3 }, identity);
  expect(plain).toEqual([{ x: 4, y: 1.5 }, { x: 6, y: 1.5 }, { x: 6, y: 2.5 }, { x: 4, y: 2.5 }]);
  const flipped = previewOutline({ ...base, flipX: true }, { x: 3, y: 3 }, identity);
  expect(flipped).toEqual(plain);
  expect(selectionHandlePositions(flipped).handles.e).toEqual({ x: 6, y: 2 });
});
test('previewOutline skips the engine LocPin lookup for a shape with no height', () => {
  const start = { canvas: { x: 0, y: 0 }, model: { x: 0, y: 0 }, resize: false, pin: { x: 2, y: 3 }, size: { width: 4, height: 0 }, locPinAtSize: () => { throw new Error('invalid resize dimensions'); } };
  expect(previewOutline(start, { x: 1, y: 0 }, identity)).toEqual([{ x: 1, y: 3 }, { x: 5, y: 3 }, { x: 5, y: 3 }, { x: 1, y: 3 }]);
});
test('an engine LocPin refusal leaves the gesture on the stored LocPin', () => {
  const refuse = () => { throw new Error('cannot evaluate LocPinX for resize'); };
  const start = { canvas: { x: 0, y: 0 }, model: { x: 0, y: 0 }, resize: false, handle: 'e' as const, pin: { x: 2, y: 3 }, locPin: { x: 1, y: 2.5 }, size: { width: 2, height: 5 }, locPinAtSize: refuse };
  const geometry = resolveDragGeometry(start, { x: 1, y: 0 });
  expect(geometry.width).toBeCloseTo(3, 10);
  expect(geometry.x - 1).toBeCloseTo(1, 10);
  expect(() => previewOutline(start, { x: 1, y: 0 }, identity)).not.toThrow();
});
test('a non-finite engine LocPin leaves the gesture on the stored LocPin', () => {
  const start = { canvas: { x: 0, y: 0 }, model: { x: 0, y: 0 }, resize: false, handle: 'e' as const, pin: { x: 2, y: 3 }, locPin: { x: 1, y: 2.5 }, size: { width: 2, height: 5 }, locPinAtSize: () => ({ x: Number.NaN, y: 2.5 }) };
  const corners = previewOutline(start, { x: 1, y: 0 }, identity);
  expect(corners.every((corner) => Number.isFinite(corner.x) && Number.isFinite(corner.y))).toBe(true);
  expect(resolveDragGeometry(start, { x: 1, y: 0 }).x).toBeCloseTo(2, 10);
});
test('previewOutline maps the box through the group transform forward', () => {
  const start = { canvas: { x: 0, y: 0 }, model: { x: 10, y: 20 }, resize: false, pin: { x: 2, y: 3 }, size: { width: 4, height: 5 }, parentTransforms: [{ a: 0, b: 2, c: -2, d: 0, e: 10, f: 20 }] };
  expect(previewOutline(start, { x: 8, y: 24 }, identity)).toEqual([{ x: 7, y: 24 }, { x: 7, y: 32 }, { x: -3, y: 32 }, { x: -3, y: 24 }]);
});
test('preview geometry and commit geometry describe the same rectangle', () => {
  const starts = [
    { canvas: { x: 0, y: 0 }, model: { x: 3, y: 3 }, resize: false, pin: { x: 5, y: 2 }, size: { width: 2, height: 1 } },
    { canvas: { x: 0, y: 0 }, model: { x: 0, y: 0 }, resize: false, pin: { x: 2, y: 3 }, size: { width: 4, height: 5 }, angle: Math.PI / 2 },
    { canvas: { x: 0, y: 0 }, model: { x: 10, y: 20 }, resize: false, pin: { x: 2, y: 3 }, size: { width: 4, height: 5 }, parentTransforms: [{ a: 0, b: 2, c: -2, d: 0, e: 10, f: 20 }] },
    { canvas: { x: 0, y: 0 }, model: { x: 3, y: 3 }, resize: true, pin: { x: 5, y: 2 }, size: { width: 2, height: 1 } },
  ];
  const releases = [{ x: 4, y: 4 }, { x: 1, y: -1 }, { x: 8, y: 24 }, { x: 4, y: 5 }];
  starts.forEach((start, index) => {
    const release = releases[index];
    const geometry = resolveDragGeometry(start, release);
    const corners = previewOutline(start, release, identity);
    const centre = { x: (corners[0].x + corners[2].x) / 2, y: (corners[0].y + corners[2].y) / 2 };
    const back = (start.parentTransforms ?? []).reduce((local, transform) => canvasPointToModel(transform, local.x, local.y), centre);
    void back;
    const parentCentre = (start.parentTransforms ?? []).length ? (start.parentTransforms ?? []).reduceRight((point, transform) => ({ x: transform.a * point.x + transform.c * point.y + transform.e, y: transform.b * point.x + transform.d * point.y + transform.f }), { x: geometry.x, y: geometry.y }) : { x: geometry.x, y: geometry.y };
    expect(centre.x).toBeCloseTo(parentCentre.x, 10); expect(centre.y).toBeCloseTo(parentCentre.y, 10);
    const edgeA = Math.hypot(corners[1].x - corners[0].x, corners[1].y - corners[0].y);
    const edgeB = Math.hypot(corners[3].x - corners[0].x, corners[3].y - corners[0].y);
    const sorted = [edgeA, edgeB].sort((left, right) => left - right);
    const expected = [Math.min(geometry.width, geometry.height), Math.max(geometry.width, geometry.height)];
    const parentScale = (start.parentTransforms ?? []).length ? 2 : 1;
    expect(sorted[0]).toBeCloseTo(expected[0] * parentScale, 8); expect(sorted[1]).toBeCloseTo(expected[1] * parentScale, 8);
  });
});
test('preview outline matches page content at zoom 2 through the device transform', () => {
  const paintTransform = { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 };
  const start = { canvas: { x: 0, y: 0 }, model: { x: 4, y: 3 }, resize: false, pin: { x: 5, y: 2 }, size: { width: 2, height: 1 } };
  const release = { x: 4, y: 3 };
  const geometry = resolveDragGeometry(start, release);
  const corners = previewOutline(start, release, paintTransform);
  const rescaled = (previewOutline as unknown as (start: unknown, release: ModelPoint, paintTransform: unknown, scale: number) => ModelPoint[])(start, release, paintTransform, 2);
  expect(rescaled).toEqual(corners);
  const centre = { x: (corners[0].x + corners[2].x) / 2, y: (corners[0].y + corners[2].y) / 2 };
  const page = modelPointToCanvas(paintTransform, geometry.x, geometry.y);
  expect(centre.x).toBeCloseTo(page.x, 10); expect(centre.y).toBeCloseTo(page.y, 10);
  expect({ x: centre.x * 2, y: centre.y * 2 }).toEqual({ x: 960, y: 1056 });
});
test('paintDragPreview strokes a dashed brand outline and restores state', () => {
  const calls: string[] = [];
  const context = new Proxy({ canvas: {} }, {
    get(target, key) {
      if (key in target) return Reflect.get(target, key);
      return (...args: unknown[]) => { calls.push(`${String(key)}:${args.join(',')}`); };
    },
    set(target, key, value) { calls.push(`${String(key)}=${String(value)}`); Reflect.set(target, key, value); return true; },
  }) as unknown as CanvasRenderingContext2D;
  paintDragPreview(context, [{ x: 1, y: 2 }, { x: 3, y: 2 }, { x: 3, y: 4 }, { x: 1, y: 4 }], 2, 1);
  expect(calls).toContain('setTransform:2,0,0,2,0,0');
  expect(calls).toContain('strokeStyle=#0f6cbd');
  expect(calls).toContain('lineWidth=1');
  expect(calls.some((entry) => entry.startsWith('setLineDash:'))).toBe(true);
  expect(calls.some((entry) => entry.startsWith('stroke:'))).toBe(true);
  expect(calls[calls.length - 1].startsWith('restore:')).toBe(true);
});
test('resize vocabulary maps handles to cursors and bounds', () => {
  expect(RESIZE_HANDLES).toEqual(['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']);
  const bounds = { x: 100, y: 200, width: 300, height: 100 };
  expect(resizeCursor('nw')).toBe('nwse-resize');
  expect(resizeCursor('ne')).toBe('nesw-resize');
  expect(resizeCursor('n')).toBe('ns-resize');
  expect(resizeCursor('w')).toBe('ew-resize');
  expect(resizedBounds(bounds, 'se', { x: 50, y: 20 }, 0)).toEqual({ x: 100, y: 200, width: 350, height: 120 });
  expect(resizedBounds(bounds, 'nw', { x: 50, y: 20 }, 0)).toEqual({ x: 150, y: 220, width: 250, height: 80 });
});
test('handle anchors follow a rotated shape', () => {
  const start = { canvas: { x: 0, y: 0 }, model: { x: 0, y: 0 }, resize: false, pin: { x: 10, y: 10 }, size: { width: 20, height: 10 }, angle: Math.PI / 2 };
  const corners = previewOutline(start, { x: 0, y: 0 }, identity);
  expect(corners[0].x).toBeCloseTo(15, 8); expect(corners[0].y).toBeCloseTo(0, 8);
  expect(corners[1].x).toBeCloseTo(15, 8); expect(corners[1].y).toBeCloseTo(20, 8);
  expect(corners[2].x).toBeCloseTo(5, 8); expect(corners[2].y).toBeCloseTo(20, 8);
  expect(corners[3].x).toBeCloseTo(5, 8); expect(corners[3].y).toBeCloseTo(0, 8);
  const positions = selectionHandlePositions(corners);
  expect(positions.handles.n.x).toBeCloseTo(5, 8); expect(positions.handles.n.y).toBeCloseTo(10, 8);
  expect(positions.handles.s.x).toBeCloseTo(15, 8); expect(positions.handles.s.y).toBeCloseTo(10, 8);
  expect(positions.handles.e.x).toBeCloseTo(10, 8); expect(positions.handles.e.y).toBeCloseTo(20, 8);
  expect(positions.handles.w.x).toBeCloseTo(10, 8); expect(positions.handles.w.y).toBeCloseTo(0, 8);
  expect(hitTestSelection({ x: 5, y: 10 }, corners, 1)).toBe('n');
  expect(hitTestSelection({ x: 15, y: 10 }, corners, 1)).toBe('s');
  expect(hitTestSelection({ x: 10, y: 10 }, corners, 1, 2)).toBeNull();
});
test('a resize from nw keeps the se corner fixed', () => {
  const start = { canvas: { x: 0, y: 0 }, model: { x: 0, y: 0 }, resize: false, handle: 'nw' as const, pin: { x: 5, y: 2 }, locPin: { x: 1, y: 0.5 }, size: { width: 2, height: 1 } };
  const geometry = resolveDragGeometry(start, { x: -1, y: 1 });
  expect(geometry.width).toBeCloseTo(3, 10);
  expect(geometry.height).toBeCloseTo(2, 10);
  expect(geometry.x).toBeCloseTo(4, 10);
  expect(geometry.y).toBeCloseTo(2, 10);
  expect(geometry.x - 1 + geometry.width).toBeCloseTo(6, 10);
  expect(geometry.y - 0.5).toBeCloseTo(1.5, 10);
});
test('the rotation grip produces the expected angle', () => {
  const start = { canvas: { x: 0, y: 0 }, model: { x: 1, y: 0 }, resize: false, rotate: true, pin: { x: 0, y: 0 }, size: { width: 2, height: 1 }, angle: 0 };
  expect(resolveRotationAngle(start, { x: 0, y: 1 })).toBeCloseTo(Math.PI / 2, 10);
  const tilted = { canvas: { x: 0, y: 0 }, model: { x: 1, y: 0 }, resize: false, rotate: true, pin: { x: 0, y: 0 }, size: { width: 2, height: 1 }, angle: 0 };
  const seventeen = { x: Math.cos(17 * Math.PI / 180), y: Math.sin(17 * Math.PI / 180) };
  expect(resolveRotationAngle(tilted, seventeen)).toBeCloseTo(17 * Math.PI / 180, 10);
  expect(resolveRotationAngle(tilted, seventeen, true)).toBeCloseTo(15 * Math.PI / 180, 10);
  const rotated = previewOutline(start, { x: 0, y: 1 }, identity);
  expect(rotated[0].x).toBeCloseTo(0.5, 10);
  expect(rotated[0].y).toBeCloseTo(-1, 10);
});
test('the selection frame paints at a zoom other than 1', () => {
  const calls: string[] = [];
  const context = new Proxy({ canvas: {} }, {
    get(target, key) {
      if (key in target) return Reflect.get(target, key);
      return (...args: unknown[]) => { calls.push(`${String(key)}:${args.join(',')}`); };
    },
    set(target, key, value) { calls.push(`${String(key)}=${String(value)}`); Reflect.set(target, key, value); return true; },
  }) as unknown as CanvasRenderingContext2D;
  const corners = [{ x: 10, y: 40 }, { x: 30, y: 40 }, { x: 30, y: 20 }, { x: 10, y: 20 }];
  paintSelectionFrame(context, corners, 2, 2);
  expect(SELECTION_STROKE).not.toBe('#0f6cbd');
  expect(calls).toContain('setTransform:4,0,0,4,0,0');
  expect(calls).toContain(`strokeStyle=${SELECTION_STROKE}`);
  expect(calls.some((entry) => entry === 'strokeStyle=#0f6cbd')).toBe(false);
  expect(calls).toContain('lineWidth=0.5');
  expect(calls).toContain('moveTo:10,40');
  expect(calls.some((entry) => entry.startsWith('fillRect:'))).toBe(false);
  expect(calls.some((entry) => entry.startsWith('strokeRect:'))).toBe(false);
  expect(calls.filter((entry) => entry.startsWith('arc:')).length).toBeGreaterThanOrEqual(9);
  const grip = rotationGripPosition(corners, 2);
  expect(grip.y).toBeLessThan(20);
  expect(calls.some((entry) => entry === `lineTo:${grip.x},${grip.y}`)).toBe(true);
  expect(calls.some((entry) => entry.startsWith(`arc:${grip.x},${grip.y},`) && entry.endsWith(',0,6.283185307179586'))).toBe(true);
  expect(hitTestSelection(grip, corners, 2)).toBe('rotate');
});
test('locPin governs the handle box instead of cancelling out', () => {
  const centred = { canvas: { x: 0, y: 0 }, model: { x: 0, y: 0 }, resize: false, handle: 'e' as const, pin: { x: 5, y: 2 }, locPin: { x: 1, y: 0.5 }, size: { width: 2, height: 1 } };
  const edge = { canvas: { x: 0, y: 0 }, model: { x: 0, y: 0 }, resize: false, handle: 'e' as const, pin: { x: 5, y: 2 }, locPin: { x: 0, y: 0.5 }, size: { width: 2, height: 1 } };
  const grownCentred = resolveDragGeometry(centred, { x: 1, y: 0 });
  expect(grownCentred.width).toBeCloseTo(3, 10);
  expect(grownCentred.x).toBeCloseTo(5, 10);
  const grownEdge = resolveDragGeometry(edge, { x: 1, y: 0 });
  expect(grownEdge.width).toBeCloseTo(3, 10);
  expect(grownEdge.x).toBeCloseTo(5, 10);
  expect(grownEdge.y).toBeCloseTo(2, 10);
  const centredCorners = previewOutline(centred, { x: 1, y: 0 }, identity);
  const edgeCorners = previewOutline(edge, { x: 1, y: 0 }, identity);
  expect(Math.min(...centredCorners.map((corner) => corner.x))).toBeCloseTo(4, 10);
  expect(Math.max(...centredCorners.map((corner) => corner.x))).toBeCloseTo(7, 10);
  expect(Math.min(...edgeCorners.map((corner) => corner.x))).toBeCloseTo(5, 10);
  expect(Math.max(...edgeCorners.map((corner) => corner.x))).toBeCloseTo(8, 10);
  const lowPin = { canvas: { x: 0, y: 0 }, model: { x: 0, y: 0 }, resize: false, handle: 'n' as const, pin: { x: 5, y: 2 }, locPin: { x: 1, y: 0 }, size: { width: 2, height: 1 } };
  const grownNorth = resolveDragGeometry(lowPin, { x: 0, y: 1 });
  expect(grownNorth.height).toBeCloseTo(2, 10);
  expect(grownNorth.y).toBeCloseTo(2, 10);
  expect(grownNorth.x).toBeCloseTo(5, 10);
  const northCorners = previewOutline(lowPin, { x: 0, y: 1 }, identity);
  expect(Math.min(...northCorners.map((corner) => corner.y))).toBeCloseTo(2, 10);
  expect(Math.max(...northCorners.map((corner) => corner.y))).toBeCloseTo(4, 10);
});
test('a flipped handle resize grows outward on both axes', () => {
  const flipX = { canvas: { x: 0, y: 0 }, model: { x: 0, y: 0 }, resize: false, handle: 'e' as const, pin: { x: 0, y: 0 }, size: { width: 20, height: 10 }, flipX: true };
  const grownX = resolveDragGeometry(flipX, { x: 2, y: 0 });
  expect(grownX.width).toBeCloseTo(22, 10);
  expect(grownX.height).toBeCloseTo(10, 10);
  const flipY = { canvas: { x: 0, y: 0 }, model: { x: 0, y: 0 }, resize: false, handle: 'n' as const, pin: { x: 0, y: 0 }, size: { width: 20, height: 10 }, flipY: true };
  const grownY = resolveDragGeometry(flipY, { x: 0, y: 2 });
  expect(grownY.height).toBeCloseTo(12, 10);
  expect(grownY.width).toBeCloseTo(20, 10);
  const base = { canvas: { x: 0, y: 0 }, model: { x: 0, y: 0 }, resize: false, pin: { x: 5, y: 2 }, size: { width: 2, height: 1 } };
  const plain = previewOutline(base, { x: 0, y: 0 }, identity);
  const flipped = previewOutline({ ...base, flipX: true }, { x: 0, y: 0 }, identity);
  expect(flipped).toEqual(plain);
  expect(selectionHandlePositions(flipped).handles.e).toEqual({ x: 6, y: 2 });
  const flippedY = previewOutline({ ...base, flipY: true }, { x: 0, y: 0 }, identity);
  expect(flippedY).toEqual(plain);
  expect(selectionHandlePositions(flippedY).handles.n).toEqual({ x: 5, y: 2.5 });
});
test('hit testing returns the nearest handle and keeps tiny shapes draggable', () => {
  const start = { canvas: { x: 0, y: 0 }, model: { x: 0, y: 0 }, resize: false, pin: { x: 1, y: 1 }, size: { width: 0.12, height: 0.12 } };
  const corners = previewOutline(start, { x: 0, y: 0 }, pagePaintTransform);
  const positions = selectionHandlePositions(corners);
  const nearEast = { x: positions.handles.e.x, y: positions.handles.e.y - 2 };
  expect(hitTestSelection(nearEast, corners, 1)).toBe('e');
  const tinyStart = { canvas: { x: 0, y: 0 }, model: { x: 0, y: 0 }, resize: false, pin: { x: 1, y: 1 }, size: { width: 0.05, height: 0.05 } };
  const tiny = previewOutline(tinyStart, { x: 0, y: 0 }, pagePaintTransform);
  const centre = { x: (tiny[0].x + tiny[2].x) / 2, y: (tiny[0].y + tiny[2].y) / 2 };
  expect(hitTestSelection(centre, tiny, 1)).toBeNull();
  expect(hitTestSelection(centre, tiny, 4)).toBeNull();
});
test('the rotation grip follows local north instead of the screen top', () => {
  for (const degrees of [0, 90, 190]) {
    const angle = degrees * Math.PI / 180;
    const start = { canvas: { x: 0, y: 0 }, model: { x: 0, y: 0 }, resize: false, pin: { x: 0, y: 0 }, size: { width: 4, height: 2 }, angle };
    const corners = previewOutline(start, { x: 0, y: 0 }, identity);
    const positions = selectionHandlePositions(corners);
    const north = { x: (corners[2].x + corners[3].x) / 2, y: (corners[2].y + corners[3].y) / 2 };
    expect(positions.topCenter.x).toBeCloseTo(north.x, 8);
    expect(positions.topCenter.y).toBeCloseTo(north.y, 8);
    const grip = rotationGripPosition(corners, 1);
    const centre = { x: (corners[0].x + corners[2].x) / 2, y: (corners[0].y + corners[2].y) / 2 };
    const direction = { x: grip.x - centre.x, y: grip.y - centre.y };
    const length = Math.hypot(direction.x, direction.y);
    const expected = { x: -Math.sin(angle), y: Math.cos(angle) };
    expect((direction.x / length) * expected.x + (direction.y / length) * expected.y).toBeCloseTo(1, 6);
  }
  const lower = previewOutline({ canvas: { x: 0, y: 0 }, model: { x: 0, y: 0 }, resize: false, pin: { x: 0, y: 0 }, size: { width: 4, height: 2 }, angle: 40 * Math.PI / 180 }, { x: 0, y: 0 }, identity);
  const upper = previewOutline({ canvas: { x: 0, y: 0 }, model: { x: 0, y: 0 }, resize: false, pin: { x: 0, y: 0 }, size: { width: 4, height: 2 }, angle: 50 * Math.PI / 180 }, { x: 0, y: 0 }, identity);
  const lowerNorth = { x: (lower[2].x + lower[3].x) / 2, y: (lower[2].y + lower[3].y) / 2 };
  const upperNorth = { x: (upper[2].x + upper[3].x) / 2, y: (upper[2].y + upper[3].y) / 2 };
  expect(selectionHandlePositions(lower).topCenter.x).toBeCloseTo(lowerNorth.x, 8);
  expect(selectionHandlePositions(upper).topCenter.x).toBeCloseTo(upperNorth.x, 8);
});
test('a literal off-centre LocPin previews exactly what the commit renders', () => {
  const frame = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
  const locPin = { x: 0.5, y: 0.5 };
  const east = { canvas: { x: 0, y: 0 }, model: { x: 0, y: 0 }, resize: false, handle: 'e' as const, pin: { x: 5, y: 2 }, locPin, size: { width: 2, height: 1 } };
  const grown = resolveDragGeometry(east, { x: 1, y: 0 });
  expect(grown.width).toBeCloseTo(3, 10);
  const committedLeft = grown.x - locPin.x;
  expect(committedLeft).toBeCloseTo(4.5, 10);
  expect(committedLeft + grown.width).toBeCloseTo(7.5, 10);
  const eastCorners = previewOutline(east, { x: 1, y: 0 }, frame);
  expect(Math.min(...eastCorners.map((corner) => corner.x))).toBeCloseTo(committedLeft, 10);
  expect(Math.max(...eastCorners.map((corner) => corner.x))).toBeCloseTo(committedLeft + grown.width, 10);
  const northWest = { ...east, handle: 'nw' as const };
  const stretched = resolveDragGeometry(northWest, { x: -1, y: 1 });
  expect(stretched.width).toBeCloseTo(3, 10);
  expect(stretched.height).toBeCloseTo(2, 10);
  const committedRight = 5 - locPin.x + 2;
  expect(stretched.x - locPin.x + stretched.width).toBeCloseTo(committedRight, 10);
  expect(stretched.x - locPin.x).toBeCloseTo(committedRight - stretched.width, 10);
  const northCorners = previewOutline(northWest, { x: -1, y: 1 }, frame);
  expect(Math.min(...northCorners.map((corner) => corner.x))).toBeCloseTo(stretched.x - locPin.x, 10);
  expect(Math.max(...northCorners.map((corner) => corner.x))).toBeCloseTo(committedRight, 10);
  expect(Math.min(...northCorners.map((corner) => corner.y))).toBeCloseTo(stretched.y - locPin.y, 10);
  expect(Math.max(...northCorners.map((corner) => corner.y))).toBeCloseTo(stretched.y - locPin.y + stretched.height, 10);
});
test('canvas keyboard maps history, delete and escape intents', () => {
  expect(canvasKeyboardIntent({ key: 'z', ctrlKey: true }, 1)).toEqual({ kind: 'undo' });
  expect(canvasKeyboardIntent({ key: 'Z', metaKey: true, shiftKey: true }, 1)).toEqual({ kind: 'redo' });
  expect(canvasKeyboardIntent({ key: 'y', ctrlKey: true }, 1)).toEqual({ kind: 'redo' });
  expect(canvasKeyboardIntent({ key: 'Y', metaKey: true }, 1)).toEqual({ kind: 'redo' });
  expect(canvasKeyboardIntent({ key: 'Delete' }, 1)).toEqual({ kind: 'delete' });
  expect(canvasKeyboardIntent({ key: 'Backspace' }, 1)).toEqual({ kind: 'delete' });
  expect(canvasKeyboardIntent({ key: 'Escape' }, 1)).toEqual({ kind: 'escape' });
  expect(canvasKeyboardIntent({ key: 'a', ctrlKey: true }, 1)).toBeNull();
  expect(canvasKeyboardIntent({ key: 'z', ctrlKey: true, altKey: true }, 1)).toBeNull();
  expect(canvasKeyboardIntent({ key: 'Delete', ctrlKey: true }, 1)).toBeNull();
  expect(canvasKeyboardIntent({ key: 'Escape', ctrlKey: true }, 1)).toBeNull();
});
test('resolveNudgeGeometry turns a screen nudge into the parent-local pin delta', () => {
  const start = { canvas: { x: 0, y: 0 }, model: { x: 0, y: 0 }, resize: false, pin: { x: 2, y: 3 }, size: { width: 4, height: 5 }, parentTransforms: [{ a: 0, b: 1, c: -1, d: 0, e: 0, f: 0 }] };
  expect(resolveNudgeGeometry(start, 1, 0)).toEqual({ x: 2, y: 2, width: 4, height: 5 });
});
test('canvas keyboard nudges one screen pixel with Y up and ten with shift', () => {
  expect(keyboardNudgeStep(1)).toBeCloseTo(1 / 96, 10);
  expect(keyboardNudgeStep(2)).toBeCloseTo(1 / 192, 10);
  expect(canvasKeyboardIntent({ key: 'ArrowUp' }, 1)).toEqual({ kind: 'nudge', dx: 0, dy: 1 / 96 });
  expect(canvasKeyboardIntent({ key: 'ArrowDown' }, 1)).toEqual({ kind: 'nudge', dx: 0, dy: -1 / 96 });
  expect(canvasKeyboardIntent({ key: 'ArrowLeft' }, 2)).toEqual({ kind: 'nudge', dx: -1 / 192, dy: 0 });
  const right = canvasKeyboardIntent({ key: 'ArrowRight', shiftKey: true }, 1);
  expect(right?.kind).toBe('nudge');
  if (right?.kind === 'nudge') { expect(right.dx).toBeCloseTo(10 / 96, 10); expect(right.dy).toBe(0); }
  const up = canvasKeyboardIntent({ key: 'ArrowUp', shiftKey: true }, 2);
  expect(up?.kind).toBe('nudge');
  if (up?.kind === 'nudge') { expect(up.dy).toBeCloseTo(10 / 192, 10); expect(up.dx).toBe(0); }
  expect(canvasKeyboardIntent({ key: 'ArrowUp', ctrlKey: true }, 1)).toBeNull();
  expect(canvasKeyboardIntent({ key: 'ArrowUp', altKey: true }, 1)).toBeNull();
});
test('canvas keyboard produces no intent from editable targets', () => {
  const input = { tagName: 'INPUT' };
  const textarea = { tagName: 'textarea' };
  const editable = { tagName: 'DIV', isContentEditable: true };
  expect(isEditableKeyboardTarget(input)).toBe(true);
  expect(isEditableKeyboardTarget(textarea)).toBe(true);
  expect(isEditableKeyboardTarget(editable)).toBe(true);
  expect(isEditableKeyboardTarget({ tagName: 'CANVAS' })).toBe(false);
  expect(canvasKeyboardIntent({ key: 'Delete', target: input }, 1)).toBeNull();
  expect(canvasKeyboardIntent({ key: 'ArrowUp', target: input }, 1)).toBeNull();
  expect(canvasKeyboardIntent({ key: 'z', ctrlKey: true, target: textarea }, 1)).toBeNull();
  expect(canvasKeyboardIntent({ key: 'Escape', target: editable }, 1)).toBeNull();
});

const textFrame: PageDisplayList = { contractVersion: 4, width: 816, height: 1056, paintTransform: pagePaintTransform, primitives: [] };

function textBox(transform?: Affine): TextBoxPrimitive {
  return {
    kind: 'textBox', id: 'visio/pages/page1.xml:1', zOrder: 0, x: 1, y: 2, width: 3, height: 0.5, transform,
    paragraphs: [{ runs: [{ text: 'label', family: 'Segoe UI', sizeIn: 0.25, bold: true, italic: false, underline: false, smallCaps: false, superscript: false, subscript: false, letterSpacing: 0, color: '#112233', diagnostics: [] }] }],
    lines: [],
  };
}

function cssCorner(sceneX: number, sceneY: number, chain: Affine[], zoom: number): ModelPoint {
  let point = { x: sceneX, y: sceneY };
  for (const transform of [...chain, pagePaintTransform]) point = { x: transform.a * point.x + transform.c * point.y + transform.e, y: transform.b * point.x + transform.d * point.y + transform.f };
  return { x: point.x * zoom, y: point.y * zoom };
}

function overlayCorner(overlay: { matrix: Affine }, cx: number, cy: number): ModelPoint {
  return { x: overlay.matrix.a * cx + overlay.matrix.c * cy + overlay.matrix.e, y: overlay.matrix.b * cx + overlay.matrix.d * cy + overlay.matrix.f };
}

test('places the text editor over an unrotated text box and scales it with the zoom', () => {
  const overlay = textEditOverlay({ ...textFrame, primitives: [textBox()] }, 'visio/pages/page1.xml:1', 1.5);
  if (!overlay) throw new Error('no overlay');
  expect(overlay.width).toBeCloseTo(3 * 96 * 1.5, 5);
  expect(overlay.height).toBeCloseTo(0.5 * 96 * 1.5, 5);
  expect(overlay.font).toEqual({ family: 'Segoe UI', sizePx: 0.25 * 96 * 1.5, bold: true, italic: false, color: '#112233' });
  expect(overlayCorner(overlay, 0, 0)).toEqual(cssCorner(1, 2.5, [], 1.5));
  expect(overlayCorner(overlay, overlay.width, overlay.height)).toEqual(cssCorner(4, 2, [], 1.5));
});

test('rotates the text editor with the text box instead of using its bounding box', () => {
  const rotation: Affine = { a: Math.cos(0.4), b: Math.sin(0.4), c: -Math.sin(0.4), d: Math.cos(0.4), e: 2, f: 1 };
  const overlay = textEditOverlay({ ...textFrame, primitives: [textBox(rotation)] }, 'visio/pages/page1.xml:1', 1);
  if (!overlay) throw new Error('no overlay');
  expect(overlay.matrix.b).not.toBeCloseTo(0, 3);
  for (const [cx, cy, sceneX, sceneY] of [[0, 0, 1, 2.5], [overlay.width, 0, 4, 2.5], [0, overlay.height, 1, 2], [overlay.width, overlay.height, 4, 2]] as const) {
    const placed = overlayCorner(overlay, cx, cy);
    const expected = cssCorner(sceneX, sceneY, [rotation], 1);
    expect(placed.x).toBeCloseTo(expected.x, 5);
    expect(placed.y).toBeCloseTo(expected.y, 5);
  }
});

test('carries the transforms of enclosing groups into the text editor position', () => {
  const group: Affine = { a: 2, b: 0, c: 0, d: 2, e: 1, f: 3 };
  const primitives = [{ kind: 'group' as const, id: 'visio/pages/page1.xml:9', zOrder: 0, transform: group, primitives: [textBox()] }];
  const overlay = textEditOverlay({ ...textFrame, primitives }, 'visio/pages/page1.xml:1', 1);
  if (!overlay) throw new Error('no overlay');
  expect(overlay.width).toBeCloseTo(3 * 2 * 96, 5);
  const placed = overlayCorner(overlay, 0, 0);
  const expected = cssCorner(1, 2.5, [group], 1);
  expect(placed.x).toBeCloseTo(expected.x, 5);
  expect(placed.y).toBeCloseTo(expected.y, 5);
});

test('hides only the edited shape text from the painted page', () => {
  const other: TextBoxPrimitive = { ...textBox(), id: 'visio/pages/page1.xml:2' };
  const nested = { kind: 'group' as const, id: 'visio/pages/page1.xml:9', zOrder: 0, primitives: [textBox(), other] };
  const kept = withoutTextBox([nested, textBox()], 'visio/pages/page1.xml:1');
  expect(kept).toHaveLength(1);
  expect((kept[0] as typeof nested).primitives).toEqual([other]);
});

test('enters text edit on a printable key but not on a shortcut or an editable target', () => {
  expect(isPrintableEntryKey({ key: 'a' })).toBe(true);
  expect(isPrintableEntryKey({ key: 'Enter' })).toBe(false);
  expect(isPrintableEntryKey({ key: 'a', metaKey: true })).toBe(false);
  expect(isPrintableEntryKey({ key: 'a', ctrlKey: true })).toBe(false);
  expect(isPrintableEntryKey({ key: 'a', target: { tagName: 'TEXTAREA' } })).toBe(false);
});
