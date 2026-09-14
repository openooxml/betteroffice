import { expect, test } from 'bun:test';
import type { DiagramSnapshot, PageDisplayList } from '@betteroffice/vsdx';
import type { PointerEvent } from 'react';
import { canvasPointerPosition, inchFormula, resolveDragGeometry, selectionCorners, stillSelectable } from './VsdxEditor';
import { previewOutline, resolveNudgeGeometry, resolveRotationAngle } from './interactions';

const frame: PageDisplayList = {
  contractVersion: 4,
  width: 816,
  height: 1056,
  paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 1056 },
  primitives: [],
};

function pointerAt(clientX: number, clientY: number, cssScale = 1): PointerEvent<HTMLCanvasElement> {
  return {
    clientX,
    clientY,
    currentTarget: { getBoundingClientRect: () => ({ left: 0, top: 0, width: frame.width * cssScale, height: frame.height * cssScale }) },
  } as unknown as PointerEvent<HTMLCanvasElement>;
}

test('maps a canvas pointer onto Y-up inches for the save projection', () => {
  const bottomLeft = canvasPointerPosition(pointerAt(0, 1056), frame);
  expect(bottomLeft.canvas).toEqual({ x: 0, y: 1056 });
  expect(bottomLeft.model).toEqual({ x: 0, y: 0 });
  const topLeft = canvasPointerPosition(pointerAt(0, 0), frame);
  expect(topLeft.model).toEqual({ x: 0, y: 11 });
  const inside = canvasPointerPosition(pointerAt(192, 864), frame);
  expect(inside.model).toEqual({ x: 2, y: 2 });
});

test('keeps the pointer mapping stable while the canvas is zoomed', () => {
  const zoomed = canvasPointerPosition(pointerAt(384, 1728, 2), frame);
  expect(zoomed.model).toEqual({ x: 2, y: 2 });
});

test('formats inch formulas without exponent noise or negative zero', () => {
  expect(inchFormula(2)).toBe('2');
  expect(inchFormula(-0)).toBe('0');
  expect(inchFormula(-1.5)).toBe('-1.5');
  expect(inchFormula(1 / 3)).toBe('0.333333');
  expect(() => inchFormula(Number.NaN)).toThrow('Shape geometry must be finite');
  expect(() => inchFormula(Number.POSITIVE_INFINITY)).toThrow('Shape geometry must be finite');
});

test('moves a shape by the pointer delta instead of teleporting its pin to release position', () => {
  const start = { canvas: { x: 0, y: 0 }, model: { x: 3, y: 3 }, resize: false, pin: { x: 5, y: 2 }, size: { width: 2, height: 1 } };
  const grabbedAwayFromPin = resolveDragGeometry(start, { x: 4, y: 4 });
  expect(grabbedAwayFromPin).toEqual({ x: 6, y: 3, width: 2, height: 1 });
});

test('resizes by adjusting existing dimensions with the pointer delta, not the raw pointer distance', () => {
  const start = { canvas: { x: 0, y: 0 }, model: { x: 3, y: 3 }, resize: true, pin: { x: 5, y: 2 }, size: { width: 2, height: 1 } };
  const grown = resolveDragGeometry(start, { x: 4, y: 5 });
  expect(grown).toEqual({ x: 5, y: 2, width: 3, height: 3 });
  const shrunkBelowMinimum = resolveDragGeometry(start, { x: -50, y: -50 });
  expect(shrunkBelowMinimum.width).toBeCloseTo(0.01, 5);
  expect(shrunkBelowMinimum.height).toBeCloseTo(0.01, 5);
});

test('drops a selection whose shape a peer removed from the page', () => {
  const withChild: DiagramSnapshot = {
    pages: [{ id: 'page', sourcePartPath: 'page', name: 'Page', shapes: [{ id: 'group', sourceId: 1, name: null, children: [{ id: 'inner', sourceId: 2, name: null, children: [], cells: [] }], cells: [] }] }],
  };
  const withoutChild: DiagramSnapshot = {
    pages: [{ id: 'page', sourcePartPath: 'page', name: 'Page', shapes: [{ id: 'group', sourceId: 1, name: null, children: [], cells: [] }] }],
  };
  const selection = { pageId: 'page', shapeId: 'inner', hit: { kind: 'shape' as const, shapeId: 'inner' } };
  expect(stillSelectable(withChild, 0, selection)).toBe(true);
  expect(stillSelectable(withoutChild, 0, selection)).toBe(false);
  expect(stillSelectable(withChild, 1, selection)).toBe(false);
});

test('moves within a rotated and scaled group using the parent coordinates', () => {
  const geometry = resolveDragGeometry({ canvas: { x: 0, y: 0 }, model: { x: 10, y: 20 }, resize: false, pin: { x: 2, y: 3 }, size: { width: 4, height: 5 }, parentTransforms: [{ a: 0, b: 2, c: -2, d: 0, e: 10, f: 20 }] }, { x: 8, y: 24 });
  expect(geometry).toEqual({ x: 4, y: 4, width: 4, height: 5 });
});

test('resizes along the rotated shape axes', () => {
  const geometry = resolveDragGeometry({ canvas: { x: 0, y: 0 }, model: { x: 0, y: 0 }, resize: true, pin: { x: 2, y: 3 }, size: { width: 4, height: 5 }, angle: Math.PI / 2 }, { x: -2, y: 3 });
  expect(geometry.width).toBeCloseTo(7);
  expect(geometry.height).toBeCloseTo(7);
});

test('preview outline and commit geometry agree for the same pointer position', () => {
  const identity = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
  const start = { canvas: { x: 0, y: 0 }, model: { x: 3, y: 3 }, resize: false, pin: { x: 5, y: 2 }, size: { width: 2, height: 1 } };
  const release = { x: 4, y: 4 };
  const geometry = resolveDragGeometry(start, release);
  const corners = previewOutline(start, release, identity);
  expect(corners).toEqual([{ x: 5, y: 2.5 }, { x: 7, y: 2.5 }, { x: 7, y: 3.5 }, { x: 5, y: 3.5 }]);
  const centre = { x: (corners[0].x + corners[2].x) / 2, y: (corners[0].y + corners[2].y) / 2 };
  expect(centre.x).toBeCloseTo(geometry.x, 10); expect(centre.y).toBeCloseTo(geometry.y, 10);
  expect(Math.abs(corners[1].x - corners[0].x)).toBeCloseTo(geometry.width, 10);
  expect(Math.abs(corners[2].y - corners[1].y)).toBeCloseTo(geometry.height, 10);
  const rotated = { canvas: { x: 0, y: 0 }, model: { x: 0, y: 0 }, resize: false, pin: { x: 2, y: 3 }, size: { width: 4, height: 5 }, angle: Math.PI / 2 };
  const rotatedGeometry = resolveDragGeometry(rotated, { x: 0, y: 0 });
  const rotatedCorners = previewOutline(rotated, { x: 0, y: 0 }, identity);
  const rotatedCentre = { x: (rotatedCorners[0].x + rotatedCorners[2].x) / 2, y: (rotatedCorners[0].y + rotatedCorners[2].y) / 2 };
  expect(rotatedCentre.x).toBeCloseTo(rotatedGeometry.x, 10); expect(rotatedCentre.y).toBeCloseTo(rotatedGeometry.y, 10);
  const grouped = { canvas: { x: 0, y: 0 }, model: { x: 10, y: 20 }, resize: false, pin: { x: 2, y: 3 }, size: { width: 4, height: 5 }, parentTransforms: [{ a: 0, b: 2, c: -2, d: 0, e: 10, f: 20 }] };
  const groupedGeometry = resolveDragGeometry(grouped, { x: 8, y: 24 });
  const groupedCorners = previewOutline(grouped, { x: 8, y: 24 }, identity);
  const groupedCentre = { x: (groupedCorners[0].x + groupedCorners[2].x) / 2, y: (groupedCorners[0].y + groupedCorners[2].y) / 2 };
  const forward = (point: { x: number; y: number }) => ({ x: 0 * point.x + -2 * point.y + 10, y: 2 * point.x + 0 * point.y + 20 });
  const expectedCentre = forward({ x: groupedGeometry.x, y: groupedGeometry.y });
  expect(groupedCentre.x).toBeCloseTo(expectedCentre.x, 10); expect(groupedCentre.y).toBeCloseTo(expectedCentre.y, 10);
});
test('a handle resize from nw moves the pin so the se corner stays put', () => {
  const start = { canvas: { x: 0, y: 0 }, model: { x: 0, y: 0 }, resize: false, handle: 'nw' as const, pin: { x: 5, y: 2 }, locPin: { x: 1, y: 0.5 }, size: { width: 2, height: 1 } };
  const geometry = resolveDragGeometry(start, { x: -1, y: 1 });
  expect(geometry.width).toBeCloseTo(3, 10);
  expect(geometry.height).toBeCloseTo(2, 10);
  expect(geometry.x).toBeCloseTo(4, 10);
  expect(geometry.y).toBeCloseTo(2, 10);
  expect(geometry.x - 1 + geometry.width).toBeCloseTo(6, 10);
  expect(geometry.y - 0.5).toBeCloseTo(1.5, 10);
  const identity = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
  const corners = previewOutline(start, { x: -1, y: 1 }, identity);
  const se = { x: (corners[1].x + corners[1].x) / 2, y: corners[1].y };
  expect(se.x).toBeCloseTo(6, 8);
  expect(corners[1].y).toBeCloseTo(1.5, 8);
  expect(corners[0].y).toBeCloseTo(1.5, 8);
});
test('a rotate drag produces the expected angle with optional shift snap', () => {
  const start = { canvas: { x: 0, y: 0 }, model: { x: 1, y: 0 }, resize: false, rotate: true, pin: { x: 0, y: 0 }, size: { width: 2, height: 1 }, angle: 0 };
  expect(resolveRotationAngle(start, { x: 0, y: 1 })).toBeCloseTo(Math.PI / 2, 10);
  const seventeen = { x: Math.cos(17 * Math.PI / 180), y: Math.sin(17 * Math.PI / 180) };
  expect(resolveRotationAngle(start, seventeen)).toBeCloseTo(17 * Math.PI / 180, 10);
  expect(resolveRotationAngle(start, seventeen, true)).toBeCloseTo(15 * Math.PI / 180, 10);
});
test('selection corners resolve the rotated box in canvas coordinates', () => {
  const page = {
    id: 'page',
    sourcePartPath: 'page',
    name: 'Page',
    shapes: [{ id: 'shape', sourceId: 1, name: null, children: [], cells: [
      { locator: { sheet: 'document' as const, shapeId: null, section: null, row: null, cellName: 'PinX' }, name: 'PinX', formula: null, value: '2' },
      { locator: { sheet: 'document' as const, shapeId: null, section: null, row: null, cellName: 'PinY' }, name: 'PinY', formula: null, value: '3' },
      { locator: { sheet: 'document' as const, shapeId: null, section: null, row: null, cellName: 'Width' }, name: 'Width', formula: null, value: '20' },
      { locator: { sheet: 'document' as const, shapeId: null, section: null, row: null, cellName: 'Height' }, name: 'Height', formula: null, value: '10' },
      { locator: { sheet: 'document' as const, shapeId: null, section: null, row: null, cellName: 'LocPinX' }, name: 'LocPinX', formula: null, value: '10' },
      { locator: { sheet: 'document' as const, shapeId: null, section: null, row: null, cellName: 'LocPinY' }, name: 'LocPinY', formula: null, value: '5' },
      { locator: { sheet: 'document' as const, shapeId: null, section: null, row: null, cellName: 'Angle' }, name: 'Angle', formula: null, value: String(Math.PI / 2) },
    ] }],
  };
  const corners = selectionCorners(page as never, frame, { pageId: 'page', shapeId: 'shape', hit: { kind: 'shape', shapeId: 'shape' } });
  expect(corners).not.toBeNull();
  const centre = { x: (corners![0].x + corners![2].x) / 2, y: (corners![0].y + corners![2].y) / 2 };
  expect(centre.x).toBeCloseTo(2 * 96, 4);
  expect(centre.y).toBeCloseTo(-3 * 96 + 1056, 4);
});
test('a handle resize honours LocPinX at the edge and a non-centred LocPinY', () => {
  const edge = { canvas: { x: 0, y: 0 }, model: { x: 0, y: 0 }, resize: false, handle: 'e' as const, pin: { x: 5, y: 2 }, locPin: { x: 0, y: 0.5 }, size: { width: 2, height: 1 } };
  const grown = resolveDragGeometry(edge, { x: 1, y: 0 });
  expect(grown.width).toBeCloseTo(3, 10);
  expect(grown.x).toBeCloseTo(5, 10);
  expect(grown.y).toBeCloseTo(2, 10);
  const identity = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
  const corners = previewOutline(edge, { x: 1, y: 0 }, identity);
  expect(Math.min(...corners.map((corner) => corner.x))).toBeCloseTo(5, 8);
  expect(Math.max(...corners.map((corner) => corner.x))).toBeCloseTo(8, 8);
  const low = { canvas: { x: 0, y: 0 }, model: { x: 0, y: 0 }, resize: false, handle: 'n' as const, pin: { x: 5, y: 2 }, locPin: { x: 1, y: 0 }, size: { width: 2, height: 1 } };
  const grownNorth = resolveDragGeometry(low, { x: 0, y: 1 });
  expect(grownNorth.height).toBeCloseTo(2, 10);
  expect(grownNorth.y).toBeCloseTo(2, 10);
  const northCorners = previewOutline(low, { x: 0, y: 1 }, identity);
  expect(Math.min(...northCorners.map((corner) => corner.y))).toBeCloseTo(2, 8);
  expect(Math.max(...northCorners.map((corner) => corner.y))).toBeCloseTo(4, 8);
});
test('a flipped handle resize grows outward on both axes', () => {
  const flipX = { canvas: { x: 0, y: 0 }, model: { x: 0, y: 0 }, resize: false, handle: 'e' as const, pin: { x: 0, y: 0 }, size: { width: 20, height: 10 }, flipX: true };
  expect(resolveDragGeometry(flipX, { x: 2, y: 0 }).width).toBeCloseTo(22, 10);
  const flipY = { canvas: { x: 0, y: 0 }, model: { x: 0, y: 0 }, resize: false, handle: 'n' as const, pin: { x: 0, y: 0 }, size: { width: 20, height: 10 }, flipY: true };
  expect(resolveDragGeometry(flipY, { x: 0, y: 2 }).height).toBeCloseTo(12, 10);
});
test('a nudge inside a rotated and scaled group matches the equivalent drag', () => {
  const parentTransforms = [{ a: 0, b: 2, c: -2, d: 0, e: 10, f: 20 }];
  const base = { canvas: { x: 0, y: 0 }, model: { x: 10, y: 20 }, resize: false, pin: { x: 2, y: 3 }, size: { width: 4, height: 5 }, parentTransforms };
  const dx = 1 / 96;
  const dy = 0;
  const dragged = resolveDragGeometry(base, { x: 10 + dx, y: 20 + dy });
  const nudged = resolveNudgeGeometry(base, dx, dy);
  expect(nudged.x).toBeCloseTo(dragged.x, 10);
  expect(nudged.y).toBeCloseTo(dragged.y, 10);
  expect(nudged.x).not.toBeCloseTo(2 + dx, 6);
});
