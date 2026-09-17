import { expect, test } from 'bun:test';
import type { PageDisplayList, ShapeSnapshot } from '@betteroffice/vsdx';
import { collectSnapTargets, paintGrid, paintSmartGuides, snapRelease, snapThresholdModel } from './snap';

const dragStart = { model: { x: 0, y: 0 }, pin: { x: 1, y: 1 }, locPin: { x: 0.5, y: 0.5 }, size: { width: 1, height: 1 } };

const live = (overrides = {}) => ({
  zoom: 1,
  snapEnabled: true,
  suppressed: false,
  isRotate: false,
  targets: { x: [], y: [] },
  ...overrides,
});

const shape = (id: string, pinX: number, pinY: number, width: number, height: number): ShapeSnapshot => {
  const cell = (name: string, value: number) => ({ locator: { sheet: { page: 1 }, shapeId: 1, section: null, row: null, cellName: name }, name, formula: null, value: String(value) });
  return { id, sourceId: 1, name: id, cells: [cell('PinX', pinX), cell('PinY', pinY), cell('Width', width), cell('Height', height), cell('LocPinX', width / 2), cell('LocPinY', height / 2)], children: [] };
};

test('snap threshold is six screen pixels in model inches', () => {
  expect(snapThresholdModel(1)).toBeCloseTo(6 / 96, 12);
  expect(snapThresholdModel(2)).toBeCloseTo(6 / 192, 12);
  expect(snapThresholdModel(0.5)).toBeCloseTo(6 / 48, 12);
});

test('drag snaps to the quarter-inch grid without guides', () => {
  const result = snapRelease(dragStart, { x: 0.06, y: 0 }, live());
  expect(result.point.x).toBeCloseTo(0, 10);
  expect(result.point.y).toBeCloseTo(0, 10);
  expect(result.guides).toEqual({ x: [], y: [] });
});

test('drag outside the threshold is untouched', () => {
  const result = snapRelease(dragStart, { x: 0.1, y: 0 }, live());
  expect(result.point).toEqual({ x: 0.1, y: 0 });
  expect(result.guides).toEqual({ x: [], y: [] });
});

test('shape edges win ties with the grid and raise guides', () => {
  const targets = { x: [2], y: [2] };
  const result = snapRelease(dragStart, { x: 0.5, y: 0.5 }, live({ targets }));
  expect(result.point.x).toBeCloseTo(0.5, 10);
  expect(result.guides).toEqual({ x: [2], y: [2] });
});

test('shape centres snap and raise guides', () => {
  const result = snapRelease(dragStart, { x: 0.03, y: 0 }, live({ targets: { x: [1], y: [] } }));
  expect(result.point.x).toBeCloseTo(0, 10);
  expect(result.guides).toEqual({ x: [1], y: [] });
});

test('rotate, suppress and disable bypass snapping', () => {
  const release = { x: 0.06, y: 0 };
  expect(snapRelease(dragStart, release, live({ isRotate: true })).point).toEqual(release);
  expect(snapRelease(dragStart, release, live({ suppressed: true })).point).toEqual(release);
  expect(snapRelease(dragStart, release, live({ snapEnabled: false })).point).toEqual(release);
});

test('targets skip the dragged shape', () => {
  const targets = collectSnapTargets([shape('a', 1, 1, 1, 1), shape('b', 5, 5, 2, 2)], 'a');
  expect(targets.x).toEqual([4, 5, 6]);
  expect(targets.y).toEqual([4, 5, 6]);
});

const frame = (): PageDisplayList => ({
  contractVersion: 7,
  width: 816,
  height: 1056,
  printWidth: 816,
  printHeight: 1056,
  paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 1056 },
  primitives: [],
});

const recordingContext = () => {
  const calls: string[] = [];
  const context = new Proxy({}, {
    get(_target, key) {
      return (...args: unknown[]) => { calls.push(`${String(key)}:${args.join(',')}`); };
    },
    set(target, key, value) { calls.push(`${String(key)}=${String(value)}`); Reflect.set(target, key, value); return true; },
  }) as unknown as CanvasRenderingContext2D;
  return { context, calls };
};

test('paintGrid strokes vertical and horizontal lines', () => {
  const { context, calls } = recordingContext();
  paintGrid(context, frame(), 1, 1);
  expect(calls).toContain('setTransform:1,0,0,1,0,0');
  expect(calls.filter((entry) => entry.startsWith('stroke:')).length).toBeGreaterThan(10);
});

test('paintGrid skips dense minors when zoomed out', () => {
  const { context, calls } = recordingContext();
  paintGrid(context, frame(), 1, 0.01);
  expect(calls.some((entry) => entry.startsWith('stroke:'))).toBe(false);
});

test('paintSmartGuides draws nothing without guides and lines with them', () => {
  const empty = recordingContext();
  paintSmartGuides(empty.context, frame(), 1, 1, { x: [], y: [] });
  expect(empty.calls.some((entry) => entry.startsWith('stroke:'))).toBe(false);
  const { context, calls } = recordingContext();
  paintSmartGuides(context, frame(), 1, 1, { x: [2], y: [3] });
  expect(calls.filter((entry) => entry.startsWith('stroke:'))).toHaveLength(2);
});
