import { expect, test } from 'bun:test';
import type { DiagramSnapshot, PageDisplayList } from '@betteroffice/vsdx';
import type { PointerEvent } from 'react';
import { canvasPointerPosition, inchFormula, resolveDragGeometry, stillSelectable } from './VsdxEditor';

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
  expect(inchFormula(Number.NaN)).toBe('0');
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
