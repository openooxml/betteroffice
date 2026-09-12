import { expect, test } from 'bun:test';
import type { DiagramSnapshot, PageDisplayList } from '@betteroffice/vsdx';
import type { PointerEvent } from 'react';
import { canvasPointerPosition, inchFormula, stillSelectable } from './VsdxEditor';

const frame: PageDisplayList = {
  contractVersion: 3,
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
