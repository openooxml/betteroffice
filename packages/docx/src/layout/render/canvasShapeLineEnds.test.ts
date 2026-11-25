import { expect, test } from 'bun:test';
import { drawPrimitive } from './canvasBackend';
import type { ShapePrimitive } from './displayList';

test('connector line ends follow the first and last nonzero tangents', async () => {
  const bent: ShapePrimitive['geometryPath'] = [
    { type: 'move', x: 0, y: 0 },
    { type: 'line', x: 100, y: 0 },
    { type: 'line', x: 100, y: 0 },
  ];
  const curved: ShapePrimitive['geometryPath'] = [
    { type: 'move', x: 0, y: 0 },
    { type: 'cubic', cp1x: 50, cp1y: 0, cp2x: 100, cp2y: 40, x: 100, y: 80 },
  ];
  for (const [geometryPath, height, expected] of [
    [bent, 0, [[[0, 0], [8, 4], [8, -4]], [[100, 0], [92, -4], [92, 4]]]],
    [curved, 80, [[[0, 0], [8, 4], [8, -4]], [[100, 80], [104, 72], [96, 72]]]],
  ] as const) {
    let path: number[][] = [];
    const filled: number[][][] = [];
    const ctx = {
      beginPath() { path = []; },
      moveTo(x: number, y: number) { path.push([x, y]); },
      lineTo(x: number, y: number) { path.push([x, y]); },
      bezierCurveTo() {},
      closePath() {},
      save() {},
      restore() {},
      setLineDash() {},
      stroke() {},
      fill() { filled.push(path); },
    } as unknown as CanvasRenderingContext2D;
    await drawPrimitive(ctx, {
      kind: 'shape', x: 0, y: 0, w: 100, h: height, geometryPath,
      fillPaint: { kind: 'none' },
      strokePaint: { width: 2, headEnd: { type: 'triangle' }, tailEnd: { type: 'triangle' } },
    });
    expect(filled.map((points) => points.map(([x, y]) => [Math.round(x), Math.round(y)])))
      .toEqual(expected.map((head) => head.map((point) => [...point])));
  }
});
