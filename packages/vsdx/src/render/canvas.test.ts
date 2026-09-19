import { expect, test } from 'bun:test';
import { MAX_CANVAS_AREA, MAX_CANVAS_DIMENSION, canvasPointToModel, effectiveDprForSurface, modelPointToCanvas, paintPage, sizeCanvasForPage } from './canvas';
import type { PageDisplayList, ShapePrimitive } from '../types';

function context(log: string[]): CanvasRenderingContext2D {
  return new Proxy({
    createLinearGradient: () => ({ addColorStop: () => {} }),
  }, {
    get(target, key) {
      if (key in target) return Reflect.get(target, key);
      return (...args: unknown[]) => { log.push(`${String(key)}:${args.join(',')}`); };
    },
    set(_, key, value) { log.push(`${String(key)}=${String(value)}`); return true; },
  }) as unknown as CanvasRenderingContext2D;
}

const transform = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };

test('replays primitives in z order and paints placeholders', async () => {
  const log: string[] = [];
  const list: PageDisplayList = {
    contractVersion: 7, width: 100, height: 100, printWidth: 100, printHeight: 100, paintTransform: transform,
    primitives: [
      { kind: 'placeholder', id: 'late', zOrder: 2, x: 10, y: 10, width: 20, height: 20, reason: 'missing image' },
      { kind: 'shape', id: 'early', zOrder: 1, path: [{ type: 'move', x: 0, y: 0 }, { type: 'line', x: 1, y: 1 }], fill: { kind: 'solid', color: '#000' } },
    ],
  };
  await paintPage(context(log), list);
  expect(log.findIndex(entry => entry.startsWith('moveTo'))).toBeLessThan(log.findIndex(entry => entry.startsWith('strokeRect')));
  expect(log.some(entry => entry.startsWith('strokeRect'))).toBe(true);
  expect(log.some(entry => entry.startsWith('fillText:missing image'))).toBe(true);
});

test('rejects display-list versions other than v7', async () => {
  await expect(paintPage(context([]), { contractVersion: 2, width: 1, height: 1, paintTransform: transform, primitives: [] } as unknown as PageDisplayList)).rejects.toThrow('unsupported VSDX display-list contract version 2');
});

test('replays positioned text runs at their line caret positions', async () => {
  const log: string[] = [];
  const list: PageDisplayList = {
    contractVersion: 7, width: 100, height: 100, printWidth: 100, printHeight: 100, paintTransform: transform,
    primitives: [{
      kind: 'textBox', id: 'text', zOrder: 1, x: 1, y: 2, width: 90, height: 80,
      paragraphs: [
        { runs: [{ text: 'left', family: 'Arial', sizeIn: 12, bold: false, italic: false, underline: false, smallCaps: false, superscript: false, subscript: false, letterSpacing: 0, color: '#111', diagnostics: [] }] },
        { runs: [{ text: 'right', family: 'Arial', sizeIn: 12, bold: true, italic: false, underline: false, smallCaps: false, superscript: false, subscript: false, letterSpacing: 0, color: '#222', diagnostics: [] }] },
      ],
      lines: [
        { x: 30, y: 20, width: 20, height: 12, start: 0, end: 4, caretStops: [{ position: 0, x: 30, y: 20 }, { position: 4, x: 50, y: 20 }] },
        { x: 60, y: 45, width: 25, height: 12, start: 4, end: 9, caretStops: [{ position: 4, x: 60, y: 45 }, { position: 9, x: 85, y: 45 }] },
      ],
    }],
  };
  await paintPage(context(log), list);
  expect(log).toContain('translate:0,84');
  expect(log).toContain('scale:1,-1');
  expect(log).toContain('textBaseline=top');
  expect(log.filter(entry => entry.startsWith('fillText:'))).toEqual(['fillText:left,30,20', 'fillText:right,60,45']);
  expect(log.some(entry => entry.startsWith('clip'))).toBe(false);
});

test('a Letter page at zoom 4 clamps its backing store instead of its CSS size', () => {
  const canvas = { width: 0, height: 0, style: { width: '', height: '' } };
  const effective = sizeCanvasForPage(canvas, { width: 816, height: 1056 }, 2, 4);
  expect(effective).toBe(effectiveDprForSurface(816 * 4, 1056 * 4, 2));
  expect(effective).toBeLessThan(2);
  expect(canvas.style.width).toBe('3264px');
  expect(canvas.style.height).toBe('4224px');
  expect(canvas.width).toBeLessThanOrEqual(MAX_CANVAS_DIMENSION);
  expect(canvas.height).toBeLessThanOrEqual(MAX_CANVAS_DIMENSION);
  expect(canvas.width * canvas.height).toBeLessThanOrEqual(MAX_CANVAS_AREA);
});

test('the area budget keeps a zoomed page near 128 MiB per canvas', () => {
  const canvas = { width: 0, height: 0, style: { width: '', height: '' } };
  sizeCanvasForPage(canvas, { width: 816, height: 1056 }, 2, 4);
  expect((canvas.width * canvas.height * 4) / 1048576).toBeLessThanOrEqual(128);
});

test('a normal page at zoom 1 keeps its full backing store', () => {
  const canvas = { width: 0, height: 0, style: { width: '', height: '' } };
  const effective = sizeCanvasForPage(canvas, { width: 816, height: 1056 }, 2, 1);
  expect(effective).toBe(2);
  expect(canvas.width).toBe(1632);
  expect(canvas.height).toBe(2112);
  expect(canvas.style.width).toBe('816px');
  expect(canvas.style.height).toBe('1056px');
});

const pagePaintTransform = { a: 96, b: 0, c: 0, d: -96, e: 0, f: 768 };

test('canvasPointToModel inverts the page paint transform onto Y-up inches', () => {
  expect(canvasPointToModel(pagePaintTransform, 0, 768)).toEqual({ x: 0, y: 0 });
  expect(canvasPointToModel(pagePaintTransform, 0, 0)).toEqual({ x: 0, y: 8 });
  expect(canvasPointToModel(pagePaintTransform, 96, 672)).toEqual({ x: 1, y: 1 });
});

test('canvasPointToModel divides out the canvas scale', () => {
  expect(canvasPointToModel(pagePaintTransform, 192, 1344, 2)).toEqual({ x: 1, y: 1 });
});

test('canvasPointToModel rejects a degenerate transform and a non-positive scale', () => {
  expect(() => canvasPointToModel({ a: 0, b: 0, c: 0, d: 0, e: 0, f: 0 }, 1, 1)).toThrow();
  expect(() => canvasPointToModel(pagePaintTransform, 1, 1, 0)).toThrow();
});

test('modelPointToCanvas round-trips through canvasPointToModel', () => {
  const points = [{ x: 0, y: 0 }, { x: 1, y: 1 }, { x: 2.5, y: -3.25 }, { x: -1, y: 8 }];
  const rotated = { a: 0, b: 2, c: -2, d: 0, e: 10, f: 20 };
  for (const point of points) {
    const canvas = modelPointToCanvas(pagePaintTransform, point.x, point.y);
    const back = canvasPointToModel(pagePaintTransform, canvas.x, canvas.y);
    expect(back.x).toBeCloseTo(point.x, 10); expect(back.y).toBeCloseTo(point.y, 10);
    const scaled = modelPointToCanvas(pagePaintTransform, point.x, point.y, 2);
    const backScaled = canvasPointToModel(pagePaintTransform, scaled.x, scaled.y, 2);
    expect(backScaled.x).toBeCloseTo(point.x, 10); expect(backScaled.y).toBeCloseTo(point.y, 10);
    const turned = modelPointToCanvas(rotated, point.x, point.y);
    const backTurned = canvasPointToModel(rotated, turned.x, turned.y);
    expect(backTurned.x).toBeCloseTo(point.x, 10); expect(backTurned.y).toBeCloseTo(point.y, 10);
    const turnedScaled = modelPointToCanvas(rotated, point.x, point.y, 2);
    const backTurnedScaled = canvasPointToModel(rotated, turnedScaled.x, turnedScaled.y, 2);
    expect(backTurnedScaled.x).toBeCloseTo(point.x, 10); expect(backTurnedScaled.y).toBeCloseTo(point.y, 10);
  }
  expect(modelPointToCanvas(pagePaintTransform, 1, 1)).toEqual({ x: 96, y: 672 });
  expect(() => modelPointToCanvas(pagePaintTransform, 1, 1, 0)).toThrow('VSDX canvas scale must be a positive number');
  expect(() => modelPointToCanvas(pagePaintTransform, 1, 1, Number.NaN)).toThrow('VSDX canvas scale must be a positive number');
});

test('a delayed image cannot overwrite a newer page or disturb its canvas state', async () => {
  const log: string[] = [];
  const ctx = context(log);
  let finish: (image: CanvasImageSource) => void = () => {};
  const oldPage: PageDisplayList = { contractVersion: 7, width: 100, height: 100, printWidth: 100, printHeight: 100, paintTransform: transform, primitives: [{ kind: 'image', id: 'old', zOrder: 0, assetId: 'slow', x: 0, y: 0, width: 1, height: 1 }] };
  const oldPaint = paintPage(ctx, oldPage, 1, 1, { resolveImage: () => new Promise(resolve => { finish = resolve; }) });
  expect(log).toEqual([]);
  await paintPage(ctx, { ...oldPage, primitives: [] });
  const current = [...log];
  finish({} as CanvasImageSource);
  await oldPaint;
  expect(log).toEqual(current);
});

test('an aborted page never touches the canvas after its images load', async () => {
  const log: string[] = [];
  const controller = new AbortController();
  controller.abort();
  await paintPage(context(log), { contractVersion: 7, width: 1, height: 1, printWidth: 1, printHeight: 1, paintTransform: transform, primitives: [] }, 1, 1, { signal: controller.signal });
  expect(log).toEqual([]);
});


test('places the top of an image above its bottom in a Y-up diagram', async () => {
  let yScale = 1, yOffset = 0;
  const stack: Array<[number, number]> = [];
  let top: number | undefined;
  let bottom: number | undefined;
  const ctx = {
    clearRect: () => {},
    save: () => { stack.push([yScale, yOffset]); },
    restore: () => { [yScale, yOffset] = stack.pop()!; },
    setTransform: (_a: number, _b: number, _c: number, d: number, _e: number, f: number) => { yScale = d; yOffset = f; },
    transform: (_a: number, _b: number, _c: number, d: number, _e: number, f: number) => { yOffset += yScale * f; yScale *= d; },
    translate: (_x: number, y: number) => { yOffset += yScale * y; },
    scale: (_x: number, y: number) => { yScale *= y; },
    drawImage: (_source: unknown, _x: number, y: number, _width: number, height: number) => {
      top = y * yScale + yOffset;
      bottom = (y + height) * yScale + yOffset;
    },
  } as unknown as CanvasRenderingContext2D;
  await paintPage(ctx, { contractVersion: 7, width: 192, height: 192, printWidth: 192, printHeight: 192, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 192 }, primitives: [{ kind: 'image', id: 'picture', assetId: 'picture', zOrder: 0, x: 0, y: 0, width: 2, height: 2 }] }, 1, 1, { resolveImage: () => ({} as CanvasImageSource) });
  expect(top).toBe(0);
  expect(bottom).toBe(192);
});

test('paints a linear gradient across the shape box along its angle', async () => {
  const gradients: Array<{ args: number[]; stops: Array<[number, string]> }> = [];
  const painted: unknown[] = [];
  let fillStyle: unknown;
  const ctx = {
    save: () => {}, restore: () => {}, setTransform: () => {}, clearRect: () => {}, beginPath: () => {}, moveTo: () => {}, lineTo: () => {}, closePath: () => {}, fill: () => {},
    transform: () => {},
    createLinearGradient: (...args: number[]) => {
      const entry = { args, stops: [] as Array<[number, string]> };
      gradients.push(entry);
      const gradient = { addColorStop: (position: number, color: string) => { entry.stops.push([position, color]); } };
      painted.push(gradient);
      return gradient;
    },
    set fillStyle(value: unknown) { fillStyle = value; },
    get fillStyle() { return fillStyle; },
  } as unknown as CanvasRenderingContext2D;
  const list: PageDisplayList = {
    contractVersion: 7, width: 100, height: 100, printWidth: 100, printHeight: 100, paintTransform: transform,
    primitives: [{
      kind: 'shape', id: 'graded', zOrder: 0,
      path: [{ type: 'move', x: 0, y: 0 }, { type: 'line', x: 2, y: 0 }, { type: 'line', x: 2, y: 1 }, { type: 'close' }],
      fill: { kind: 'gradient', angleDeg: 0, stops: [{ position: 0, color: '#ff0000' }, { position: 1, color: '#0000ff' }] },
    }],
  };
  await paintPage(ctx, list);
  expect(gradients).toHaveLength(1);
  expect(gradients[0].args).toEqual([0, 0.5, 2, 0.5]);
  expect(gradients[0].stops).toEqual([[0, '#ff0000'], [1, '#0000ff']]);
  expect(fillStyle).toBe(painted[0]);

  gradients.length = 0;
  const upright: PageDisplayList = { ...list, primitives: [{ ...(list.primitives[0] as ShapePrimitive), fill: { kind: 'gradient', angleDeg: 90, stops: [{ position: 0, color: '#ff0000' }, { position: 1, color: '#0000ff' }] } }] };
  await paintPage(ctx, upright);
  for (const [index, expected] of [1, 0, 1, 1].entries()) expect(gradients[0].args[index]).toBeCloseTo(expected, 10);
});

const shadow = { color: '#11223380', blurIn: 0.5, offsetXIn: 0.125, offsetYIn: -0.125 };

function shadowedPage(primitives: PageDisplayList['primitives']): PageDisplayList {
  return { contractVersion: 7, width: 768, height: 768, printWidth: 768, printHeight: 768, paintTransform: pagePaintTransform, primitives };
}

test('casts a shape shadow in device pixels and clears it before the stroke', async () => {
  const log: string[] = [];
  await paintPage(context(log), shadowedPage([{
    kind: 'shape', id: 'boxed', zOrder: 0, path: [{ type: 'move', x: 0, y: 0 }, { type: 'line', x: 1, y: 1 }, { type: 'close' }],
    fill: { kind: 'solid', color: '#ffffff' }, stroke: { color: '#000000', width: 1, dashed: false }, shadow,
  }]));
  expect(log).toContain('shadowColor=#11223380');
  expect(log).toContain('shadowOffsetX=12');
  expect(log).toContain('shadowOffsetY=12');
  expect(log).toContain('shadowBlur=48');
  expect(log.indexOf('shadowColor=#11223380')).toBeLessThan(log.findIndex(entry => entry.startsWith('fill:')));
  expect(log.findIndex(entry => entry === 'shadowBlur=0')).toBeLessThan(log.findIndex(entry => entry.startsWith('stroke:')));
});

test('a shape shadow follows the device scale and the group transform', async () => {
  const log: string[] = [];
  const shaded: ShapePrimitive = { kind: 'shape', id: 'boxed', zOrder: 0, path: [{ type: 'move', x: 0, y: 0 }], fill: { kind: 'solid', color: '#ffffff' }, shadow };
  await paintPage(context(log), shadowedPage([shaded]), 2, 3);
  expect(log).toContain('shadowOffsetX=72');
  expect(log).toContain('shadowBlur=288');
  const nested: string[] = [];
  await paintPage(context(nested), shadowedPage([{ kind: 'group', id: 'g', zOrder: 0, transform: { a: 2, b: 0, c: 0, d: 2, e: 0, f: 0 }, primitives: [shaded] }]));
  expect(nested).toContain('shadowOffsetX=24');
  expect(nested).toContain('shadowBlur=96');
});

test('a shape without a fill strokes with its shadow, and never casts it twice', async () => {
  const log: string[] = [];
  await paintPage(context(log), shadowedPage([{
    kind: 'shape', id: 'connector', zOrder: 0, path: [{ type: 'move', x: 0, y: 0 }, { type: 'line', x: 1, y: 1 }],
    stroke: { color: '#000000', width: 1, dashed: false }, shadow,
  }]));
  expect(log.some(entry => entry === 'shadowBlur=0')).toBe(false);
  expect(log.indexOf('shadowColor=#11223380')).toBeLessThan(log.findIndex(entry => entry.startsWith('stroke:')));
  expect(log.filter(entry => entry === 'shadowColor=#11223380')).toHaveLength(1);
});

test('an unshadowed shape never touches the canvas shadow state', async () => {
  const log: string[] = [];
  await paintPage(context(log), shadowedPage([{ kind: 'shape', id: 'plain', zOrder: 0, path: [{ type: 'move', x: 0, y: 0 }], fill: { kind: 'solid', color: '#ffffff' } }]));
  expect(log.some(entry => entry.startsWith('shadow'))).toBe(false);
});

test('a degenerate gradient box falls back to its first stop', async () => {
  let fillStyle: unknown;
  const ctx = {
    save: () => {}, restore: () => {}, setTransform: () => {}, clearRect: () => {}, beginPath: () => {}, moveTo: () => {}, fill: () => {},
    transform: () => {},
    createLinearGradient: () => { throw new Error('must not create a gradient for a point box'); },
    set fillStyle(value: unknown) { fillStyle = value; },
    get fillStyle() { return fillStyle; },
  } as unknown as CanvasRenderingContext2D;
  await paintPage(ctx, {
    contractVersion: 7, width: 100, height: 100, printWidth: 100, printHeight: 100, paintTransform: transform,
    primitives: [{
      kind: 'shape', id: 'point', zOrder: 0,
      path: [{ type: 'move', x: 3, y: 4 }],
      fill: { kind: 'gradient', angleDeg: 90, stops: [{ position: 0, color: '#112233' }, { position: 1, color: '#445566' }] },
    }],
  });
  expect(fillStyle).toBe('#112233');
});
