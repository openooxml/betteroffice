import { expect, test } from 'bun:test';
import { canvasPointToModel, paintPage } from './canvas';
import type { PageDisplayList } from '../types';

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
    contractVersion: 4, width: 100, height: 100, paintTransform: transform,
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

test('rejects display-list versions other than v4', async () => {
  await expect(paintPage(context([]), { contractVersion: 2, width: 1, height: 1, paintTransform: transform, primitives: [] } as unknown as PageDisplayList)).rejects.toThrow('unsupported VSDX display-list contract version 2');
});

test('replays positioned text runs at their line caret positions', async () => {
  const log: string[] = [];
  const list: PageDisplayList = {
    contractVersion: 4, width: 100, height: 100, paintTransform: transform,
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

test('a delayed image cannot overwrite a newer page or disturb its canvas state', async () => {
  const log: string[] = [];
  const ctx = context(log);
  let finish: (image: CanvasImageSource) => void = () => {};
  const oldPage: PageDisplayList = { contractVersion: 4, width: 100, height: 100, paintTransform: transform, primitives: [{ kind: 'image', id: 'old', zOrder: 0, assetId: 'slow', x: 0, y: 0, width: 1, height: 1 }] };
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
  await paintPage(context(log), { contractVersion: 4, width: 1, height: 1, paintTransform: transform, primitives: [] }, 1, 1, { signal: controller.signal });
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
  await paintPage(ctx, { contractVersion: 4, width: 192, height: 192, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 192 }, primitives: [{ kind: 'image', id: 'picture', assetId: 'picture', zOrder: 0, x: 0, y: 0, width: 2, height: 2 }] }, 1, 1, { resolveImage: () => ({} as CanvasImageSource) });
  expect(top).toBe(0);
  expect(bottom).toBe(192);
});
