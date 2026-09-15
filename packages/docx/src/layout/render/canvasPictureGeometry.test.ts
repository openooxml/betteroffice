import { expect, test } from 'bun:test';
import { drawPrimitive } from './canvasBackend';
import type { ImagePrimitive } from './displayList';

type Call = [string, ...unknown[]];

function recordingContext() {
  const calls: Call[] = [];
  const images: { arguments: unknown[]; clips: Call[] }[] = [];
  const strokes: Call[][] = [];
  let clips: Call[] = [];
  const stack: Call[][] = [];
  let path: Call = ['empty'];
  const ctx = new Proxy({
    globalAlpha: 1,
    save() { stack.push([...clips]); calls.push(['save']); },
    restore() { clips = stack.pop()!; calls.push(['restore']); },
    beginPath() { path = ['empty']; calls.push(['beginPath']); },
    ellipse(...args: unknown[]) { path = ['ellipse', ...args]; calls.push(path); },
    rect(...args: unknown[]) { path = ['rect', ...args]; calls.push(path); },
    clip() { clips.push(path); calls.push(['clip']); },
    drawImage(...args: unknown[]) { images.push({ arguments: args, clips: [...clips] }); calls.push(['drawImage']); },
    stroke() { strokes.push([...clips]); calls.push(['stroke']); },
  }, {
    get(target, key) {
      if (key in target) return target[key as keyof typeof target];
      return (...args: unknown[]) => calls.push([String(key), ...args]);
    },
  }) as unknown as CanvasRenderingContext2D;
  return { ctx, calls, images, strokes, depth: () => stack.length };
}

function image(shapeType?: string): ImagePrimitive {
  return { kind: 'image', relId: 'picture', x: 10, y: 20, w: 100, h: 80, shapeType };
}

const bitmap = { width: 200, height: 100 } as CanvasImageSource;
const options = { resolveImage: async () => bitmap };

test('ellipse images intersect their outer clip and do not clip later pictures', async () => {
  const recorder = recordingContext();
  await drawPrimitive(recorder.ctx, {
    ...image('ellipse'),
    clipGroup: { clip: { x: 15, y: 25, w: 80, h: 50 } },
    border: { width: 4, color: '#123456' },
    revision: { kind: 'del', author: '', date: '', revisionId: '1' },
  }, options);
  expect(recorder.images[0].clips).toEqual([
    ['rect', 15, 25, 80, 50],
    ['ellipse', 60, 60, 50, 40, 0, 0, 2 * Math.PI],
  ]);
  expect(recorder.strokes).toEqual([
    [['rect', 15, 25, 80, 50]],
    [['rect', 15, 25, 80, 50]],
  ]);
  expect(recorder.calls).toContainEqual(['strokeRect', 10, 20, 100, 80]);
  await drawPrimitive(recorder.ctx, image(), options);
  expect(recorder.images[1].clips).toEqual([]);
  expect(recorder.depth()).toBe(0);
});

test('ellipse clipping uses the content frame after image transforms and preserves source crop', async () => {
  const recorder = recordingContext();
  await drawPrimitive(recorder.ctx, {
    ...image('ellipse'),
    contentFrame: { x: 30, y: 40, w: 120, h: 60 },
    rotationDeg: 30,
    flipH: true,
    crop: { left: 0.1, top: 0.2, right: 0.15, bottom: 0.3 },
    border: { width: 2, color: '#123456' },
  }, options);
  const ellipse: Call = ['ellipse', 90, 70, 60, 30, 0, 0, 2 * Math.PI];
  expect(recorder.images[0]).toEqual({
    arguments: [bitmap, 20, 20, 150, 50, 30, 40, 120, 60],
    clips: [ellipse],
  });
  expect(recorder.calls.indexOf(recorder.calls.find((call) => call[0] === 'ellipse')!))
    .toBeGreaterThan(recorder.calls.findIndex((call) => call[0] === 'scale'));
  expect(recorder.calls).toContainEqual(['translate', 90, 70]);
  expect(recorder.calls).toContainEqual(['rotate', Math.PI / 6]);
  expect(recorder.calls).toContainEqual(['scale', -1, 1]);
  expect(recorder.calls.filter((call) => call[0] === 'ellipse')).toEqual([ellipse, ellipse]);
  expect(recorder.strokes).toEqual([[]]);
  expect(recorder.depth()).toBe(0);
});

test.each([undefined, 'rect', 'roundRect'])('preset %s retains rectangular picture rendering', async (shapeType) => {
  const recorder = recordingContext();
  await drawPrimitive(recorder.ctx, { ...image(shapeType), border: { width: 2 } }, options);
  expect(recorder.images[0]).toEqual({ arguments: [bitmap, 10, 20, 100, 80], clips: [] });
  expect(recorder.calls.some((call) => call[0] === 'ellipse')).toBe(false);
  expect(recorder.calls).toContainEqual(['strokeRect', 10, 20, 100, 80]);
  expect(recorder.depth()).toBe(0);
});
