import { expect, test } from 'bun:test';
import { rasterizeDisplayPageToBackBuffer } from '../layout/render/canvasBackend';

test('echo background allocation is bounded by the damage band', async () => {
  const translations: number[][] = [];
  const ctx = { resetTransform() {}, beginPath() {}, rect() {}, clip() {}, scale() {}, clearRect() {},
    translate(...values: number[]) { translations.push(values); } };
  const canvas = { width: 1, height: 1, getContext: () => ctx };
  const clip = { x: 50, y: 500, width: 125, height: 27 };
  await rasterizeDisplayPageToBackBuffer(canvas as unknown as OffscreenCanvas,
    { pageIndex: 0, width: 816, height: 1056, primitives: [] }, {}, 1.25, 1, clip);
  expect(canvas.width * canvas.height).toBe(clip.width * clip.height);
  expect(translations).toEqual([[-50, -500]]);
});
