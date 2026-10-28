import { expect, test } from 'bun:test';
import { applyPatch, type PatchContext } from './applyPatch';
import type { DeviceRect, Patch } from './types';

function bitmap(pixels = [0, 0]) { return { width: pixels.length, height: 1, pixels, close() {} }; }
type Bitmap = ReturnType<typeof bitmap>;
function context() {
  const ops: string[] = [], pixels = [1, 2, 3, 4, 5, 6];
  let composite: GlobalCompositeOperation = 'source-over';
  let previous: GlobalCompositeOperation = 'source-over';
  let clip: DeviceRect | undefined, previousClip: DeviceRect | undefined;
  const ctx: PatchContext<Bitmap> = {
    save() { ops.push('save'); previous = composite; previousClip = clip; },
    beginPath() { ops.push('beginPath'); },
    rect(x, y, width, height) { ops.push(`rect:${x},${y},${width},${height}`); clip = { x, y, width, height }; },
    clip() { ops.push('clip'); },
    get globalCompositeOperation() { return composite; },
    set globalCompositeOperation(value) { ops.push(`composite:${value}`); composite = value; },
    drawImage(image, x, y) {
      ops.push(`draw:${x},${y}`);
      for (let index = 0; index < pixels.length; index += 1) {
        if (clip && (index < clip.x || index >= clip.x + clip.width)) continue;
        if (composite === 'copy') pixels[index] = image.pixels[index - x] ?? 0;
      }
    },
    restore() { ops.push('restore'); composite = previous; clip = previousClip; },
  };
  return { ctx, ops, pixels };
}
function patch(full: boolean): Patch<Bitmap> {
  return { surface: { kind: 'page', id: '1' }, rect: { x: 2, y: 0, width: 2, height: 1 }, bitmap: bitmap(), full };
}

test('full copies start at zero and restore the composite property', () => {
  const { ctx, ops } = context();
  applyPatch(ctx, patch(true));
  expect(ops).toEqual(['save', 'composite:copy', 'draw:0,0', 'restore']);
  expect(ctx.globalCompositeOperation).toBe('source-over');
});

test('partial copies clip before drawing and preserve outside pixels across replacements', () => {
  const { ctx, ops, pixels } = context();
  applyPatch(ctx, { ...patch(false), bitmap: bitmap([9, 8]) });
  expect(pixels).toEqual([1, 2, 9, 8, 5, 6]);
  expect(ops).toEqual(['save', 'beginPath', 'rect:2,0,2,1', 'clip', 'composite:copy', 'draw:2,0', 'restore']);
  applyPatch(ctx, patch(false));
  expect(pixels).toEqual([1, 2, 0, 0, 5, 6]);
  expect(ctx.globalCompositeOperation).toBe('source-over');
});

test('draw failure still restores the context', () => {
  const { ctx, ops } = context();
  ctx.drawImage = () => { throw new Error('draw failed'); };
  expect(() => applyPatch(ctx, patch(false))).toThrow('draw failed');
  expect(ops[ops.length - 1]).toBe('restore');
  expect(ctx.globalCompositeOperation).toBe('source-over');
});
