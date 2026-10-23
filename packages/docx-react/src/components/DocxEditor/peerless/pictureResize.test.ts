import { expect, test } from 'bun:test';
import { beginPictureResize, movePictureResize, pictureResizeIntent, type PictureResizeHandle } from './pictureResize';

const target = { handle: 'image', version: 'v1', pageIndex: 0, rect: { x: 100, y: 50, width: 40, height: 20 } };

for (const [handle, rect] of Object.entries({
  nw: { x: 110, y: 55, width: 30, height: 15 },
  n: { x: 100, y: 55, width: 40, height: 15 },
  ne: { x: 100, y: 55, width: 50, height: 15 },
  e: { x: 100, y: 50, width: 50, height: 20 },
  se: { x: 100, y: 50, width: 50, height: 25 },
  s: { x: 100, y: 50, width: 40, height: 25 },
  sw: { x: 110, y: 50, width: 30, height: 25 },
  w: { x: 110, y: 50, width: 30, height: 20 },
})) {
  test(`${handle} resizes in page pixels and anchors the opposite edge`, () => {
    expect(movePictureResize(beginPictureResize(target, handle as PictureResizeHandle, 20, 30), 40, 40, 2, false).rect).toEqual(rect);
  });
}

test('corners honor the picture lock or Shift; unlocked corners are free', () => {
  for (const handle of ['nw', 'ne', 'se', 'sw'] as const) {
    const x = handle.includes('w') ? -10 : 10;
    const y = handle.includes('n') ? -2 : 2;
    for (const [noChangeAspect, shift] of [[true, false], [false, true], [true, true]]) {
      const resized = movePictureResize(beginPictureResize({ ...target, noChangeAspect }, handle, 0, 0), x, y, 1, shift!).rect;
      expect(resized.width).toBe(50);
      expect(resized.height).toBe(25);
    }
    expect(movePictureResize(beginPictureResize(target, handle, 0, 0), x, y, 1, false).rect).toMatchObject({ width: 50, height: 22 });
  }
});

test('side handles change only one dimension even with locks and Shift', () => {
  for (const handle of ['n', 'e', 's', 'w'] as const) {
    const resized = movePictureResize(beginPictureResize({ ...target, noChangeAspect: true }, handle, 0, 0), 10, 5, 1, true).rect;
    expect(handle === 'n' || handle === 's' ? resized.width : resized.height).toBe(handle === 'n' || handle === 's' ? 40 : 20);
  }
});

test('minimum dimensions are one pixel and locked corners preserve the ratio at the minimum', () => {
  expect(movePictureResize(beginPictureResize(target, 'nw', 0, 0), 100, 100, 1, false).rect).toEqual({ x: 139, y: 69, width: 1, height: 1 });
  expect(movePictureResize(beginPictureResize(target, 'se', 0, 0), -100, -100, 1, true).rect).toEqual({ x: 100, y: 50, width: 2, height: 1 });
});

test('drop rounds integer EMUs and ignores unchanged sizes and invalid zoom', () => {
  const start = beginPictureResize(target, 'se', 0, 0);
  expect(pictureResizeIntent(start)).toBeNull();
  expect(movePictureResize(start, 5, 5, 0, false)).toBe(start);
  expect(movePictureResize(start, 5, 5, NaN, false)).toBe(start);
  const resized = movePictureResize(start, 1 / 3, 2 / 7, 2, false);
  expect(pictureResizeIntent(resized)).toEqual({ kind: 'resizeObject', payload: {
    handle: 'image', expectVersion: 'v1', widthEmu: Math.round((40 + 1 / 6) * 9525), heightEmu: Math.round((20 + 1 / 7) * 9525),
  } });
  expect(pictureResizeIntent(movePictureResize(resized, 0, 0, 1, false))).toBeNull();
  expect(pictureResizeIntent(movePictureResize(start, 0.00001, 0, 1, false))).toBeNull();
});

for (const handle of ['nw', 'ne', 'se', 'sw'] as const) {
  test(`locked ${handle} corners shrink along either axis`, () => {
    const start = beginPictureResize({ ...target, noChangeAspect: true }, handle, 0, 0);
    const dx = handle.includes('w') ? 10 : -10;
    const dy = handle.includes('n') ? 5 : -5;
    expect(movePictureResize(start, dx, 0, 1, false).rect).toMatchObject({ width: 30, height: 15 });
    expect(movePictureResize(start, 0, dy, 1, false).rect).toMatchObject({ width: 30, height: 15 });
  });
}

test('northwest floating resize commits the preview position with the dimensions', () => {
  const resize = movePictureResize(beginPictureResize({ ...target, floating: true }, 'nw', 0, 0), 10, 5, 1, false);
  expect(pictureResizeIntent(resize)).toEqual({ kind: 'resizeObject', payload: {
    handle: 'image', expectVersion: 'v1', widthEmu: 30 * 9525, heightEmu: 15 * 9525,
    hOffsetEmu: 110 * 9525, vOffsetEmu: 55 * 9525,
  } });
});
