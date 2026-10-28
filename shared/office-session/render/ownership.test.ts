import { expect, test } from 'bun:test';
import { createBitmapOwnership } from './ownership';

test('last release closes once, unknown and repeated releases do nothing', () => {
  let closed = 0;
  const bitmap = { width: 1, height: 1, close() { closed += 1; } };
  const ownership = createBitmapOwnership<typeof bitmap>();
  const a = {}, b = {};
  ownership.release(bitmap, a);
  ownership.retain(bitmap, a);
  ownership.retain(bitmap, a);
  ownership.retain(bitmap, b);
  expect(ownership.openCount).toBe(1);
  ownership.release(bitmap, a);
  ownership.release(bitmap, a);
  expect(closed).toBe(0);
  ownership.release(bitmap, b);
  ownership.release(bitmap, b);
  ownership.retain(bitmap, a);
  ownership.release(bitmap, a);
  expect(closed).toBe(1);
  expect(ownership.openCount).toBe(0);
});

test('the injected closer receives the bitmap', () => {
  const bitmap = { width: 1, height: 1, close() { throw new Error('default closer'); } };
  const closed: typeof bitmap[] = [], holder = {};
  const ownership = createBitmapOwnership<typeof bitmap>((value) => { closed.push(value); });
  ownership.retain(bitmap, holder);
  ownership.release(bitmap, holder);
  expect(closed).toEqual([bitmap]);
  expect(ownership.openCount).toBe(0);
});
