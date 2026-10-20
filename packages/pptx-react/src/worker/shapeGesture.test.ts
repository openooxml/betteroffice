import { expect, test } from 'bun:test';
import { peerlessGestureRect } from './usePeerlessShapeGestures';

const object = {
  kind: 'object' as const, slideId: 'slide', shapeId: 'picture', storyId: '', token: 'picture',
  anchor: 0, focus: 0, affinity: 'downstream' as const,
  bounds: { x: 100, y: 200, width: 240, height: 120 },
  rect: { x: 1_000_000, y: 2_000_000, width: 2_400_000, height: 1_200_000 },
  canMove: true, canResize: true, units: { x: 10_000, y: 10_000 }, aspectRatioLocked: true,
};

test('worker picture corner drags preserve the locked aspect ratio', () => {
  expect(peerlessGestureRect(object, { x: 80, y: 10 }, 'se')).toEqual({
    ...object.rect, width: 3_080_000, height: 1_540_000,
  });
  expect(peerlessGestureRect(object, { x: 10, y: 80 }, 'nw')).toEqual({
    x: 1_400_000, y: 2_200_000, width: 2_000_000, height: 1_000_000,
  });
});

test('locked picture corner drags keep both dimensions above the minimum', () => {
  const rect = peerlessGestureRect(object, { x: -1000, y: -1000 }, 'se')!;
  expect(rect.height).toBe(76_200);
  expect(rect.width).toBe(152_400);
});

test('locked picture edge drags and unlocked corner drags retain their behavior', () => {
  expect(peerlessGestureRect(object, { x: 80, y: 10 }, 'e')).toEqual({
    ...object.rect, width: 3_200_000,
  });
  expect(peerlessGestureRect({ ...object, aspectRatioLocked: false }, { x: 80, y: 10 }, 'se')).toEqual({
    ...object.rect, width: 3_200_000, height: 1_300_000,
  });
});
