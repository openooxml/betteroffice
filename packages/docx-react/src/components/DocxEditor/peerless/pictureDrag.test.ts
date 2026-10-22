import { expect, test } from 'bun:test';
import { beginPictureDrag, floatingPictureAt, movePictureDrag, pictureDropIntent } from './pictureDrag';
import type { PeerlessFrameGeometry } from './viewStore';

const rect = (x: number, y: number) => ({ pageId: '1', pageIndex: 0, x, y, width: 40, height: 20 });

function geometry(overrides: { version?: string; floating?: boolean[] } = {}): PeerlessFrameGeometry {
  const floating = overrides.floating ?? [true, false];
  return {
    documentVersion: 'v1', frameEpoch: 2, layoutRevision: 3,
    objects: {
      version: overrides.version ?? 'v1', frameEpoch: 2, layoutRevision: 3, selection: [], unbuiltPages: [],
      objects: [
        { handle: 'text', kind: 'text', story: 'body', rects: [{ ...rect(0, 0), width: 400, height: 400 }] },
        { handle: 'floating', kind: 'image', story: 'body', rects: [rect(100, 50)], floating: floating[0] },
        { handle: 'inline', kind: 'image', story: 'body', rects: [rect(10, 10)], floating: floating[1] },
      ],
    },
  } as unknown as PeerlessFrameGeometry;
}

test('a press hits only floating pictures from geometry that matches the frame', () => {
  expect(floatingPictureAt(geometry(), 0, 120, 60)).toEqual({ handle: 'floating', version: 'v1', pageIndex: 0, rect: { x: 100, y: 50, width: 40, height: 20 } });
  expect(floatingPictureAt(geometry(), 0, 20, 20)).toBeNull();
  expect(floatingPictureAt(geometry(), 1, 120, 60)).toBeNull();
  expect(floatingPictureAt(geometry(), 0, 99, 60)).toBeNull();
  expect(floatingPictureAt(geometry({ version: 'v0' }), 0, 120, 60)).toBeNull();
  expect(floatingPictureAt(geometry({ floating: [false, false] }), 0, 120, 60)).toBeNull();
});

test('text over a behind-text picture keeps the press; the picture outside the text is still draggable', () => {
  const behind = geometry();
  const objects = behind.objects!.objects as { kind: string; behindText?: boolean; rects: { x: number; y: number; width: number; height: number }[] }[];
  objects[1]!.behindText = true;
  objects[0]!.rects[0] = { ...objects[0]!.rects[0]!, width: 115, height: 400 };
  expect(floatingPictureAt(behind, 0, 110, 60)).toBeNull();
  expect(floatingPictureAt(behind, 0, 130, 60)?.handle).toBe('floating');
  objects[1]!.behindText = false;
  expect(floatingPictureAt(behind, 0, 110, 60)?.handle).toBe('floating');
});

test('a drag starts past the slop and drops the dialog move at the new page offset', () => {
  const target = floatingPictureAt(geometry(), 0, 120, 60)!;
  let drag = beginPictureDrag(target, 500, 300);
  drag = movePictureDrag(drag, 502, 301, 2);
  expect(drag.dragging).toBe(false);
  expect(pictureDropIntent(drag)).toBeNull();
  drag = movePictureDrag(drag, 520, 340, 2);
  expect(drag).toMatchObject({ dragging: true, dx: 10, dy: 20 });
  drag = movePictureDrag(drag, 501, 300, 2);
  expect(drag.dragging).toBe(true);
  drag = movePictureDrag(drag, 560, 280, 2);
  expect(pictureDropIntent(drag)).toEqual({ kind: 'moveObject', payload: {
    handle: 'floating', expectVersion: 'v1', hOffsetEmu: 130 * 9525, vOffsetEmu: 40 * 9525,
  } });
});

test('a drag that returns to its start drops nothing', () => {
  const drag = movePictureDrag(movePictureDrag(beginPictureDrag(floatingPictureAt(geometry(), 0, 120, 60)!, 0, 0), 30, 0, 1), 0, 0, 1);
  expect(drag.dragging).toBe(true);
  expect(pictureDropIntent(drag)).toBeNull();
});
