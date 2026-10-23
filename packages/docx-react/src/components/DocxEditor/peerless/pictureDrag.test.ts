import { expect, test } from 'bun:test';
import { beginPictureDrag, floatingPictureAt, inlinePictureAt, pictureInSelection, movePictureDrag, pictureAt, pictureDropIntent, selectedPicture } from './pictureDrag';
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

test('a behind-text picture is draggable where its own anchor box is the only text rect over it', () => {
  const behind = geometry();
  const objects = behind.objects!.objects as { kind: string; behindText?: boolean; rects: { pageIndex: number; x: number; y: number; width: number; height: number }[] }[];
  objects[1]!.behindText = true;
  // As the worker reports it: line rects for the text, plus the anchored picture's full box.
  objects[0]!.rects = [{ ...rect(0, 40), width: 115, height: 16 }, rect(100, 50)];
  expect(floatingPictureAt(behind, 0, 130, 65)?.handle).toBe('floating');
  expect(floatingPictureAt(behind, 0, 110, 50)).toBeNull();
});

test('the selected picture follows the worker flag, even when an empty-line marker joins its selection rects', () => {
  const selected = geometry();
  const objects = selected.objects!.objects as { kind: string; selected?: boolean }[];
  (selected.objects as { selection: unknown[] }).selection = [rect(100, 50), { ...rect(96, 50), width: 4 }];
  objects[1]!.selected = true;
  expect(pictureInSelection(selected)?.handle).toBe('floating');
  objects[1]!.selected = undefined;
  expect(pictureInSelection(selected)).toBeNull();
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

test('an inline drag requires a drop point and preserves its current frame identity', () => {
  const target = inlinePictureAt(geometry(), 0, 20, 20)!;
  expect(target).toMatchObject({ handle: 'inline', version: 'v1', floating: false });
  const drop = {
    page: { pageId: 'next-page', pageIndex: 1 }, x: 120, y: 65, frameEpoch: 7, layoutRevision: 8,
    clickCount: 1, modifiers: { shift: false, alt: false, ctrl: false, meta: false },
  };
  const press = beginPictureDrag(target, 500, 300);
  expect(pictureDropIntent(press, drop)).toBeNull();
  expect(pictureDropIntent(movePictureDrag(press, 502, 301, 2), drop)).toBeNull();
  const drag = movePictureDrag(press, 520, 340, 2);
  expect(pictureDropIntent(drag)).toBeNull();
  expect(pictureDropIntent(drag, null)).toBeNull();
  expect(pictureDropIntent(drag, drop)).toEqual({ kind: 'moveEmbed', payload: {
    handle: 'inline', expectVersion: 'v1', drop,
  } });
});

test('inline targets exclude floating, nested, split, and stale picture geometry', () => {
  expect(inlinePictureAt(geometry(), 0, 120, 60)).toBeNull();
  expect(inlinePictureAt(geometry(), 1, 20, 20)).toBeNull();
  expect(inlinePictureAt(geometry({ version: 'v0' }), 0, 20, 20)).toBeNull();
  expect(inlinePictureAt({ ...geometry(), frameEpoch: 4 }, 0, 20, 20)).toBeNull();
  expect(inlinePictureAt({ ...geometry(), layoutRevision: 4 }, 0, 20, 20)).toBeNull();
  const current = geometry();
  const image = current.objects!.objects[2]!;
  image.nested = true;
  expect(inlinePictureAt(current, 0, 20, 20)).toBeNull();
  image.nested = false;
  image.rects.push(rect(60, 10));
  expect(inlinePictureAt(current, 0, 20, 20)).toBeNull();
});

test('picture selection hits inline and floating images with the same behind-text rule', () => {
  expect(pictureAt(geometry(), 0, 20, 20)?.handle).toBe('inline');
  expect(pictureAt(geometry(), 0, 120, 60)?.handle).toBe('floating');
  const behind = geometry();
  behind.objects!.objects[1]!.behindText = true;
  expect(pictureAt(behind, 0, 120, 60)).toBeNull();
  expect(pictureAt(geometry({ version: 'v0' }), 0, 20, 20)).toBeNull();
  expect(pictureAt({ ...geometry(), frameEpoch: 4 }, 0, 20, 20)).toBeNull();
  expect(pictureAt({ ...geometry(), layoutRevision: 4 }, 0, 20, 20)).toBeNull();
});

test('selection re-finds the sticky image handle in current geometry and clears when absent', () => {
  const current = geometry();
  const image = current.objects!.objects[1]!;
  image.noChangeAspect = true;
  image.rects[0]!.x = 200;
  expect(selectedPicture(current, 'floating')).toMatchObject({ rect: { x: 200 }, noChangeAspect: true });
  current.objects!.objects.splice(1, 1);
  expect(selectedPicture(current, 'floating')).toBeNull();
  expect(selectedPicture(geometry({ version: 'old' }), 'floating')).toBeNull();
  expect(selectedPicture(geometry(), 'reissued-handle')).toBeNull();
});

for (const [name, first, last] of [
  ['foreground over a later behind-text picture', { behindText: false, paintOrder: 2 }, { behindText: true, paintOrder: 0 }],
  ['higher relativeHeight over later document order', { behindText: true, paintOrder: 1 }, { behindText: true, paintOrder: 0 }],
] as const) {
  test(`picture drag picks the painted ${name}`, () => {
    const frame = geometry();
    frame.objects!.objects = [
      { handle: 'visible', kind: 'image', story: 'body', floating: true, rects: [rect(100, 50)], ...first },
      { handle: 'hidden', kind: 'image', story: 'body', floating: true, rects: [rect(100, 50)], ...last },
    ];
    expect(floatingPictureAt(frame, 0, 120, 60)?.handle).toBe('visible');
    expect(pictureAt(frame, 0, 120, 60)?.handle).toBe('visible');
  });
}

for (const story of ['body:t0:r0c0', 'body:tb0']) {
  test(`nested picture in ${story} cannot be selected for dragging or resizing`, () => {
    const frame = geometry();
    frame.objects!.objects = [{ handle: 'nested-picture', kind: 'image', story,
      floating: true, nested: true, selected: true, rects: [rect(100, 50)] }];
    expect(floatingPictureAt(frame, 0, 120, 60)).toBeNull();
    expect(pictureAt(frame, 0, 120, 60)).toBeNull();
    expect(selectedPicture(frame, 'nested-picture')).toBeNull();
    expect(pictureInSelection(frame)).toBeNull();
  });
}

test('an inline picture painted above a behind-text picture keeps the press', () => {
  const frame = geometry();
  frame.objects!.objects = [
    { handle: 'inline', kind: 'image', story: 'body', floating: false, paintOrder: 1, rects: [rect(100, 50)] },
    { handle: 'behind', kind: 'image', story: 'body', floating: true, behindText: true, paintOrder: 0, rects: [rect(100, 50)] },
  ];
  expect(floatingPictureAt(frame, 0, 120, 60)).toBeNull();
});
