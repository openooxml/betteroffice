import { describe, expect, test } from 'bun:test';
import { FRAME_DELTA_VERSION, type RetainedFrame } from '../layout/render/frameDelta';
import type { YrsResidentCaretSnapshot } from './index';
import { residentCaretDeviceRect, residentCaretSnapshotForFrame } from './residentCaret';

const page = { pageIndex: 0, width: 100, height: 100, primitives: [] };
const frame: RetainedFrame = {
  protocolVersion: FRAME_DELTA_VERSION,
  docEpoch: 1,
  layoutEpoch: 1,
  frameEpoch: 9,
  pages: [
    {
      pageIndex: 0,
      pageId: 41n,
      fingerprint: 1n,
      primitiveIds: new BigUint64Array(),
      page,
    },
  ],
  damagedPageIds: new Set([41n]),
  removedPageIds: new Set(),
  displayList: { pages: [page] },
};

const current: YrsResidentCaretSnapshot = {
  frameEpoch: 9,
  caretRect: {
    pageIndex: 0,
    pageId: '41',
    x: 12,
    y: 34,
    height: 18,
  },
};

describe('resident caret epoch guard', () => {
  test('accepts geometry from the matching frame and page', () => {
    expect(residentCaretSnapshotForFrame(current, frame)).toBe(current);
  });

  test('drops stale epochs and stale page identities', () => {
    expect(
      residentCaretSnapshotForFrame({ ...current, frameEpoch: current.frameEpoch - 1 }, frame)
    ).toBeNull();
    expect(
      residentCaretSnapshotForFrame(
        { ...current, caretRect: { ...current.caretRect!, pageId: '42' } },
        frame
      )
    ).toBeNull();
  });

  test('keeps a same-epoch null caret for non-collapsed selections', () => {
    const range = { frameEpoch: 9, caretRect: null };
    expect(residentCaretSnapshotForFrame(range, frame)).toBe(range);
  });
});

describe('resident caret device rect', () => {
  test('scales geometry by dpr*zoom and stroke width by dpr only', () => {
    expect(
      residentCaretDeviceRect({ x: 10, y: 20, height: 15 }, { color: '#000', width: 2 }, 2, 1.5)
    ).toEqual({ x: 30, y: 60, width: 4, height: 45 });
  });

  test('keeps at least one device pixel of stroke', () => {
    expect(
      residentCaretDeviceRect({ x: 0, y: 0, height: 10 }, { color: '#000', width: 2 }, 0.25, 1)
    ).toEqual({ x: 0, y: 0, width: 1, height: 2.5 });
  });
});
