import { describe, expect, test } from 'bun:test';
import { createDocxSelectionPlacement } from './selectionPlacement';
import { pluginSelectionKeys } from './selectionIdentity';
import type { DocxAnchorGeometryResult, DocxAnchorRect, DocxPluginGeometry, DocxPluginRect, DocxPluginSnapshot } from './types';

const viewport = { x: 0, y: 0, width: 800, height: 600 };
const page = { x: 0, y: 0, width: 800, height: 1000 };
const pending: DocxAnchorGeometryResult = { ok: false, failure: { code: 'pending', message: 'waiting' } };

function placed(rects: DocxAnchorRect[], pageRect: DocxPluginRect = page): DocxAnchorGeometryResult {
  const last = rects.at(-1)!;
  return { ok: true, version: '1', previewVersion: 0, layoutId: 'layout', rects,
    anchor: { ...last, x: last.x + last.width, width: 0 }, pageRect };
}

function frame(key: string, result: DocxAnchorGeometryResult | null, view: DocxPluginRect = viewport) {
  const snapshot = { selection: { formatting: null, displayRange: result?.ok ? { story: 'body', from: 1, to: 3, layoutId: 'layout' } : null } } as unknown as DocxPluginSnapshot;
  pluginSelectionKeys.set(snapshot, key);
  const geometry = result ? { viewport: view, getAnchorGeometry: () => result } as unknown as DocxPluginGeometry : null;
  return { snapshot, geometry };
}

const size = { width: 100, height: 30 };

describe('createDocxSelectionPlacement', () => {
  test('places above the selection box on the anchor page', () => {
    const place = createDocxSelectionPlacement();
    expect(place(frame('a', placed([
      { x: 100, y: 200, width: 300, height: 20, pageIndex: 0 },
      { x: 50, y: 220, width: 100, height: 20, pageIndex: 0 },
    ])), size)).toEqual({
      x: 175, y: 162, width: 100, height: 30, side: 'above', flipped: false, clamped: false, hidden: false,
      anchor: { x: 50, y: 200, width: 350, height: 40 }, kept: false,
    });
  });

  test('a selection across pages is placed from its part on the anchor page', () => {
    const place = createDocxSelectionPlacement();
    const second = { x: 0, y: 1040, width: 800, height: 1000 };
    expect(place(frame('a', placed([
      { x: 100, y: 900, width: 300, height: 20, pageIndex: 0 },
      { x: 100, y: 1100, width: 200, height: 20, pageIndex: 1 },
    ], second), { x: 0, y: 800, width: 800, height: 600 }), size)).toMatchObject({
      anchor: { x: 100, y: 1100, width: 200, height: 20 }, y: 1062, hidden: false, kept: false,
    });
  });

  test('keeps the last box while the range is null and the selection key is unchanged', () => {
    const place = createDocxSelectionPlacement();
    const first = place(frame('a', placed([{ x: 100, y: 200, width: 300, height: 20, pageIndex: 0 }])), size)!;
    expect(place(frame('a', pending), size)).toEqual({ ...first, kept: true });
    expect(place(frame('a', null), size)).toEqual({ ...first, kept: true });
    expect(place(frame('a', pending, { x: 0, y: 190, width: 800, height: 600 }), size))
      .toMatchObject({ anchor: first.anchor, side: 'below', flipped: true, kept: true });
  });

  test('drops the last box once the selection key changes', () => {
    const place = createDocxSelectionPlacement();
    place(frame('a', placed([{ x: 100, y: 200, width: 300, height: 20, pageIndex: 0 }])), size);
    expect(place(frame('b', pending), size)).toBeNull();
    expect(place(frame('a', pending), size)).toBeNull();
  });

  test('a new box replaces the kept one', () => {
    const place = createDocxSelectionPlacement();
    place(frame('a', placed([{ x: 100, y: 200, width: 300, height: 20, pageIndex: 0 }])), size);
    place(frame('a', pending), size);
    expect(place(frame('a', placed([{ x: 200, y: 300, width: 100, height: 20, pageIndex: 0 }])), size))
      .toMatchObject({ anchor: { x: 200, y: 300, width: 100, height: 20 }, kept: false });
    expect(place(frame('a', pending), size)).toMatchObject({ anchor: { x: 200, y: 300, width: 100, height: 20 }, kept: true });
  });

  test('no selection box and nothing kept answers null', () => {
    const place = createDocxSelectionPlacement();
    expect(place(frame('a', pending), size)).toBeNull();
    expect(place(frame('a', null), size)).toBeNull();
  });

  test('an explicit geometry wins over the context geometry', () => {
    const place = createDocxSelectionPlacement();
    const { snapshot, geometry } = frame('a', pending);
    const shown = frame('a', placed([{ x: 100, y: 200, width: 300, height: 20, pageIndex: 0 }])).geometry;
    expect(place({ snapshot, geometry }, { ...size, geometry: shown })).toMatchObject({ kept: false, anchor: { x: 100, y: 200 } });
  });

  test('placers keep separate boxes', () => {
    const one = createDocxSelectionPlacement();
    const two = createDocxSelectionPlacement();
    one(frame('a', placed([{ x: 100, y: 200, width: 300, height: 20, pageIndex: 0 }])), size);
    expect(two(frame('a', pending), size)).toBeNull();
    expect(one(frame('a', pending), size)).toMatchObject({ kept: true });
  });

  test.each(['stale-version', 'layout-unavailable'] as const)('keeps the last box while the geometry is %s', (code) => {
    const place = createDocxSelectionPlacement();
    const first = place(frame('a', placed([{ x: 100, y: 200, width: 300, height: 20, pageIndex: 0 }])), size)!;
    expect(place(frame('a', { ok: false, failure: { code, message: 'not shown' } }), size)).toEqual({ ...first, kept: true });
  });

  test.each(['missing-target', 'unsupported', 'read-failed'] as const)('a %s answer drops the last box', (code) => {
    const place = createDocxSelectionPlacement();
    place(frame('a', placed([{ x: 100, y: 200, width: 300, height: 20, pageIndex: 0 }])), size);
    expect(place(frame('a', { ok: false, failure: { code, message: 'failed' } }), size)).toBeNull();
    expect(place(frame('a', pending), size)).toBeNull();
  });
});
