import { describe, expect, test } from 'bun:test';
import { placeFloating } from './floatingPlacement';

const viewport = { x: 0, y: 0, width: 800, height: 600 };
const size = { width: 200, height: 40 };

describe('placeFloating', () => {
  test('centres the box above the anchor with the gap', () => {
    expect(placeFloating({ anchor: { x: 300, y: 200, width: 100, height: 20 }, size, viewport })).toEqual({
      x: 250, y: 152, width: 200, height: 40, side: 'above', flipped: false, clamped: false, hidden: false,
    });
  });

  test('places below when asked, with a custom gap', () => {
    expect(placeFloating({ anchor: { x: 300, y: 200, width: 100, height: 20 }, size, viewport, side: 'below', gap: 4 }))
      .toMatchObject({ x: 250, y: 224, side: 'below', flipped: false, clamped: false, hidden: false });
  });

  test('flips below when there is no room above', () => {
    expect(placeFloating({ anchor: { x: 300, y: 30, width: 100, height: 20 }, size, viewport }))
      .toMatchObject({ y: 58, side: 'below', flipped: true, clamped: false, hidden: false });
  });

  test('flips above when there is no room below', () => {
    expect(placeFloating({ anchor: { x: 300, y: 560, width: 100, height: 20 }, size, viewport, side: 'below' }))
      .toMatchObject({ y: 512, side: 'above', flipped: true, clamped: false });
  });

  test('keeps the preferred side and clamps when neither side has room', () => {
    const short = { x: 0, y: 0, width: 800, height: 90 };
    expect(placeFloating({ anchor: { x: 300, y: 30, width: 100, height: 20 }, size, viewport: short }))
      .toMatchObject({ y: 8, side: 'above', flipped: false, clamped: true, hidden: false });
  });

  test('clamps to the left and right edges with the margin', () => {
    expect(placeFloating({ anchor: { x: 0, y: 200, width: 20, height: 20 }, size, viewport, margin: 12 }))
      .toMatchObject({ x: 12, y: 152, clamped: true, hidden: false });
    expect(placeFloating({ anchor: { x: 790, y: 200, width: 10, height: 20 }, size, viewport }))
      .toMatchObject({ x: 592, clamped: true, hidden: false });
  });

  test('a box wider than the viewport starts at the left margin', () => {
    expect(placeFloating({ anchor: { x: 300, y: 200, width: 10, height: 20 }, size: { width: 900, height: 40 }, viewport }))
      .toMatchObject({ x: 8, clamped: true });
  });

  test('an anchor partly off the page is placed from its part on the page', () => {
    const page = { x: 100, y: 100, width: 400, height: 500 };
    expect(placeFloating({ anchor: { x: 400, y: 200, width: 200, height: 20 }, size, viewport, page }))
      .toMatchObject({ x: 350, y: 152, clamped: true, hidden: false });
  });

  test('an anchor wholly off its page is hidden', () => {
    const page = { x: 100, y: 100, width: 400, height: 300 };
    expect(placeFloating({ anchor: { x: 300, y: 450, width: 50, height: 20 }, size, viewport, page }))
      .toMatchObject({ hidden: true });
  });

  test('an anchor partly above the viewport is placed from its visible part', () => {
    const scrolled = { x: 0, y: 1000, width: 800, height: 600 };
    expect(placeFloating({ anchor: { x: 300, y: 990, width: 100, height: 40 }, size, viewport: scrolled }))
      .toMatchObject({ y: 1038, side: 'below', flipped: true, clamped: true, hidden: false });
  });

  test('an anchor wholly outside the viewport is hidden and the box stays inside', () => {
    const below = placeFloating({ anchor: { x: 300, y: 900, width: 100, height: 20 }, size, viewport });
    expect(below).toMatchObject({ hidden: true, clamped: true });
    expect(below.y).toBeGreaterThanOrEqual(8);
    expect(below.y + below.height).toBeLessThanOrEqual(592);
    expect(placeFloating({ anchor: { x: -400, y: 200, width: 100, height: 20 }, size, viewport }))
      .toMatchObject({ x: 8, hidden: true, clamped: true });
  });

  test('scrolling moves the placement with the viewport', () => {
    const anchor = { x: 300, y: 1230, width: 100, height: 20 };
    expect(placeFloating({ anchor, size, viewport: { x: 0, y: 1200, width: 800, height: 600 } }))
      .toMatchObject({ y: 1258, side: 'below', flipped: true, hidden: false });
    expect(placeFloating({ anchor, size, viewport: { x: 0, y: 1100, width: 800, height: 600 } }))
      .toMatchObject({ y: 1182, side: 'above', flipped: false, hidden: false });
    expect(placeFloating({ anchor, size, viewport: { x: 0, y: 0, width: 800, height: 600 } }))
      .toMatchObject({ hidden: true });
  });

  test('zoom scales the anchor while the gap, margin and box stay in pixels', () => {
    const at = (zoom: number) => placeFloating({
      anchor: { x: 300 * zoom, y: 200 * zoom, width: 100 * zoom, height: 20 * zoom },
      size, viewport: { x: 0, y: 0, width: 800 * zoom, height: 600 * zoom },
    });
    expect(at(1)).toMatchObject({ x: 250, y: 152, width: 200, height: 40 });
    expect(at(2)).toMatchObject({ x: 600, y: 352, width: 200, height: 40, side: 'above', clamped: false });
    expect(at(0.5)).toMatchObject({ x: 75, y: 52, width: 200, height: 40, side: 'above', clamped: false });
  });

  test('a collapsed caret anchor is placed', () => {
    expect(placeFloating({ anchor: { x: 400, y: 300, width: 0, height: 18 }, size, viewport }))
      .toMatchObject({ x: 300, y: 252, hidden: false, clamped: false });
  });
});
