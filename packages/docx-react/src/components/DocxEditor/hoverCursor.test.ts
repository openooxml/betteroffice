import { describe, expect, test } from 'bun:test';
import type { DisplayListRegionHit } from '@betteroffice/docx/layout/render';
import { canvasHoverCursor, createCursorPainter } from './hoverCursor';

const hit = (over: Partial<DisplayListRegionHit>): DisplayListRegionHit => ({
  region: 'body',
  pos: 12,
  target: 'none',
  ...over,
});

const editing = { readOnly: false };

describe('canvas hover cursor', () => {
  test('follows the engine target inside the body', () => {
    expect(canvasHoverCursor(editing, hit({ target: 'text' }))).toBe('text');
    expect(canvasHoverCursor(editing, hit({ target: 'none' }))).toBe('default');
    expect(canvasHoverCursor(editing, hit({ target: 'image' }))).toBe('default');
  });

  test('a resolved position off the typeable area is still not text', () => {
    expect(canvasHoverCursor(editing, hit({ pos: 40, target: 'none' }))).toBe('default');
  });

  test('an idle header or footer band reads as text over its own runs', () => {
    expect(canvasHoverCursor(editing, hit({ region: 'header', target: 'text' }))).toBe('text');
    expect(canvasHoverCursor(editing, hit({ region: 'footer', target: 'none' }))).toBe('default');
  });

  test('editing a band makes only that band typeable', () => {
    const mode = { readOnly: false, hfEditMode: 'header' as const };
    expect(canvasHoverCursor(mode, hit({ region: 'header', target: 'none' }))).toBe('text');
    expect(canvasHoverCursor(mode, hit({ region: 'header', target: 'text' }))).toBe('text');
    expect(canvasHoverCursor(mode, hit({ region: 'header', target: 'image' }))).toBe('default');
    expect(canvasHoverCursor(mode, hit({ region: 'body', target: 'text' }))).toBe('default');
    expect(canvasHoverCursor(mode, hit({ region: 'footer', target: 'text' }))).toBe('default');
  });

  test('read-only and off-page points never type', () => {
    expect(canvasHoverCursor({ readOnly: true }, hit({ target: 'text' }))).toBe('default');
    expect(
      canvasHoverCursor({ readOnly: true, hfEditMode: 'footer' }, hit({ region: 'footer' }))
    ).toBe('default');
    expect(canvasHoverCursor(editing, null)).toBe('default');
  });

  test('a wasm build without the target degrades to the pre-hover behaviour', () => {
    const older: DisplayListRegionHit = { region: 'body', pos: 12 };
    expect(canvasHoverCursor(editing, older)).toBe('default');
    // an open band types regardless: the whole band is its typeable area
    expect(
      canvasHoverCursor({ readOnly: false, hfEditMode: 'header' }, { ...older, region: 'header' })
    ).toBe('text');
  });
});

describe('cursor painter', () => {
  const fakeHost = () => ({ style: { cursor: '' } }) as unknown as HTMLElement;

  test('writes once per change and skips repeats', () => {
    const paint = createCursorPainter();
    const host = fakeHost();

    expect(paint(host, 'text')).toBe(true);
    expect(host.style.cursor).toBe('text');
    expect(paint(host, 'text')).toBe(false);
    expect(paint(host, 'default')).toBe(true);
    expect(host.style.cursor).toBe('default');
  });

  test('re-stamps a replaced host holding the same cursor', () => {
    const paint = createCursorPainter();
    const first = fakeHost();
    const second = fakeHost();

    paint(first, 'text');
    expect(paint(second, 'text')).toBe(true);
    expect(second.style.cursor).toBe('text');
  });

  test('a missing host is skipped without poisoning the guard', () => {
    const paint = createCursorPainter();
    const host = fakeHost();

    expect(paint(null, 'text')).toBe(false);
    expect(paint(host, 'text')).toBe(true);
  });
});
