/**
 * The hover cursor's imperative path: real pointer events on a real host,
 * through the hook's delegated listeners. The mapping and the write guard are
 * unit-tested in `hoverCursor.test.ts`; only a mounted hook shows the wiring.
 */

import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, describe, expect, test } from 'bun:test';
import { createRef } from 'react';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import type { YrsSession } from '@betteroffice/docx/yrs';
import { usePagesPointer, type UsePagesPointerOptions } from './usePagesPointer';
import type { YrsPositionProjection } from '../internals/yrsPositionProjection';

// Registered only if nobody else did, and handed back at the end: a leaked
// happy-dom replaces globalThis.fetch with one that rejects `file:`, which
// kills wasm init in every suite that runs after this one.
const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');

const PAGE = { width: 800, height: 1000 };
/** x below this is "margin" in the fake engine; at or past it is body text. */
const TEXT_FROM = 100;

let host: HTMLDivElement;

function fakeQueries(): DisplayListQueries {
  return {
    displayList: {
      pages: [{ pageIndex: 0, width: PAGE.width, height: PAGE.height, primitives: [] }],
    },
    pageCount: () => 1,
    pageSize: () => PAGE,
    hitTestRegions: (_pageIndex: number, x: number) => ({
      region: 'body' as const,
      pos: 1,
      target: x >= TEXT_FROM ? ('text' as const) : ('none' as const),
    }),
    imageAtPoint: () => null,
  } as unknown as DisplayListQueries;
}

/**
 * Every identity here is stable across renders, as the editor's own are: a ref
 * that changed identity per render would re-run effects keyed on it and could
 * repaint the cursor for reasons a mode change never causes.
 */
function stableOptions(): (overrides?: Partial<UsePagesPointerOptions>) => UsePagesPointerOptions {
  const canvasHostRef = createRef<HTMLDivElement>() as { current: HTMLDivElement | null };
  canvasHostRef.current = host;
  const projection = {
    size: 10,
    targetAt: () => null,
    tableAtPosition: () => null,
    cellPosition: () => null,
  } as unknown as YrsPositionProjection;
  const base: UsePagesPointerOptions = {
    pagesContainerRef: { current: null },
    yrsInputRef: { current: null },
    yrsSession: null as unknown as YrsSession,
    yrsRootStory: 'body',
    getYrsPositionProjection: () => projection,
    applyYrsCommand: () => false,
    syncYrsInputState: () => false,
    readOnly: false,
    displayListQueries: fakeQueries(),
    canvasHostRef,
    setSelectionRects: () => {},
    setCaretPosition: () => {},
    setIsFocused: () => {},
    scrollToPositionImpl: () => {},
  };
  return (overrides = {}) => ({ ...base, ...overrides });
}

function mouse(type: string, clientX: number, clientY: number, target: EventTarget): void {
  act(() => {
    target.dispatchEvent(
      new MouseEvent(type, { bubbles: true, cancelable: true, clientX, clientY, button: 0 })
    );
  });
}

beforeEach(() => {
  host = document.createElement('div');
  host.className = 'canvas-pages';
  const canvas = document.createElement('canvas');
  canvas.dataset.pageIndex = '0';
  canvas.getBoundingClientRect = () =>
    ({
      left: 0,
      top: 0,
      right: PAGE.width,
      bottom: PAGE.height,
      width: PAGE.width,
      height: PAGE.height,
    }) as DOMRect;
  host.append(canvas);
  document.body.append(host);
});

afterEach(() => {
  cleanup();
  host.remove();
});

afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const canvasOf = () => host.firstElementChild as HTMLCanvasElement;

describe('pages pointer hover cursor', () => {
  test('a move over text paints the caret cursor, and the margin drops it', () => {
    const options = stableOptions();
    renderHook(() => usePagesPointer(options()));
    const canvas = canvasOf();

    mouse('mousemove', 400, 500, canvas);
    expect(host.style.cursor).toBe('text');

    mouse('mousemove', 40, 500, canvas);
    expect(host.style.cursor).toBe('default');
  });

  test('a drag holds its cursor, and the release re-reads the point under it', () => {
    const options = stableOptions();
    renderHook(() => usePagesPointer(options()));
    const canvas = canvasOf();

    mouse('mousemove', 400, 500, canvas);
    expect(host.style.cursor).toBe('text');

    mouse('mousedown', 400, 500, canvas);
    // dragging into the margin keeps the cursor the gesture started with
    mouse('mousemove', 40, 500, canvas);
    expect(host.style.cursor).toBe('text');

    // released there without ever moving again
    mouse('mouseup', 40, 500, window);
    expect(host.style.cursor).toBe('default');
  });

  test('leaving the pages drops the cursor', () => {
    const options = stableOptions();
    renderHook(() => usePagesPointer(options()));

    mouse('mousemove', 400, 500, canvasOf());
    expect(host.style.cursor).toBe('text');

    mouse('mousemove', 400, 500, document.body);
    expect(host.style.cursor).toBe('default');
  });

  test('a drag that runs off the pages keeps its cursor', () => {
    const options = stableOptions();
    renderHook(() => usePagesPointer(options()));

    mouse('mousemove', 400, 500, canvasOf());
    mouse('mousedown', 400, 500, canvasOf());
    mouse('mousemove', 400, 2000, document.body);

    expect(host.style.cursor).toBe('text');
  });

  test('unmounting takes back the cursor it stamped on the host', () => {
    const options = stableOptions();
    const { unmount } = renderHook(() => usePagesPointer(options()));

    mouse('mousemove', 400, 500, canvasOf());
    expect(host.style.cursor).toBe('text');

    act(() => unmount());
    expect(host.style.cursor).toBe('default');
  });

  test('a mode flip mid-drag waits for the release', () => {
    const options = stableOptions();
    const { rerender } = renderHook(
      (hfEditMode: 'header' | 'footer' | null) => usePagesPointer(options({ hfEditMode })),
      { initialProps: null as 'header' | 'footer' | null }
    );

    mouse('mousemove', 400, 500, canvasOf());
    mouse('mousedown', 400, 500, canvasOf());

    // the body becomes inert behind an open band editor, but not mid-gesture
    act(() => rerender('header'));
    expect(host.style.cursor).toBe('text');

    mouse('mouseup', 400, 500, window);
    expect(host.style.cursor).toBe('default');
  });

  test('a re-render that changes no mode leaves the cursor alone', () => {
    // options rebuilt per render, as an editor's own props may be
    const { rerender } = renderHook(() => usePagesPointer(stableOptions()()));

    mouse('mousemove', 400, 500, canvasOf());
    expect(host.style.cursor).toBe('text');

    act(() => rerender());
    expect(host.style.cursor).toBe('text');
  });

  test('a read-only document never paints the caret cursor', () => {
    const options = stableOptions();
    renderHook(() => usePagesPointer(options({ readOnly: true })));

    mouse('mousemove', 400, 500, canvasOf());
    expect(host.style.cursor).toBe('default');
  });

  test('opening a header editor re-resolves a pointer that has not moved', () => {
    const options = stableOptions();
    const { rerender } = renderHook(
      (hfEditMode: 'header' | 'footer' | null) => usePagesPointer(options({ hfEditMode })),
      { initialProps: null as 'header' | 'footer' | null }
    );

    mouse('mousemove', 400, 500, canvasOf());
    expect(host.style.cursor).toBe('text');

    // the body is inert behind an open band editor, under the same pointer
    act(() => rerender('header'));
    expect(host.style.cursor).toBe('default');
  });
});
