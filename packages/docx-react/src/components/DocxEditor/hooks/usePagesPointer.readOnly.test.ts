import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, expect, test } from 'bun:test';
import type { DisplayListQueries, DisplayListRegionHit } from '@betteroffice/docx/layout/render';
import type { YrsSession } from '@betteroffice/docx/yrs';
import type { YrsInputRef } from '../YrsInput';
import type { YrsPositionProjection } from '../internals/yrsPositionProjection';
import { usePagesPointer, type UsePagesPointerOptions } from './usePagesPointer';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');

const PAGE = { width: 800, height: 1000 };
const HEADER_BOTTOM = 80;
const NOTE_TOP = 900;

let host: HTMLDivElement;

/** Body text at y 80..900, one position per 10 px of x; a header above and a footnote below. */
function fakeQueries(): DisplayListQueries {
  return {
    displayList: {
      pages: [{ pageIndex: 0, width: PAGE.width, height: PAGE.height, primitives: [] }],
    },
    pageCount: () => 1,
    pageSize: () => PAGE,
    hitTestRegions: (_pageIndex: number, x: number, y: number): DisplayListRegionHit =>
      y < HEADER_BOTTOM
        ? { region: 'header', rId: 'rId7', pos: 1, target: 'text' }
        : y >= NOTE_TOP
          ? { region: 'footnote', noteId: 2, pos: 1, target: 'text' }
          : { region: 'body', pos: Math.floor(x / 10), target: 'text' },
    imageAtPoint: () => ({ pos: 5 }),
  } as unknown as DisplayListQueries;
}

function options(overrides: Partial<UsePagesPointerOptions> = {}) {
  const selections: Array<[number, number]> = [];
  let focused = 0;
  const input = {
    focus: () => {
      focused += 1;
    },
    setSelectionFromDisplay: (anchor: number, head: number) => selections.push([anchor, head]),
    displaySelection: () => null,
  } as unknown as YrsInputRef;
  const projection = {
    size: 100,
    targetAt: (position: number) => ({ story: 'body', displayPosition: position }),
    tableAtPosition: () => null,
    cellPosition: () => null,
    nodeAt: () => null,
  } as unknown as YrsPositionProjection;
  const noteClicks: unknown[] = [];
  const opts: UsePagesPointerOptions = {
    pagesContainerRef: { current: null },
    yrsInputRef: { current: input },
    yrsSession: { cellSelection: () => null } as unknown as YrsSession,
    yrsRootStory: 'body',
    getYrsPositionProjection: () => projection,
    applyYrsCommand: () => false,
    syncYrsInputState: () => false,
    readOnly: true,
    displayListQueries: fakeQueries(),
    canvasHostRef: { current: host },
    onNoteClick: (note) => noteClicks.push(note),
    setSelectionRects: () => {},
    setCaretPosition: () => {},
    setIsFocused: () => {},
    scrollToPositionImpl: () => {},
    ...overrides,
  };
  return { opts, selections, noteClicks, focused: () => focused };
}

function mouse(type: string, clientX: number, clientY: number, target: EventTarget): void {
  act(() => {
    target.dispatchEvent(
      new MouseEvent(type, { bubbles: true, cancelable: true, clientX, clientY, button: 0 })
    );
  });
}

const nextFrame = () =>
  act(() => new Promise<void>((resolve) => requestAnimationFrame(() => resolve())));

beforeEach(() => {
  host = document.createElement('div');
  host.className = 'canvas-pages';
  const canvas = document.createElement('canvas');
  canvas.dataset.pageIndex = '0';
  canvas.getBoundingClientRect = () =>
    ({ left: 0, top: 0, right: PAGE.width, bottom: PAGE.height, ...PAGE }) as DOMRect;
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

test('a read-only press places the caret and a drag extends the selection', async () => {
  const { opts, selections, focused } = options();
  renderHook(() => usePagesPointer(opts));

  mouse('mousedown', 200, 400, canvasOf());
  expect(selections).toEqual([[20, 20]]);
  expect(focused()).toBe(1);

  mouse('mousemove', 450, 600, window);
  await nextFrame();
  expect(selections.at(-1)).toEqual([20, 45]);
  mouse('mouseup', 450, 600, window);

  mouse('mousemove', 700, 600, window);
  await nextFrame();
  expect(selections.at(-1)).toEqual([20, 45]);
});

test('a read-only press selects text only: no image, note or header', () => {
  const { opts, selections, noteClicks } = options();
  renderHook(() => usePagesPointer(opts));

  mouse('mousedown', 50, 40, canvasOf());
  mouse('mousedown', 50, 950, canvasOf());
  expect(selections).toEqual([]);
  expect(noteClicks).toEqual([]);

  mouse('mousedown', 50, 400, canvasOf());
  expect(selections).toEqual([[5, 5]]);
});

test('an editable press still selects the image under it', () => {
  const { opts, selections } = options({ readOnly: false });
  renderHook(() => usePagesPointer(opts));

  mouse('mousedown', 50, 400, canvasOf());
  expect(selections).toEqual([[5, 6]]);
});
