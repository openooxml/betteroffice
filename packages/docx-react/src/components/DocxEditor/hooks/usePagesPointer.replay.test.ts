import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, expect, mock, spyOn, test } from 'bun:test';
import { createRef } from 'react';
import type { DisplayListQueries, DisplayListRegionHit } from '@betteroffice/docx/layout/render';
import type { YrsCellLoc, YrsSession } from '@betteroffice/docx/yrs';
import type { YrsInputRef } from '../YrsInput';
import type { PagedEditorRef } from '../PagedEditor';
import type { YrsPositionProjection } from '../internals/yrsPositionProjection';
import { usePagesPointer, type UsePagesPointerOptions } from './usePagesPointer';
import { usePagedEditorRefApi } from './usePagedEditorRefApi';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');

const PAGE = { width: 800, height: 1000 };
let host: HTMLDivElement;

function fakeQueries(): DisplayListQueries {
  return {
    displayList: {
      pages: [{ pageIndex: 0, width: PAGE.width, height: PAGE.height, primitives: [] }],
    },
    pageCount: () => 1,
    pageSize: () => PAGE,
    hitTestRegions: (_pageIndex: number, x: number, y: number): DisplayListRegionHit =>
      y < 80
        ? { region: 'header', rId: 'rId7', pos: 1, target: 'text' }
        : { region: 'body', pos: Math.floor(x / 10), target: 'text' },
    imageAtPoint: () => null,
  } as unknown as DisplayListQueries;
}

function options(overrides: Partial<UsePagesPointerOptions> = {}) {
  const selections: Array<[number, number, string]> = [];
  const words: Array<[number, string]> = [];
  const paragraphs: Array<[number, string]> = [];
  let focused = 0;
  const input = {
    focus: () => {
      focused += 1;
    },
    isFocused: () => false,
    setSelectionFromDisplay: (anchor: number, head: number, story: string) =>
      selections.push([anchor, head, story]),
    displaySelection: () => null,
    selectAll: () => {},
    selectWordAtDisplay: (position: number, story: string) => words.push([position, story]),
    selectParagraphAtDisplay: (position: number, story: string) =>
      paragraphs.push([position, story]),
  } as unknown as YrsInputRef;
  const projection = {
    size: 100,
    targetAt: (position: number) => ({ story: 'body', displayPosition: position }),
    tableAtPosition: () => null,
    cellPosition: () => null,
    nodeAt: () => null,
  } as unknown as YrsPositionProjection;
  const opts: UsePagesPointerOptions = {
    pagesContainerRef: { current: null },
    yrsInputRef: { current: input },
    yrsSession: { cellSelection: () => null } as unknown as YrsSession,
    yrsRootStory: 'body',
    getYrsPositionProjection: () => (opts.replicaReady ? projection : null),
    applyYrsCommand: () => false,
    syncYrsInputState: () => false,
    readOnly: true,
    replicaReady: false,
    replicaPending: () => !opts.replicaReady,
    displayListQueries: fakeQueries(),
    canvasHostRef: { current: host },
    setSelectionRects: () => {},
    setCaretPosition: () => {},
    setIsFocused: () => {},
    scrollToPositionImpl: () => {},
    ...overrides,
  };
  return { opts, selections, words, paragraphs, projection, focused: () => focused };
}

function mouse(type: string, clientX: number, clientY = 400, detail = 1): void {
  act(() => {
    const target = type === 'mousedown' || type === 'click' ? host.firstElementChild! : window;
    target.dispatchEvent(
      new MouseEvent(type, { bubbles: true, cancelable: true, clientX, clientY, button: 0, detail })
    );
  });
}

function click(detail: number, clientX = 200, clientY = 400): void {
  mouse('mousedown', clientX, clientY, detail);
  mouse('mouseup', clientX, clientY, detail);
  mouse('click', clientX, clientY, detail);
}

beforeEach(() => {
  host = document.createElement('div');
  host.className = 'canvas-pages';
  const canvas = document.createElement('canvas');
  canvas.className = 'canvas-page';
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

test('a double-click during replica loading selects the word when ready', () => {
  const { opts, selections, words, focused } = options();
  const view = renderHook(() => usePagesPointer(opts));

  click(1);
  click(2);
  expect(selections).toEqual([]);
  expect(words).toEqual([]);
  const beforeReplay = focused();

  opts.replicaReady = true;
  view.rerender();
  expect(words).toEqual([[20, 'body']]);
  expect(selections).toEqual([]);
  expect(focused()).toBe(beforeReplay + 1);
});

test('a triple-click during replica loading selects the paragraph when ready', () => {
  const { opts, selections, words, paragraphs } = options();
  const view = renderHook(() => usePagesPointer(opts));

  click(1);
  click(2);
  click(3);
  expect(paragraphs).toEqual([]);

  opts.replicaReady = true;
  view.rerender();
  expect(paragraphs).toEqual([[20, 'body']]);
  expect(words).toEqual([]);
  expect(selections).toEqual([]);
});

test('a single click during replica loading replays the caret', () => {
  const { opts, selections } = options();
  const view = renderHook(() => usePagesPointer(opts));

  click(1);
  expect(selections).toEqual([]);

  opts.replicaReady = true;
  view.rerender();
  expect(selections).toEqual([[20, 20, 'body']]);
});

test('a pending drag replays the latest mousemove before an animation frame', () => {
  const { opts, selections } = options();
  const view = renderHook(() => usePagesPointer(opts));

  mouse('mousedown', 200);
  mouse('mousemove', 350);
  mouse('mousemove', 450);
  expect(selections).toEqual([]);

  opts.replicaReady = true;
  view.rerender();
  expect(selections).toEqual([[20, 45, 'body']]);
  mouse('mouseup', 450);
});

test('a pending drag captures its final mouseup position', () => {
  const { opts, selections } = options();
  const view = renderHook(() => usePagesPointer(opts));

  mouse('mousedown', 450);
  mouse('mousemove', 350);
  mouse('mouseup', 200);
  expect(selections).toEqual([]);

  opts.replicaReady = true;
  view.rerender();
  expect(selections).toEqual([[45, 20, 'body']]);
});

test('a pending drag across cells in one table replays a cell selection', () => {
  const anchor: YrsCellLoc = { story: 'body', tableIndex: 0, row: 0, column: 0 };
  const head: YrsCellLoc = { ...anchor, column: 1 };
  const setCellSelection = mock(() => {});
  const syncYrsInputState = mock(() => false);
  const setSelectionRects = mock(() => {});
  const setCaretPosition = mock(() => {});
  const { opts, projection, selections } = options({
    yrsSession: { cellSelection: () => null, setCellSelection } as unknown as YrsSession,
    syncYrsInputState,
    setSelectionRects,
    setCaretPosition,
  });
  projection.targetAt = (position) => ({
    story: position === 20 ? 'body:t0:r0c0' : 'body:t0:r0c1',
    displayPosition: 1,
    cell: position === 20 ? anchor : head,
  });
  const view = renderHook(() => usePagesPointer(opts));

  mouse('mousedown', 200);
  mouse('mousemove', 450);
  mouse('mouseup', 450);
  expect(setCellSelection).not.toHaveBeenCalled();

  opts.replicaReady = true;
  view.rerender();
  expect(setCellSelection).toHaveBeenCalledTimes(1);
  expect(setCellSelection).toHaveBeenCalledWith({ anchor, head });
  expect(syncYrsInputState).toHaveBeenCalledTimes(1);
  expect(syncYrsInputState).toHaveBeenCalledWith(false);
  expect(setSelectionRects).toHaveBeenCalledTimes(1);
  expect(setSelectionRects).toHaveBeenCalledWith([]);
  expect(setCaretPosition).toHaveBeenCalledTimes(1);
  expect(setCaretPosition).toHaveBeenCalledWith(null);
  expect(selections).toEqual([]);
});

test.each([70, null])('a pending bookmark link replays with bookmark position %s', (bookmark) => {
  const queries = fakeQueries();
  queries.displayList.pages[0]!.primitives = [{
    kind: 'text', text: 'link', x: 0, baselineY: 410, width: 800,
    font: '400 16px Calibri', color: '#000000', docStart: 1, docEnd: 5,
    href: '#bookmark',
  }];
  const scrollToPositionImpl = mock(() => {});
  const { opts, projection, selections } = options({
    displayListQueries: queries,
    scrollToPositionImpl,
  });
  projection.bookmarkPosition = mock((name: string) => name === 'bookmark' ? bookmark : null);
  const view = renderHook(() => usePagesPointer(opts));

  click(1, 200, 405);
  expect(selections).toEqual([]);
  expect(scrollToPositionImpl).not.toHaveBeenCalled();
  expect(projection.bookmarkPosition).not.toHaveBeenCalled();

  opts.replicaReady = true;
  view.rerender();
  expect(projection.bookmarkPosition).toHaveBeenCalledTimes(1);
  expect(projection.bookmarkPosition).toHaveBeenCalledWith('bookmark');
  if (bookmark === null) {
    expect(scrollToPositionImpl).not.toHaveBeenCalled();
    expect(selections).toEqual([[20, 20, 'body']]);
  } else {
    expect(scrollToPositionImpl).toHaveBeenCalledTimes(1);
    expect(scrollToPositionImpl).toHaveBeenCalledWith(bookmark);
    expect(selections).toEqual([[bookmark + 1, bookmark + 1, 'body']]);
  }
});

test.each(['pointerdown', 'keydown'])('outside %s drops a pending gesture', (type) => {
  const { opts, selections, words, focused } = options();
  const view = renderHook(() => usePagesPointer(opts));
  const outside = document.createElement('button');
  document.body.append(outside);
  try {
    click(2);
    act(() => outside.dispatchEvent(new Event(type, { bubbles: true })));
    const beforeReplay = focused();

    opts.replicaReady = true;
    view.rerender();
    expect(selections).toEqual([]);
    expect(words).toEqual([]);
    expect(focused()).toBe(beforeReplay);
  } finally {
    outside.remove();
  }
});

test('input on the pages or the focused hidden input keeps a pending gesture', () => {
  const { opts, selections, focused } = options();
  const hiddenInput = document.createElement('textarea');
  hiddenInput.className = 'paged-editor__yrs-input';
  document.body.append(hiddenInput);
  opts.yrsInputRef.current!.isFocused = () => document.activeElement === hiddenInput;
  const view = renderHook(() => usePagesPointer(opts));
  try {
    click(1);
    act(() => {
      host.firstElementChild!.dispatchEvent(new Event('pointerdown', { bubbles: true }));
      hiddenInput.focus();
      hiddenInput.dispatchEvent(new Event('pointerdown', { bubbles: true }));
      hiddenInput.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Shift' }));
    });
    const beforeReplay = focused();

    opts.replicaReady = true;
    view.rerender();
    expect(selections).toEqual([[20, 20, 'body']]);
    expect(focused()).toBe(beforeReplay + 1);
  } finally {
    hiddenInput.remove();
  }
});

test('replaying a gesture preserves focus in another input', () => {
  const { opts, selections, focused } = options();
  const view = renderHook(() => usePagesPointer(opts));
  const outside = document.createElement('input');
  document.body.append(outside);
  try {
    click(1);
    act(() => outside.focus());
    const beforeReplay = focused();

    opts.replicaReady = true;
    view.rerender();
    expect(selections).toEqual([[20, 20, 'body']]);
    expect(document.activeElement).toBe(outside);
    expect(focused()).toBe(beforeReplay);
  } finally {
    outside.remove();
  }
});

test.each(['replay', 'outside input', 'unmount'])('pending input listeners are removed on %s', (end) => {
  const { opts } = options();
  const view = renderHook(() => usePagesPointer(opts));
  const add = spyOn(document, 'addEventListener');
  const remove = spyOn(document, 'removeEventListener');
  try {
    click(1);
    const pointerListener = add.mock.calls.find(([type]) => type === 'pointerdown')![1];
    const keyListener = add.mock.calls.find(([type]) => type === 'keydown')![1];
    if (end === 'replay') {
      opts.replicaReady = true;
      view.rerender();
    } else if (end === 'outside input') {
      act(() => document.body.dispatchEvent(new Event('pointerdown', { bubbles: true })));
    } else {
      view.unmount();
    }
    expect(remove).toHaveBeenCalledWith('pointerdown', pointerListener, true);
    expect(remove).toHaveBeenCalledWith('keydown', keyListener, true);
  } finally {
    add.mockRestore();
    remove.mockRestore();
  }
});

test('newer keyboard input drops the pending gesture without refocusing', () => {
  const { opts, selections, words, focused } = options();
  const view = renderHook(() => usePagesPointer(opts));

  click(2);
  act(() => view.result.current.bumpInputEpoch());
  const beforeReplay = focused();

  opts.replicaReady = true;
  view.rerender();
  expect(words).toEqual([]);
  expect(selections).toEqual([]);
  expect(focused()).toBe(beforeReplay);
});

test('newer input stops a pending drag from restoring a stale range', () => {
  const { opts, selections } = options();
  const view = renderHook(() => usePagesPointer(opts));

  mouse('mousedown', 200);
  mouse('mousemove', 350);
  act(() => view.result.current.bumpInputEpoch());
  mouse('mousemove', 450);
  mouse('mouseup', 450);

  opts.replicaReady = true;
  view.rerender();
  expect(selections).toEqual([]);
});

test('a newer mousedown replaces the pending word selection with its caret', () => {
  const { opts, selections, words } = options();
  const view = renderHook(() => usePagesPointer(opts));

  click(2);
  click(1, 450);

  opts.replicaReady = true;
  view.rerender();
  expect(words).toEqual([]);
  expect(selections).toEqual([[45, 45, 'body']]);
});

test('a newer mousedown with no body position drops the pending gesture', () => {
  const { opts, selections, words } = options();
  const view = renderHook(() => usePagesPointer(opts));

  click(2);
  click(1, 200, 40);

  opts.replicaReady = true;
  view.rerender();
  expect(words).toEqual([]);
  expect(selections).toEqual([]);
});

test('a consumed gesture cannot replay on a later readiness transition', () => {
  const { opts, words, focused } = options();
  const view = renderHook(() => usePagesPointer(opts));

  click(2);
  opts.replicaReady = true;
  view.rerender();
  const afterReplay = focused();

  opts.replicaReady = false;
  view.rerender();
  opts.replicaReady = true;
  view.rerender();
  expect(words).toEqual([[20, 'body']]);
  expect(focused()).toBe(afterReplay);
});

test('a document swap clears the pending gesture', () => {
  const { opts, selections, words } = options();
  const view = renderHook(() => usePagesPointer(opts));

  click(2);
  opts.yrsSession = { cellSelection: () => null } as unknown as YrsSession;
  opts.replicaReady = true;
  view.rerender();
  expect(words).toEqual([]);
  expect(selections).toEqual([]);
});

test('without a pending replica, selection stays immediate and is never replayed', () => {
  const { opts, projection, selections, words, focused } = options({ replicaPending: () => false });
  opts.getYrsPositionProjection = () => projection;
  const view = renderHook(() => usePagesPointer(opts));

  click(2);
  expect(selections).toEqual([[20, 20, 'body']]);
  expect(words).toEqual([[20, 'body']]);
  const beforeReady = focused();

  opts.replicaReady = true;
  view.rerender();
  expect(selections).toEqual([[20, 20, 'body']]);
  expect(words).toEqual([[20, 'body']]);
  expect(focused()).toBe(beforeReady);
});

test('a missing projection alone does not record a gesture', () => {
  const { opts, selections, words } = options({ replicaPending: () => false });
  const view = renderHook(() => usePagesPointer(opts));

  click(2);
  opts.replicaReady = true;
  view.rerender();
  expect(words).toEqual([]);
  expect(selections).toEqual([]);
});

for (const [name, navigate] of [
  ['setSelection', (ref: PagedEditorRef) => ref.setSelection(45)],
  ['selectAll', (ref: PagedEditorRef) => ref.selectAll()],
  ['scrollToPosition', (ref: PagedEditorRef) => ref.scrollToPosition(45)],
  ['scrollToParaId', (ref: PagedEditorRef) => ref.scrollToParaId('paragraph')],
  ['scrollToCommentId', (ref: PagedEditorRef) => ref.scrollToCommentId(1)],
  ['scrollToChangeId', (ref: PagedEditorRef) => ref.scrollToChangeId(1)],
  ['highlightRange', (ref: PagedEditorRef) => ref.highlightRange(20, 45)],
  ['scrollToPage', (ref: PagedEditorRef) => ref.scrollToPage(1)],
  ['revealDisplayPosition', (ref: PagedEditorRef) => ref.revealDisplayPosition(45)],
] as const) {
  test(`${name} supersedes a gesture while the replica is loading`, () => {
    const { opts, words, selections } = options({
      yrsSession: {
        cellSelection: () => null,
        resolveComment: () => [],
        listRevisions: () => [],
      } as unknown as YrsSession,
    });
    const ref = createRef<PagedEditorRef>();
    const view = renderHook(() => {
      const pointer = usePagesPointer(opts);
      usePagedEditorRefApi({
        ref,
        bumpInputEpoch: pointer.bumpInputEpoch,
        yrsInputRef: opts.yrsInputRef,
        layout: null,
        runLayoutPipeline: () => {},
        getLayoutRequest: () => null,
        scrollToPositionImpl: () => {},
        revealPositionImpl: () => 'layout-unavailable',
        scrollToParaIdImpl: () => false,
        scrollToPageImpl: () => {},
        setIsFocused: () => {},
        onReadyRef: { current: undefined },
        documentFromYrs: () => null,
        yrsSession: opts.yrsSession,
        replicaReady: opts.replicaReady,
        yrsLocToDisplayPosition: () => null,
        syncYrsInputState: () => false,
        applyYrsFormatting: () => false,
        applyYrsCommand: () => false,
        getYrsPositionProjection: () => opts.getYrsPositionProjection('body'),
        displayPositionToYrsLoc: () => null,
        getPositionAtPoint: () => null,
      });
      return pointer;
    });

    click(2);
    act(() => navigate(ref.current!));

    opts.replicaReady = true;
    view.rerender();
    expect(words).toEqual([]);
    expect(selections).toEqual([]);
  });
}
