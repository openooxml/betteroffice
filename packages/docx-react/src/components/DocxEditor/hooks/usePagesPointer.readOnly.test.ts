import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, expect, test } from 'bun:test';
import { createElement, createRef, useState } from 'react';
import type { DisplayListQueries, DisplayListRegionHit } from '@betteroffice/docx/layout/render';
import type { ResidentDocumentRead, ResidentEngineWorkerClient, YrsSession } from '@betteroffice/docx/yrs';
import { ViewerInput } from '../ViewerInput';
import { stampWorkerFrameVersion } from '../internals/layoutProvenance';
import type { YrsInputRef } from '../YrsInput';
import type { YrsPositionProjection } from '../internals/yrsPositionProjection';
import { partEditStory, type PartEdit } from '../partEdit';
import { usePagesPointer, type UsePagesPointerOptions } from './usePagesPointer';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, renderHook } = await import('@testing-library/react');

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

function options(
  overrides: Partial<UsePagesPointerOptions> = {},
  shown: { anchor: number; head: number } | null = null
) {
  const selections: Array<[number, number]> = [];
  const words: number[] = [];
  let focused = 0;
  const input = {
    focus: () => {
      focused += 1;
    },
    setSelectionFromDisplay: (anchor: number, head: number) => selections.push([anchor, head]),
    displaySelection: () => shown,
    selectWordAtDisplay: (position: number) => words.push(position),
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
  return { opts, selections, noteClicks, words, focused: () => focused };
}

function mouse(type: string, clientX: number, clientY: number, target: EventTarget, detail = 0): void {
  act(() => {
    target.dispatchEvent(
      new MouseEvent(type, { bubbles: true, cancelable: true, clientX, clientY, button: 0, detail })
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

test('P1-6: a pending viewer bookmark cannot override a newer keyboard selection', async () => {
  const linked = fakeQueries();
  linked.isReady = () => true;
  linked.visualLinesOnPage = (pageIndex) => (pageIndex === 0 ? [{ pageIndex: 0, from: 1, to: 80 }] : []) as unknown as
    ReturnType<DisplayListQueries['visualLinesOnPage']>;
  linked.visualLineAtPosition = (position) => (position >= 1 && position <= 80
    ? { pageIndex: 0, from: 1, to: 80 }
    : null) as
    ReturnType<DisplayListQueries['visualLineAtPosition']>;
  (linked.displayList.pages[0] as { primitives: unknown[] }).primitives = [{
    kind: 'text', text: 'linked text', x: 0, baselineY: 410, width: 800,
    font: '400 16px Calibri', color: '#000000', docStart: 1, docEnd: 80, href: '#target',
  }];
  stampWorkerFrameVersion(linked, 'A', false, false);
  const pending: Array<{ request: ResidentDocumentRead; resolve(value: { version: string; value: unknown }): void }> = [];
  const read = ((request: ResidentDocumentRead) => new Promise<{ version: string; value: unknown }>((resolve) => {
    pending.push({ request, resolve });
  })) as unknown as ResidentEngineWorkerClient['documentRead'];
  const ref = createRef<YrsInputRef>();
  const input = render(createElement(ViewerInput, {
    ref, read, story: 'body', queries: linked, document: { isDisplayOnly: () => false },
    onSelectionChange: () => {},
  }));
  const scrolled: number[] = [];
  const { opts } = options({
    viewerSelection: true,
    yrsSession: null,
    yrsInputRef: ref,
    displayListQueries: linked,
    getYrsPositionProjection: () => null,
    scrollToPositionImpl: (position) => { scrolled.push(position); },
    resolveBookmarkPosition: async (name) => (await read({
      kind: 'bookmarkPosition', story: 'body', name, expectVersion: 'A',
    })).value,
  });
  renderHook(() => usePagesPointer(opts));
  mouse('mousedown', 200, 405, canvasOf(), 1);
  mouse('mouseup', 200, 405, window, 1);
  mouse('click', 200, 405, canvasOf(), 1);
  const bookmark = pending.find((entry) => entry.request.kind === 'bookmarkPosition');
  expect(bookmark).toBeDefined();
  fireEvent.keyDown(input.getByTestId('yrs-input'), { key: 'ArrowRight', shiftKey: true });
  expect(ref.current!.displaySelection()).toEqual({ anchor: 20, head: 21 });
  await act(async () => { bookmark!.resolve({ version: 'A', value: 60 }); });
  expect(ref.current!.displaySelection()).toEqual({ anchor: 20, head: 21 });
  expect(scrolled).toEqual([]);
});

test('a right-click inside a viewer selection keeps it for the context-menu copy', async () => {
  const queries = fakeQueries();
  stampWorkerFrameVersion(queries, 'A', false, false);
  const pending: Array<{ request: ResidentDocumentRead; resolve(value: { version: string; value: unknown }): void }> = [];
  const read = ((request: ResidentDocumentRead) => new Promise<{ version: string; value: unknown }>((resolve) => {
    pending.push({ request, resolve });
  })) as unknown as ResidentEngineWorkerClient['documentRead'];
  const ref = createRef<YrsInputRef>();
  render(createElement(ViewerInput, {
    ref, read, story: 'body', queries, document: { isDisplayOnly: () => false },
    onSelectionChange: () => {},
  }));
  let gesture = 0;
  const beginGesture = ref.current!.beginGesture!;
  ref.current!.beginGesture = () => { gesture = beginGesture(); return gesture; };
  const { opts } = options({
    viewerSelection: true,
    yrsSession: null,
    yrsInputRef: ref,
    displayListQueries: queries,
    getYrsPositionProjection: () => null,
  });
  renderHook(() => usePagesPointer(opts));
  mouse('mousedown', 200, 400, canvasOf());
  mouse('mousemove', 450, 600, window);
  await nextFrame();
  mouse('mouseup', 450, 600, window);
  expect(ref.current!.displaySelection()).toEqual({ anchor: 20, head: 45 });
  const selectionGesture = gesture;
  expect(ref.current!.isGestureCurrent!(selectionGesture)).toBe(true);
  const pendingCopy = ref.current!.readSelectedText!()!.catch((error: Error) => error);
  fireEvent.mouseDown(canvasOf(), { button: 2, clientX: 300, clientY: 400 });
  expect(gesture).toBe(selectionGesture);
  expect(ref.current!.isGestureCurrent!(selectionGesture)).toBe(true);
  expect(ref.current!.displaySelection()).toEqual({ anchor: 20, head: 45 });
  const copy = ref.current!.readSelectedText!()!;
  const caret = pending.shift()!;
  expect(caret.request).toEqual({
    kind: 'selectionText', story: 'body', anchor: 20, head: 20, expectVersion: 'A',
  });
  await act(async () => { caret.resolve({ version: 'A', value: { text: '', range: null } }); });
  const range = pending.shift()!;
  expect(range.request).toEqual({
    kind: 'selectionText', story: 'body', anchor: 20, head: 45, expectVersion: 'A',
  });
  await act(async () => { range.resolve({ version: 'A', value: { text: 'selected text', range: null } }); });
  expect(await pendingCopy).toBe('selected text');
  expect(await copy).toBe('selected text');
});

const ignoredInputs: Array<[string, (handleEditorKeyDown: (e: never) => void) => void]> = [
  ['a header press', () => mouse('mousedown', 50, 40, canvasOf())],
  ['a footnote press', () => mouse('mousedown', 50, 950, canvasOf())],
  ['a middle-button press', () => { fireEvent.mouseDown(canvasOf(), { button: 1, clientX: 300, clientY: 400 }); }],
  ['a key outside the hidden input', (handleEditorKeyDown) => {
    act(() => handleEditorKeyDown({ target: document.body } as never));
  }],
];
for (const [name, ignore] of ignoredInputs) {
  test(`${name} hides the viewer selection it does not replace`, async () => {
    const queries = fakeQueries();
    stampWorkerFrameVersion(queries, 'A', false, false);
    const read = (() => new Promise(() => {})) as unknown as ResidentEngineWorkerClient['documentRead'];
    const ref = createRef<YrsInputRef>();
    render(createElement(ViewerInput, {
      ref, read, story: 'body', queries, document: { isDisplayOnly: () => false },
      onSelectionChange: () => {},
    }));
    const { opts } = options({
      viewerSelection: true,
      yrsSession: null,
      yrsInputRef: ref,
      displayListQueries: queries,
      getYrsPositionProjection: () => null,
    });
    const { result } = renderHook(() => usePagesPointer(opts));
    mouse('mousedown', 200, 400, canvasOf());
    mouse('mousemove', 450, 600, window);
    await nextFrame();
    mouse('mouseup', 450, 600, window);
    expect(ref.current!.displaySelection()).toEqual({ anchor: 20, head: 45 });
    ignore(result.current.handleEditorKeyDown);
    expect(ref.current!.displaySelection()).toBeNull();
    expect(ref.current!.readSelectedText!()).toBeNull();
  });
}

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

/** A header above y 80 and a footer below y 900; `shift` moves the content under the pointer. */
function bandQueries(shift: () => number = () => 0): DisplayListQueries {
  const bands = fakeQueries();
  bands.hitTestRegions = (_pageIndex: number, x: number, y: number): DisplayListRegionHit => {
    const at = y + shift();
    return at < HEADER_BOTTOM
      ? { region: 'header', rId: 'rId7', pos: 1, target: 'text' }
      : at >= NOTE_TOP
        ? { region: 'footer', rId: 'rId8', pos: 1, target: 'text' }
        : { region: 'body', pos: Math.floor(x / 10), target: 'text' };
  };
  return bands;
}

function press(y: number, detail: number): void {
  mouse('mousedown', 400, y, canvasOf(), detail);
  mouse('mouseup', 400, y, window, detail);
  mouse('click', 400, y, canvasOf(), detail);
}

test('a header or footer click opens part editing only when editable, with or without a pending replica', () => {
  const bands = bandQueries();
  for (const readOnly of [true, false]) {
    for (const pending of [false, true]) {
      const opened: Array<['header' | 'footer', number | undefined]> = [];
      const { opts, selections, words } = options({
        readOnly,
        displayListQueries: bands,
        replicaPending: () => pending,
        onHeaderFooterDoubleClick: (region, pageNumber) => opened.push([region, pageNumber]),
      });
      const view = renderHook(() => usePagesPointer(opts));
      press(40, 1);
      press(950, 1);
      expect(opened).toEqual([]);
      press(40, 2);
      press(950, 2);
      expect(opened).toEqual(readOnly ? [] : [['header', 1], ['footer', 1]]);
      if (readOnly) expect({ selections, words }).toEqual({ selections: [], words: [] });
      view.unmount();
    }
  }
});

test('an editable double-click opens the header or footer it lands on and leaves the caret', () => {
  for (const pending of [false, true]) {
    let caretMoved = () => false;
    // a caret the press places scrolls the page, so the release lands lower
    const bands = bandQueries(() => (caretMoved() ? NOTE_TOP : 0));
    const opened: Array<['header' | 'footer', number | undefined]> = [];
    const { opts, selections, words } = options({
      readOnly: false,
      displayListQueries: bands,
      replicaPending: () => pending,
      onHeaderFooterDoubleClick: (region, pageNumber) => opened.push([region, pageNumber]),
    });
    caretMoved = () => selections.length > 0;
    const view = renderHook(() => usePagesPointer(opts));
    press(40, 1);
    press(950, 1);
    expect({ opened, selections, words }).toEqual({ opened: [], selections: [], words: [] });
    press(40, 1);
    press(40, 2);
    expect(opened).toEqual([['header', 1]]);
    press(950, 1);
    press(950, 2);
    expect({ opened, selections, words }).toEqual({
      opened: [
        ['header', 1],
        ['footer', 1],
      ],
      selections: [],
      words: [],
    });
    view.unmount();
  }
});

/**
 * A host that opens whatever part the pointer asks for. Closing a part returns
 * the body caret, which scrolls the page, so later presses land lower.
 */
function partHost(initial: PartEdit, queries: (shift: () => number) => DisplayListQueries) {
  let scrolled = false;
  let closed = 0;
  const opened: Array<['header' | 'footer', number | undefined]> = [];
  const carets: Array<[number, string]> = [];
  const words: Array<[number, string]> = [];
  const input = {
    focus: () => {},
    displaySelection: () => null,
    setSelectionFromDisplay: (anchor: number, _head: number, story: string) =>
      carets.push([anchor, story]),
    selectWordAtDisplay: (position: number, story: string) => words.push([position, story]),
  } as unknown as YrsInputRef;
  const bands = queries(() => (scrolled ? NOTE_TOP : 0));
  bands.imageAtPoint = () => null;
  const { opts } = options({
    readOnly: false,
    displayListQueries: bands,
    yrsInputRef: { current: input },
    getYrsPositionProjection: (story) =>
      ({
        size: 100,
        targetAt: (displayPosition: number) => ({ story, displayPosition }),
        tableAtPosition: () => null,
        cellPosition: () => null,
        nodeAt: () => null,
      }) as unknown as YrsPositionProjection,
  });
  renderHook(() => {
    const [partEdit, setPartEdit] = useState<PartEdit | null>(initial);
    return usePagesPointer({
      ...opts,
      partEdit,
      yrsRootStory: partEditStory(partEdit),
      onBodyClick: () => {
        closed += 1;
        scrolled = true;
        setPartEdit(null);
      },
      onHeaderFooterDoubleClick: (region, pageNumber) => {
        opened.push([region, pageNumber]);
        setPartEdit({ kind: region, rId: region === 'header' ? 'rId7' : 'rId8' });
      },
    });
  });
  return { opened, carets, words, closed: () => closed };
}

test('with the footer open, the header opens directly: no return to the body, no scroll', () => {
  const { opened, carets, words, closed } = partHost({ kind: 'footer', rId: 'rId8' }, bandQueries);
  press(40, 1);
  press(40, 2);
  expect({ opened, carets, words, closed: closed() }).toEqual({
    opened: [['header', 1]],
    carets: [
      [1, 'hf:rId7'],
      [1, 'hf:rId7'],
    ],
    words: [[1, 'hf:rId7']],
    closed: 0,
  });
  press(950, 1);
  expect(opened.at(-1)).toEqual(['footer', 1]);
  expect(carets.at(-1)).toEqual([1, 'hf:rId8']);
  press(400, 1);
  expect(closed()).toBe(1);
});

test('with a note open, a band click leaves the note open and a double-click opens the band', () => {
  const { opened, carets, closed } = partHost({ kind: 'footnote', noteId: 2 }, (shift) => {
    const queries = fakeQueries();
    const hitTest = queries.hitTestRegions;
    queries.hitTestRegions = (pageIndex, x, y) => hitTest(pageIndex, x, y + shift());
    return queries;
  });
  press(40, 1);
  expect({ opened, carets, closed: closed() }).toEqual({ opened: [], carets: [], closed: 0 });
  press(40, 1);
  press(40, 2);
  expect({ opened, closed: closed() }).toEqual({ opened: [['header', 1]], closed: 0 });
});

test('the table insert button hides and inserts nothing once the editor turns read-only', () => {
  const grid = fakeQueries();
  const cell = (row: number, col: number) => ({
    kind: 'rect',
    x: 10 + col * 100,
    y: 200 + row * 20,
    w: 100,
    h: 20,
    fill: '#ffffff',
    cell: { row, col, rowSpan: 1, colSpan: 1 },
    table: { tableId: 'A' },
    docStart: 5 + row * 4 + col * 2,
    docEnd: 6 + row * 4 + col * 2,
  });
  (grid.displayList.pages[0] as { primitives: unknown[] }).primitives = [
    cell(0, 0),
    cell(0, 1),
    cell(1, 0),
    cell(1, 1),
  ];
  const at = { story: 'body', tableIndex: 0, row: 0, column: 0 };
  const projection = {
    size: 100,
    targetAt: (position: number) => ({ story: 'body', displayPosition: position, cell: at }),
    tableAtPosition: () => ({ start: 3 }),
    cellPosition: (_start: number, row: number, col: number) => 5 + row * 4 + col * 2,
    nodeAt: () => null,
  } as unknown as YrsPositionProjection;
  for (const switched of [false, true]) {
    const commands: unknown[] = [];
    const { opts } = options({
      displayListQueries: grid,
      canvasOverlayTarget: document.body,
      getYrsPositionProjection: () => projection,
      yrsSession: { cellSelection: () => null, setCellSelection: () => {} } as unknown as YrsSession,
      applyYrsCommand: (command) => {
        commands.push(command);
        return true;
      },
    });
    const view = renderHook(
      ({ readOnly }: { readOnly: boolean }) => usePagesPointer({ ...opts, readOnly }),
      { initialProps: { readOnly: false } }
    );
    mouse('mousemove', 10, 220, canvasOf());
    expect(view.result.current.tableInsertButton?.type).toBe('row');
    if (switched) view.rerender({ readOnly: true });
    expect(view.result.current.tableInsertButton === null).toBe(switched);
    act(() =>
      view.result.current.handleTableInsertClick({
        preventDefault: () => {},
        stopPropagation: () => {},
      } as unknown as React.MouseEvent)
    );
    expect(commands).toHaveLength(switched ? 0 : 1);
    view.unmount();
  }
});

test('an editable press still selects the image under it', () => {
  const { opts, selections } = options({ readOnly: false });
  renderHook(() => usePagesPointer(opts));

  mouse('mousedown', 50, 400, canvasOf());
  expect(selections).toEqual([[5, 6]]);
});

test('a right-click inside a selected cell range offers Copy only when read-only', () => {
  const cell = (row: number, column: number) => ({ story: 'body', tableIndex: 0, row, column });
  const table = {
    size: 100,
    targetAt: (position: number) => ({ story: 'body:t0:r0c0', displayPosition: position, cell: cell(0, 0) }),
    tableAtPosition: () => null,
    cellPosition: () => null,
    nodeAt: () => null,
  } as unknown as YrsPositionProjection;
  const session = {
    cellSelection: () => ({ anchor: cell(0, 0), head: cell(1, 1) }),
  } as unknown as YrsSession;
  for (const readOnly of [true, false]) {
    const menus: Array<{ hasSelection: boolean }> = [];
    const { opts } = options({
      readOnly,
      yrsSession: session,
      getYrsPositionProjection: () => table,
      onContextMenu: (data) => menus.push(data),
    });
    const view = renderHook(() => usePagesPointer(opts));
    mouse('contextmenu', 200, 400, canvasOf());
    expect(menus.map((menu) => menu.hasSelection)).toEqual([readOnly]);
    view.unmount();
  }
});

test('read-only selection wins over a link: double-click and a drag ending on it', () => {
  const linked = fakeQueries();
  (linked.displayList.pages[0] as { primitives: unknown[] }).primitives = [
    {
      kind: 'text',
      text: 'linked text',
      x: 0,
      baselineY: 410,
      width: 800,
      font: '400 16px Calibri',
      color: '#000000',
      docStart: 1,
      docEnd: 12,
      href: 'https://example.com/',
    },
  ];
  const click = (detail: number) =>
    act(() => {
      canvasOf().dispatchEvent(
        new MouseEvent('click', { bubbles: true, clientX: 200, clientY: 405, button: 0, detail })
      );
    });
  const run = (readOnly: boolean, detail: number, shown: { anchor: number; head: number } | null) => {
    const links: string[] = [];
    const { opts, words } = options(
      {
        readOnly,
        displayListQueries: linked,
        canvasOverlayTarget: document.body,
        onHyperlinkClick: (link) => links.push(link.href),
      },
      shown
    );
    const view = renderHook(() => usePagesPointer(opts));
    click(detail);
    view.unmount();
    return { links, words };
  };

  expect(run(true, 1, null)).toEqual({ links: ['https://example.com/'], words: [] });
  expect(run(true, 2, null)).toEqual({ links: [], words: [20] });
  expect(run(true, 1, { anchor: 3, head: 9 })).toEqual({ links: [], words: [] });
  expect(run(false, 2, null).words).toEqual([]);
});

test('read-only selection wins over a link inside a selected cell range', () => {
  const cell = (row: number, column: number) => ({ story: 'body', tableIndex: 0, row, column });
  const table = {
    size: 100,
    targetAt: (position: number) => ({ story: 'body:t0:r1c1', displayPosition: position, cell: cell(1, 1) }),
    tableAtPosition: () => null,
    cellPosition: () => null,
    nodeAt: () => null,
  } as unknown as YrsPositionProjection;
  const linked = fakeQueries();
  (linked.displayList.pages[0] as { primitives: unknown[] }).primitives = [
    { kind: 'text', text: 'link', x: 0, baselineY: 410, width: 800, font: '400 16px Calibri', color: '#000000', docStart: 1, docEnd: 5, href: 'https://example.com/' },
  ];
  for (const [head, followed] of [
    [cell(1, 1), false],
    [cell(0, 0), true],
  ] as const) {
    const links: string[] = [];
    const { opts } = options({
      displayListQueries: linked,
      canvasOverlayTarget: document.body,
      getYrsPositionProjection: () => table,
      yrsSession: {
        cellSelection: () => ({ anchor: cell(0, 0), head }),
        setCellSelection: () => {},
      } as unknown as YrsSession,
      onHyperlinkClick: (link) => links.push(link.href),
    });
    const view = renderHook(() => usePagesPointer(opts));
    act(() => {
      canvasOf().dispatchEvent(
        new MouseEvent('click', { bubbles: true, clientX: 200, clientY: 405, button: 0, detail: 1 })
      );
    });
    view.unmount();
    expect(links.length > 0).toBe(followed);
  }
});
