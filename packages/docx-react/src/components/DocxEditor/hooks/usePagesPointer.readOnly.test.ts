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

test('a header or footer click opens part editing only when editable, with or without a pending replica', () => {
  const bands = fakeQueries();
  bands.hitTestRegions = (_pageIndex: number, x: number, y: number): DisplayListRegionHit =>
    y < HEADER_BOTTOM
      ? { region: 'header', rId: 'rId7', pos: 1, target: 'text' }
      : y >= NOTE_TOP
        ? { region: 'footer', rId: 'rId8', pos: 1, target: 'text' }
        : { region: 'body', pos: Math.floor(x / 10), target: 'text' };
  const press = (y: number, detail: number) => {
    mouse('mousedown', 400, y, canvasOf(), detail);
    mouse('mouseup', 400, y, window, detail);
    mouse('click', 400, y, canvasOf(), detail);
  };
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
