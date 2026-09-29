import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { createRef } from 'react';
import { createRenderedDomContext } from '@betteroffice/docx/plugin-api/RenderedDomContext';
import {
  applyFrameDeltaOwned,
  FRAME_DELTA_VERSION,
  type DecodedFrameDelta,
  type DisplayList,
  type DisplayListQueries,
  type DisplayPage,
  type DisplayPrimitive,
  type RetainedFrame,
} from '@betteroffice/docx/layout/render';
import { CanvasInteractiveOverlay } from './CanvasInteractiveOverlay';
import { CanvasPageMirror } from './CanvasPageMirror';
import { CanvasPagesView } from './CanvasPagesView';

const { act, cleanup, render } = await import('@testing-library/react');

/** Lets chrome that waits for idle time build. */
const idle = () => act(() => new Promise<void>((resolve) => setTimeout(resolve, 200)));

const blankPages = (
  count: number,
  primitives: (index: number) => DisplayPrimitive[] = () => []
): DisplayPage[] =>
  Array.from({ length: count }, (_, pageIndex) => ({
    pageIndex,
    width: 100,
    height: 100,
    primitives: primitives(pageIndex),
  }));

/** The pages under `host` whose mirror holds content. */
const mirroredPages = (host: HTMLElement) =>
  Array.from(host.querySelectorAll<HTMLElement>('.canvas-page'))
    .filter((page) => page.querySelector('.canvas-page-mirror')?.firstElementChild)
    .map((page) => Number(page.dataset.pageIndex));

afterEach(() => {
  cleanup();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

test('only pages in the window and the page holding focus carry a mirror', async () => {
  const hostRef = createRef<HTMLDivElement>();
  const displayList: DisplayList = { pages: blankPages(40) };
  render(
    <CanvasPagesView displayList={displayList} hostRef={hostRef} glyphOutlineProvider={() => ''} />
  );
  await idle();
  const host = hostRef.current!;
  const withMirror = () => mirroredPages(host);

  const mirrored = withMirror();
  expect(mirrored[0]).toBe(0);
  expect(mirrored.length).toBeLessThan(displayList.pages.length);
  expect(host.querySelectorAll('.canvas-page canvas')).toHaveLength(displayList.pages.length);

  const far = host.querySelector<HTMLElement>('.canvas-page[data-page-index="35"] canvas')!;
  far.tabIndex = 0;
  await act(async () => far.focus());
  await idle();
  expect(withMirror()).toContain(35);
  await act(async () => far.blur());
  expect(withMirror()).not.toContain(35);
});

function framed(pageIds: bigint[]): RetainedFrame {
  const pages = pageIds.map((_, pageIndex) => ({
    pageIndex,
    width: 100,
    height: 100,
    primitives: [],
  }));
  return {
    protocolVersion: 1,
    docEpoch: 1,
    layoutEpoch: 1,
    frameEpoch: 1,
    pages: pageIds.map((pageId, pageIndex) => ({
      pageId,
      pageIndex,
      fingerprint: 0n,
      primitiveIds: new BigUint64Array(),
      page: pages[pageIndex]!,
    })),
    damagedPageIds: new Set(pageIds),
    removedPageIds: new Set(),
    displayList: { pages },
  };
}

test('the focus pin follows its page when pages before it are inserted', async () => {
  const hostRef = createRef<HTMLDivElement>();
  const ids = Array.from({ length: 40 }, (_, index) => BigInt(index + 1));
  const view = (frame: RetainedFrame) => (
    <CanvasPagesView
      displayList={frame.displayList}
      frame={frame}
      hostRef={hostRef}
      glyphOutlineProvider={() => ''}
    />
  );
  const { rerender } = render(view(framed(ids)));
  await idle();
  const host = hostRef.current!;
  const mirroredKeys = () =>
    Array.from(host.querySelectorAll<HTMLElement>('.canvas-page'))
      .filter((page) => page.querySelector('.canvas-page-mirror')?.firstElementChild)
      .map((page) => page.dataset.pageKey);
  const pinned = host.querySelector<HTMLElement>('.canvas-page[data-page-key="36"] canvas')!;
  pinned.tabIndex = 0;
  await act(async () => pinned.focus());
  await idle();
  expect(mirroredKeys()).toContain('36');

  rerender(view(framed([99n, ...ids])));
  await idle();
  expect(mirroredKeys()).toContain('36');
  expect(mirroredKeys()).not.toContain('35');
});

test('changing when chrome builds keeps what it already built', async () => {
  const page: DisplayPage = { pageIndex: 0, width: 100, height: 100, primitives: [] };
  for (const Chrome of [CanvasPageMirror, CanvasInteractiveOverlay]) {
    const { container, rerender, unmount } = render(<Chrome page={page} defer={false} />);
    await act(async () => {});
    const built = container.firstElementChild!.firstElementChild;
    expect(built).not.toBeNull();
    rerender(<Chrome page={page} defer />);
    await act(async () => {});
    expect(container.firstElementChild!.firstElementChild).toBe(built);
    unmount();
  }
});

test('a page in the window rebuilds its chrome at once, and one outside it at idle', async () => {
  const page = (): DisplayPage => ({ pageIndex: 0, width: 100, height: 100, primitives: [] });
  for (const [Chrome, visible, atOnce] of [
    [CanvasInteractiveOverlay, false, true],
    [CanvasPageMirror, true, true],
    [CanvasPageMirror, false, false],
  ] as const) {
    const { container, rerender, unmount } = render(
      <Chrome page={page()} defer visible={visible} />
    );
    await idle();
    const built = container.firstElementChild!.firstElementChild;
    expect(built).not.toBeNull();
    rerender(<Chrome page={page()} defer visible={visible} />);
    await act(async () => {});
    expect(container.firstElementChild!.firstElementChild === built).toBe(!atOnce);
    unmount();
  }
});

const widget = (groupId: string): DisplayPrimitive =>
  ({
    kind: 'rect',
    x: 10,
    y: 10,
    w: 10,
    h: 10,
    inlineSdtWidget: { kind: 'checkbox', groupId, pos: 1 },
  }) as DisplayPrimitive;

test('Tab reaches a content control on a page whose chrome is not built', async () => {
  const hostRef = createRef<HTMLDivElement>();
  const displayList: DisplayList = {
    pages: blankPages(40, (index) =>
      index === 0
        ? [widget('first'), { ...widget('second'), x: 50 } as DisplayPrimitive]
        : index === 30
          ? [widget('far')]
          : []
    ),
  };
  render(
    <CanvasPagesView
      displayList={displayList}
      hostRef={hostRef}
      interactive
      glyphOutlineProvider={() => ''}
    />
  );
  await idle();
  const host = hostRef.current!;
  const control = (groupId: string) =>
    host.querySelector<HTMLButtonElement>(`button[data-sdt-group-id="${groupId}"]`);
  expect(control('far')).toBeNull();
  const press = (groupId: string) => {
    const event = new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true });
    control(groupId)!.dispatchEvent(event);
    return event;
  };
  // The next control on the same page is the browser's to reach.
  expect(press('first').defaultPrevented).toBe(false);
  expect(control('far')).toBeNull();
  await act(async () => control('second')!.focus());
  let tab!: KeyboardEvent;
  await act(async () => {
    tab = press('second');
  });
  expect(tab.defaultPrevented).toBe(true);
  expect(document.activeElement).toBe(control('far'));
  await idle();
  expect(document.activeElement).toBe(control('far'));
});

test('a link to a note on a page whose chrome is not built builds that page first', async () => {
  const hostRef = createRef<HTMLDivElement>();
  const reference = {
    kind: 'text',
    x: 10,
    y: 10,
    text: '1',
    font: '11px sans-serif',
    color: '#000',
    noteRef: { id: 7, kind: 'endnote' },
  } as unknown as DisplayPrimitive;
  const pages = blankPages(40, (index) => (index === 0 ? [reference] : []));
  pages[30] = {
    ...pages[30]!,
    noteAreas: [
      {
        kind: 'endnote',
        y: 50,
        height: 20,
        noteIds: [7],
        primitives: [{ ...reference, noteRef: undefined, groupId: 'endnote-7' } as DisplayPrimitive],
      },
    ],
  };
  render(
    <CanvasPagesView displayList={{ pages }} hostRef={hostRef} glyphOutlineProvider={() => ''} />
  );
  await idle();
  const host = hostRef.current!;
  expect(document.getElementById('oox-endnote-7')).toBeNull();
  const link = host.querySelector<HTMLAnchorElement>('a[href="#oox-endnote-7"]')!;
  await act(async () => {
    link.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
  });
  expect(document.getElementById('oox-endnote-7')).not.toBeNull();
});

test('a plugin DOM query for a range on a page whose chrome is not built finds its elements', async () => {
  const hostRef = createRef<HTMLDivElement>();
  const run = (docStart: number) =>
    ({
      kind: 'text',
      x: 10,
      y: 10,
      text: 'far',
      font: '11px sans-serif',
      color: '#000',
      blockId: `p${docStart}`,
      docStart,
      docEnd: docStart + 3,
    }) as unknown as DisplayPrimitive;
  const pages = blankPages(40, (index) => [run(index * 10)]);
  render(<CanvasPagesView displayList={{ pages }} hostRef={hostRef} glyphOutlineProvider={() => ''} />);
  await idle();
  const host = hostRef.current!;
  expect(mirroredPages(host)).not.toContain(30);
  const queries = {
    rangeRects: (from: number) => [{ pageIndex: Math.floor(from / 10), x: 0, y: 0, width: 1, height: 1 }],
  } as unknown as DisplayListQueries;
  const context = createRenderedDomContext(host, 1, {
    displayListQueries: queries,
    projector: { projectRect: () => null, getPageBounds: () => null },
  });
  let elements: Element[] = [];
  await act(async () => {
    elements = context.findElementsForRange(300, 303);
  });
  expect(elements).toHaveLength(1);
  expect(elements[0]!.isConnected).toBe(true);
  expect(mirroredPages(host)).toContain(30);
});

test('a position shift in place moves a control at once, and the mirror by idle time', async () => {
  const page = {
    pageIndex: 0,
    width: 100,
    height: 100,
    primitives: [
      { ...widget('shifted'), docStart: 1, docEnd: 2, blockId: 'p' } as DisplayPrimitive,
    ],
  } as DisplayPage;
  const first: RetainedFrame = {
    protocolVersion: FRAME_DELTA_VERSION,
    docEpoch: 1,
    layoutEpoch: 1,
    frameEpoch: 1,
    pages: [
      { pageId: 1n, pageIndex: 0, fingerprint: 1n, primitiveIds: new BigUint64Array([1n]), page },
    ],
    damagedPageIds: new Set([1n]),
    removedPageIds: new Set(),
    displayList: { pages: [page] },
  };
  const hostRef = createRef<HTMLDivElement>();
  const view = (frame: RetainedFrame) => (
    <CanvasPagesView
      displayList={frame.displayList}
      frame={frame}
      hostRef={hostRef}
      interactive
      glyphOutlineProvider={() => ''}
    />
  );
  const { rerender } = render(view(first));
  await act(async () => {});
  const host = hostRef.current!;
  const controlPos = () =>
    host.querySelector<HTMLElement>('button[data-sdt-group-id="shifted"]')?.dataset.sdtPos;
  const mirrorStart = () =>
    host.querySelector<HTMLElement>('.canvas-page-mirror [data-doc-start]')?.dataset.docStart;
  expect(controlPos()).toBe('1');
  expect(mirrorStart()).toBe('1');

  const shift: DecodedFrameDelta = {
    protocolVersion: FRAME_DELTA_VERSION,
    full: false,
    docEpoch: 2,
    layoutEpoch: 2,
    frameEpoch: 2,
    baseFrameEpoch: 1,
    pageCount: 1,
    operations: [
      {
        kind: 'shift-positions',
        pageIndex: 0,
        pageId: 1n,
        fingerprint: 2n,
        // docStart, docEnd and the inline widget's position
        runs: [{ start: 0, count: 1, changedMask: 1 | 2 | 16, delta: 5 }],
        anchors: [],
      },
    ],
    bytes: new Uint8Array(),
  };
  const second = applyFrameDeltaOwned(first, shift);
  expect(second.displayList.pages[0]).toBe(page);
  rerender(view(second));
  await act(async () => {});
  expect(controlPos()).toBe('6');
  await idle();
  expect(mirrorStart()).toBe('6');
});
