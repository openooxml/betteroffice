import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, expect, test } from 'bun:test';

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
import type { PageChromeHandle } from './usePageChrome';

const { act, cleanup, render } = await import('@testing-library/react');

const originalIdle = globalThis.requestIdleCallback;
const originalCancelIdle = globalThis.cancelIdleCallback;
const idleWork = new Map<number, { callback: IdleRequestCallback; timeout?: number }>();
let nextIdle = 1;

const flushIdle = (deadline: IdleDeadline): void => {
  for (const [id, { callback }] of Array.from(idleWork)) {
    if (!idleWork.delete(id)) continue;
    callback(deadline);
  }
};

/** Lets chrome and queued fallbacks finish their idle work. */
const idle = () =>
  act(async () => {
    await new Promise<void>((resolve) => setTimeout(resolve, 200));
    while (idleWork.size > 0) {
      flushIdle({ didTimeout: false, timeRemaining: () => 50 });
      await Promise.resolve();
    }
  });

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

const hasFullMirror = (page: Element): boolean => {
  const mirror = page.querySelector<HTMLElement>('.canvas-page-mirror > .layout-page-mirror');
  return Boolean(mirror && mirror.style.contentVisibility !== 'auto');
};

/** The pages under `host` whose full mirror is built. */
const mirroredPages = (host: HTMLElement) =>
  Array.from(host.querySelectorAll<HTMLElement>('.canvas-page'))
    .filter(hasFullMirror)
    .map((page) => Number(page.dataset.pageIndex));

beforeEach(() => {
  globalThis.requestIdleCallback = (callback, options) => {
    const id = nextIdle++;
    idleWork.set(id, { callback, timeout: options?.timeout });
    return id;
  };
  globalThis.cancelIdleCallback = (id) => {
    idleWork.delete(id);
  };
});
afterEach(() => {
  cleanup();
  idleWork.clear();
  globalThis.requestIdleCallback = originalIdle;
  globalThis.cancelIdleCallback = originalCancelIdle;
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
      .filter(hasFullMirror)
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


test('controls on every page and links on pages outside the window stay built', async () => {
  const hostRef = createRef<HTMLDivElement>();
  const link = {
    kind: 'text',
    x: 10,
    y: 10,
    text: 'far link',
    font: '11px sans-serif',
    color: '#000',
    href: '#somewhere',
  } as unknown as DisplayPrimitive;
  const words = { ...link, text: 'plain words', href: undefined } as unknown as DisplayPrimitive;
  const pages = blankPages(40, (index) => (index === 30 ? [widget('far'), link, words] : []));
  render(
    <CanvasPagesView
      displayList={{ pages }}
      hostRef={hostRef}
      interactive
      glyphOutlineProvider={() => ''}
    />
  );
  await idle();
  const page = hostRef.current!.querySelector<HTMLElement>('.canvas-page[data-page-index="30"]')!;
  expect(page.querySelector('button[data-sdt-group-id="far"]')).not.toBeNull();
  const mirror = page.querySelector('.canvas-page-mirror')!;
  expect(mirror.querySelector('a[href="#somewhere"]')?.textContent).toBe('far link');
  expect(mirror.textContent).toContain('plain words');
});

test('a far page keeps the header cells that cells on other pages name', async () => {
  const hostRef = createRef<HTMLDivElement>();
  const header = {
    kind: 'text',
    x: 10,
    y: 10,
    text: 'Account',
    font: '11px sans-serif',
    color: '#000',
    blockKey: 't',
    table: { tableId: 't' },
    cell: { row: 0, col: 0, rowSpan: 1, colSpan: 1, cellId: 'account', isHeader: true },
  } as unknown as DisplayPrimitive;
  const pages = blankPages(40, (index) => (index === 30 ? [header] : []));
  render(<CanvasPagesView displayList={{ pages }} hostRef={hostRef} glyphOutlineProvider={() => ''} />);
  await idle();
  const page = hostRef.current!.querySelector<HTMLElement>('.canvas-page[data-page-index="30"]')!;
  expect(page.querySelector('[id="account"]')?.textContent).toBe('Account');
});

test('pages in the window hold their links and controls before idle time', async () => {
  const hostRef = createRef<HTMLDivElement>();
  const link = {
    kind: 'text',
    x: 10,
    y: 10,
    text: 'near link',
    font: '11px sans-serif',
    color: '#000',
    href: '#near',
  } as unknown as DisplayPrimitive;
  const pages = blankPages(40, (index) => (index === 0 ? [widget('near'), link] : []));
  render(
    <CanvasPagesView
      displayList={{ pages }}
      hostRef={hostRef}
      interactive
      glyphOutlineProvider={() => ''}
    />
  );
  await act(async () => {});
  const page = hostRef.current!.querySelector<HTMLElement>('.canvas-page[data-page-index="0"]')!;
  expect(page.querySelector('button[data-sdt-group-id="near"]')).not.toBeNull();
  expect(page.querySelector('.canvas-page-mirror a[href="#near"]')).not.toBeNull();
});

test('a page leaving the window keeps the links it built', async () => {
  const link = {
    kind: 'text',
    x: 10,
    y: 10,
    text: 'kept link',
    font: '11px sans-serif',
    color: '#000',
    href: '#kept',
  } as unknown as DisplayPrimitive;
  const words = { ...link, text: 'plain words', href: undefined } as unknown as DisplayPrimitive;
  const page: DisplayPage = { pageIndex: 0, width: 100, height: 100, primitives: [link, words] };
  const { container, rerender } = render(<CanvasPageMirror page={page} />);
  await act(async () => {});
  const built = container.querySelector('a[href="#kept"]');
  expect(built).not.toBeNull();
  expect(container.textContent).toContain('plain words');
  rerender(<CanvasPageMirror page={page} active={false} />);
  await act(async () => {});
  expect(container.querySelector('a[href="#kept"]')).toBe(built);
  expect(container.textContent).not.toContain('plain words');
});

test('a rebuild keeps focus on the same link when links before it change', async () => {
  // Links in a content control carry its attributes too.
  for (const sdt of [
    {},
    { sdt: { groupId: 'sdt@5', sdtType: 'richText' } },
    { inlineSdtWidget: { kind: 'dropdown', groupId: 'sdt@5', pos: 5 } },
  ]) {
    const link = (href: string, y: number) =>
      ({
        kind: 'text',
        x: 10,
        y,
        text: `link ${href}`,
        font: '11px sans-serif',
        color: '#000',
        href,
        ...sdt,
      }) as unknown as DisplayPrimitive;
    const page = (primitives: DisplayPrimitive[]): DisplayPage => ({
      pageIndex: 0,
      width: 100,
      height: 100,
      primitives,
    });
    const { container, rerender, unmount } = render(
      <CanvasPageMirror page={page([link('#a', 10), link('#b', 30), link('#c', 50)])} />
    );
    await act(async () => {});
    container.querySelector<HTMLElement>('a[href="#b"]')!.focus();
    rerender(<CanvasPageMirror page={page([link('#b', 10), link('#c', 30)])} />);
    await act(async () => {});
    expect(document.activeElement?.getAttribute('href')).toBe('#b');
    expect(container.contains(document.activeElement)).toBe(true);
    unmount();
  }
});

test('a rebuild keeps focus on a control that changed, moved or lost a neighbour', async () => {
  const checkbox = (controlId: number, pos: number, checked = false): DisplayPrimitive =>
    ({
      kind: 'rect',
      x: pos,
      y: 10,
      w: 10,
      h: 10,
      inlineSdtWidget: { kind: 'checkbox', groupId: `sdt@${pos}`, pos, controlId, checked },
    }) as DisplayPrimitive;
  const page = (...primitives: DisplayPrimitive[]): DisplayPage => ({
    pageIndex: 0,
    width: 100,
    height: 100,
    primitives,
  });
  const focusedControl = () => (document.activeElement as HTMLElement | null)?.dataset.sdtControlId;
  const { container, rerender } = render(
    <CanvasInteractiveOverlay page={page(checkbox(101, 10), checkbox(102, 40))} />
  );
  await act(async () => {});
  container.querySelector<HTMLElement>('[data-sdt-control-id="102"]')!.focus();
  // Toggled.
  rerender(<CanvasInteractiveOverlay page={page(checkbox(101, 10), checkbox(102, 40, true))} />);
  await act(async () => {});
  expect(focusedControl()).toBe('102');
  // Moved by content inserted before both: its old group is now the other's.
  rerender(<CanvasInteractiveOverlay page={page(checkbox(101, 40), checkbox(102, 70, true))} />);
  await act(async () => {});
  expect(focusedControl()).toBe('102');
  // The control before it removed.
  rerender(<CanvasInteractiveOverlay page={page(checkbox(102, 40, true))} />);
  await act(async () => {});
  expect(focusedControl()).toBe('102');
});

test('a link to a note outside the window finds its target in the text fallback', async () => {
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
  expect(document.getElementById('oox-endnote-7')).not.toBeNull();
  expect(mirroredPages(host)).not.toContain(30);
  const link = host.querySelector<HTMLAnchorElement>('a[href="#oox-endnote-7"]')!;
  await act(async () => {
    link.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
  });
  expect(document.getElementById('oox-endnote-7')).not.toBeNull();
  expect(mirroredPages(host)).not.toContain(30);
});

test('every page exposes text outside the window and range queries build full chrome on demand', async () => {
  const hostRef = createRef<HTMLDivElement>();
  const pages = blankPages(40, (index) => [{
    kind: 'text',
    x: 10,
    baselineY: 20,
    width: 30,
    font: '11px sans-serif',
    color: '#000',
    text: `Page ${index}`,
    blockKey: `p${index}`,
    docStart: index * 10,
    docEnd: index * 10 + 7,
  }]);
  render(<CanvasPagesView displayList={{ pages }} hostRef={hostRef} glyphOutlineProvider={() => ''} />);
  await idle();
  const host = hostRef.current!;
  const far = host.querySelector<HTMLElement>('.canvas-page[data-page-index="30"]')!;
  expect(far.querySelector('.canvas-page-mirror')?.textContent).toBe('Page 30');
  expect(mirroredPages(host)).not.toContain(30);
  expect(far.querySelector('.layout-run-text, [data-doc-start], [data-para-id]')).toBeNull();
  const queries = {
    rangeRects: () => [{ pageIndex: 30, x: 0, y: 0, width: 1, height: 1 }],
  } as unknown as DisplayListQueries;
  const context = createRenderedDomContext(host, 1, {
    displayListQueries: queries,
    projector: { projectRect: () => null, getPageBounds: () => null },
  });
  let elements: Element[] = [];
  await act(async () => {
    elements = context.findElementsForRange(300, 307);
  });
  expect(elements).toHaveLength(1);
  expect(elements[0]).toBe(far.querySelector('span.layout-run-text[data-doc-start="300"]'));
  expect(elements[0]!.textContent).toBe('Page 30');
  expect(elements[0]!.isConnected).toBe(true);
  expect(mirroredPages(host)).toContain(30);
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
      { ...widget('shifted'), docStart: 1, docEnd: 2, blockId: 'p' } as unknown as DisplayPrimitive,
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
  const control = () => host.querySelector<HTMLElement>('button[data-sdt-group-id="shifted"]');
  const controlPos = () => control()?.dataset.sdtPos;
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
  await act(async () => control()!.focus());
  const second = applyFrameDeltaOwned(first, shift);
  expect(second.displayList.pages[0]).toBe(page);
  rerender(view(second));
  await act(async () => {});
  expect(controlPos()).toBe('6');
  // The rebuilt control keeps focus.
  expect(document.activeElement).toBe(control());
  await idle();
  expect(mirrorStart()).toBe('6');
});

test('elements a plugin query returned stay connected through other queries', async () => {
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
  const queries = {
    rangeRects: (from: number) => [
      { pageIndex: Math.floor(from / 10), x: 0, y: 0, width: 1, height: 1 },
    ],
  } as unknown as DisplayListQueries;
  const context = createRenderedDomContext(host, 1, {
    displayListQueries: queries,
    projector: { projectRect: () => null, getPageBounds: () => null },
  });
  let first: Element[] = [];
  await act(async () => {
    first = context.findElementsForRange(300, 303);
  });
  for (let page = 31; page < 39; page += 1) {
    await act(async () => {
      context.findElementsForRange(page * 10, page * 10 + 3);
    });
  }
  await idle();
  expect(first).toHaveLength(1);
  expect(first[0]!.isConnected).toBe(true);
});

test('a query across many far pages keeps their chrome only for the most recent', async () => {
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
  const pages = blankPages(60, (index) => [run(index * 10)]);
  render(<CanvasPagesView displayList={{ pages }} hostRef={hostRef} glyphOutlineProvider={() => ''} />);
  await idle();
  const host = hostRef.current!;
  const before = mirroredPages(host);
  const queries = {
    rangeRects: (from: number, to: number) =>
      Array.from({ length: Math.ceil((to - from) / 10) }, (_, i) => ({
        pageIndex: Math.floor(from / 10) + i,
        x: 0,
        y: 0,
        width: 1,
        height: 1,
      })),
  } as unknown as DisplayListQueries;
  const context = createRenderedDomContext(host, 1, {
    displayListQueries: queries,
    projector: { projectRect: () => null, getPageBounds: () => null },
  });
  let elements: Element[] = [];
  await act(async () => {
    elements = context.findElementsForRange(300, 600);
    expect(elements).toHaveLength(30);
    expect(elements.every((element) => element.isConnected)).toBe(true);
  });
  await idle();
  const far = mirroredPages(host).filter((index) => !before.includes(index));
  expect(far).toEqual(Array.from({ length: 12 }, (_, i) => 48 + i));
});

test('chrome built on demand for an inactive page is cleared on release', async () => {
  const page: DisplayPage = { pageIndex: 0, width: 100, height: 100, primitives: [] };
  let handle: PageChromeHandle | null = null;
  const register = (next: PageChromeHandle | null) => {
    handle = next;
  };
  const { container, rerender } = render(
    <CanvasPageMirror page={page} active={false} register={register} />
  );
  await idle();
  const content = () => container.firstElementChild!.firstElementChild as HTMLElement;
  expect(content().style.contentVisibility).toBe('auto');
  act(() => handle!.build());
  expect(content().style.contentVisibility).not.toBe('auto');
  act(() => handle!.release());
  expect(content().style.contentVisibility).toBe('auto');

  rerender(<CanvasPageMirror page={page} active register={register} />);
  await act(async () => {});
  act(() => handle!.release());
  expect(content().style.contentVisibility).not.toBe('auto');
});

test('first fallbacks share a FIFO and honor idle deadlines and timeouts', async () => {
  const pages = blankPages(4, (index) => [{
    kind: 'text',
    x: 10,
    baselineY: 20,
    width: 30,
    font: '11px sans-serif',
    color: '#000',
    text: `Page ${index}`,
  }]);
  const { container } = render(
    <>{pages.map((page) => (
      <CanvasPageMirror key={page.pageIndex} page={page} active={false} />
    ))}</>
  );
  expect(idleWork.size).toBe(1);
  expect(Array.from(idleWork.values())[0]!.timeout).toBe(5000);
  const text = () =>
    Array.from(container.querySelectorAll('.canvas-page-mirror'), (host) => host.textContent);
  await act(async () => flushIdle({ didTimeout: false, timeRemaining: () => 1 }));
  expect(text()).toEqual(['Page 0', '', '', '']);
  expect(idleWork.size).toBe(1);
  await act(async () => flushIdle({ didTimeout: true, timeRemaining: () => 50 }));
  expect(text()).toEqual(['Page 0', 'Page 1', '', '']);
  expect(idleWork.size).toBe(1);
  await idle();
  expect(text()).toEqual(['Page 0', 'Page 1', 'Page 2', 'Page 3']);
  expect(idleWork.size).toBe(0);
});

test('unmounting an inactive page cancels its queued first fallback', async () => {
  const pages = blankPages(2);
  const view = (showFirst: boolean) => (
    <>
      {showFirst && <CanvasPageMirror key="first" page={pages[0]!} active={false} />}
      <CanvasPageMirror key="second" page={pages[1]!} active={false} />
    </>
  );
  const { container, rerender, unmount } = render(view(true));
  expect(idleWork.size).toBe(1);
  rerender(view(false));
  await act(async () => flushIdle({ didTimeout: true, timeRemaining: () => 0 }));
  expect(container.querySelectorAll('.layout-page-mirror')).toHaveLength(1);
  expect(idleWork.size).toBe(0);
  unmount();
  render(<CanvasPageMirror page={pages[0]!} active={false} />).unmount();
  expect(idleWork.size).toBe(0);
});

test('deferred full chrome keeps its own idle schedule alongside queued fallbacks', async () => {
  const pages = blankPages(2);
  const { container } = render(
    <>
      <CanvasPageMirror page={pages[0]!} active={false} />
      <CanvasPageMirror page={pages[1]!} defer />
    </>
  );
  expect(Array.from(idleWork.values(), ({ timeout }) => timeout)).toEqual([5000, 1500]);
  await idle();
  const mirrors = container.querySelectorAll<HTMLElement>('.layout-page-mirror');
  expect(mirrors).toHaveLength(2);
  expect(mirrors[0]!.style.contentVisibility).toBe('auto');
  expect(mirrors[1]!.style.contentVisibility).not.toBe('auto');
});
