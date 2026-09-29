import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { createRef } from 'react';
import type { DisplayList, DisplayPage, RetainedFrame } from '@betteroffice/docx/layout/render';
import { CanvasInteractiveOverlay } from './CanvasInteractiveOverlay';
import { CanvasPageMirror } from './CanvasPageMirror';
import { CanvasPagesView } from './CanvasPagesView';

const { act, cleanup, render } = await import('@testing-library/react');

afterEach(() => {
  cleanup();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

test('only pages in the window and the page holding focus carry a mirror', async () => {
  const hostRef = createRef<HTMLDivElement>();
  const displayList: DisplayList = {
    pages: Array.from({ length: 40 }, (_, pageIndex) => ({
      pageIndex,
      width: 100,
      height: 100,
      primitives: [],
    })),
  };
  render(
    <CanvasPagesView displayList={displayList} hostRef={hostRef} glyphOutlineProvider={() => ''} />
  );
  await act(async () => {});
  const host = hostRef.current!;
  const withMirror = () =>
    Array.from(host.querySelectorAll<HTMLElement>('.canvas-page'))
      .filter((page) => page.querySelector('.canvas-page-mirror'))
      .map((page) => Number(page.dataset.pageIndex));

  const mirrored = withMirror();
  expect(mirrored[0]).toBe(0);
  expect(mirrored.length).toBeLessThan(displayList.pages.length);
  expect(host.querySelectorAll('.canvas-page canvas')).toHaveLength(displayList.pages.length);

  const far = host.querySelector<HTMLElement>('.canvas-page[data-page-index="35"] canvas')!;
  far.tabIndex = 0;
  await act(async () => far.focus());
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
  await act(async () => {});
  const host = hostRef.current!;
  const mirroredKeys = () =>
    Array.from(host.querySelectorAll<HTMLElement>('.canvas-page'))
      .filter((page) => page.querySelector('.canvas-page-mirror'))
      .map((page) => page.dataset.pageKey);
  const pinned = host.querySelector<HTMLElement>('.canvas-page[data-page-key="36"] canvas')!;
  pinned.tabIndex = 0;
  await act(async () => pinned.focus());
  expect(mirroredKeys()).toContain('36');

  rerender(view(framed([99n, ...ids])));
  await act(async () => {});
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
