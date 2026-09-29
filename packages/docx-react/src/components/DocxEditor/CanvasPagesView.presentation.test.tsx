import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { createRef } from 'react';
import type { DisplayList, RetainedFrame } from '@betteroffice/docx/layout/render';
import { CanvasPagesView } from './CanvasPagesView';
import { isPresented } from './internals/layoutProvenance';

const { act, cleanup, render } = await import('@testing-library/react');

afterEach(() => {
  cleanup();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function frame(epoch: number): RetainedFrame {
  const page = { pageIndex: 0, width: 100, height: 100, primitives: [] };
  return {
    protocolVersion: 1,
    docEpoch: 1,
    layoutEpoch: epoch,
    frameEpoch: epoch,
    pages: [{ pageId: 1n, pageIndex: 0, fingerprint: 0n, primitiveIds: new BigUint64Array(), page }],
    damagedPageIds: new Set([1n]),
    removedPageIds: new Set(),
    displayList: { pages: [page] },
  };
}

function workerReplay() {
  const attaches: Array<{ zoom: number; resolve: (attached: boolean) => void }> = [];
  return {
    attaches,
    offscreenReplay: {
      attach: (_pages: unknown, _ids: string[], _dpr: number, zoom: number) =>
        new Promise<boolean>((resolve) => attaches.push({ zoom, resolve })),
    },
  };
}

test('worker pages count as presented only once their attach has painted them', async () => {
  const { attaches, offscreenReplay } = workerReplay();
  const hostRef = createRef<HTMLDivElement>();
  const first = frame(1);
  const view = (current: RetainedFrame, displayList: DisplayList, zoom: number) => (
    <CanvasPagesView
      displayList={displayList}
      frame={current}
      hostRef={hostRef}
      zoom={zoom}
      glyphOutlineProvider={() => ''}
      offscreenReplay={offscreenReplay}
    />
  );
  const shows = (displayList: DisplayList) => isPresented(hostRef.current, displayList);
  const { rerender } = render(view(first, first.displayList, 1));
  await act(async () => {});
  expect(attaches.map(({ zoom }) => zoom)).toEqual([1]);
  expect(shows(first.displayList)).toBe(false);

  await act(async () => attaches[0].resolve(true));
  expect(shows(first.displayList)).toBe(true);

  rerender(view(first, first.displayList, 2));
  await act(async () => {});
  expect(attaches.map(({ zoom }) => zoom)).toEqual([1, 2]);
  expect(shows(first.displayList)).toBe(false);

  rerender(view(first, first.displayList, 3));
  await act(async () => {});
  await act(async () => attaches[1].resolve(true));
  expect(shows(first.displayList)).toBe(false);
  await act(async () => attaches[2].resolve(true));
  expect(shows(first.displayList)).toBe(true);

  const second = frame(2);
  rerender(view(first, first.displayList, 4));
  await act(async () => {});
  rerender(view(second, second.displayList, 4));
  await act(async () => {});
  expect(attaches).toHaveLength(4);
  expect(shows(second.displayList)).toBe(false);
  await act(async () => attaches[3].resolve(true));
  expect(shows(second.displayList)).toBe(true);

  const third = frame(3);
  rerender(view(third, third.displayList, 4));
  await act(async () => {});
  expect(attaches).toHaveLength(4);
  expect(shows(third.displayList)).toBe(true);
});

test('main-thread pages stop counting as presented at a new zoom until they repaint', async () => {
  const hostRef = createRef<HTMLDivElement>();
  const shown = frame(1);
  const view = (zoom: number) => (
    <CanvasPagesView
      displayList={shown.displayList}
      frame={shown}
      hostRef={hostRef}
      zoom={zoom}
      glyphOutlineProvider={() => ''}
    />
  );
  const shows = () => isPresented(hostRef.current, shown.displayList);
  const { rerender } = render(view(1));
  await act(async () => {});
  expect(shows()).toBe(true);
  rerender(view(2));
  expect(shows()).toBe(false);
  await act(async () => {});
  expect(shows()).toBe(true);
});
