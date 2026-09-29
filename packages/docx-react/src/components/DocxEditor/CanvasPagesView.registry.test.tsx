import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { createRef } from 'react';
import {
  bindDisplayPageRegistry,
  displayPageCanvas,
  displayPageCanvases,
  DisplayPageRegistry,
  type RetainedFrame,
} from '@betteroffice/docx/layout/render';
import { CanvasPagesView } from './CanvasPagesView';

const { act, cleanup, render } = await import('@testing-library/react');

afterEach(() => {
  cleanup();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
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

/** Registry lookups answer as a subtree search does, without searching. */
function expectSearchAnswers(host: HTMLElement, pages: number) {
  const byIndex = Array.from({ length: pages + 1 }, (_, index) =>
    host.querySelector<HTMLCanvasElement>(`canvas[data-page-index="${index}"]`)
  );
  const all = Array.from(host.querySelectorAll<HTMLCanvasElement>('canvas[data-page-index]'));
  const search = { querySelector: host.querySelector, querySelectorAll: host.querySelectorAll };
  host.querySelector = () => {
    throw new Error('searched the host');
  };
  host.querySelectorAll = host.querySelector as typeof host.querySelectorAll;
  try {
    byIndex.forEach((canvas, index) => expect(displayPageCanvas(host, index)).toBe(canvas));
    const listed = displayPageCanvases(host);
    expect(listed).toHaveLength(all.length);
    listed.forEach((canvas, index) => expect(canvas).toBe(all[index]!));
  } finally {
    Object.assign(host, search);
  }
}

test('page lookups read the mounted canvases as a search of the host would', async () => {
  const hostRef = createRef<HTMLDivElement>();
  const view = (frame: RetainedFrame) => (
    <CanvasPagesView
      displayList={frame.displayList}
      frame={frame}
      hostRef={hostRef}
      glyphOutlineProvider={() => ''}
    />
  );
  const { rerender } = render(view(framed([1n, 2n])));
  await act(async () => {});
  expectSearchAnswers(hostRef.current!, 2);

  rerender(view(framed([3n, 1n, 2n])));
  await act(async () => {});
  expectSearchAnswers(hostRef.current!, 3);

  rerender(view(framed([2n])));
  await act(async () => {});
  expectSearchAnswers(hostRef.current!, 1);
});

test('page lookups follow pages the view reorders without rerendering them', async () => {
  const hostRef = createRef<HTMLDivElement>();
  const pages = [0, 1, 2].map((pageIndex) => ({
    pageIndex,
    width: 100,
    height: 100,
    primitives: [],
  }));
  const view = (order: typeof pages) => (
    <CanvasPagesView
      displayList={{ pages: order }}
      hostRef={hostRef}
      glyphOutlineProvider={() => ''}
    />
  );
  const { rerender } = render(view(pages));
  await act(async () => {});
  expectSearchAnswers(hostRef.current!, 3);

  rerender(view([...pages].reverse()));
  await act(async () => {});
  expectSearchAnswers(hostRef.current!, 3);
});

test('a registry answers for renumbered canvases and prefers the first in document order', () => {
  const host = document.createElement('div');
  const canvases = [0, 1, 1].map((pageIndex) => {
    const canvas = document.createElement('canvas');
    canvas.dataset.pageIndex = String(pageIndex);
    host.append(canvas);
    return canvas;
  });
  const registry = new DisplayPageRegistry();
  for (const canvas of [...canvases].reverse()) registry.add(canvas);
  bindDisplayPageRegistry(host, registry);
  expect(displayPageCanvas(host, 1)).toBe(canvases[1]!);
  expect(displayPageCanvases(host)).toEqual(canvases);

  canvases[1]!.dataset.pageIndex = '2';
  expect(displayPageCanvas(host, 1)).toBe(canvases[2]!);
  expect(displayPageCanvas(host, 2)).toBe(canvases[1]!);

  bindDisplayPageRegistry(host, null);
  expect(displayPageCanvas(host, 0)).toBe(canvases[0]!);
});
