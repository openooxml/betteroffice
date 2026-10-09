import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { createRef } from 'react';
import {
  applyFrameDeltaOwned,
  FRAME_DELTA_VERSION,
  type DecodedFrameDelta,
  type DisplayPage,
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

test('the mirror shows note anchors an owned shift moved in place', async () => {
  const page = {
    pageIndex: 0,
    width: 100,
    height: 100,
    primitives: [],
    noteAreas: [
      {
        kind: 'footnote',
        y: 80,
        height: 20,
        noteIds: [1],
        notes: [{ id: 1, anchorDocStart: 3, anchorDocEnd: 4 }],
        primitives: [
          {
            kind: 'text',
            text: 'note',
            x: 10,
            baselineY: 90,
            width: 20,
            font: '400 10px Calibri',
            color: '#000000',
            groupId: 'footnote-1',
          },
        ],
      },
    ],
  } as unknown as DisplayPage;
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
      glyphOutlineProvider={() => ''}
    />
  );
  const anchor = () =>
    hostRef.current?.querySelector<HTMLElement>('aside[data-note-id="1"]')?.dataset.anchorDocStart;
  const { rerender } = render(view(first));
  await act(async () => {});
  expect(anchor()).toBe('3');

  const shifted: DecodedFrameDelta = {
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
        runs: [],
        anchors: [{ area: 0, note: 0, start: 13, end: 14 }],
      },
    ],
    bytes: new Uint8Array(),
  };
  const second = applyFrameDeltaOwned(first, shifted);
  expect(second.displayList.pages[0]).toBe(page);
  rerender(view(second));
  await act(async () => {});
  expect(anchor()).toBe('13');
});
