import { expect, mock, test } from 'bun:test';
import { createWorkerPrintCapture } from './printPages';
import type { DocxCommandFrame } from '@betteroffice/docx/yrs';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';

function fixture() {
  const snapshot = { version: 3, geometry: { frameEpoch: 7, pages: [0, 1, 2].map((pageIndex) => ({ page: { pageIndex, pageId: String(pageIndex + 1) } })) } };
  const binding = { session: { flush: async () => {} }, admission: {}, viewStore: { getSnapshot: () => snapshot } } as unknown as PeerlessEditSessionBinding;
  return { snapshot, binding };
}

test('print captures worker pixels for every page including pages outside the viewport', async () => {
  const { binding } = fixture();
  const render = mock((bitmap: ImageBitmap) => `png:${bitmap.width}`);
  const capture = createWorkerPrintCapture(binding, render, () => true);
  const build = mock(async (pages: number[]) => {
    capture.accept({ payload: { stagedPresentation: { pages: pages.map((pageIndex) => ({ page: { pageIndex, pageId: String(pageIndex + 1) }, bitmap: { width: pageIndex + 1 } })) } } } as unknown as DocxCommandFrame);
  });
  const complete = mock(async () => {});
  expect(await capture.collect(build, complete)).toEqual(['png:1', 'png:2', 'png:3']);
  expect(build.mock.calls.map(([pages]) => pages)).toEqual([[0], [1], [2]]);
  expect(complete).not.toHaveBeenCalled();
});

test('a document edit while printing rejects the mixed-version job', async () => {
  const { binding, snapshot } = fixture();
  const capture = createWorkerPrintCapture(binding, () => 'png', () => true);
  await expect(capture.collect(async () => { snapshot.version += 1; }, async () => {})).rejects.toThrow('changed');
});

test('print completes a provisional prefix first and prints the whole document', async () => {
  const { binding, snapshot } = fixture();
  Object.assign(snapshot.geometry, { layoutProvisional: true });
  const capture = createWorkerPrintCapture(binding, (bitmap) => `png:${bitmap.width}`, () => true);
  let completing = false;
  const complete = mock(async (frameEpoch: number) => {
    expect(frameEpoch).toBe(7);
    completing = true;
  });
  const build = mock(async (pages: number[]) => {
    if (completing) {
      completing = false;
      Object.assign(snapshot.geometry, { layoutProvisional: false, frameEpoch: 8,
        pages: [0, 1, 2, 3, 4].map((pageIndex) => ({ page: { pageIndex, pageId: String(pageIndex + 1) } })) });
    }
    capture.accept({ payload: { stagedPresentation: { pages: pages.map((pageIndex) => ({
      page: { pageIndex, pageId: String(pageIndex + 1) }, bitmap: { width: pageIndex + 1 },
    })) } } } as unknown as DocxCommandFrame);
  });
  expect(await capture.collect(build, complete)).toEqual(['png:1', 'png:2', 'png:3', 'png:4', 'png:5']);
  expect(complete).toHaveBeenCalledTimes(1);
});
