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
  expect(await capture.collect(build)).toEqual(['png:1', 'png:2', 'png:3']);
  expect(build.mock.calls.map(([pages]) => pages)).toEqual([[0], [1], [2]]);
});

test('a document edit while printing rejects the mixed-version job', async () => {
  const { binding, snapshot } = fixture();
  const capture = createWorkerPrintCapture(binding, () => 'png', () => true);
  await expect(capture.collect(async () => { snapshot.version += 1; })).rejects.toThrow('changed');
});
