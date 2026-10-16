import { expect, test } from 'bun:test';
import type { RetainedFrame } from '../layout/render/frameDelta';
import type { DocxCommandEchoContext, DocxCommandFrame } from './residentCommandProtocol';
import { discardCommandPresentation, stageCommandPresentation } from './residentCommandPresentation';
import { buildCommandFrame } from './residentCommandFrame';

const page = { pageId: '1', pageIndex: 0 };
const rect = { x: 0, y: 0, width: 10, height: 10 };
const layout = { frame: { frameEpoch: 3 } as RetainedFrame, layoutRevision: 1 };
const echoContext: DocxCommandEchoContext = {
  page, frameEpoch: 3, layoutRevision: 1, serializedBytes: 900, text: 'a',
  styleRuns: [{ startUtf16: 0, endUtf16: 1, fontFamily: 'Arial', fontSizePx: 8,
    color: '#000', bold: false, italic: false, underline: false, strike: false,
    transform: [1, 0, 0, 1, 0, 0], baseline: { x: 0, y: 8 } }],
  boundaries: [0, 1].map((utf16) => ({ utf16, grapheme: utf16,
    caret: { x: utf16 * 5, y: 0, width: 1, height: 10 } })),
  selection: { anchorGrapheme: 1, focusGrapheme: 1 }, transform: [1, 0, 0, 1, 0, 0],
  damage: [{ page, rect, backgroundId: 'background', omittedPrimitiveIds: ['shade', 'media'] }],
  shaping: 'simple-ltr',
};

test('d2c_covering_frame_preserves_complete_echo_context', async () => {
  const bitmap = { close() {} } as ImageBitmap;
  const staged = await stageCommandPresentation({ layout, pages: [page], pixelRatio: 1, echoContext,
    async rasterizePage() { return bitmap; },
    async rasterizeBackground() { return bitmap; },
  });
  const packed = buildCommandFrame({ frame: new Uint8Array([1]),
    metadata: { full: true, frameEpoch: 3, baseFrameEpoch: 0 }, coversThroughSeq: 1,
    caret: { frameEpoch: 3, caretRect: null }, layoutRevision: 1,
    echoContext, stagedPresentation: staged.value });
  expect(packed.value.payload.echoContext).toEqual(echoContext);
  expect(packed.value.payload.echoContext?.frameEpoch).toBe(packed.value.revision);
  expect(packed.value.payload.echoContext?.layoutRevision).toBe(packed.value.payload.layoutRevision);
  expect(packed.value.payload.echoContext?.damage[0]?.backgroundId)
    .toBe(packed.value.payload.stagedPresentation?.pages[0]?.backgrounds?.[0]?.id);
});

test('d2c_staging_transfers_authoritative_and_text_free_bitmaps', async () => {
  const calls: string[] = [];
  const bitmap = { close() {} } as ImageBitmap;
  const background = { close() {} } as ImageBitmap;
  const staged = await stageCommandPresentation({ layout, pages: [page], pixelRatio: 2, echoContext,
    async rasterizePage() { calls.push('page'); return bitmap; },
    async rasterizeBackground(dependency) { calls.push(dependency.omittedPrimitiveIds.join(',')); return background; },
  });
  expect(calls).toEqual(['page', 'shade,media']);
  expect(staged.value).toEqual({ frameEpoch: 3, pages: [{ page, frameEpoch: 3,
    bitmap, backgrounds: [{ id: 'background', rect, bitmap: background }] }] });
  expect(staged.transfer).toEqual([bitmap, background]);
  const packed = buildCommandFrame({ frame: new Uint8Array([1]),
    metadata: { full: true, frameEpoch: 3, baseFrameEpoch: 0 }, coversThroughSeq: 2,
    caret: { frameEpoch: 3, caretRect: null }, layoutRevision: 1, stagedPresentation: staged.value });
  expect(packed.transfer).toContain(bitmap);
  expect(packed.transfer).toContain(background);
});

test('d2c_failed_staging_closes_all_untransferred_bitmaps', async () => {
  let closed = 0;
  await expect(stageCommandPresentation({ layout, pages: [page], pixelRatio: 1, echoContext,
    async rasterizePage() { return { close() { closed += 1; } } as ImageBitmap; },
    async rasterizeBackground() { throw new Error('raster failed'); },
  })).rejects.toThrow('raster failed');
  expect(closed).toBe(1);
});

test('d2c_discard_closes_aliased_page_and_background_bitmaps_once', () => {
  const closed: string[] = [];
  const bitmap = { close() { closed.push('page'); } } as ImageBitmap;
  const background = { close() { closed.push('background'); } } as ImageBitmap;
  const frame: DocxCommandFrame = { revision: 3, coversThroughSeq: 0, payload: {
    frame: new ArrayBuffer(0), layoutRevision: 1, caret: { frameEpoch: 3, caretRect: null },
    stagedPresentation: { frameEpoch: 3, pages: [{ page, frameEpoch: 3, bitmap,
      backgrounds: [{ id: 'page', rect, bitmap }, { id: 'background', rect, bitmap: background }] },
      { page, frameEpoch: 3, bitmap: background }] },
  } };
  discardCommandPresentation(frame);
  discardCommandPresentation({ ...frame });
  expect(closed).toEqual(['page', 'background']);
});
