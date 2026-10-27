import { expect, test } from 'bun:test';
import type { EchoEntry } from '../../../../../../shared/office-session/commands/types';
import { createPresenter } from '../../../../../../shared/office-session/commands/presenter';
import type { DocxCommandFrame } from '@betteroffice/docx/worker';
import { createEchoPainter, type PeerlessEchoIntent } from './echoPainter';
import { createFramePresentation } from './framePresentation';
import { createPeerlessViewStore, type PeerlessFrameGeometry } from './viewStore';
import { registerBrowserHarness } from './browserHarness';
import { glyphChar, installGlyphPaths, testEchoGlyphLine, testEchoGlyphSet } from './echoGlyphs.testing';

const identity = { session: 'harness', incarnation: 1, epoch: 1 };
const page = { pageId: '1', pageIndex: 0 };

function entry(seq: number, text: string): EchoEntry<PeerlessEchoIntent> {
  return { seq, ticket: { id: seq, seq, ack: Promise.resolve({ seq, status: 'applied' }),
    presented: Promise.resolve('presented') }, value: { originalIntentCount: 1,
    contextFrameEpoch: 1, intent: { kind: 'insertText', payload: { text } } } };
}

function frame(revision: number, coverage: number, text: string): DocxCommandFrame {
  const bytes = new ArrayBuffer(88);
  const view = new DataView(bytes);
  new Uint8Array(bytes).set([0x46, 0x44, 0x56, 0x31]);
  view.setUint16(4, 1, true);
  view.setUint16(6, 80, true);
  view.setUint32(8, 88, true);
  view.setUint32(12, 1, true);
  view.setBigUint64(16, 1n, true);
  view.setBigUint64(24, 1n, true);
  view.setBigUint64(32, BigInt(revision), true);
  for (const offset of [56, 60]) view.setUint32(offset, 80, true);
  view.setUint32(64, 4, true);
  view.setUint32(68, 88, true);
  const rect = { x: 0, y: 0, width: 100, height: 20 };
  const bitmap = (pixels: string) => ({ width: 100, height: 20, pixels, close() {} }) as unknown as ImageBitmap;
  return { revision, coversThroughSeq: coverage, payload: {
    frame: bytes, layoutRevision: revision, caret: { frameEpoch: revision, caretRect: null },
    stagedPresentation: { frameEpoch: revision, pages: [{ page, frameEpoch: revision,
      bitmap: bitmap(text), backgrounds: [{ id: 'background', rect, bitmap: bitmap('') }] }] },
    echoContext: { page, frameEpoch: revision, layoutRevision: revision, serializedBytes: 900,
      text, shaping: 'simple-ltr', transform: [1, 0, 0, 1, 0, 0],
      styleRuns: [{ startUtf16: 0, endUtf16: text.length, fontFamily: 'Arial', fontSizePx: 12,
        color: '#000', bold: false, italic: false, underline: false, strike: false,
        transform: [1, 0, 0, 1, 0, 0], baseline: { x: 0, y: 14 } }],
      boundaries: Array.from({ length: text.length + 1 }, (_, utf16) => ({ utf16, grapheme: utf16,
        caret: { x: utf16 * 10, y: 0, width: 1, height: 20 } })),
      selection: { anchorGrapheme: text.length, focusGrapheme: text.length },
      damage: [{ page, rect, backgroundId: 'background', omittedPrimitiveIds: [] }],
      glyphs: testEchoGlyphLine(),
    },
    echoGlyphSets: [testEchoGlyphSet],
  } };
}

test('checkpoints_capture_one_synchronous_handoff', async () => {
  const checkpoints: { stage: string; revision: number; retired: number[]; remaining: number[]; pixels: string }[] = [];
  const scope = globalThis as typeof globalThis & {
    __DOCX_PEERLESS_HARNESS__?: boolean;
    __docxPeerlessHarness?: { presentationProbe(stage: string, frame: DocxCommandFrame,
      retired: readonly EchoEntry<PeerlessEchoIntent>[], remaining: readonly EchoEntry<PeerlessEchoIntent>[],
      canvases: readonly HTMLCanvasElement[]): void };
  };
  const previousFlag = scope.__DOCX_PEERLESS_HARNESS__;
  const previousHarness = scope.__docxPeerlessHarness;
  let pixels = '';
  const canvas = { width: 100, height: 20, toDataURL: () => pixels, getContext: () => ({
    save() {}, restore() {}, setTransform() {}, clearRect() { pixels = ''; },
    getImageData: () => ({ width: 100, height: 20 }) as ImageData,
    putImageData() {}, transform() {}, beginPath() {}, rect() {}, clip() {}, fillRect() {},
    translate() {}, scale() {}, fill(path: unknown) { pixels += glyphChar(path); },
    drawImage(bitmap: ImageBitmap) { pixels = (bitmap as ImageBitmap & { pixels: string }).pixels; },
  }) } as unknown as HTMLCanvasElement;
  const restorePaths = installGlyphPaths();
  const store = createPeerlessViewStore({ identity, version: 0, throughSeq: 0, presentedThroughSeq: -1,
    workerSelection: null, anchors: null, editState: undefined, geometry: null });
  const painter = createEchoPainter({ canvasForPage: () => canvas });
  const presentation = createFramePresentation({ viewStore: store, echoPainter: painter,
    canvasForPage: () => canvas, requestFullFrame() { throw new Error('Unexpected full frame request'); },
    prepareGeometry(candidate): PeerlessFrameGeometry { return { identity,
      coversThroughSeq: candidate.coversThroughSeq, frameEpoch: candidate.revision,
      layoutRevision: candidate.payload.layoutRevision,
      pages: [{ page, bounds: { x: 0, y: 0, width: 100, height: 20 }, transform: [1, 0, 0, 1, 0, 0] }],
      caret: null, selectionRects: [] }; },
  });
  const presenter = createPresenter({ echo: presentation.hooks, transactionalPresentation: true,
    sameCoverageFrames: true, requestFullFrame() {}, now: () => 0 });
  try {
    scope.__DOCX_PEERLESS_HARNESS__ = true;
    expect(presenter.present(frame(1, 0, ''))).toBe('presented');
    const a = entry(1, 'A');
    const b = entry(2, 'B');
    presenter.add(a);
    presenter.add(b);
    expect(pixels).toBe('AB');
    scope.__docxPeerlessHarness = {
      presentationProbe(stage, candidate, retired, remaining, canvases) {
        checkpoints.push({ stage, revision: candidate.revision, retired: retired.map((item) => item.seq!),
          remaining: remaining.map((item) => item.seq!), pixels: canvases[0]!.toDataURL() });
      },
    };
    expect(presenter.present(frame(2, 1, 'A'))).toBe('presented');
    expect(checkpoints).toEqual([
      { stage: 'installed', revision: 2, retired: [1], remaining: [2], pixels: 'A' },
      { stage: 'reconciled', revision: 2, retired: [1], remaining: [2], pixels: 'AB' },
    ]);
    await Promise.resolve();
    expect(pixels).toBe('AB');
    expect(checkpoints).toHaveLength(2);
    expect(presenter.present(frame(3, 2, 'AB'))).toBe('presented');
    expect(checkpoints.slice(2)).toEqual([
      { stage: 'installed', revision: 3, retired: [2], remaining: [], pixels: 'AB' },
      { stage: 'reconciled', revision: 3, retired: [2], remaining: [], pixels: 'AB' },
    ]);
    scope.__DOCX_PEERLESS_HARNESS__ = false;
    expect(presenter.present(frame(4, 2, 'AB'))).toBe('presented');
    expect(checkpoints).toHaveLength(4);
  } finally {
    restorePaths();
    presenter.dispose();
    presentation.dispose();
    painter.dispose();
    store.dispose();
    if (previousFlag === undefined) delete scope.__DOCX_PEERLESS_HARNESS__;
    else scope.__DOCX_PEERLESS_HARNESS__ = previousFlag;
    if (previousHarness === undefined) delete scope.__docxPeerlessHarness;
    else scope.__docxPeerlessHarness = previousHarness;
  }
});

test('viewport_driver_requests_evicted_page_again', async () => {
  const scope = globalThis as typeof globalThis & {
    __DOCX_PEERLESS_HARNESS__?: boolean;
    __docxPeerlessHarness?: {
      capabilities: { viewport?: boolean };
      registerWorker(worker: Worker): void;
      viewportProbe?(geometry: PeerlessFrameGeometry): void;
      setViewport?(pages: number[]): Promise<void>;
    };
  };
  const previousFlag = scope.__DOCX_PEERLESS_HARNESS__;
  const previousHarness = scope.__docxPeerlessHarness;
  const requests: { id: number; type: string; pages: number[]; expectedFrameEpoch: number }[] = [];
  const worker = new EventTarget();
  const geometry = (revision: number): PeerlessFrameGeometry => ({ identity, coversThroughSeq: 0,
    frameEpoch: revision, layoutRevision: 1, caret: null, selectionRects: [], pages: [0, 39].map((pageIndex) => ({
      page: { pageId: String(pageIndex + 1), pageIndex }, bounds: { x: 0, y: 0, width: 100, height: 20 },
      transform: [1, 0, 0, 1, 0, 0] as const,
    })) });
  Object.assign(worker, { postMessage(request: typeof requests[number]) {
    requests.push(request);
    queueMicrotask(() => worker.dispatchEvent(new MessageEvent('message', { data: { id: request.id, ok: true } })));
  } });
  let unregister: (() => void) | undefined;
  try {
    scope.__DOCX_PEERLESS_HARNESS__ = true;
    unregister = registerBrowserHarness({ setMode() {} });
    const harness = scope.__docxPeerlessHarness!;
    harness.registerWorker(worker as unknown as Worker);
    const initial = geometry(1);
    harness.viewportProbe?.(initial);
    expect(harness.capabilities.viewport).toBe(true);
    await harness.setViewport!([39]);
    const away = geometry(2);
    harness.viewportProbe!(away);
    expect(away.pages.map(({ page }) => page.pageIndex)).toEqual([39]);
    await harness.setViewport!([0]);
    const returned = geometry(3);
    harness.viewportProbe!(returned);
    expect(returned.pages.map(({ page }) => page.pageIndex)).toEqual([0]);
    expect(requests.map(({ type, pages, expectedFrameEpoch }) => ({ type, pages, expectedFrameEpoch }))).toEqual([
      { type: 'buildPages', pages: [39], expectedFrameEpoch: 1 },
      { type: 'buildPages', pages: [0], expectedFrameEpoch: 2 },
    ]);
  } finally {
    unregister?.();
    if (previousFlag === undefined) delete scope.__DOCX_PEERLESS_HARNESS__;
    else scope.__DOCX_PEERLESS_HARNESS__ = previousFlag;
    if (previousHarness === undefined) delete scope.__docxPeerlessHarness;
    else scope.__docxPeerlessHarness = previousHarness;
  }
});
