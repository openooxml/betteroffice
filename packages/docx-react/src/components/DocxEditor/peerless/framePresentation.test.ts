import { expect, test } from 'bun:test';
import docxPackage from '../../../../../docx/package.json';
import type { EchoEntry } from '../../../../../../shared/office-session/commands/types';
import { createPresenter } from '../../../../../../shared/office-session/commands/presenter';
import { discardCommandPresentation } from '@betteroffice/docx/peerless';
import type { DocxCommandFrame } from '@betteroffice/docx/peerless';
import { createEchoPainter, type PeerlessEchoIntent, type PeerlessEchoPainter } from './echoPainter';
import { glyphChar, installGlyphPaths, testEchoGlyphLine, testEchoGlyphSet } from './echoGlyphs.testing';
import { createFramePresentation, present } from './framePresentation';
import type { PeerlessFrameGeometry, PeerlessViewSnapshot, PeerlessViewStore } from './viewStore';

installGlyphPaths();

const identity = { session: 'test', incarnation: 1, epoch: 1 };
const page = { pageId: '1', pageIndex: 0 };
const secondPage = { pageId: '2', pageIndex: 1 };
const thirdPage = { pageId: '3', pageIndex: 2 };

function entry(seq: number): EchoEntry<PeerlessEchoIntent> {
  return { seq, ticket: { id: seq, seq, ack: Promise.resolve({ seq, status: 'applied' }),
    presented: Promise.resolve('presented') }, value: { originalIntentCount: 1,
    contextFrameEpoch: 1, intent: { kind: 'insertText', payload: { text: 'x' } } } };
}

function frame(revision: number, coverage: number, base?: number): DocxCommandFrame {
  const bytes = new ArrayBuffer(88);
  const view = new DataView(bytes);
  new Uint8Array(bytes).set([0x46, 0x44, 0x56, 0x31]);
  view.setUint16(4, 1, true);
  view.setUint16(6, 80, true);
  view.setUint32(8, 88, true);
  view.setUint32(12, base === undefined ? 1 : 0, true);
  view.setBigUint64(16, 1n, true);
  view.setBigUint64(24, 1n, true);
  view.setBigUint64(32, BigInt(revision), true);
  view.setBigUint64(40, BigInt(base ?? 0), true);
  for (const offset of [56, 60]) view.setUint32(offset, 80, true);
  view.setUint32(64, 4, true);
  view.setUint32(68, 88, true);
  return { revision, coversThroughSeq: coverage, ...(base === undefined ? {} : { baseRevision: base }), payload: {
    frame: bytes, layoutRevision: revision, caret: { frameEpoch: revision, caretRect: null },
    stagedPresentation: { frameEpoch: revision, pages: [{ page, frameEpoch: revision,
      bitmap: { width: 100, height: 20, close() {} } as ImageBitmap }] },
  } };
}

function setup(realPainter = false, makeCanvas?: (item: typeof page) => HTMLCanvasElement, holdRequests = false) {
  const events: string[] = [];
  const held: (() => void)[] = [];
  const draws: ImageBitmap[] = [];
  const pixels = new Map<string, ImageBitmap | ImageData>();
  const readbacks = { getImageData: 0, putImageData: 0 };
  const closes = new Map<ImageBitmap, number>();
  const drawnAfterClose: ImageBitmap[] = [];
  let onDraw: (() => void) | undefined;
  let mounted = true;
  let geometryPages = [page, secondPage];
  const snapshots = new WeakMap<ImageData, ImageBitmap | ImageData>();
  const reconciled: { retired: readonly EchoEntry<PeerlessEchoIntent>[]; remaining: readonly EchoEntry<PeerlessEchoIntent>[] }[] = [];
  let snapshot: PeerlessViewSnapshot = { identity, version: 0, throughSeq: 0, presentedThroughSeq: -1,
    workerSelection: null, anchors: null, editState: undefined, geometry: null };
  const store: PeerlessViewStore = {
    subscribe: () => () => {}, getSnapshot: () => snapshot, getServerSnapshot: () => snapshot,
    acceptWorkerState(_identity, state) { snapshot = { ...snapshot, version: state.version,
      throughSeq: state.throughSeq, workerSelection: state.selection }; },
    installFrameGeometry(geometry) { events.push('geometry'); snapshot = { ...snapshot,
      geometry, presentedThroughSeq: geometry.coversThroughSeq }; }, setProvisionalCaret() {}, dispose() {},
  };
  function createCanvas(item: typeof page): HTMLCanvasElement {
    const canvas = { width: 100, height: 20, getContext: () => ({
      save() {}, restore() {}, setTransform() {}, clearRect() { pixels.delete(item.pageId); },
      getImageData() {
        readbacks.getImageData += 1;
        const value = { width: canvas.width, height: canvas.height } as ImageData;
        const bitmap = pixels.get(item.pageId);
        if (bitmap) snapshots.set(value, bitmap);
        return value;
      },
      putImageData(value: ImageData) {
        readbacks.putImageData += 1;
        pixels.set(item.pageId, snapshots.get(value) ?? value);
      },
      transform() {}, beginPath() {}, rect() {}, clip() {}, fillRect() {},
      translate() {}, scale() {},
      fill(path: unknown) {
        const last = events.at(-1);
        if (last?.startsWith('echo:')) events[events.length - 1] = last + glyphChar(path);
        else events.push(`echo:${glyphChar(path)}`);
      },
      drawImage(bitmap: ImageBitmap) {
        if (closes.has(bitmap)) drawnAfterClose.push(bitmap);
        events.push('pixels'); draws.push(bitmap); onDraw?.(); pixels.set(item.pageId, bitmap);
      },
    }) } as unknown as HTMLCanvasElement;
    return canvas;
  }
  const create = makeCanvas ?? createCanvas;
  const canvases = new Map([page, secondPage].map((item) => [item.pageId, create(item)] as const));
  const canvasForPage = (item: typeof page) => mounted ? canvases.get(item.pageId) ?? null : null;
  const painter: PeerlessEchoPainter = realPainter ? createEchoPainter({ canvasForPage }) : {
    installContext() {}, restoreBackground: () => true, preview: () => undefined,
    paint(entries) { events.push(`paint:${entries.length}`); },
    drop(entries, reason) { events.push(`drop:${reason}:${entries.length}`); },
    reconcile(_frame, retired, remaining) { events.push('reconcile'); reconciled.push({ retired, remaining }); },
    disable() {}, dispose() {},
  };
  const presentation = createFramePresentation({ viewStore: store, echoPainter: painter,
    canvasForPage, onError(error) { events.push(`error:${error.message}`); },
    requestFullFrame(settled?: () => void) { events.push('full'); if (holdRequests) held.push(settled!); else settled?.(); },
    prepareGeometry(candidate): PeerlessFrameGeometry { return { identity: snapshot.identity,
      coversThroughSeq: candidate.coversThroughSeq, frameEpoch: candidate.revision,
      layoutRevision: candidate.payload.layoutRevision, pages: geometryPages.map((page) => ({ page,
        bounds: { x: 0, y: 0, width: 100, height: 20 }, transform: [1, 0, 0, 1, 0, 0] as const })),
      caret: null, selectionRects: [] }; },
  });
  return { presentation, events, held, reconciled, store, draws, pixels, canvases, painter, readbacks, closes, drawnAfterClose,
    observeDraw(listener: () => void) { onDraw = listener; },
    mountCanvas(value: boolean) { mounted = value; },
    addCanvas(item: typeof page) { const canvas = create(item); canvases.set(item.pageId, canvas); return canvas; },
    track(width = 100, height = 20): ImageBitmap {
      const value = { width, height, close() { closes.set(value, (closes.get(value) ?? 0) + 1); } } as ImageBitmap;
      return value;
    },
    setPages(pages: typeof geometryPages) { geometryPages = pages; },
    setIdentity(next: PeerlessViewSnapshot['identity']) { snapshot = { ...snapshot, identity: next,
      geometry: null, presentedThroughSeq: -1 }; },
  };
}

test('d2c_runtime_docx_imports_use_package_exports', async () => {
  const exported = new Set(Object.keys(docxPackage.exports));
  const imports: string[] = [];
  for (const name of ['framePresentation.ts', 'framePresentation.test.ts', 'echoPainter.ts', 'echoPainter.test.ts', 'bitmapOwnership.ts']) {
    const source = await Bun.file(new URL(`./${name}`, import.meta.url)).text();
    for (const match of source.matchAll(/import\s+(?!type\b)[^;]*?\bfrom\s*['"](@betteroffice\/docx(?:\/[^'"]*)?)['"]/g)) {
      imports.push(match[1]!);
      expect(exported.has(match[1]!.replace('@betteroffice/docx', '.'))).toBe(true);
    }
  }
  expect(imports).toContain('@betteroffice/docx/layout/render');
});

function echoFrame(revision: number, coverage: number, text: string): DocxCommandFrame {
  const candidate = frame(revision, coverage);
  const rect = { x: 0, y: 0, width: 100, height: 20 };
  candidate.payload.stagedPresentation!.pages[0]!.backgrounds = [
    { id: 'background', rect, bitmap: { width: 100, height: 20, close() {} } as ImageBitmap },
  ];
  candidate.payload.echoContext = { page, frameEpoch: revision, layoutRevision: revision,
    serializedBytes: 900, text, shaping: 'simple-ltr', transform: [1, 0, 0, 1, 0, 0],
    styleRuns: [{ startUtf16: 0, endUtf16: text.length, fontFamily: 'Arial', fontSizePx: 12,
      color: '#000', bold: false, italic: false, underline: false, strike: false,
      transform: [1, 0, 0, 1, 0, 0], baseline: { x: 0, y: 14 } }],
    boundaries: Array.from({ length: text.length + 1 }, (_, utf16) => ({ utf16, grapheme: utf16,
      caret: { x: utf16 * 10, y: 0, width: 1, height: 20 } })),
    selection: { anchorGrapheme: text.length, focusGrapheme: text.length },
    damage: [{ page, rect, backgroundId: 'background', omittedPrimitiveIds: [] }],
    glyphs: testEchoGlyphLine(),
  };
  candidate.payload.echoGlyphSets = [testEchoGlyphSet];
  return candidate;
}

test('scaled bitmaps install atomically and partial frames retain other page pixels', () => {
  for (const scale of [2, 3]) {
    const { presentation, canvases, pixels, store, events } = setup(true);
    const initial = echoFrame(1, 0, 'ab');
    expect(present(presentation, initial, [], [])).toBe(true);
    const canvas = canvases.get('1')!;
    const original = pixels.get('1');
    const scaled = echoFrame(2, 0, 'ab');
    const staged = scaled.payload.stagedPresentation!.pages[0]!;
    staged.bitmap = { width: 100 * scale, height: 20 * scale, close() {} } as ImageBitmap;
    staged.backgrounds![0]!.bitmap = { width: 100 * scale, height: 20 * scale, close() {} } as ImageBitmap;
    scaled.payload.layoutMeta = { v: 1, layoutRevision: 2, pageCount: 2, partial: false, provisional: false,
      notesConverged: true, pageSizes: new Float64Array([100, 20, 100, 20]), layoutShell: '{}', headersFootersEpoch: 0 };
    expect(canvas.width).toBe(100);
    expect(pixels.get('1')).toBe(original);
    expect(present(presentation, scaled, [], [entry(1)])).toBe(true);
    expect([canvas.width, canvas.height]).toEqual([100 * scale, 20 * scale]);
    expect(store.getSnapshot().geometry!.pages[0]!.bounds).toEqual({ x: 0, y: 0, width: 100, height: 20 });
    expect(events.at(-1)).toBe('echo:abx');
    const partial = frame(3, 0, 2);
    partial.payload.stagedPresentation!.pages[0]!.page = secondPage;
    partial.payload.stagedPresentation!.pages[0]!.bitmap = { width: 100 * scale, height: 20 * scale, close() {} } as ImageBitmap;
    const installed = pixels.get('1');
    expect(present(presentation, partial, [], [entry(1)])).toBe(true);
    expect(pixels.get('1')).toBe(installed);
    expect(canvas.width).toBe(100 * scale);
    expect(canvases.get('2')!.width).toBe(100 * scale);
  }
});

test('d2c_bare_ack_retains_echo_until_covering_frame', async () => {
  const { presentation, events, store } = setup();
  const pending = entry(1);
  presentation.hooks.paint([pending]);
  await pending.ticket.ack;
  store.acceptWorkerState(identity, { epoch: 1, version: 1, throughSeq: 1,
    presentedThroughSeq: -1, selection: null, editState: undefined });
  expect(events).toEqual(['paint:1']);
  expect(store.getSnapshot().presentedThroughSeq).toBe(-1);
  expect(present(presentation, frame(1, 1), [pending], [])).toBe(true);
  expect(events).toEqual(['paint:1', 'pixels', 'geometry', 'reconcile']);
});

test('d2c_covering_frame_retires_only_covered_echo', () => {
  const { presentation, events, draws, observeDraw } = setup(true);
  const presenter = createPresenter({ echo: presentation.hooks, transactionalPresentation: true,
    sameCoverageFrames: true, requestFullFrame() {}, now: () => 0 });
  const initial = echoFrame(1, 0, 'a');
  initial.payload.stagedPresentation!.pages[0]!.bitmap.close = () => { events.push('retire'); };
  expect(presenter.present(initial)).toBe('presented');
  const covered = entry(1);
  const later = entry(2);
  covered.value.intent = { kind: 'insertText', payload: { text: 'b' } };
  later.value.intent = { kind: 'insertText', payload: { text: 'c' } };
  presenter.add(covered);
  presenter.add(later);
  expect(events.at(-1)).toBe('echo:abc');
  events.length = 0;
  draws.length = 0;
  const pendingAtDraw: (readonly EchoEntry<PeerlessEchoIntent>[])[] = [];
  observeDraw(() => { pendingAtDraw.push(presenter.pending); });
  const candidate = echoFrame(2, 1, 'ab');
  expect(presenter.present(candidate)).toBe('presented');
  expect(events).toEqual(['pixels', 'geometry', 'retire', 'pixels', 'echo:abc']);
  expect(draws).toEqual([candidate.payload.stagedPresentation!.pages[0]!.bitmap,
    candidate.payload.stagedPresentation!.pages[0]!.backgrounds![0]!.bitmap]);
  expect(pendingAtDraw).toEqual([[covered, later], [covered, later]]);
  expect(presenter.pending).toEqual([later]);
  expect(presenter.present(candidate)).toBe('stale');
  expect(events.filter((event) => event.startsWith('echo:'))).toEqual(['echo:abc']);
});

test('d2c_recovery_resets_local_revision_and_coverage_before_stale_checks', async () => {
  for (const next of [{ ...identity, incarnation: 2 }, { ...identity, epoch: 2 }]) {
    const { presentation, events, store, setIdentity } = setup();
    const presenter = createPresenter({ echo: presentation.hooks, transactionalPresentation: true,
      sameCoverageFrames: true, requestFullFrame() {}, now: () => 0 });
    expect(presenter.present(frame(10, 5))).toBe('presented');
    const pending = entry(6);
    presenter.add(pending);
    let status: string | undefined;
    const waiting = presenter.whenPresented(6).then((value) => { status = value; return value; });
    setIdentity(next);
    presenter.reset();
    events.length = 0;
    expect(presenter.present(frame(1, 0))).toBe('presented');
    await Promise.resolve();
    expect(status).toBeUndefined();
    expect(store.getSnapshot().geometry?.identity).toEqual(next);
    expect(store.getSnapshot().presentedThroughSeq).toBe(0);
    expect(presenter.present(frame(2, 6, 1))).toBe('presented');
    expect(await waiting).toBe('presented');
    expect(presenter.pending).toEqual([]);
    expect(events).toEqual(['pixels', 'geometry', 'reconcile', 'pixels', 'geometry', 'reconcile']);
  }
});

test('missing canvases retain echo and request fresh bitmaps before presentation', () => {
  const { presentation, events, store, reconciled, mountCanvas } = setup();
  const presenter = createPresenter({ echo: presentation.hooks, transactionalPresentation: true,
    discardFrame: discardCommandPresentation, requestFullFrame() {}, now: () => 0 });
  presenter.add(entry(1));
  events.length = 0;
  let closes = 0;
  const candidate = frame(1, 1);
  candidate.payload.stagedPresentation!.pages[0]!.bitmap.close = () => { closes += 1; };
  mountCanvas(false);
  expect(presenter.present(candidate)).toBe('stale');
  expect(presenter.present(candidate)).toBe('stale');
  expect(presenter.presentedThroughSeq).toBe(-1);
  expect(store.getSnapshot().geometry).toBeNull();
  expect(reconciled).toEqual([]);
  expect(presenter.pending.length).toBe(1);
  expect(events).toEqual(['full', 'full']);
  expect(closes).toBe(1);
  mountCanvas(true);
  expect(presenter.present(frame(2, 1))).toBe('presented');
  expect(events).toEqual(['full', 'full', 'pixels', 'geometry', 'reconcile']);
  expect(presenter.pending).toEqual([]);
  expect(store.getSnapshot().presentedThroughSeq).toBe(1);
  expect(closes).toBe(1);
});

test('the painter and shared presenter close rejected bitmaps once', () => {
  const { presentation } = setup();
  const presenter = createPresenter({ echo: presentation.hooks, transactionalPresentation: true,
    discardFrame: discardCommandPresentation, requestFullFrame() {}, now: () => 0 });
  const candidate = frame(1, 1);
  let closes = 0;
  candidate.payload.stagedPresentation!.pages[0]!.bitmap.close = () => { closes += 1; };
  candidate.payload.stagedPresentation!.frameEpoch = 99;
  expect(presenter.present(candidate)).toBe('stale');
  expect(closes).toBe(1);
  presentation.dispose();
  presenter.dispose();
  expect(closes).toBe(1);
});

test('d2c_rejected_or_superseded_mutation_restores_worker_selection', () => {
  for (const reason of ['refused', 'superseded']) {
    const { presentation, store, events } = setup();
    const selection = { anchor: { story: 'body', paraId: 'worker', offset: 7 },
      focus: { story: 'body', paraId: 'worker', offset: 7 } };
    presentation.hooks.paint([entry(1), entry(2)]);
    presentation.hooks.drop([entry(1), entry(2)], reason);
    store.acceptWorkerState(identity, { epoch: 1, version: 0, throughSeq: 2,
      presentedThroughSeq: -1, selection, editState: undefined });
    expect(events).toEqual(['paint:2', `drop:${reason}:2`]);
    expect(store.getSnapshot().workerSelection).toBe(selection);
    expect(store.getSnapshot().geometry).toBeNull();
  }
});

test('d2c_delta_mismatch_retains_echo_and_requests_full_frame', () => {
  for (const binary of [false, true]) {
    const { presentation, events, store, reconciled } = setup();
    expect(presentation.present(frame(1, 0), [], [entry(1)])).toBe(true);
    events.length = 0;
    const candidate = frame(2, 1, binary ? 1 : 99);
    if (binary) new DataView(candidate.payload.frame).setBigUint64(40, 99n, true);
    expect(presentation.present(candidate, [entry(1)], [])).toBe(false);
    expect(events).toEqual(['full']);
    expect(store.getSnapshot().presentedThroughSeq).toBe(0);
    expect(reconciled.length).toBe(1);
  }
});

const stalled = 'Full frame responses did not satisfy requested coverage and revision';

function lossyCanvas(state: { lost: boolean }): HTMLCanvasElement {
  const ctx = { save() {}, restore() {}, setTransform() {}, clearRect() {}, drawImage() {} };
  return { width: 100, height: 20, getContext: () => state.lost ? null : ctx } as unknown as HTMLCanvasElement;
}

test('a presentation that keeps failing re-requests twice, reports once and stops requesting', () => {
  const { presentation, events } = setup(false, () => lossyCanvas({ lost: true }));
  let revision = 0;
  do expect(presentation.present(frame(++revision, 0), [], [])).toBe(false);
  while (events.at(-1) === 'full' && revision < 20);
  expect(events).toEqual(['full', 'full', `error:${stalled}`]);
  expect(presentation.present(frame(++revision, 0), [], [])).toBe(false);
  expect(events).toEqual(['full', 'full', `error:${stalled}`]);
});

test('a stalled presentation re-arms one request on invalidation and recovers on the next installable frame', () => {
  for (const clear of ['invalidate', 'frame'] as const) {
    const state = { lost: true };
    const { presentation, events } = setup(false, () => lossyCanvas(state));
    let revision = 0;
    do presentation.present(frame(++revision, 0), [], []);
    while (events.at(-1) === 'full' && revision < 20);
    expect(events.at(-1)).toBe(`error:${stalled}`);
    events.length = 0;
    presentation.invalidate();
    expect(presentation.present(frame(++revision, 0), [], [])).toBe(false);
    expect(events).toEqual(['full']);
    state.lost = false;
    events.length = 0;
    if (clear === 'invalidate') presentation.invalidate();
    expect(events).toEqual(clear === 'invalidate' ? ['full'] : []);
    expect(presentation.present(frame(++revision, 1), [], [])).toBe(true);
    events.length = 0;
    presentation.invalidate();
    state.lost = true;
    expect(presentation.present(frame(++revision, 1), [], [])).toBe(false);
    expect(events).toEqual(['full']);
  }
});

test('frames the presentation did not request spend no retry budget while its own request is pending', () => {
  const { presentation, events, held } = setup(false, () => lossyCanvas({ lost: true }), true);
  let revision = 0;
  const failUnrequested = () => {
    for (let count = 0; count < 3; count += 1) expect(presentation.present(frame(++revision, 0), [], [])).toBe(false);
  };
  expect(presentation.present(frame(++revision, 0), [], [])).toBe(false);
  failUnrequested();
  expect(events).toEqual(['full']);
  held.shift()!();
  expect(presentation.present(frame(++revision, 0), [], [])).toBe(false);
  failUnrequested();
  expect(events).toEqual(['full', 'full']);
  held.shift()!();
  expect(presentation.present(frame(++revision, 0), [], [])).toBe(false);
  expect(events).toEqual(['full', 'full', `error:${stalled}`]);
  failUnrequested();
  expect(events).toEqual(['full', 'full', `error:${stalled}`]);
});

test('a newer frame that failed while a request was pending is fetched again after an older frame installs', () => {
  const state = { lost: true };
  const { presentation, events, held } = setup(false, () => lossyCanvas(state), true);
  const requested = () => events.filter((event) => event === 'full' || event.startsWith('error:'));
  expect(presentation.present(frame(1, 0), [], [])).toBe(false);
  expect(presentation.present(frame(3, 0), [], [])).toBe(false);
  expect(requested()).toEqual(['full']);
  state.lost = false;
  held.shift()!();
  expect(presentation.present(frame(2, 0), [], [])).toBe(true);
  expect(requested()).toEqual(['full', 'full']);
  held.shift()!();
  expect(presentation.present(frame(4, 0), [], [])).toBe(true);
  state.lost = true;
  expect(presentation.present(frame(5, 0), [], [])).toBe(false);
  expect(presentation.present(frame(6, 0), [], [])).toBe(false);
  state.lost = false;
  held.shift()!();
  expect(presentation.present(frame(7, 0), [], [])).toBe(true);
  expect(requested()).toEqual(['full', 'full', 'full']);
  expect(held).toEqual([]);
});

test('transient failures each re-request one full frame and recover without a report', () => {
  const { presentation, events } = setup();
  expect(presentation.present(frame(1, 0), [], [])).toBe(true);
  for (let revision = 2; revision < 8; revision += 2) {
    expect(presentation.present(frame(revision, 0, 99), [], [])).toBe(false);
    expect(presentation.present(frame(revision + 1, 0), [], [])).toBe(true);
  }
  expect(events.filter((event) => event === 'full' || event.startsWith('error:'))).toEqual(['full', 'full', 'full']);
});

test('d2c_new_revision_at_same_coverage_installs_without_extra_retirement', () => {
  const { presentation, events, reconciled } = setup();
  const pending = entry(2);
  expect(presentation.present(frame(1, 1), [entry(1)], [pending])).toBe(true);
  events.length = 0;
  expect(presentation.present(frame(2, 1, 1), [], [pending])).toBe(true);
  expect(events).toEqual(['pixels', 'geometry', 'reconcile']);
  expect(reconciled.at(-1)).toEqual({ retired: [], remaining: [pending] });
  expect(presentation.present(frame(2, 1), [], [pending])).toBe(false);
});

test('d2c_failed_second_page_draw_restores_pages_from_partial_frames', () => {
  const { presentation, store, pixels, canvases, observeDraw, painter, readbacks } = setup(true);
  const closed: ImageBitmap[] = [];
  function bitmap(width = 100, height = 20): ImageBitmap {
    const value = { width, height, close() { closed.push(value); } } as ImageBitmap;
    return value;
  }
  const initial = frame(1, 0);
  const first = bitmap();
  const second = bitmap();
  initial.payload.stagedPresentation!.pages = [
    { page, frameEpoch: 1, bitmap: first }, { page: secondPage, frameEpoch: 1, bitmap: second },
  ];
  expect(presentation.present(initial, [], [])).toBe(true);
  const partialFirst = frame(2, 0, 1);
  const firstBackup = bitmap(120, 30);
  partialFirst.payload.stagedPresentation!.pages[0]!.bitmap = firstBackup;
  expect(presentation.present(partialFirst, [], [])).toBe(true);
  const partialSecond = frame(3, 0, 2);
  const secondBackup = bitmap(110, 25);
  partialSecond.payload.stagedPresentation!.pages = [
    { page: secondPage, frameEpoch: 3, bitmap: secondBackup },
  ];
  expect(presentation.present(partialSecond, [], [])).toBe(true);
  expect(closed).toEqual([first, second]);
  const candidate = frame(4, 1, 3);
  const rejectedFirst = bitmap(200, 40);
  const rejectedSecond = bitmap(210, 45);
  candidate.payload.stagedPresentation!.pages = [
    { page, frameEpoch: 4, bitmap: rejectedFirst },
    { page: secondPage, frameEpoch: 4, bitmap: rejectedSecond },
  ];
  observeDraw(() => {
    if (pixels.get(page.pageId) === rejectedFirst && !pixels.has(secondPage.pageId)) {
      throw new Error('draw-failed');
    }
  });
  expect(presentation.present(candidate, [], [])).toBe(false);
  expect(pixels.get(page.pageId)).toBe(firstBackup);
  expect(pixels.get(secondPage.pageId)).toBe(secondBackup);
  expect([canvases.get('1')!.width, canvases.get('1')!.height]).toEqual([120, 30]);
  expect([canvases.get('2')!.width, canvases.get('2')!.height]).toEqual([110, 25]);
  expect(store.getSnapshot().geometry?.frameEpoch).toBe(3);
  expect(store.getSnapshot().presentedThroughSeq).toBe(0);
  expect(closed).toEqual([first, second, rejectedFirst, rejectedSecond]);
  expect(readbacks).toEqual({ getImageData: 0, putImageData: 0 });
  presentation.dispose();
  painter.dispose();
  expect(closed).toEqual([first, second, rejectedFirst, rejectedSecond, firstBackup, secondBackup]);
});

test('d2c_viewport_changes_evict_backups_and_restore_active_echo_after_failed_draw', () => {
  const { presentation, painter, canvases, pixels, draws, events, observeDraw, addCanvas, setPages, readbacks } = setup(true);
  const closed = new Map<ImageBitmap, number>();
  const resources: ImageBitmap[] = [];
  const initial = echoFrame(1, 0, 'a');
  for (const value of [initial.payload.stagedPresentation!.pages[0]!.bitmap,
    initial.payload.stagedPresentation!.pages[0]!.backgrounds![0]!.bitmap]) {
    resources.push(value);
    value.close = () => { closed.set(value, (closed.get(value) ?? 0) + 1); };
  }
  expect(presentation.present(initial, [], [])).toBe(true);
  const pending = entry(1);
  presentation.hooks.paint([pending]);
  const detached: HTMLCanvasElement[] = [];
  let previous = secondPage;
  for (let revision = 2; revision <= 20; revision++) {
    detached.push(canvases.get(previous.pageId)!);
    canvases.delete(previous.pageId);
    const nextPage = { pageId: String(revision + 1), pageIndex: revision };
    addCanvas(nextPage);
    setPages([page, nextPage]);
    const candidate = frame(revision, 0, revision - 1);
    const value = candidate.payload.stagedPresentation!.pages[0]!.bitmap;
    resources.push(value);
    value.close = () => { closed.set(value, (closed.get(value) ?? 0) + 1); };
    candidate.payload.stagedPresentation!.pages[0]!.page = nextPage;
    expect(presentation.present(candidate, [], [pending])).toBe(true);
    expect(resources.filter((resource) => !closed.has(resource)).length).toBe(3);
    previous = nextPage;
  }
  for (const canvas of detached) canvas.getContext = () => { throw new Error('detached-canvas'); };
  const candidate = frame(21, 1, 20);
  const rejected = candidate.payload.stagedPresentation!.pages[0]!.bitmap;
  resources.push(rejected);
  rejected.close = () => { closed.set(rejected, (closed.get(rejected) ?? 0) + 1); };
  observeDraw(() => { if (draws.at(-1) === rejected) throw new Error('draw-failed'); });
  expect(presentation.present(candidate, [pending], [])).toBe(false);
  expect(events.at(-2)).toBe('echo:ax');
  presentation.hooks.drop([pending], 'refused');
  expect(pixels.get(page.pageId)).toBe(initial.payload.stagedPresentation!.pages[0]!.bitmap);
  setPages([previous]);
  const deleted = frame(22, 1, 20);
  const bytes = new ArrayBuffer(136);
  new Uint8Array(bytes).set(new Uint8Array(deleted.payload.frame).subarray(0, 80));
  const view = new DataView(bytes);
  view.setUint32(8, 136, true);
  view.setUint32(48, 1, true);
  view.setUint32(52, 1, true);
  view.setUint32(60, 128, true);
  view.setUint32(68, 136, true);
  view.setUint8(80, 2);
  view.setBigUint64(88, 1n, true);
  deleted.payload.frame = bytes;
  const replacement = deleted.payload.stagedPresentation!.pages[0]!;
  replacement.page = previous;
  resources.push(replacement.bitmap);
  replacement.bitmap.close = () => { closed.set(replacement.bitmap, (closed.get(replacement.bitmap) ?? 0) + 1); };
  expect(presentation.present(deleted, [], [])).toBe(true);
  expect(closed.get(initial.payload.stagedPresentation!.pages[0]!.bitmap)).toBe(1);
  expect(readbacks).toEqual({ getImageData: 0, putImageData: 0 });
  presentation.dispose();
  painter.dispose();
  expect(resources.map((value) => closed.get(value))).toEqual(resources.map(() => 1));
});

function staged(revision: number, base: number | undefined,
  pages: readonly (readonly [typeof page, ImageBitmap])[]): DocxCommandFrame {
  const candidate = frame(revision, 0, base);
  candidate.payload.stagedPresentation!.pages = pages.map(([item, bitmap]) => ({ page: item, frameEpoch: revision, bitmap }));
  return candidate;
}

type PixelBitmap = ImageBitmap & { data: Uint8ClampedArray; failing: boolean };
type PixelCanvas = HTMLCanvasElement & { readonly pixels: Uint8ClampedArray };

function pixelBitmap(width: number, height: number, seed: number, failing = false): PixelBitmap {
  const data = new Uint8ClampedArray(width * height * 4);
  for (let index = 0; index < width * height; index++) {
    if ((index + seed) % 7 === 0) continue;
    data.set([(index * 3 + seed) % 256, (index * 5 + seed * 7) % 256, (seed * 31) % 256, 255], index * 4);
  }
  return { width, height, data, failing, close() {} } as unknown as PixelBitmap;
}

function pixelCanvas(calls: { getImageData: number; putImageData: number }): PixelCanvas {
  let width = 300;
  let height = 150;
  let data = new Uint8ClampedArray(width * height * 4);
  const copy = (source: Uint8ClampedArray, sourceWidth: number, sourceHeight: number, x: number, y: number, opaqueOnly: boolean) => {
    for (let row = 0; row < sourceHeight && row + y < height; row++) {
      for (let column = 0; column < sourceWidth && column + x < width; column++) {
        const from = (row * sourceWidth + column) * 4;
        if (opaqueOnly && source[from + 3] === 0) continue;
        data.set(source.subarray(from, from + 4), ((row + y) * width + column + x) * 4);
      }
    }
  };
  const ctx = {
    save() {}, restore() {}, setTransform() {}, transform() {}, beginPath() {}, rect() {}, clip() {},
    fillRect() {}, fillText() {}, measureText: (text: string) => ({ width: text.length * 10 }),
    clearRect(x: number, y: number, w: number, h: number) {
      for (let row = Math.max(0, y); row < Math.min(height, y + h); row++) {
        data.fill(0, (row * width + Math.max(0, x)) * 4, (row * width + Math.min(width, x + w)) * 4);
      }
    },
    drawImage(bitmap: PixelBitmap, x = 0, y = 0, ...size: number[]) {
      if (bitmap.failing) throw new Error('draw-failed');
      if (size.length && (size[0] !== bitmap.width || size[1] !== bitmap.height)) throw new Error('scaled-draw');
      copy(bitmap.data, bitmap.width, bitmap.height, x, y, true);
    },
    getImageData(x: number, y: number, w: number, h: number) {
      calls.getImageData += 1;
      const out = new Uint8ClampedArray(w * h * 4);
      for (let row = 0; row < h; row++) out.set(data.subarray(((row + y) * width + x) * 4, ((row + y) * width + x + w) * 4), row * w * 4);
      return { width: w, height: h, data: out } as unknown as ImageData;
    },
    putImageData(image: ImageData, x: number, y: number) {
      calls.putImageData += 1;
      copy(image.data, image.width, image.height, x, y, false);
    },
  };
  return {
    get width() { return width; },
    set width(value: number) { width = value; data = new Uint8ClampedArray(width * height * 4); },
    get height() { return height; },
    set height(value: number) { height = value; data = new Uint8ClampedArray(width * height * 4); },
    get pixels() { return data; },
    getContext: () => ctx,
  } as unknown as PixelCanvas;
}

function expectPixels(canvas: PixelCanvas, bitmap: PixelBitmap): void {
  expect([canvas.width, canvas.height]).toEqual([bitmap.width, bitmap.height]);
  expect(canvas.pixels.length === bitmap.data.length && canvas.pixels.every((value, index) => value === bitmap.data[index])).toBe(true);
}

function pixelSetup() {
  const calls = { getImageData: 0, putImageData: 0 };
  const canvases = new Map<string, PixelCanvas>();
  const context = setup(true, (item) => {
    const canvas = pixelCanvas(calls);
    canvases.set(item.pageId, canvas);
    return canvas;
  });
  context.addCanvas(thirdPage);
  context.setPages([page, secondPage, thirdPage]);
  return { ...context, calls, canvas: (item: typeof page) => canvases.get(item.pageId)! };
}

test('installs leave each canvas with exactly the staged pixels at the staged size without readback', () => {
  const { presentation, canvas, calls } = pixelSetup();
  const single = pixelBitmap(100, 20, 1);
  expect(presentation.present(staged(1, undefined, [[page, single]]), [], [])).toBe(true);
  expectPixels(canvas(page), single);
  const multi = [pixelBitmap(100, 20, 2), pixelBitmap(100, 20, 3), pixelBitmap(100, 20, 4)] as const;
  expect(presentation.present(staged(2, 1, [[page, multi[0]], [secondPage, multi[1]], [thirdPage, multi[2]]]), [], [])).toBe(true);
  expectPixels(canvas(page), multi[0]);
  expectPixels(canvas(secondPage), multi[1]);
  expectPixels(canvas(thirdPage), multi[2]);
  const resized = [pixelBitmap(150, 30, 5), pixelBitmap(80, 16, 6)] as const;
  expect(presentation.present(staged(3, 2, [[page, resized[0]], [secondPage, resized[1]]]), [], [])).toBe(true);
  expectPixels(canvas(page), resized[0]);
  expectPixels(canvas(secondPage), resized[1]);
  expectPixels(canvas(thirdPage), multi[2]);
  expect(calls).toEqual({ getImageData: 0, putImageData: 0 });
});

test('a failed install restores retained pixels exactly and blanks a canvas without a frame at its previous size', () => {
  const { presentation, canvas, calls, events } = pixelSetup();
  const first = pixelBitmap(150, 30, 1);
  const second = pixelBitmap(200, 40, 2);
  expect(presentation.present(staged(1, undefined, [[page, first], [secondPage, second]]), [], [])).toBe(true);
  canvas(thirdPage).width = 64;
  canvas(thirdPage).height = 16;
  events.length = 0;
  expect(presentation.present(staged(2, 1, [[page, pixelBitmap(120, 24, 3)],
    [secondPage, pixelBitmap(210, 42, 4, true)]]), [], [])).toBe(false);
  expectPixels(canvas(page), first);
  expectPixels(canvas(secondPage), second);
  expect(presentation.present(staged(2, 1, [[thirdPage, pixelBitmap(90, 18, 5)],
    [secondPage, pixelBitmap(210, 42, 6, true)]]), [], [])).toBe(false);
  expect([canvas(thirdPage).width, canvas(thirdPage).height]).toEqual([64, 16]);
  expect(canvas(thirdPage).pixels.every((value) => value === 0)).toBe(true);
  expectPixels(canvas(page), first);
  expectPixels(canvas(secondPage), second);
  expect(events.filter((event) => event === 'full')).toHaveLength(2);
  expect(calls).toEqual({ getImageData: 0, putImageData: 0 });
});

test('a failed second-page draw redraws both pages from retained bitmaps before echo repaints', () => {
  const { presentation, events, draws, observeDraw, readbacks } = setup(true);
  const initial = echoFrame(1, 0, 'a');
  const first = initial.payload.stagedPresentation!.pages[0]!.bitmap;
  const second = { width: 100, height: 20, close() {} } as ImageBitmap;
  initial.payload.stagedPresentation!.pages = [...initial.payload.stagedPresentation!.pages,
    { page: secondPage, frameEpoch: 1, bitmap: second }];
  expect(presentation.present(initial, [], [])).toBe(true);
  const pending = entry(1);
  presentation.hooks.paint([pending]);
  expect(events.at(-1)).toBe('echo:ax');
  const rejectedFirst = { width: 100, height: 20, close() {} } as ImageBitmap;
  const rejectedSecond = { width: 100, height: 20, close() {} } as ImageBitmap;
  observeDraw(() => { if (draws.at(-1) === rejectedSecond) throw new Error('draw-failed'); });
  draws.length = 0;
  events.length = 0;
  expect(presentation.present(staged(2, 1, [[page, rejectedFirst], [secondPage, rejectedSecond]]), [], [pending])).toBe(false);
  expect(draws.slice(0, 4)).toEqual([rejectedFirst, rejectedSecond, first, second]);
  expect(events.slice(-2)).toEqual(['echo:ax', 'full']);
  expect(readbacks).toEqual({ getImageData: 0, putImageData: 0 });
});

test('staged bitmaps close exactly once and are never drawn after close', () => {
  const { presentation, painter, canvases, draws, pixels, observeDraw, addCanvas, setPages, setIdentity,
    track, closes, drawnAfterClose, readbacks } = setup(true);
  const tracked: ImageBitmap[] = [];
  const bitmap = () => { const value = track(); tracked.push(value); return value; };
  const failOn = (value: ImageBitmap) => observeDraw(() => { if (draws.at(-1) === value) throw new Error('draw-failed'); });
  const initial = echoFrame(1, 0, 'a');
  const installed = initial.payload.stagedPresentation!.pages[0]!;
  installed.bitmap = bitmap();
  installed.backgrounds![0]!.bitmap = bitmap();
  const second = bitmap();
  initial.payload.stagedPresentation!.pages = [...initial.payload.stagedPresentation!.pages,
    { page: secondPage, frameEpoch: 1, bitmap: second }];
  expect(presentation.present(initial, [], [])).toBe(true);
  const replacement = bitmap();
  expect(presentation.present(staged(2, 1, [[page, replacement]]), [], [])).toBe(true);
  expect(closes.get(installed.bitmap)).toBe(1);
  canvases.delete(secondPage.pageId);
  addCanvas(thirdPage);
  setPages([page, thirdPage]);
  const third = bitmap();
  expect(presentation.present(staged(3, 2, [[thirdPage, third]]), [], [])).toBe(true);
  expect(closes.get(second)).toBe(1);
  const mismatched = staged(4, 3, [[page, bitmap()]]);
  mismatched.payload.stagedPresentation!.frameEpoch = 99;
  expect(presentation.present(mismatched, [], [])).toBe(false);
  const failedPage = bitmap();
  const failedThird = bitmap();
  failOn(failedThird);
  expect(presentation.present(staged(4, 3, [[page, failedPage], [thirdPage, failedThird]]), [], [])).toBe(false);
  expect(new Set(draws.slice(-2))).toEqual(new Set([replacement, third]));
  expect([closes.get(failedPage), closes.get(failedThird)]).toEqual([1, 1]);
  setIdentity({ ...identity, epoch: 2 });
  const reset = bitmap();
  expect(presentation.present(staged(1, undefined, [[page, reset]]), [], [])).toBe(true);
  expect([closes.get(replacement), closes.get(installed.backgrounds![0]!.bitmap)]).toEqual([1, 1]);
  expect(closes.has(third)).toBe(false);
  const failedAfterReset = bitmap();
  failOn(failedAfterReset);
  expect(presentation.present(staged(2, 1, [[thirdPage, failedAfterReset]]), [], [])).toBe(false);
  expect(new Set(draws.slice(-2))).toEqual(new Set([reset, third]));
  expect(pixels.get(thirdPage.pageId)).toBe(third);
  presentation.dispose();
  painter.dispose();
  expect(tracked.map((value) => closes.get(value))).toEqual(tracked.map(() => 1));
  expect(drawnAfterClose).toEqual([]);
  expect(readbacks).toEqual({ getImageData: 0, putImageData: 0 });
});

test('a page moved by geometry alone keeps its backup for a later rollback', () => {
  const { presentation, painter, pixels, draws, observeDraw, setPages, track, closes, drawnAfterClose } = setup(true);
  const tracked: ImageBitmap[] = [];
  const bitmap = () => { const value = track(); tracked.push(value); return value; };
  const first = bitmap();
  const second = bitmap();
  expect(presentation.present(staged(1, undefined, [[page, first], [secondPage, second]]), [], [])).toBe(true);
  const moved = { pageId: secondPage.pageId, pageIndex: 2 };
  setPages([page, moved]);
  const replacement = bitmap();
  expect(presentation.present(staged(2, 1, [[page, replacement]]), [], [])).toBe(true);
  const restaged = bitmap();
  const failing = bitmap();
  observeDraw(() => { if (draws.at(-1) === failing) throw new Error('draw-failed'); });
  expect(presentation.present(staged(3, 2, [[moved, restaged], [page, failing]]), [], [])).toBe(false);
  expect(pixels.get(secondPage.pageId)).toBe(second);
  expect(pixels.get(page.pageId)).toBe(replacement);
  presentation.dispose();
  painter.dispose();
  expect(tracked.map((value) => closes.get(value))).toEqual(tracked.map(() => 1));
  expect(drawnAfterClose).toEqual([]);
});

test('reinstalling the authoritative bitmap object keeps it open for rollback', () => {
  const { presentation, painter, pixels, draws, observeDraw, setIdentity, track, closes, drawnAfterClose } = setup(true);
  const tracked: ImageBitmap[] = [];
  const bitmap = () => { const value = track(); tracked.push(value); return value; };
  const second = bitmap();
  expect(presentation.present(staged(1, undefined, [[page, bitmap()], [secondPage, second]]), [], [])).toBe(true);
  setIdentity({ ...identity, epoch: 2 });
  expect(presentation.present(staged(1, undefined, [[page, bitmap()]]), [], [])).toBe(true);
  expect(presentation.present(staged(2, 1, [[secondPage, second]]), [], [])).toBe(true);
  expect(closes.has(second)).toBe(false);
  const failing = bitmap();
  observeDraw(() => { if (draws.at(-1) === failing) throw new Error('draw-failed'); });
  expect(presentation.present(staged(3, 2, [[secondPage, failing]]), [], [])).toBe(false);
  expect(pixels.get(secondPage.pageId)).toBe(second);
  presentation.dispose();
  painter.dispose();
  expect(tracked.map((value) => closes.get(value))).toEqual(tracked.map(() => 1));
  expect(drawnAfterClose).toEqual([]);
});

test('a failed echo reconciliation keeps the installed frame and its bitmaps open', () => {
  const { presentation, painter, pixels, draws, observeDraw, track, closes, drawnAfterClose } = setup(true);
  const presenter = createPresenter({ echo: presentation.hooks, transactionalPresentation: true, sameCoverageFrames: true,
    discardFrame: discardCommandPresentation, requestFullFrame() {}, now: () => 0 });
  const tracked: ImageBitmap[] = [];
  const bitmap = () => { const value = track(); tracked.push(value); return value; };
  const initial = echoFrame(1, 0, 'a');
  initial.payload.stagedPresentation!.pages[0]!.bitmap = bitmap();
  initial.payload.stagedPresentation!.pages[0]!.backgrounds![0]!.bitmap = bitmap();
  expect(presenter.present(initial)).toBe('presented');
  presenter.add(entry(1));
  const next = echoFrame(2, 0, 'a');
  const installed = bitmap();
  const background = bitmap();
  next.payload.stagedPresentation!.pages[0]!.bitmap = installed;
  next.payload.stagedPresentation!.pages[0]!.backgrounds![0]!.bitmap = background;
  observeDraw(() => { if (draws.at(-1) === background) throw new Error('draw-failed'); });
  presenter.present(next);
  expect(closes.has(installed)).toBe(false);
  expect(pixels.get(page.pageId)).toBe(installed);
  const failing = bitmap();
  observeDraw(() => { if (draws.at(-1) === failing) throw new Error('draw-failed'); });
  expect(presenter.present(staged(3, undefined, [[page, failing]]))).toBe('stale');
  expect(pixels.get(page.pageId)).toBe(installed);
  presenter.dispose();
  presentation.dispose();
  painter.dispose();
  expect(tracked.map((value) => closes.get(value))).toEqual(tracked.map(() => 1));
  expect(drawnAfterClose).toEqual([]);
});
