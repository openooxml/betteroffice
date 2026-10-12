import { expect, test } from 'bun:test';
import docxPackage from '../../../../../docx/package.json';
import type { EchoEntry } from '../../../../../../shared/office-session/commands/types';
import { createPresenter } from '../../../../../../shared/office-session/commands/presenter';
import { discardCommandPresentation } from '@betteroffice/docx/peerless';
import type { DocxCommandFrame } from '@betteroffice/docx/peerless';
import { createEchoPainter, type PeerlessEchoIntent, type PeerlessEchoPainter } from './echoPainter';
import { createFramePresentation, present } from './framePresentation';
import type { PeerlessFrameGeometry, PeerlessViewSnapshot, PeerlessViewStore } from './viewStore';

const identity = { session: 'test', incarnation: 1, epoch: 1 };
const page = { pageId: '1', pageIndex: 0 };
const secondPage = { pageId: '2', pageIndex: 1 };

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

function setup(realPainter = false) {
  const events: string[] = [];
  const draws: ImageBitmap[] = [];
  const pixels = new Map<string, ImageBitmap | ImageData>();
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
      geometry, presentedThroughSeq: geometry.coversThroughSeq }; }, dispose() {},
  };
  function createCanvas(item: typeof page): HTMLCanvasElement {
    const canvas = { width: 100, height: 20, getContext: () => ({
      save() {}, restore() {}, setTransform() {}, clearRect() { pixels.delete(item.pageId); },
      getImageData() {
        const value = { width: canvas.width, height: canvas.height } as ImageData;
        const bitmap = pixels.get(item.pageId);
        if (bitmap) snapshots.set(value, bitmap);
        return value;
      },
      putImageData(value: ImageData) { pixels.set(item.pageId, snapshots.get(value) ?? value); },
      transform() {}, beginPath() {}, rect() {}, clip() {}, fillRect() {},
      measureText(text: string) { return { width: text.length * 10 }; },
      fillText(text: string) { events.push(`echo:${text}`); },
      drawImage(bitmap: ImageBitmap) {
        events.push('pixels'); draws.push(bitmap); onDraw?.(); pixels.set(item.pageId, bitmap);
      },
    }) } as unknown as HTMLCanvasElement;
    return canvas;
  }
  const canvases = new Map([page, secondPage].map((item) => [item.pageId, createCanvas(item)] as const));
  const canvasForPage = (item: typeof page) => mounted ? canvases.get(item.pageId) ?? null : null;
  const painter: PeerlessEchoPainter = realPainter ? createEchoPainter({ canvasForPage }) : {
    installContext() {}, restoreBackground: () => true, preview: () => undefined,
    paint(entries) { events.push(`paint:${entries.length}`); },
    drop(entries, reason) { events.push(`drop:${reason}:${entries.length}`); },
    reconcile(_frame, retired, remaining) { events.push('reconcile'); reconciled.push({ retired, remaining }); },
    disable() {}, dispose() {},
  };
  const presentation = createFramePresentation({ viewStore: store, echoPainter: painter,
    canvasForPage, requestFullFrame() { events.push('full'); },
    prepareGeometry(candidate): PeerlessFrameGeometry { return { identity: snapshot.identity,
      coversThroughSeq: candidate.coversThroughSeq, frameEpoch: candidate.revision,
      layoutRevision: candidate.payload.layoutRevision, pages: geometryPages.map((page) => ({ page,
        bounds: { x: 0, y: 0, width: 100, height: 20 }, transform: [1, 0, 0, 1, 0, 0] as const })),
      caret: null, selectionRects: [] }; },
  });
  return { presentation, events, reconciled, store, draws, pixels, canvases, painter,
    observeDraw(listener: () => void) { onDraw = listener; },
    mountCanvas(value: boolean) { mounted = value; },
    addCanvas(item: typeof page) { const canvas = createCanvas(item); canvases.set(item.pageId, canvas); return canvas; },
    setPages(pages: typeof geometryPages) { geometryPages = pages; },
    setIdentity(next: PeerlessViewSnapshot['identity']) { snapshot = { ...snapshot, identity: next,
      geometry: null, presentedThroughSeq: -1 }; },
  };
}

test('d2c_runtime_docx_imports_use_package_exports', async () => {
  const exported = new Set(Object.keys(docxPackage.exports));
  const imports: string[] = [];
  for (const name of ['framePresentation.ts', 'framePresentation.test.ts', 'echoPainter.ts', 'echoPainter.test.ts']) {
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
    damage: [{ page, rect, backgroundId: 'background', primitiveIds: [] }],
  };
  return candidate;
}

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
  const { presentation, store, pixels, canvases, observeDraw, painter } = setup(true);
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
  presentation.dispose();
  painter.dispose();
  expect(closed).toEqual([first, second, rejectedFirst, rejectedSecond, firstBackup, secondBackup]);
});

test('d2c_viewport_changes_evict_backups_and_restore_active_echo_after_failed_draw', () => {
  const { presentation, painter, canvases, pixels, draws, events, observeDraw, addCanvas, setPages } = setup(true);
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
  presentation.dispose();
  painter.dispose();
  expect(resources.map((value) => closed.get(value))).toEqual(resources.map(() => 1));
});
