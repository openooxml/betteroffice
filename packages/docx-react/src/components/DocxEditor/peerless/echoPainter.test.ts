import { expect, test } from 'bun:test';
import type {
  CommandAck, CommandConnection, CommandEnvelope, CommandTicket, EchoEntry,
} from '../../../../../../shared/office-session/commands/types';
import { createWorkerEditSession } from '../../../../../../shared/office-session/commands/session';
import type { DocxCommandEchoContext, DocxCommandFrame } from '@betteroffice/docx/peerless';
import { createEchoPainter, type PeerlessEchoIntent, type PeerlessEchoPainter } from './echoPainter';
import type { PeerlessTextIntent } from './inputEvents';

const page = { pageId: '1', pageIndex: 0 };
const rect = { x: 0, y: 0, width: 100, height: 20 };

function context(): DocxCommandEchoContext {
  return {
    page, frameEpoch: 1, layoutRevision: 1, serializedBytes: 900, text: 'ab',
    styleRuns: [{ startUtf16: 0, endUtf16: 2, fontFamily: 'Arial', fontSizePx: 12,
      color: '#000', bold: false, italic: false, underline: false, strike: false,
      transform: [1, 0, 0, 1, 0, 0], baseline: { x: 0, y: 14 } }],
    boundaries: [0, 1, 2].map((utf16) => ({ utf16, grapheme: utf16,
      caret: { x: utf16 * 10, y: 0, width: 1, height: 20 } })),
    selection: { anchorGrapheme: 2, focusGrapheme: 2 }, transform: [1, 0, 0, 1, 0, 0],
    damage: [{ page, rect, backgroundId: 'background', primitiveIds: ['shading', 'media'] }],
    shaping: 'simple-ltr',
  };
}

function entry(id: number, value: PeerlessEchoIntent): EchoEntry<PeerlessEchoIntent> {
  return { seq: id, value, ticket: { id, seq: id,
    ack: Promise.resolve({ status: 'applied', seq: id }), presented: Promise.resolve('presented') } };
}

function setup(characterWidth = 10, canEchoFont?: (family: string) => boolean) {
  const draws: unknown[][] = [];
  const texts: string[] = [];
  const transforms: number[][] = [];
  const ctx = {
    save() {}, restore() {}, setTransform(...values: number[]) { transforms.push(values); },
    transform() {}, beginPath() {}, rect() {}, clip() {},
    clearRect() {}, fillRect() {}, fillText(text: string) { texts.push(text); },
    measureText(text: string) { return { width: text.length * characterWidth }; },
    drawImage(...args: unknown[]) { draws.push(args); },
  };
  const canvas = { width: 100, height: 20, getContext: () => ctx } as unknown as HTMLCanvasElement;
  let mountedPages: ReadonlySet<string> | undefined;
  const painter = createEchoPainter({ canEchoFont, canvasForPage: (item) =>
    mountedPages === undefined || mountedPages.has(item.pageId) ? canvas : null });
  const bitmap = { width: 100, height: 20, close() {} } as ImageBitmap;
  const background = { width: 100, height: 20, close() {} } as ImageBitmap;
  const frame: DocxCommandFrame = { coversThroughSeq: 0, revision: 1, payload: {
    frame: new ArrayBuffer(0), caret: { frameEpoch: 1, caretRect: null }, layoutRevision: 1,
    echoContext: context(), stagedPresentation: { frameEpoch: 1, pages: [
      { page, frameEpoch: 1, bitmap, backgrounds: [{ id: 'background', rect, bitmap: background }] },
    ] },
  } };
  painter.reconcile(frame, [], []);
  return { painter, draws, texts, transforms, canvas, background, bitmap, frame,
    mountPages(pages: ReadonlySet<string>) { mountedPages = pages; },
  };
}

test('worker-only fonts decline main-canvas echo while retaining authoritative pixels', () => {
  const { painter, texts } = setup(10, () => false);
  expect(painter.preview({ kind: 'insertText', payload: { text: 'x' } })).toBeUndefined();
  expect(texts).toEqual([]);
});

function sessionWithHeldReplies(painter: PeerlessEchoPainter) {
  type Payload = PeerlessTextIntent['payload'];
  type Ack = CommandAck<null, undefined, DocxCommandFrame['payload']>;
  const calls: { envelope: CommandEnvelope<Payload>; resolve(value: Ack): void }[] = [];
  const queued: (() => void)[] = [];
  const connection: CommandConnection<Payload, null, undefined, DocxCommandFrame['payload']> = {
    initial: { session: 'test', incarnation: 1, epoch: 1, version: 0, throughSeq: 0, outcomes: [] },
    client: {
      call: {
        edit(envelope) { return new Promise<Ack>((resolve) => { calls.push({ envelope, resolve }); }); },
        checkpoint: async () => ({ state: new Uint8Array(), epoch: 1, version: 0, throughSeq: 0 }),
      },
      failure: undefined, onFailure: () => () => {}, dispose: async () => {},
    },
  };
  const session = createWorkerEditSession<string, Payload, null, undefined,
    DocxCommandFrame['payload'], PeerlessEchoIntent>({
    source: 'test', connection, recovery: { reopen: async () => { throw new Error('unexpected recovery'); } },
    maxRecoveryAttempts: 0, maxInFlight: 512, maxEchoEntries: 512, checkpointPolicy: () => false,
    requestFullFrame: async () => connection.initial, now: () => 0,
    enqueue(callback) { queued.push(callback); },
    echo: { paint: painter.paint, drop: painter.drop, present: painter.reconcile },
  });
  let tail: CommandTicket | undefined;
  return { session, calls,
    submit(intent: PeerlessTextIntent) {
      const echo = painter.preview(intent);
      const ticket = session.submit(intent.kind, intent.payload,
        { ...(tail ? { after: tail } : {}), ...(echo ? { echo } : {}) });
      tail = ticket;
      return { ticket, echo };
    },
    pump() { while (queued.length) queued.shift()!(); },
    acknowledge() {
      for (const { envelope, resolve } of calls) resolve({ ...connection.initial,
        version: envelope.seq, throughSeq: envelope.seq, outcomes: [{ seq: envelope.seq, status: 'applied' }] });
    },
  };
}

test('d2c_echo_caps_disable_preview_without_dropping_input', async () => {
  for (const cap of ['serialized-context', 'actual-context', 'pending-intents']) {
    const { painter, texts } = setup(0);
    if (cap === 'serialized-context') painter.installContext({ ...context(), serializedBytes: 16 * 1024 + 1 });
    if (cap === 'actual-context') painter.installContext({ ...context(), text: 'x'.repeat(16 * 1024), serializedBytes: 1 });
    const h = sessionWithHeldReplies(painter);
    const intents: PeerlessTextIntent[] = Array.from({ length: 258 }, (_, index) =>
      ({ kind: 'insertText', payload: { text: String(index) } }));
    const admitted = intents.map((intent) => h.submit(intent));
    let acknowledgements = 0;
    for (const { ticket } of admitted) void ticket.ack.then(() => { acknowledgements += 1; });
    h.pump();
    expect(acknowledgements).toBe(0);
    expect(h.calls.map(({ envelope }) => ({ kind: envelope.kind, payload: envelope.payload }))).toEqual(intents);
    expect(h.calls.map(({ envelope }) => envelope.seq)).toEqual(intents.map((_, index) => index + 1));
    expect(h.calls.map(({ envelope }) => envelope.afterSeq)).toEqual(intents.map((_, index) => index || undefined));
    if (cap === 'pending-intents') {
      expect(admitted.slice(0, 256).every(({ echo }) => echo !== undefined)).toBe(true);
      expect(admitted.slice(256).every(({ echo }) => echo === undefined)).toBe(true);
      expect(texts.length).toBe(256);
    } else {
      expect(admitted.every(({ echo }) => echo === undefined)).toBe(true);
      expect(texts).toEqual([]);
    }
    h.acknowledge();
    expect(await Promise.all(admitted.map(({ ticket }) => ticket.ack))).toEqual(
      intents.map((_, index) => ({ seq: index + 1, status: 'applied' })));
    expect(acknowledgements).toBe(intents.length);
    expect(h.session.state.throughSeq).toBe(intents.length);
    await h.session.dispose();
  }
});

test('d2c_echo_cap_counts_original_intents_in_merged_drafts', () => {
  const { painter, texts } = setup();
  const value: PeerlessEchoIntent = { intent: { kind: 'insertText', payload: { text: 'x' } },
    originalIntentCount: 256, contextFrameEpoch: 1 };
  painter.paint([entry(1, value)]);
  expect(texts.length).toBeGreaterThan(0);
  const painted = texts.length;
  painter.paint([entry(1, { ...value, originalIntentCount: 257 })]);
  expect(texts.length).toBe(painted);
  expect(painter.preview({ kind: 'insertText', payload: { text: 'last' } })).toBeUndefined();
});

test('d2c_echo_deletes_code_points_in_combining_text_and_surrogate_pairs', () => {
  for (const [kind, text, grapheme, expected] of [
    ['deleteBackward', 'e\u0301', 1, 'e'],
    ['deleteForward', 'e\u0301', 0, '\u0301'],
    ['deleteBackward', '\u{1f600}e\u0301', 1, 'e\u0301'],
    ['deleteForward', '\u{1f600}e\u0301', 0, 'e\u0301'],
  ] as const) {
    const { painter, texts } = setup();
    const boundaries = [...new Intl.Segmenter(undefined, { granularity: 'grapheme' }).segment(text)]
      .map((segment) => segment.index).concat(text.length);
    painter.installContext({ ...context(), text,
      boundaries: boundaries.map((utf16, grapheme) => ({ utf16, grapheme, caret: rect })),
      selection: { anchorGrapheme: grapheme, focusGrapheme: grapheme },
    });
    const preview = painter.preview({ kind, payload: {} });
    expect(preview).toBeDefined();
    painter.paint([entry(1, preview!)]);
    expect(texts.at(-1)).toBe(expected);
  }
});

test('d2c_cross_paragraph_delete_declines_echo_and_keeps_submitted_input', async () => {
  for (const [kind, offset, count] of [
    ['deleteBackward', 0, 1], ['deleteForward', 2, 1],
    ['deleteBackward', 2, 3], ['deleteForward', 0, 3],
  ] as const) {
    const { painter, texts } = setup();
    painter.installContext({ ...context(), selection: { anchorGrapheme: offset, focusGrapheme: offset } });
    const h = sessionWithHeldReplies(painter);
    const intent: PeerlessTextIntent = { kind, payload: { count } };
    const admitted = h.submit(intent);
    expect(admitted.echo).toBeUndefined();
    expect(texts).toEqual([]);
    h.pump();
    expect(h.calls.map(({ envelope }) => ({ kind: envelope.kind, payload: envelope.payload }))).toEqual([intent]);
    h.acknowledge();
    expect(await admitted.ticket.ack).toEqual({ seq: 1, status: 'applied' });
    await h.session.dispose();
  }
});

test('d2c_delete_after_pending_deletion_declines_cross_paragraph_echo', async () => {
  const { painter, texts, draws, bitmap } = setup();
  const h = sessionWithHeldReplies(painter);
  const first = h.submit({ kind: 'deleteBackward', payload: { count: 2 } });
  expect(first.echo).toBeDefined();
  expect(texts).toEqual(['']);
  const second = h.submit({ kind: 'deleteBackward', payload: {} });
  expect(second.echo).toBeUndefined();
  expect(texts).toEqual(['']);
  expect(draws.at(-1)?.[0]).toBe(bitmap);
  h.pump();
  expect(h.calls.map(({ envelope }) => ({ kind: envelope.kind, payload: envelope.payload }))).toEqual([
    { kind: 'deleteBackward', payload: { count: 2 } }, { kind: 'deleteBackward', payload: {} },
  ]);
  expect(h.calls.map(({ envelope }) => envelope.afterSeq)).toEqual([undefined, 1]);
  h.acknowledge();
  expect(await Promise.all([first.ticket.ack, second.ticket.ack])).toEqual([
    { seq: 1, status: 'applied' }, { seq: 2, status: 'applied' },
  ]);
  await h.session.dispose();
});

test('d2c_damage_restore_preserves_shading_and_media', () => {
  const { painter, draws, texts, background, bitmap } = setup();
  expect(painter.restoreBackground(page, 'background', rect)).toBe(true);
  expect(draws.at(-1)?.[0]).toBe(background);
  const preview = painter.preview({ kind: 'insertText', payload: { text: 'c' } });
  expect(preview).toBeDefined();
  painter.paint([entry(1, preview!)]);
  expect(draws.at(-1)?.[0]).toBe(background);
  expect(texts.at(-1)).toBe('abc');
  const deletion = painter.preview({ kind: 'deleteBackward', payload: {} });
  expect(deletion).toBeDefined();
  painter.paint([entry(1, preview!), entry(2, deletion!)]);
  expect(texts.at(-1)).toBe('ab');
  painter.drop([entry(1, preview!)], 'superseded');
  expect(draws.at(-1)?.[0]).toBe(bitmap);
  expect(painter.restoreBackground(page, 'missing', rect)).toBe(false);
});

test('d2c_other_page_partial_frame_retains_echo_restoration_until_refusal', () => {
  for (const reason of ['refused', 'superseded']) {
    const { painter, draws, texts, bitmap, background, frame } = setup();
    const closed: ImageBitmap[] = [];
    bitmap.close = () => { closed.push(bitmap); };
    background.close = () => { closed.push(background); };
    const preview = painter.preview({ kind: 'insertText', payload: { text: 'c' } });
    expect(preview).toBeDefined();
    const pending = entry(1, preview!);
    painter.paint([pending]);
    expect(texts.at(-1)).toBe('abc');
    const other = { width: 100, height: 20, close() { closed.push(other); } } as ImageBitmap;
    painter.reconcile({ ...frame, revision: 2, payload: { ...frame.payload, echoContext: undefined,
      caret: { frameEpoch: 2, caretRect: null }, layoutRevision: 2,
      stagedPresentation: { frameEpoch: 2, pages: [
        { page: { pageId: '2', pageIndex: 1 }, frameEpoch: 2, bitmap: other },
      ] },
    } }, [], [pending]);
    expect(closed).toEqual([]);
    expect(texts.at(-1)).toBe('abc');
    expect(painter.preview({ kind: 'insertText', payload: { text: 'd' } })).toBeDefined();
    painter.drop([pending], reason);
    expect(draws.at(-1)?.[0]).toBe(bitmap);
    expect(painter.preview({ kind: 'insertText', payload: { text: 'e' } })).toBeUndefined();
    painter.dispose();
    expect(closed).toEqual([bitmap, background, other]);
  }
});

test('d2c_viewport_eviction_bounds_bitmaps_and_preserves_active_echo_on_refusal', () => {
  const { painter, frame, bitmap, background, draws, mountPages } = setup();
  const mounted = new Set([page.pageId]);
  mountPages(mounted);
  const closes = new Map<ImageBitmap, number>();
  const resources: ImageBitmap[] = [bitmap, background];
  for (const value of resources) value.close = () => { closes.set(value, (closes.get(value) ?? 0) + 1); };
  const pending = entry(1, painter.preview({ kind: 'insertText', payload: { text: 'c' } })!);
  painter.paint([pending]);
  for (let revision = 2; revision <= 30; revision++) {
    mounted.clear();
    mounted.add(page.pageId);
    mounted.add(String(revision));
    const value = { width: 100, height: 20,
      close() { closes.set(value, (closes.get(value) ?? 0) + 1); } } as ImageBitmap;
    resources.push(value);
    painter.reconcile({ ...frame, revision, payload: { ...frame.payload, echoContext: undefined,
      stagedPresentation: { frameEpoch: revision, pages: [
        { page: { pageId: String(revision), pageIndex: revision - 1 }, frameEpoch: revision, bitmap: value },
      ] },
    } }, [], [pending]);
    expect(resources.filter((resource) => !closes.has(resource)).length).toBe(3);
    expect(closes.has(bitmap)).toBe(false);
    expect(closes.has(background)).toBe(false);
  }
  painter.drop([pending], 'refused');
  expect(draws.at(-1)?.[0]).toBe(bitmap);
  mounted.clear();
  painter.paint([]);
  painter.dispose();
  painter.dispose();
  expect(resources.map((value) => closes.get(value))).toEqual(resources.map(() => 1));
});

test('d2c_fractional_raster_scale_uses_page_dimensions_for_crop_and_text', () => {
  for (const [scale, zoom] of [[1.25, 1], [1.25, 1.1], [2, 1], [3, 1]] as const) {
    const { painter, frame, canvas, bitmap, background, draws, transforms } = setup();
    const damage = { x: 40, y: 400, width: 100, height: 21 };
    const next = context();
    next.frameEpoch = 2;
    next.damage = [{ ...next.damage[0]!, rect: damage }];
    next.styleRuns = [{ ...next.styleRuns[0]!, baseline: { x: 40, y: 414 } }];
    const raster = { width: 800 * scale, height: 1000 * scale, close: bitmap.close } as ImageBitmap;
    const crop = { width: Math.ceil(100 * scale), height: Math.ceil(21 * scale), close: background.close } as ImageBitmap;
    painter.reconcile({ ...frame, revision: 2, payload: { ...frame.payload, echoContext: next,
      layoutMeta: { v: 1, layoutRevision: 2, pageCount: 1, partial: false, provisional: false,
        notesConverged: true, pageSizes: new Float64Array([800, 1000]), layoutShell: '{}', headersFootersEpoch: 0 },
      stagedPresentation: { frameEpoch: 2, pages: [
        { page, frameEpoch: 2, bitmap: raster, backgrounds: [{ id: 'background', rect: damage, bitmap: crop }] },
      ] },
    } }, [], []);
    canvas.width = raster.width * zoom;
    canvas.height = raster.height * zoom;
    const pending = entry(1, painter.preview({ kind: 'insertText', payload: { text: 'c' } })!);
    painter.paint([pending]);
    expect(draws.at(-1)).toEqual([crop, 0, 0, crop.width, 21 * (crop.height / 21), 40, 400, 100, 21]);
    expect(transforms.at(-1)).toEqual([scale * zoom, 0, 0, scale * zoom, 0, 0]);
    expect(transforms.at(-2)).toEqual(transforms.at(-1));
    expect(transforms.at(-1)![3]! * damage.y).toBe(400 * scale * zoom);
  }
});

test('d2c_fractional_raster_scale_restores_partial_background_coordinates', () => {
  const { painter, frame, canvas, draws, transforms } = setup();
  const damage = { x: 40, y: 400, width: 100, height: 21 };
  const raster = { width: 1000, height: 1250, close() {} } as ImageBitmap;
  const crop = { width: 125, height: 27, close() {} } as ImageBitmap;
  painter.reconcile({ ...frame, revision: 2, payload: { ...frame.payload, echoContext: undefined,
    layoutMeta: { v: 1, layoutRevision: 2, pageCount: 1, partial: false, provisional: false,
      notesConverged: true, pageSizes: new Float64Array([800, 1000]), layoutShell: '{}', headersFootersEpoch: 0 },
    stagedPresentation: { frameEpoch: 2, pages: [
      { page, frameEpoch: 2, bitmap: raster, backgrounds: [{ id: 'background', rect: damage, bitmap: crop }] },
    ] },
  } }, [], []);
  canvas.width = 1100;
  canvas.height = 1375;
  expect(painter.restoreBackground(page, 'background', { x: 60, y: 407, width: 40, height: 7 })).toBe(true);
  expect(draws.at(-1)).toEqual([crop, 25, 9, 50, 9, 60, 407, 40, 7]);
  expect(transforms.at(-1)).toEqual([1.375, 0, 0, 1.375, 0, 0]);
  const restored = draws.length;
  expect(painter.restoreBackground(page, 'background', { ...damage, height: 22 })).toBe(false);
  expect(draws).toHaveLength(restored);
});
