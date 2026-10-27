import { expect, test } from 'bun:test';
import type {
  CommandAck, CommandConnection, CommandEnvelope, CommandTicket, EchoEntry,
} from '../../../../../../shared/office-session/commands/types';
import { createWorkerEditSession } from '../../../../../../shared/office-session/commands/session';
import type { DocxCommandEchoContext, DocxCommandFrame, DocxEchoGlyphSet } from '@betteroffice/docx/worker';
import { createEchoPainter, type PeerlessEchoIntent, type PeerlessEchoPainter } from './echoPainter';
import type { PeerlessTextIntent } from './inputEvents';
import { testEchoGlyphLine, testEchoGlyphSet } from './echoGlyphs.testing';

const page = { pageId: '1', pageIndex: 0 };
const rect = { x: 0, y: 0, width: 100, height: 20 };
const alphabet = 'abcdex0123456789e\u0301\u{1f600}';
const glyphSet: DocxEchoGlyphSet = { key: 'set', fontId: 1, alphabet: [...new Set(alphabet)].join(''), horizontal: 0.0001,
  glyphIds: [...new Set(alphabet)].map((char) => char.codePointAt(0)!), advances: [...new Set(alphabet)].map(() => 1000),
  kerning: [], unsafePairs: [], outlines: Object.fromEntries([...new Set(alphabet)].map((char) => [char.codePointAt(0)!,
    JSON.stringify({ upem: 1000, cmds: [{ t: 'M', x: char.codePointAt(0)!, y: 0 }, { t: 'L', x: 0, y: 700 }, { t: 'Z' }] })])) };

function context(): DocxCommandEchoContext {
  return {
    page, frameEpoch: 1, layoutRevision: 1, serializedBytes: 900, text: 'ab',
    styleRuns: [{ startUtf16: 0, endUtf16: 2, fontFamily: 'Arial', fontSizePx: 12,
      color: '#000', bold: false, italic: false, underline: false, strike: false,
      transform: [1, 0, 0, 1, 0, 0], baseline: { x: 0, y: 14 } }],
    boundaries: [0, 1, 2].map((utf16) => ({ utf16, grapheme: utf16,
      caret: { x: utf16 * 10, y: 0, width: 1, height: 20 } })),
    selection: { anchorGrapheme: 2, focusGrapheme: 2 }, transform: [1, 0, 0, 1, 0, 0],
    damage: [{ page, rect, backgroundId: 'background', omittedPrimitiveIds: ['shading', 'media'] }],
    shaping: 'simple-ltr',
    glyphs: { glyphSet: 'set', fontId: 1, size: 12, color: '#000', x: 0, y: 14, next: '', right: 100,
      device: { x: 0, y: 0, width: 100, height: 20 }, pixelRatio: 1 },
  };
}

function entry(id: number, value: PeerlessEchoIntent): EchoEntry<PeerlessEchoIntent> {
  return { seq: id, value, ticket: { id, seq: id,
    ack: Promise.resolve({ status: 'applied', seq: id }), presented: Promise.resolve('presented') } };
}

const scope = globalThis as { Path2D?: unknown };
scope.Path2D = class { glyph?: number; moveTo(x: number) { this.glyph ??= x; } lineTo() {} closePath() {} };

function setup(options: { requestFullFrame?(): void; onPaint?(): void } = {}) {
  const draws: unknown[][] = [];
  const texts: string[] = [];
  const transforms: number[][] = [];
  const bitmap = { width: 100, height: 20, close() {} } as ImageBitmap;
  const background = { width: 100, height: 20, close() {} } as ImageBitmap;
  const ctx = {
    save() {}, restore() {}, setTransform(...values: number[]) { transforms.push(values); },
    transform() {}, beginPath() {}, rect() {}, clip() {}, translate() {}, scale() {},
    clearRect() {}, fillRect() {}, fillText() { throw new Error('the echo paints glyph outlines only'); },
    fill(path: { glyph: number }) { texts.push(texts.pop()! + String.fromCodePoint(path.glyph)); },
    drawImage(...args: unknown[]) { draws.push(args); if (args.length === 3) texts.push(''); },
  };
  const canvas = { width: 100, height: 20, getContext: () => ctx } as unknown as HTMLCanvasElement;
  let mountedPages: ReadonlySet<string> | undefined;
  const painter = createEchoPainter({ ...options, canvasForPage: (item) =>
    mountedPages === undefined || mountedPages.has(item.pageId) ? canvas : null });
  const frame: DocxCommandFrame = { coversThroughSeq: 0, revision: 1, payload: {
    frame: new ArrayBuffer(0), caret: { frameEpoch: 1, caretRect: null }, layoutRevision: 1,
    echoContext: context(), echoGlyphSets: [glyphSet], stagedPresentation: { frameEpoch: 1, pages: [
      { page, frameEpoch: 1, bitmap, backgrounds: [{ id: 'background', rect, bitmap: background }] },
    ] },
  } };
  painter.reconcile(frame, [], []);
  const paints = () => draws.filter(([image]) => image === background).length;
  return { painter, draws, texts, transforms, canvas, background, bitmap, frame, paints,
    mountPages(pages: ReadonlySet<string>) { mountedPages = pages; },
  };
}

test('ack contexts require the staged page epoch and paint with its retained background', () => {
  const painted: number[] = [];
  const { painter, paints } = setup({ onPaint: () => { painted.push(1); } });
  const next = { ...context(), selection: { anchorGrapheme: 1, focusGrapheme: 1 } };
  expect(painter.installAckContext!({ ...next, frameEpoch: 2 }, 1)).toBe(false);
  expect(painter.installAckContext!({ ...next, page: { pageId: 'other', pageIndex: 0 } }, 1)).toBe(false);
  expect(painter.installAckContext!(next, 1)).toBe(true);
  expect(painter.contextThroughSeq).toBe(1);
  const preview = painter.preview({ kind: 'insertText', payload: { text: 'x' } });
  expect(preview).toBeDefined();
  painter.paint([entry(2, preview!)]);
  expect(paints()).toBe(1);
  expect(painted).toEqual([1]);
});

test('echo paints only lines it can place exactly and asks once for a glyph set it lacks', () => {
  const requests: number[] = [];
  const { painter, texts, frame } = setup({ requestFullFrame: () => { requests.push(1); } });
  const { glyphs: _, ...inexact } = context();
  painter.installContext(inexact);
  expect(painter.preview({ kind: 'insertText', payload: { text: 'x' } })).toBeUndefined();
  painter.installContext({ ...context(), glyphs: { ...context().glyphs!, glyphSet: 'missing' } });
  painter.installContext({ ...context(), glyphs: { ...context().glyphs!, glyphSet: 'missing' } });
  expect(painter.preview({ kind: 'insertText', payload: { text: 'x' } })).toBeUndefined();
  expect(requests).toHaveLength(1);
  painter.reconcile({ ...frame, revision: 2, payload: { ...frame.payload, echoGlyphSets: [{ ...glyphSet, key: 'missing' }],
    echoContext: { ...context(), frameEpoch: 2, glyphs: { ...context().glyphs!, glyphSet: 'missing' } },
    stagedPresentation: { frameEpoch: 2, pages: frame.payload.stagedPresentation!.pages.map((item) => ({ ...item, frameEpoch: 2 })) },
  } }, [], []);
  expect(painter.preview({ kind: 'insertText', payload: { text: 'x' } })).toBeDefined();
  expect(texts.filter((text) => text.length)).toEqual([]);
});

test('echo declines a line that would wrap or pull a word up from the next line', () => {
  const { painter } = setup();
  painter.installContext({ ...context(), glyphs: { ...context().glyphs!, right: 0.0003 } });
  expect(painter.preview({ kind: 'insertText', payload: { text: 'x' } })).toBeUndefined();
  painter.installContext({ ...context(), glyphs: { ...context().glyphs!, nextWord: 101 } });
  expect(painter.preview({ kind: 'deleteBackward', payload: {} })).toBeDefined();
  painter.installContext({ ...context(), glyphs: { ...context().glyphs!, nextWord: 100.5 } });
  expect(painter.preview({ kind: 'deleteBackward', payload: {} })).toBeUndefined();
  painter.installContext({ ...context(), glyphs: { ...context().glyphs!, previousSlack: 5 } });
  expect(painter.preview({ kind: 'deleteBackward', payload: {} })).toBeUndefined();
});

test('echo declines deleting the sole 24pt glyph above a 10pt paragraph mark', () => {
  const { painter } = setup();
  const line = { ...context(), text: 'a', selection: { anchorGrapheme: 1, focusGrapheme: 1 },
    boundaries: context().boundaries.slice(0, 2),
    glyphs: Object.assign({ ...context().glyphs!, size: 32 }, { canEmpty: false }) };
  painter.installContext(line);
  expect(painter.preview({ kind: 'deleteBackward', payload: {} })).toBeUndefined();
});

test('a backspace that lets the line rewrap onto the previous one declines', () => {
  const { painter } = setup();
  const glyphs = { glyphSet: 'set', fontId: 1, size: 12, color: '#000', x: 0, y: 14, next: '', right: 100,
    device: { x: 0, y: 0, width: 100, height: 20 }, pixelRatio: 1 };
  const line = { ...context(), text: 'cde', selection: { anchorGrapheme: 2, focusGrapheme: 2 },
    boundaries: [0, 1, 2, 3].map((utf16) => ({ utf16, grapheme: utf16, caret: { x: utf16 * 10, y: 0, width: 1, height: 20 } })) };
  painter.installContext({ ...line, glyphs });
  expect(painter.preview({ kind: 'deleteBackward', payload: {} })).toBeDefined();
  painter.installContext({ ...line, glyphs: { ...glyphs, previousSlack: 0.2 } });
  expect(painter.preview({ kind: 'deleteBackward', payload: {} })).toBeUndefined();
});

function wrappedSetup(offset: number, focus = offset, text = 'aaa bb ', next = 'c') {
  const result = setup();
  const echo: DocxCommandEchoContext = { ...context(), frameEpoch: 2, text,
    boundaries: Array.from({ length: text.length + 1 }, (_, utf16) => ({ utf16, grapheme: utf16, caret: rect })),
    selection: { anchorGrapheme: offset, focusGrapheme: focus },
    glyphs: testEchoGlyphLine({ right: 100, next, nextWord: 80 }),
  };
  result.painter.reconcile({ ...result.frame, revision: 2, payload: { ...result.frame.payload,
    echoContext: echo, echoGlyphSets: [testEchoGlyphSet], stagedPresentation: { frameEpoch: 2,
      pages: result.frame.payload.stagedPresentation!.pages.map((item) => ({ ...item, frameEpoch: 2 })),
    },
  } }, [], []);
  return result;
}

test('echo declines deleting the space separating a soft wrapped next word', () => {
  const { painter, texts } = wrappedSetup(6);
  expect(painter.preview({ kind: 'deleteForward', payload: {} })).toBeUndefined();
  expect(texts).toEqual([]);
});

test('echo declines replacing or backspacing the soft wrap separator', () => {
  for (const [offset, focus, intent] of [
    [7, 7, { kind: 'deleteBackward', payload: {} }],
    [6, 7, { kind: 'replaceSelection', payload: { text: 'x' } }],
    [6, 7, { kind: 'replaceSelection', payload: { text: '' } }],
  ] satisfies [number, number, PeerlessTextIntent][]) {
    const { painter } = wrappedSetup(offset, focus);
    expect(painter.preview(intent)).toBeUndefined();
  }
});

test('echo declines appending a word that joins the next line first word', () => {
  const { painter } = wrappedSetup(7);
  expect(painter.preview({ kind: 'insertText', payload: { text: 'x' } })).toBeUndefined();
});

test('echo declines changing a break opportunity adjacent to the soft wrap', () => {
  const { painter } = wrappedSetup(7, 7, 'aaa bb  ');
  expect(painter.preview({ kind: 'deleteForward', payload: {} })).toBeUndefined();
});

test('echo keeps mid line insertions and deletions with a stable soft wrap separator', () => {
  for (const [offset, intent, expected] of [
    [1, { kind: 'insertText', payload: { text: 'a' } }, 'aaaa bb '],
    [2, { kind: 'deleteBackward', payload: {} }, 'aa bb '],
  ] satisfies [number, PeerlessTextIntent, string][]) {
    const { painter, texts } = wrappedSetup(offset);
    const preview = painter.preview(intent);
    expect(preview).toBeDefined();
    painter.paint([entry(1, preview!)]);
    expect(texts.at(-1)).toBe(expected);
  }
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
    const { painter, paints } = setup();
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
      expect(paints()).toBe(256);
    } else {
      expect(admitted.every(({ echo }) => echo === undefined)).toBe(true);
      expect(paints()).toBe(0);
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
  const { painter, paints } = setup();
  const value: PeerlessEchoIntent = { intent: { kind: 'insertText', payload: { text: 'x' } },
    originalIntentCount: 256, contextFrameEpoch: 1 };
  painter.paint([entry(1, value)]);
  expect(paints()).toBeGreaterThan(0);
  const painted = paints();
  painter.paint([entry(1, { ...value, originalIntentCount: 257 })]);
  expect(paints()).toBe(painted);
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

test('echo reuses the preview placement and the frame baseline for each key', () => {
  const { painter, texts } = setup();
  const get = Map.prototype.get;
  let characters = 0;
  Map.prototype.get = function (key) {
    if (typeof key === 'string' && key.length === 1) characters += 1;
    return get.call(this, key);
  };
  try {
    const value = painter.preview({ kind: 'insertText', payload: { text: 'x' } });
    expect(value).toBeDefined();
    painter.paint([entry(1, value!)]);
    expect(texts.at(-1)).toBe('abx');
  } finally { Map.prototype.get = get; }
  expect(characters).toBe(3);
});

test('echo submits one glyph run for the whole draft line', () => {
  const { painter, canvas, texts } = setup();
  let styles = 0;
  Object.defineProperty(canvas.getContext('2d')!, 'fillStyle', { configurable: true, set() { styles += 1; } });
  const value = painter.preview({ kind: 'insertText', payload: { text: 'x' } });
  expect(value).toBeDefined();
  painter.paint([entry(1, value!)]);
  expect(texts.at(-1)).toBe('abx');
  expect(styles).toBe(1);
});

test('echo key paints restore the damage band once and drop restores only that band', () => {
  const { painter, canvas, bitmap, background, draws } = setup();
  canvas.width = 200;
  canvas.height = 100;
  Object.assign(bitmap, { width: 200, height: 100 });
  const first = entry(1, painter.preview({ kind: 'insertText', payload: { text: 'c' } })!);
  painter.paint([first]);
  const second = entry(2, painter.preview({ kind: 'insertText', payload: { text: 'd' } })!);
  painter.paint([first, second]);
  expect(draws.filter(([image]) => image === background)).toHaveLength(2);
  expect(draws.filter(([image]) => image === bitmap)).toHaveLength(0);
  painter.drop([first, second], 'refused');
  expect(draws.at(-1)).toEqual([bitmap, 0, 0, 100, 20, 0, 0, 100, 20]);
});

test('echo painting allocates no encoders for pending input', () => {
  const { painter, texts } = setup();
  const Encoder = globalThis.TextEncoder;
  let constructors = 0;
  globalThis.TextEncoder = class extends Encoder { constructor() { super(); constructors += 1; } };
  try {
    const value = painter.preview({ kind: 'insertText', payload: { text: 'x' } });
    expect(value).toBeDefined();
    painter.paint([entry(1, value!)]);
    expect(texts.at(-1)).toBe('abx');
  } finally { globalThis.TextEncoder = Encoder; }
  expect(constructors).toBe(0);
});

function scaledFrame(frame: DocxCommandFrame, echoContext: DocxCommandEchoContext | undefined, scale = 1.25) {
  const raster = { width: 800 * scale, height: 1000 * scale, close() {} } as ImageBitmap;
  const crop = { width: Math.ceil(100 * scale), height: Math.ceil(21.6 * scale), close() {} } as ImageBitmap;
  const damage = { x: 40, y: 400, width: 100, height: 21.6 };
  const next: DocxCommandFrame = { ...frame, revision: 2, payload: { ...frame.payload, echoContext,
    stagedPresentation: { frameEpoch: 2, pages: [
      { page, frameEpoch: 2, bitmap: raster, backgrounds: [{ id: 'background', rect: damage, bitmap: crop }] },
    ] },
  } };
  return { next, raster, crop, damage };
}

test('echo restores and paints in raster device pixels and declines on a resized canvas', () => {
  for (const [scale, zoom] of [[1.25, 1], [1.25, 1.1], [2, 1], [3, 1]] as const) {
    const { painter, frame, canvas, draws, transforms, texts } = setup();
    const echo = context();
    echo.frameEpoch = 2;
    echo.damage = [{ ...echo.damage[0]!, rect: { x: 40, y: 400, width: 100, height: 21.6 } }];
    const device = { x: 40 * scale, y: 400 * scale, width: Math.ceil(100 * scale), height: Math.ceil(21.6 * scale) };
    echo.glyphs = { ...echo.glyphs!, x: 40, y: 414, right: 140, device, pixelRatio: scale };
    const { next, raster, crop } = scaledFrame(frame, echo, scale);
    painter.reconcile(next, [], []);
    canvas.width = raster.width * zoom;
    canvas.height = raster.height * zoom;
    const pending = entry(1, painter.preview({ kind: 'insertText', payload: { text: 'c' } })!);
    painter.paint([pending]);
    if (zoom === 1) {
      expect(draws.filter(([image]) => image === crop).at(-1)).toEqual([crop, device.x, device.y]);
      expect(transforms.slice(-2)).toEqual([[1, 0, 0, 1, 0, 0], [scale, 0, 0, scale, 0, 0]]);
      expect(texts.at(-1)).toBe('abc');
    } else {
      expect(draws.some(([image]) => image === crop)).toBe(false);
      expect(texts.filter(Boolean)).toEqual([]);
    }
  }
});

test('background restore copies the whole staged crop at its device pixels', () => {
  const { painter, frame, canvas, draws, transforms } = setup();
  const { next, crop, damage } = scaledFrame(frame, undefined);
  painter.reconcile(next, [], []);
  canvas.width = 1000;
  canvas.height = 1250;
  expect(painter.restoreBackground(page, 'background', { x: 60, y: 407, width: 40, height: 7 })).toBe(true);
  expect(draws.at(-1)).toEqual([crop, 50, 500]);
  expect(transforms.at(-1)).toEqual([1, 0, 0, 1, 0, 0]);
  const restored = draws.length;
  expect(painter.restoreBackground(page, 'background', { ...damage, height: 22 })).toBe(false);
  canvas.width = 1100;
  expect(painter.restoreBackground(page, 'background', damage)).toBe(false);
  expect(draws).toHaveLength(restored);
});

test('echo recenters a centered line on every key and declines a trailing space', () => {
  const { painter, canvas } = setup();
  const xs: number[] = [];
  Object.assign(canvas.getContext('2d')!, { translate(x: number) { xs.push(x); } });
  const centered = { ...context(), glyphs: { ...context().glyphs!, center: 100 } };
  painter.installContext(centered);
  const typed = painter.preview({ kind: 'insertText', payload: { text: 'x' } });
  expect(typed).toBeDefined();
  painter.paint([entry(1, typed!)]);
  expect(xs).toEqual([49.85, 49.95, 50.05]);
  painter.installContext(centered);
  expect(painter.preview({ kind: 'insertText', payload: { text: ' ' } })).toBeUndefined();
});

test('echo declines an edit that brings a ligature pair together on either side of the caret', () => {
  const { painter, frame } = setup();
  const chars = [...glyphSet.alphabet];
  const ligating = { ...glyphSet, key: 'ligating', unsafePairs: [[chars.indexOf('a'), chars.indexOf('b')] as const] };
  const at = (text: string, caret: number): DocxCommandEchoContext => ({ ...context(), frameEpoch: 2, text,
    boundaries: [...Array(text.length + 1).keys()].map((utf16) => ({ utf16, grapheme: utf16,
      caret: { x: utf16 * 10, y: 0, width: 1, height: 20 } })),
    selection: { anchorGrapheme: caret, focusGrapheme: caret }, glyphs: { ...context().glyphs!, glyphSet: 'ligating' } });
  painter.reconcile({ ...frame, revision: 2, payload: { ...frame.payload, echoGlyphSets: [ligating], echoContext: at('a', 1),
    stagedPresentation: { frameEpoch: 2, pages: frame.payload.stagedPresentation!.pages.map((item) => ({ ...item, frameEpoch: 2 })) },
  } }, [], []);
  expect(painter.preview({ kind: 'insertText', payload: { text: 'c' } }), 'a pair that cannot ligate paints').toBeDefined();
  painter.installContext(at('a', 1));
  expect(painter.preview({ kind: 'insertText', payload: { text: 'b' } }), 'the typed glyph would ligate with its left neighbour').toBeUndefined();
  painter.installContext(at('b', 0));
  expect(painter.preview({ kind: 'insertText', payload: { text: 'a' } }), 'the typed glyph would ligate with its right neighbour').toBeUndefined();
  painter.installContext(at('acb', 2));
  expect(painter.preview({ kind: 'deleteBackward', payload: {} }), 'a deletion would join a ligature pair').toBeUndefined();
  painter.installContext(at('acb', 1));
  expect(painter.preview({ kind: 'deleteForward', payload: {} }), 'a forward deletion would join a ligature pair').toBeUndefined();
});

test('echo declines a new first glyph on a continuation line, which shapes with the line above', () => {
  const { painter } = setup();
  const continuation = { ...context(), glyphs: { ...context().glyphs!, previousSlack: -10 } };
  painter.installContext({ ...continuation, selection: { anchorGrapheme: 0, focusGrapheme: 0 } });
  expect(painter.preview({ kind: 'insertText', payload: { text: 'c' } })).toBeUndefined();
  painter.installContext({ ...continuation, selection: { anchorGrapheme: 1, focusGrapheme: 1 } });
  expect(painter.preview({ kind: 'deleteBackward', payload: {} })).toBeUndefined();
  painter.installContext(continuation);
  expect(painter.preview({ kind: 'insertText', payload: { text: 'c' } }), 'an edit after the first glyph still paints').toBeDefined();
});

test('echo declines ink that would end past the line even when the line does not grow', () => {
  const { painter, frame } = setup();
  // ten pixel glyphs, with a space the shared alphabet lacks
  const wide: DocxEchoGlyphSet = { ...glyphSet, key: 'wide', horizontal: 0.01, alphabet: glyphSet.alphabet + ' ',
    glyphIds: [...glyphSet.glyphIds, 32], advances: [...glyphSet.advances, 1000],
    outlines: { ...glyphSet.outlines, 32: JSON.stringify({ upem: 1000, cmds: [{ t: 'M', x: 32, y: 0 }, { t: 'L', x: 0, y: 700 }, { t: 'Z' }] }) } };
  const text = 'ab  ';
  const hanging: DocxCommandEchoContext = { ...context(), frameEpoch: 2, text,
    boundaries: [...Array(text.length + 1).keys()].map((utf16) => ({ utf16, grapheme: utf16,
      caret: { x: utf16 * 10, y: 0, width: 1, height: 20 } })),
    selection: { anchorGrapheme: 3, focusGrapheme: 4 }, glyphs: { ...context().glyphs!, glyphSet: 'wide', right: 25 } };
  painter.reconcile({ ...frame, revision: 2, payload: { ...frame.payload, echoGlyphSets: [wide], echoContext: hanging,
    stagedPresentation: { frameEpoch: 2, pages: frame.payload.stagedPresentation!.pages.map((item) => ({ ...item, frameEpoch: 2 })) },
  } }, [], []);
  expect(painter.preview({ kind: 'replaceSelection', payload: { text: 'c' } }), 'a trailing space replaced by ink past the line').toBeUndefined();
  painter.installContext({ ...hanging, selection: { anchorGrapheme: 2, focusGrapheme: 2 } });
  expect(painter.preview({ kind: 'deleteForward', payload: {} }), 'removing a hanging space keeps the ink inside').toBeDefined();
});
