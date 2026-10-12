import { beforeAll, describe, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import type { WorkbookAck, WorkbookCommand, WorkbookCommandFrame } from '@betteroffice/xlsx';
import { initWasm } from '@betteroffice/xlsx';
import { isClientMessage, isHostMessage, SessionFailure } from '../../../../shared/office-session';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import { createWorkbookSession } from '../../../xlsx/src/session/client';
import { createWorkbookSessionHost } from '../../../xlsx/src/session/host';
import { WorkerCellEcho } from './WorkerCellEcho';
import { PeerlessWorkerPaintSource as WorkerPaintSource, type WorkerPaintRequest, type WorkerPaintResult } from './PeerlessWorkerPaintSource';
import { XlsxWorkerSessionFailedError } from './useEditableSessionWorkbook';
import { PeerlessWorkbookSession } from './usePeerlessSessionWorkbook';

let bytes: Uint8Array;

beforeAll(async () => {
  await initWasm(await readFile(new URL('../../../xlsx/src/wasm/generated/xlsx_wasm_bg.wasm', import.meta.url)));
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>');
  zip.file('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>');
  zip.file('xl/worksheets/sheet1.xml', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>');
  bytes = await zip.generateAsync({ type: 'uint8array' });
});

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}

async function harness(maxPatches = 8) {
  const pair = createInProcessPair();
  createWorkbookSessionHost(pair.host);
  const requests: string[] = [];
  const sequences = new Map<number, number>();
  const replies = new Map<number, () => void>();
  const arrivals = new Map<number, ReturnType<typeof deferred<void>>>();
  let crash: ((error: unknown) => void) | undefined;
  const session = await createWorkbookSession(bytes, {
    includeEditState: true,
    initialViewport: { width: 384, height: 96, zoom: 1, dpr: 1, requestGeneration: 0 },
  }, {
    ...pair.client,
    post(message, transfer) {
      if (isClientMessage(message) && message.kind === 'call') {
        requests.push(message.method);
        const input = message.args[0];
        if (input && typeof input === 'object' && 'seq' in input && typeof input.seq === 'number') {
          sequences.set(message.id, input.seq);
        }
      }
      pair.client.post(message, transfer);
    },
    listen: (listener) => pair.client.listen((message) => {
      if (isHostMessage(message) && message.kind === 'reply' && sequences.has(message.id)) {
        const seq = sequences.get(message.id)!;
        replies.set(seq, () => listener(message));
        arrivals.get(seq)?.resolve();
      } else listener(message);
    }),
    onError(listener) { crash = listener; return pair.client.onError(listener); },
  });
  const text: string[] = [];
  const events: string[] = [];
  const paints: WorkerPaintResult[] = [];
  const errors: Error[] = [];
  const acks: WorkbookAck[] = [];
  const captures = new Map<number, WorkerPaintRequest>();
  const state = { sent: 0, installReplies: true, throwCanvas: false };
  const view: WorkerPaintRequest = {
    generation: 1, sheet: 0, navigation: 0, width: 384, height: 96, zoom: 1, dpr: 1,
    viewport: { x: 0, y: 0, width: 384, height: 96 },
  };
  const context = {
    save() {}, restore() {}, setTransform() {}, beginPath() {}, clip() {}, clearRect() {}, rect() {},
    moveTo() {}, lineTo() {}, stroke() {}, setLineDash() {},
    fillRect() {}, fillText(value: string) { text.push(value); },
  } as unknown as CanvasRenderingContext2D;
  let source: WorkerPaintSource | undefined;
  const paintFailure = new Error('Canvas installation failed');
  const run = new PeerlessWorkbookSession(session, 1, {
    changed: () => {}, isCurrent: () => true,
    onReady: () => { events.push('ready'); return () => { events.push('cleanup'); }; },
    onError: (error) => { errors.push(error); source?.fail(error); },
    checkpointScheduler: { now: () => 0, schedule: () => () => {} },
  });
  run.connectRecoveryInput({ draft: null, pending: false }, { compositionId: undefined, subscribe: () => () => {} });
  const paintSource = new WorkerPaintSource({
    generation: 1, commandIdentity: session.commandIdentity,
    capture: () => view, sentSequence: () => state.sent,
    requestFrame: () => { throw new Error('Unexpected scheduled frame RPC'); },
    publish: (painted) => {
      if (state.throwCanvas) throw paintFailure;
      paints.push(painted);
      events.push(`canvas:${painted.coversThroughSeq}`);
    },
    onInstalled: (painted) => {
      events.push(`presented:${painted.coversThroughSeq}`);
      session.markPresented({ ...session.commandIdentity, coversThroughSeq: painted.coversThroughSeq! });
      run.firstPaint();
    },
    onError: (error) => run.fail(error),
  });
  source = paintSource;
  const echo = new WorkerCellEcho({ source: paintSource, context: () => context, maxPatches });
  run.beforeRetire = () => { echo.dispose(); paintSource.dispose(); };
  const off = session.onAck((ack) => {
    acks.push(ack);
    echo.acknowledge(ack);
    if (state.installReplies && ack.frame) {
      const request = captures.get(ack.statuses[0].seq)!;
      paintSource.install(ack.frame, request);
    }
  });
  const submit = (seq: number, value: string | null, requestFrame = true) => {
    state.sent = seq;
    const captured = paintSource.captureCommandFrame()!;
    captures.set(seq, captured);
    if (value !== null) echo.patch(seq, [{ sheet: 0, row: 0, col: 0, text: value }]);
    const selection = { sheetId: 'sheet:0', range: { kind: 'a1' as const, a1: 'A1' } };
    const identity = { ...session.commandIdentity, seq, basis: session.confirmedBasis };
    const command: WorkbookCommand = value === null
      ? { ...identity, kind: 'select', payload: { selection } }
      : { ...identity, kind: 'editCell', payload: { args: [0, 0, 0, value], target: selection } };
    if (requestFrame) command.frame = {
      sheet: captured.sheet, viewport: captured.viewport, zoom: captured.zoom, dpr: captured.dpr,
      generation: captured.generation, requestGeneration: captured.requestGeneration,
    };
    return run.session.submit(command);
  };
  const held = (seq: number) => {
    if (replies.has(seq)) return Promise.resolve();
    const arrival = deferred<void>();
    arrivals.set(seq, arrival);
    return arrival.promise;
  };
  const deliver = (seq: number) => { replies.get(seq)!(); };
  const installCurrent = async () => {
    const request = paintSource.captureCommandFrame()!;
    const coversThroughSeq = session.throughSeq;
    const frame = await session.call.frame(request.viewport, { sheet: request.sheet });
    const covering: WorkbookCommandFrame = { ...frame, ...session.commandIdentity,
      coversThroughSeq, requestGeneration: request.requestGeneration!,
      zoom: request.zoom, dpr: request.dpr };
    return paintSource.install(covering, request);
  };
  return { session, run, source: paintSource, echo, view, state, text, events, paints, errors, acks,
    requests, context, paintFailure, submit, held, deliver, installCurrent,
    bind: () => paintSource.installInitial(session.initialPaint?.frame),
    crash: (failure: SessionFailure) => crash!(failure),
    dispose: async () => { off(); await run.dispose(); },
  };
}

describe('peerless command swap presentation transactions', () => {
  test('fused open before renderer binding waits for successful paint and runs ready cleanup once', async () => {
    const h = await harness();
    try {
      expect(h.session.initialPaint?.frame).toBeDefined();
      expect(h.run.commandsAccepted).toBe(true);
      expect(h.run.ready).toBe(false);
      expect(h.run.readOnly).toBe(true);
      let ready = false;
      const waiting = h.run.whenReady().then(() => { ready = true; });
      await Promise.resolve();
      expect(ready).toBe(false);
      expect(h.events).toEqual([]);
      expect(h.bind()).toBe(true);
      await waiting;
      h.run.firstPaint();
      expect(h.events).toEqual(['canvas:0', 'presented:0', 'ready']);
      expect(h.run.ready).toBe(true);
      expect(h.run.readOnly).toBe(false);
      expect(h.requests).toEqual(['open']);
      expect('peer' in h.run).toBe(false);
      expect('editPeer' in h.run).toBe(false);
      await h.run.dispose();
      await h.run.dispose();
      expect(h.events.filter((event) => event === 'cleanup')).toHaveLength(1);
    } finally { await h.dispose(); }
  });

  test('reversed replies retain the newest authoritative paint and retire covered echo once', async () => {
    const h = await harness();
    try {
      expect(h.bind()).toBe(true);
      const first = h.submit(1, 'first');
      const second = h.submit(2, 'second');
      expect(h.echo.size).toBe(2);
      await h.held(2);
      h.deliver(2);
      await second;
      const newest = h.source.painted;
      if (!newest) throw new Error('Missing authoritative paint');
      expect(h.session.throughSeq).toBe(2);
      expect(h.session.presentedThroughSeq).toBe(2);
      expect(h.echo.size).toBe(0);
      h.deliver(1);
      await first;
      expect(h.source.painted).toBe(newest);
      expect(h.source.coversThroughSeq).toBe(2);
      expect(h.session.confirmedBasis.version).toBe(newest.version);
      expect(h.acks.map((ack) => ack.statuses[0].seq)).toEqual([2, 1]);
      expect(h.paints.map((paint) => paint.coversThroughSeq)).toEqual([0, 2]);
      expect(h.errors).toEqual([]);
    } finally { await h.dispose(); }
  });

  test('ack followed by a stale viewport frame leaves presentation pending until a matching paint', async () => {
    const h = await harness();
    try {
      h.bind();
      const initial = h.source.painted;
      const editing = h.submit(1, 'pending');
      let presented = false;
      const printing = h.run.session.whenPresented(1).then(() => { presented = true; });
      h.view.viewport.x = 24;
      h.view.navigation++;
      await h.held(1);
      h.deliver(1);
      await editing;
      expect(h.session.throughSeq).toBe(1);
      expect(h.session.presentedThroughSeq).toBe(0);
      expect(h.source.painted).toBe(initial);
      expect(h.echo.size).toBe(1);
      expect(presented).toBe(false);
      const matching = h.submit(2, null);
      await h.held(2);
      h.deliver(2);
      await Promise.all([matching, printing]);
      expect(presented).toBe(true);
      expect(h.source.painted?.request.viewport.x).toBe(24);
      expect(h.source.coversThroughSeq).toBe(2);
      expect(h.echo.size).toBe(0);
      expect(h.errors).toEqual([]);
    } finally { await h.dispose(); }
  });

  test('same command coverage can paint a changed viewport while later echo and presentation remain pending', async () => {
    const h = await harness();
    try {
      h.bind();
      const first = h.submit(1, 'first');
      await h.held(1);
      h.deliver(1);
      await first;
      const previous = h.source.painted;
      h.echo.patch(2, [{ sheet: 0, row: 0, col: 0, text: 'later' }]);
      h.view.viewport.x = 1;
      const scrolled = h.installCurrent();
      const second = h.submit(2, 'later', false);
      let presented = false;
      const waiting = h.run.session.whenPresented(2).then(() => { presented = true; });
      await h.held(2);
      expect(await scrolled).toBe(true);
      expect(h.source.painted).not.toBe(previous);
      expect(h.source.painted?.request.viewport.x).toBe(1);
      expect(h.source.coversThroughSeq).toBe(1);
      expect(h.session.presentedThroughSeq).toBe(1);
      expect(h.echo.size).toBe(1);
      expect(h.text[h.text.length - 1]).toBe('later');
      expect(presented).toBe(false);
      h.deliver(2);
      await second;
      expect(await h.installCurrent()).toBe(true);
      await waiting;
      expect(h.source.coversThroughSeq).toBe(2);
      expect(h.echo.size).toBe(0);
    } finally { await h.dispose(); }
  });

  for (const failureAt of ['canvas', 'echo'] as const) {
    test(`${failureAt} failure rolls back paint coverage and both echoes before rejecting presentation`, async () => {
      const h = await harness();
      try {
        h.bind();
        const initial = h.source.painted;
        const first = h.submit(1, 'pending');
        const second = h.submit(2, 'later');
        const waiting = h.run.session.whenPresented(1);
        const settled = Promise.allSettled([first, second, waiting]);
        await h.held(2);
        if (failureAt === 'canvas') h.state.throwCanvas = true;
        else h.context.fillText = () => { throw h.paintFailure; };
        h.deliver(1);
        for (const outcome of await settled) {
          expect(outcome).toMatchObject({ status: 'rejected', reason: {
            name: 'XlsxWorkerSessionFailedError', cause: h.paintFailure,
          } });
        }
        expect(h.source.painted).toBe(initial);
        expect(h.source.coversThroughSeq).toBe(0);
        expect(h.session.presentedThroughSeq).toBe(0);
        expect(h.echo.size).toBe(2);
        expect(h.run.readOnly).toBe(true);
        expect(h.run.failure).toBeInstanceOf(XlsxWorkerSessionFailedError);
        h.deliver(2);
        await Promise.resolve();
        expect(h.source.painted).toBe(initial);
        expect(h.echo.size).toBe(2);
        expect(h.errors).toHaveLength(1);
        expect(h.events.filter((event) => event.startsWith('presented:'))).toEqual(['presented:0']);
        expect(h.requests.filter((method) => method === 'open')).toHaveLength(1);
      } finally { await h.dispose(); }
    });
  }

  test('hundreds of bare acknowledgements keep the newest visible echo bounded until covering paint', async () => {
    const h = await harness(3);
    try {
      h.bind();
      const initial = h.source.painted;
      for (let seq = 1; seq <= 260; seq++) {
        const editing = h.submit(seq, `value:${seq}`, false);
        await h.held(seq);
        h.deliver(seq);
        await editing;
        expect(h.echo.size).toBeLessThanOrEqual(3);
        expect(h.text[h.text.length - 1]).toBe(`value:${seq}`);
      }
      expect(h.echo.size).toBe(3);
      expect(h.source.painted).toBe(initial);
      expect(h.session.throughSeq).toBe(260);
      expect(h.session.presentedThroughSeq).toBe(0);
      expect(h.acks.every((ack) => ack.frame === undefined)).toBe(true);
      const waiting = h.run.session.whenPresented(260);
      expect(await h.installCurrent()).toBe(true);
      await waiting;
      expect(h.echo.size).toBe(0);
      expect(h.source.coversThroughSeq).toBe(260);
      expect(h.source.painted?.displayList.commands.some((entry) => entry.op === 'text' && entry.text === 'value:260')).toBe(true);
      expect(h.errors).toEqual([]);
    } finally { await h.dispose(); }
  });

  test('a trap after mutation rejects once and ignores a late frame without reopening or recovery', async () => {
    const h = await harness();
    try {
      h.bind();
      const initial = h.source.painted;
      const editing = h.submit(1, 'committed before trap');
      const waiting = h.run.session.whenPresented(1);
      const settled = Promise.allSettled([editing, waiting]);
      await h.held(1);
      expect(h.session.state.dirty).toBe(true);
      const cause = new SessionFailure('trap', 'Worker trapped before reply');
      h.crash(cause);
      for (const outcome of await settled) {
        expect(outcome).toMatchObject({ status: 'rejected', reason: { cause } });
      }
      h.deliver(1);
      h.run.firstPaint();
      await Promise.resolve();
      expect(h.source.painted).toBe(initial);
      expect(h.source.coversThroughSeq).toBe(0);
      expect(h.session.presentedThroughSeq).toBe(0);
      expect(h.run.readOnly).toBe(true);
      expect(h.errors).toHaveLength(1);
      expect(h.events.filter((event) => event === 'ready')).toHaveLength(1);
      await expect(h.run.session.save()).rejects.toBe(h.run.failure);
      expect(h.requests).toEqual(['open', 'command']);
    } finally { await h.dispose(); }
  });

  test('replacement rejects held work and keeps late old frames out of the next session', async () => {
    const old = await harness();
    const next = await harness();
    try {
      old.bind();
      const editing = old.submit(1, 'old');
      const painting = old.run.session.whenPresented(1);
      const settled = Promise.allSettled([editing, painting]);
      await old.held(1);
      await old.run.dispose();
      for (const outcome of await settled) {
        expect(outcome).toMatchObject({ status: 'rejected', reason: { code: 'document-replaced' } });
      }
      expect(next.bind()).toBe(true);
      const current = next.source.painted;
      old.deliver(1);
      old.run.firstPaint();
      await Promise.resolve();
      expect(next.source.painted).toBe(current);
      expect(next.session.throughSeq).toBe(0);
      expect(next.run.ready).toBe(true);
      expect(old.echo.size).toBe(0);
      expect(old.events.filter((event) => event === 'cleanup')).toHaveLength(1);
      expect(old.events.filter((event) => event === 'ready')).toHaveLength(1);
      expect(next.events.filter((event) => event === 'ready')).toHaveLength(1);
      expect(old.errors).toEqual([]);
      expect(next.errors).toEqual([]);
    } finally { await old.dispose(); await next.dispose(); }
  });
});
