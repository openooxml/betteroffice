import { describe, expect, it } from 'bun:test';
import { SessionFailure } from '../types';
import { createWorkerEditSession } from './session';
import type {
  Checkpoint, CommandAck, CommandConnection, CommandEnvelope, CommandOutcome,
  CommandTicket, EchoHooks, FrameEnvelope, WorkerEditSessionOptions,
} from './types';

type Ack = CommandAck<string, string, string>;

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<T>((yes, no) => { resolve = yes; reject = no; });
  return { promise, resolve, reject };
}

function frameEvents(connection: CommandConnection<string, string, string, string>) {
  const active = new Set<(frame: FrameEnvelope<string>) => void>();
  const saved: Array<(frame: FrameEnvelope<string>) => void> = [];
  let unsubscribes = 0;
  connection.client.onFrame = (listener) => {
    active.add(listener);
    saved.push(listener);
    return () => { unsubscribes += 1; active.delete(listener); };
  };
  return {
    push(frame: FrameEnvelope<string>, late = false) {
      for (const listener of late ? saved : [...active]) listener(frame);
    },
    get unsubscribes() { return unsubscribes; },
    get listeners() { return active.size; },
  };
}

function harness(overrides: Partial<WorkerEditSessionOptions<string, string, string, string, string, string>> = {}, pushed = false) {
  const calls: Array<{ envelope: CommandEnvelope<string>; reply: ReturnType<typeof deferred<Ack>> }> = [];
  const queued: Array<() => void> = [];
  const full = deferred<Ack>();
  let failures = new Set<(failure: SessionFailure) => void>();
  let disposals = 0;
  let fullRequests = 0;
  const connection: CommandConnection<string, string, string, string> = {
    initial: { session: 'test', incarnation: 0, epoch: 1, version: 0, throughSeq: 0, outcomes: [] },
    client: {
      call: {
        edit(envelope) {
          const reply = deferred<Ack>();
          calls.push({ envelope, reply });
          return reply.promise;
        },
        checkpoint: () => Promise.resolve({ state: new Uint8Array(), epoch: 1, version: 0, throughSeq: 0 }),
      },
      failure: undefined,
      onFailure(listener) { failures.add(listener); return () => { failures.delete(listener); }; },
      dispose() { disposals += 1; failures = new Set(); return Promise.resolve(); },
    },
  };
  const events = pushed ? frameEvents(connection) : undefined;
  const drops: Array<{ values: string[]; reason: string }> = [];
  const frames: Array<{ payload: string; retired: string[]; remaining: string[] }> = [];
  const echo: EchoHooks<string, string> = {
    paint() {},
    drop(entries, reason) { drops.push({ values: entries.map((entry) => entry.value), reason }); },
    present(frame, retired, remaining) {
      frames.push({ payload: frame.payload, retired: retired.map((entry) => entry.value), remaining: remaining.map((entry) => entry.value) });
    },
  };
  const session = createWorkerEditSession({
    source: 'source', connection, recovery: { reopen: () => Promise.reject(new Error('unexpected recovery')) },
    requestFullFrame: () => { fullRequests += 1; return full.promise; },
    enqueue: (callback) => { queued.push(callback); }, now: () => 0, echo, ...overrides,
  });
  function pump() { while (queued.length) queued.shift()!(); }
  function ack(index: number, outcomes?: CommandOutcome[], extra: Partial<Ack> = {}) {
    const seq = calls[index]!.envelope.seq;
    calls[index]!.reply.resolve({ ...connection.initial, throughSeq: seq, version: seq,
      outcomes: outcomes ?? [{ seq, status: 'applied' }], ...extra });
  }
  return { session, connection, calls, pump, ack, drops, frames, full, events,
    fail(failure: SessionFailure) { for (const listener of [...failures]) listener(failure); },
    get disposals() { return disposals; }, get fullRequests() { return fullRequests; } };
}

describe('worker edit session', () => {
  for (const transactionalPresentation of [false, true]) {
    it(`releases a painter-owned frame once when a ${transactionalPresentation ? 'transactional' : 'cosmetic'} presentation overrun disposes the session`, async () => {
      let closes = 0;
      let clock = 0;
      let installed = false;
      const bitmap = { close() { closes += 1; } } as ImageBitmap;
      const h = harness({
        transactionalPresentation, now: () => clock,
        echo: {
          paint() {}, drop() {},
          present() { installed = true; clock += 3; },
        },
        discardFrame() { bitmap.close(); },
        onEchoOverrun(sample) {
          expect(sample.hook).toBe('present');
          void h.session.dispose();
          if (installed) { bitmap.close(); installed = false; }
        },
      }, true);
      h.events!.push({ coversThroughSeq: 0, revision: 1, payload: 'bitmap' });
      await h.session.dispose();
      expect(h.disposals).toBe(1);
      expect(installed).toBe(false);
      expect(closes).toBe(1);
    });
  }

  it('discards rejected pushed frames and acknowledgements before they reach the presenter', async () => {
    const discarded: string[] = [];
    const h = harness({ discardFrame(candidate) { discarded.push(candidate.payload); } }, true);
    h.events!.push({ coversThroughSeq: 1, revision: 1, payload: 'ahead' });
    const first = h.session.submit('insert', 'a');
    h.pump();
    h.ack(0, undefined, { incarnation: 99,
      frame: { coversThroughSeq: 1, revision: 2, payload: 'wrong incarnation' } });
    for (let turn = 0; turn < 8; turn += 1) await Promise.resolve();
    expect(discarded).toEqual(['ahead', 'wrong incarnation']);
    expect(h.frames).toEqual([]);
    await h.session.dispose();
    expect(await first.presented).toBe('dropped');
  });

  it('discards late edit replies, full-frame replies, and pushes after disposal', async () => {
    const discarded: string[] = [];
    const h = harness({ discardFrame(candidate) { discarded.push(candidate.payload); } }, true);
    const first = h.session.submit('insert', 'a');
    const second = h.session.submit('insert', 'b');
    h.pump();
    h.ack(0, undefined, { frame: { coversThroughSeq: 1, revision: 1, baseRevision: 99, payload: 'delta' } });
    await first.ack;
    expect(h.fullRequests).toBe(1);
    await h.session.dispose();
    h.ack(1, undefined, { frame: { coversThroughSeq: 2, revision: 2, payload: 'late edit' } });
    h.full.resolve({ ...h.connection.initial, throughSeq: 1,
      frame: { coversThroughSeq: 1, revision: 3, payload: 'late full' } });
    h.events!.push({ coversThroughSeq: 1, revision: 4, payload: 'late push' }, true);
    for (let turn = 0; turn < 8; turn += 1) await Promise.resolve();
    expect(discarded).toEqual(['delta', 'late push', 'late edit', 'late full']);
    expect(h.frames).toEqual([]);
    expect(await second.presented).toBe('dropped');
  });

  for (const rejection of ['returns false', 'throws'] as const) {
    it(`keeps transactional tickets and coverage waiters pending when an ack or pushed frame hook ${rejection}`, async () => {
      for (const pushed of [false, true]) {
        let reject = true;
        const installs: string[] = [];
        const h = harness({
          transactionalPresentation: true,
          echo: {
            paint() {}, drop() {},
            present(candidate) {
              installs.push(candidate.payload);
              if (!reject) return;
              if (rejection === 'throws') throw new Error('install failed');
              return false;
            },
          },
        }, pushed);
        const first = h.session.submit('insert', 'a', { echo: 'a' });
        h.pump();
        let ticketStatus: string | undefined;
        let coverageStatus: string | undefined;
        void first.presented.then((status) => { ticketStatus = status; });
        const waiting = h.session.whenPresented(1).then((status) => { coverageStatus = status; return status; });
        const rejected = { coversThroughSeq: 1, revision: 10, payload: 'rejected' };
        h.ack(0, undefined, pushed ? {} : { frame: rejected });
        expect(await first.ack).toMatchObject({ seq: 1, status: 'applied' });
        if (pushed) h.events!.push(rejected);
        await Promise.resolve();
        expect(ticketStatus).toBeUndefined();
        expect(coverageStatus).toBeUndefined();
        expect(h.session.state.presentedThroughSeq).toBe(-1);
        expect(h.session.state.throughSeq).toBe(1);
        expect(h.fullRequests).toBe(0);
        reject = false;
        const later = h.session.submit('insert', 'b', { echo: 'b' });
        h.pump();
        h.ack(1, undefined, { frame: { coversThroughSeq: 2, revision: 11, payload: 'accepted' } });
        expect(await first.presented).toBe('presented');
        expect(await waiting).toBe('presented');
        expect(await later.presented).toBe('presented');
        expect(h.session.state.presentedThroughSeq).toBe(2);
        expect(installs).toEqual(['rejected', 'accepted']);
        expect(h.fullRequests).toBe(0);
        await h.session.dispose();
      }
    });
  }

  it('stops full-frame retries after transactional rejection and presents pending tickets on a later accepted frame', async () => {
    for (const throws of [false, true]) {
      let reject = true;
      let requests = 0;
      const h = harness({
        transactionalPresentation: true,
        requestFullFrame(connection) {
          requests += 1;
          return Promise.resolve({ ...connection.initial, throughSeq: 1, version: 1,
            frame: { coversThroughSeq: 1, revision: 10 + requests, payload: 'full' } });
        },
        echo: {
          paint() {}, drop() {},
          present() {
            if (!reject) return;
            if (throws) throw new Error('install failed');
            return false;
          },
        },
      }, true);
      const first = h.session.submit('insert', 'a', { echo: 'a' });
      h.pump();
      let ticketStatus: string | undefined;
      let coverageStatus: string | undefined;
      void first.presented.then((status) => { ticketStatus = status; });
      const waiting = h.session.whenPresented(1).then((status) => { coverageStatus = status; return status; });
      h.ack(0, undefined, { frame: { coversThroughSeq: 1, revision: 10, baseRevision: 99, payload: 'delta' } });
      expect(await first.ack).toMatchObject({ seq: 1, status: 'applied' });
      for (let attempt = 1; attempt <= 4; attempt += 1) {
        for (let turn = 0; turn < 32; turn += 1) await Promise.resolve();
        expect(requests).toBe(attempt);
        expect(h.disposals).toBe(0);
        expect(h.session.state.presentedThroughSeq).toBe(-1);
        expect(ticketStatus).toBeUndefined();
        expect(coverageStatus).toBeUndefined();
        if (attempt < 4) h.events!.push({ coversThroughSeq: 1, revision: 20 + attempt,
          baseRevision: 99, payload: 'delta' });
      }
      reject = false;
      h.events!.push({ coversThroughSeq: 1, revision: 30, payload: 'accepted' });
      expect(await first.presented).toBe('presented');
      expect(await waiting).toBe('presented');
      expect(h.session.state.presentedThroughSeq).toBe(1);
      expect(requests).toBe(4);
      expect(h.disposals).toBe(0);
      await h.session.dispose();
    }
  });

  it('fetches a newer full-frame target requested during a transactional reply hook', async () => {
    for (const rejectOuter of [false, true]) {
      const requests: Array<ReturnType<typeof deferred<Ack>>> = [];
      const installs: string[] = [];
      let push!: (frame: FrameEnvelope<string>) => void;
      const h = harness({
        transactionalPresentation: true,
        requestFullFrame() {
          const reply = deferred<Ack>();
          requests.push(reply);
          return reply.promise;
        },
        echo: {
          paint() {}, drop() {},
          present(candidate) {
            installs.push(candidate.payload);
            if (candidate.payload === 'nested') return false;
            if (candidate.payload === 'first full') {
              push({ coversThroughSeq: 2, revision: 20, payload: 'nested' });
              push({ coversThroughSeq: 2, revision: 21, baseRevision: 99, payload: 'newer delta' });
              if (rejectOuter) return false;
            }
          },
        },
      }, true);
      push = (frame) => h.events!.push(frame);
      const first = h.session.submit('insert', 'a');
      const second = h.session.submit('insert', 'b');
      h.pump();
      h.ack(0);
      h.ack(1);
      await Promise.all([first.ack, second.ack]);
      let secondPresented = false;
      void second.presented.then(() => { secondPresented = true; });
      const waiting = h.session.whenPresented(2);
      push({ coversThroughSeq: 1, revision: 10, baseRevision: 99, payload: 'first delta' });
      expect(requests).toHaveLength(1);
      requests[0]!.resolve({ ...h.connection.initial, throughSeq: 1, version: 1,
        frame: { coversThroughSeq: 1, revision: 10, payload: 'first full' } });
      for (let turn = 0; turn < 16; turn += 1) await Promise.resolve();
      expect(requests).toHaveLength(2);
      expect(secondPresented).toBe(false);
      expect(h.session.state.presentedThroughSeq).toBe(rejectOuter ? -1 : 1);
      expect(h.disposals).toBe(0);
      requests[1]!.resolve({ ...h.connection.initial, throughSeq: 2, version: 2,
        frame: { coversThroughSeq: 2, revision: 21, payload: 'second full' } });
      expect(await second.presented).toBe('presented');
      expect(await first.presented).toBe('presented');
      expect(await waiting).toBe('presented');
      expect(h.session.state.presentedThroughSeq).toBe(2);
      expect(installs).toEqual(['first full', 'nested', 'second full']);
      expect(requests).toHaveLength(2);
      expect(h.disposals).toBe(0);
      await h.session.dispose();
    }
  });

  it('avoids retrying a full-frame target satisfied by a nested transactional install', async () => {
    const full = deferred<Ack>();
    const installs: string[] = [];
    let requests = 0;
    let push!: (frame: FrameEnvelope<string>) => void;
    const h = harness({
      transactionalPresentation: true,
      sameCoverageFrames: true,
      maxRecoveryAttempts: 0,
      requestFullFrame() {
        requests += 1;
        return requests === 1 ? full.promise : Promise.reject(new Error('unexpected full-frame request'));
      },
      echo: {
        paint() {}, drop() {},
        present(candidate) {
          installs.push(candidate.payload);
          if (candidate.payload === 'first full') {
            push({ coversThroughSeq: 2, revision: 20, baseRevision: 99, payload: 'newer delta' });
            push({ coversThroughSeq: 2, revision: 20, payload: 'second full' });
            return false;
          }
        },
      },
    }, true);
    push = (frame) => h.events!.push(frame);
    const first = h.session.submit('insert', 'a');
    const second = h.session.submit('insert', 'b');
    h.pump();
    h.ack(0);
    h.ack(1);
    await Promise.all([first.ack, second.ack]);
    const waiting = h.session.whenPresented(2);
    push({ coversThroughSeq: 1, revision: 10, baseRevision: 99, payload: 'first delta' });
    expect(requests).toBe(1);
    full.resolve({ ...h.connection.initial, throughSeq: 1, version: 1,
      frame: { coversThroughSeq: 1, revision: 10, payload: 'first full' } });
    for (let turn = 0; turn < 16; turn += 1) await Promise.resolve();
    expect(requests).toBe(1);
    expect(h.disposals).toBe(0);
    expect(h.events!.listeners).toBe(1);
    expect(await second.presented).toBe('presented');
    expect(await first.presented).toBe('presented');
    expect(await waiting).toBe('presented');
    expect(h.session.state.presentedThroughSeq).toBe(2);
    expect(installs).toEqual(['first full', 'second full']);
    await h.session.flush();
    await h.session.dispose();
  });

  it('preserves the evicted echo command acknowledgement and presents it on a covering frame', async () => {
    const h = harness({ maxEchoEntries: 1, echoOverflow: 'evict-oldest' });
    const first = h.session.submit('insert', 'a', { echo: 'a' });
    h.pump();
    const waiting = h.session.whenPresented(1);
    const later = h.session.submit('insert', 'b', { echo: 'b' });
    h.pump();
    let acked = false;
    let presented = false;
    void first.ack.then(() => { acked = true; });
    void first.presented.then(() => { presented = true; });
    await Promise.resolve();
    expect(acked).toBe(false);
    expect(presented).toBe(false);
    expect(h.drops).toEqual([]);
    h.ack(0);
    expect(await first.ack).toEqual({ seq: 1, status: 'applied' });
    expect(presented).toBe(false);
    h.ack(1, undefined, { frame: { coversThroughSeq: 2, revision: 1, payload: 'ab' } });
    expect(await later.ack).toEqual({ seq: 2, status: 'applied' });
    expect(await first.presented).toBe('presented');
    expect(await later.presented).toBe('presented');
    expect(await waiting).toBe('presented');
    expect(h.frames).toEqual([{ payload: 'ab', retired: ['b'], remaining: [] }]);
    expect(h.drops).toEqual([]);
    await h.session.dispose();
  });

  it('installs opted-in pushed frames at unchanged coverage, publishes state, and leaves tickets untouched', async () => {
    const h = harness({ sameCoverageFrames: true }, true);
    const published: number[] = [];
    h.session.subscribe((state) => { published.push(state.presentedThroughSeq); });
    const first = h.session.submit('insert', 'a', { echo: 'a' });
    const later = h.session.submit('insert', 'b', { echo: 'b' });
    h.pump();
    h.ack(0, undefined, { frame: { coversThroughSeq: 1, revision: 10, payload: 'ack' } });
    expect(await first.presented).toBe('presented');
    let laterAck = false;
    let laterPresentation = false;
    let waiting = false;
    void later.ack.then(() => { laterAck = true; });
    void later.presented.then(() => { laterPresentation = true; });
    void h.session.whenPresented(2).then(() => { waiting = true; });
    h.events!.push({ coversThroughSeq: 1, revision: 11, baseRevision: 10, payload: 'viewport' });
    await Promise.resolve();
    expect(h.frames).toEqual([
      { payload: 'ack', retired: ['a'], remaining: ['b'] },
      { payload: 'viewport', retired: [], remaining: ['b'] },
    ]);
    expect(published).toEqual([1, 1]);
    expect(laterAck).toBe(false);
    expect(laterPresentation).toBe(false);
    expect(waiting).toBe(false);
    expect(h.drops).toEqual([]);
    h.ack(1, undefined, { frame: { coversThroughSeq: 2, revision: 12, payload: 'later' } });
    expect(await later.presented).toBe('presented');
    await h.session.dispose();
    expect(h.events!.unsubscribes).toBe(1);
  });

  it('ignores pushed coverage ahead of acknowledgements and presents the later ack with or without onFrame', async () => {
    for (const pushed of [false, true]) {
      const h = harness({}, pushed);
      const ticket = h.session.submit('insert', 'a', { echo: 'a' });
      h.pump();
      let settled = false;
      void ticket.presented.then(() => { settled = true; });
      if (pushed) {
        expect(h.events!.listeners).toBe(1);
        h.events!.push({ coversThroughSeq: 1, revision: 99, payload: 'too early' });
      }
      await Promise.resolve();
      expect(settled).toBe(false);
      expect(h.frames).toEqual([]);
      h.ack(0, undefined, { frame: { coversThroughSeq: 1, revision: 1, payload: 'ack' } });
      expect(await ticket.presented).toBe('presented');
      expect(h.frames).toEqual([{ payload: 'ack', retired: ['a'], remaining: [] }]);
      await h.session.dispose();
    }
  });

  it('terminates malformed pushed-frame metadata with message failure, dropped presentations, and listener removal', async () => {
    const valid = { coversThroughSeq: 1, revision: 10, payload: 'invalid' };
    const malformed: unknown[] = [
      undefined, null, 'frame',
      { ...valid, coversThroughSeq: -1 },
      { ...valid, coversThroughSeq: 0.5 },
      { ...valid, coversThroughSeq: Number.MAX_SAFE_INTEGER + 1 },
      { ...valid, coversThroughSeq: '1' },
      { ...valid, revision: undefined },
      { ...valid, revision: NaN },
      { ...valid, coversThroughSeq: 99, revision: Infinity },
      { ...valid, baseRevision: '10' },
      { ...valid, baseRevision: Infinity },
      { coversThroughSeq: 1, revision: 10 },
    ];
    for (const candidate of malformed) {
      const h = harness({ sameCoverageFrames: true }, true);
      const acked = h.session.submit('insert', 'a', { echo: 'a' });
      const pending = h.session.submit('insert', 'b', { echo: 'b' });
      h.pump();
      h.ack(0);
      await acked.ack;
      const covered = h.session.whenPresented(1);
      const future = h.session.whenPresented(3);
      h.events!.push(candidate as FrameEnvelope<string>);
      await expect(pending.ack).rejects.toMatchObject({ code: 'message' });
      await expect(h.session.flush()).rejects.toMatchObject({ code: 'message' });
      expect(await acked.presented).toBe('dropped');
      expect(await pending.presented).toBe('dropped');
      expect(await covered).toBe('dropped');
      expect(await future).toBe('dropped');
      expect(h.events!.unsubscribes).toBe(1);
      expect(h.events!.listeners).toBe(0);
      expect(h.disposals).toBe(1);
      h.events!.push({ coversThroughSeq: 1, revision: 11, payload: 'late' }, true);
      expect(h.frames).toEqual([]);
      await h.session.dispose();
      expect(h.events!.unsubscribes).toBe(1);
    }
  });

  it('uses pushed frames to present already-acked tickets and requests a full frame for same-coverage base mismatches', async () => {
    const h = harness({ sameCoverageFrames: true }, true);
    const ticket = h.session.submit('insert', 'a');
    h.pump();
    h.ack(0);
    await ticket.ack;
    h.events!.push({ coversThroughSeq: 1, revision: 10, payload: 'pushed' });
    expect(h.frames.map((frame) => frame.payload)).toEqual(['pushed']);
    expect(await ticket.presented).toBe('presented');
    h.events!.push({ coversThroughSeq: 1, revision: 11, baseRevision: 9, payload: 'delta' });
    expect(h.fullRequests).toBe(1);
    expect(h.frames.map((frame) => frame.payload)).toEqual(['pushed']);
    h.full.resolve({ ...h.connection.initial, throughSeq: 1, version: 1, outcomes: [],
      frame: { coversThroughSeq: 1, revision: 12, payload: 'full' } });
    await Promise.resolve();
    expect(h.frames.map((frame) => frame.payload)).toEqual(['pushed', 'full']);
    await h.session.dispose();
  });

  it('requests revision 14 after a mismatched push arrives while the revision 13 full frame is pending', async () => {
    const requests: Array<ReturnType<typeof deferred<Ack>>> = [];
    const h = harness({ sameCoverageFrames: true, requestFullFrame() {
      const reply = deferred<Ack>();
      requests.push(reply);
      return reply.promise;
    } }, true);
    const ticket = h.session.submit('insert', 'a');
    h.pump();
    h.ack(0, undefined, { frame: { coversThroughSeq: 1, revision: 10, payload: 'initial' } });
    expect(await ticket.presented).toBe('presented');
    h.events!.push({ coversThroughSeq: 1, revision: 13, baseRevision: 12, payload: 'first delta' });
    expect(requests).toHaveLength(1);
    h.events!.push({ coversThroughSeq: 1, revision: 14, baseRevision: 13, payload: 'newer delta' });
    expect(requests).toHaveLength(1);
    requests[0]!.resolve({ ...h.connection.initial, throughSeq: 1, version: 1,
      frame: { coversThroughSeq: 1, revision: 13, payload: 'revision 13' } });
    await Promise.resolve();
    expect(h.frames.map((frame) => frame.payload)).toEqual(['initial', 'revision 13']);
    expect(requests).toHaveLength(2);
    requests[1]!.resolve({ ...h.connection.initial, throughSeq: 1, version: 1,
      frame: { coversThroughSeq: 1, revision: 14, payload: 'revision 14' } });
    await Promise.resolve();
    expect(h.frames.map((frame) => frame.payload)).toEqual(['initial', 'revision 13', 'revision 14']);
    expect(requests).toHaveLength(2);
    expect(h.disposals).toBe(0);
    expect(h.session.state.presentedThroughSeq).toBe(1);
    expect(await ticket.presented).toBe('presented');
    await h.session.dispose();
  });

  it('bounds immediately resolved full frames that satisfy coverage but miss the requested revision', async () => {
    let requests = 0;
    const h = harness({ sameCoverageFrames: true, requestFullFrame(connection) {
      requests += 1;
      return Promise.resolve({ ...connection.initial, throughSeq: 1, version: 1,
        frame: { coversThroughSeq: 1, revision: 10 + requests, payload: 'insufficient' } });
    } }, true);
    const ticket = h.session.submit('insert', 'a');
    h.pump();
    h.ack(0, undefined, { frame: { coversThroughSeq: 1, revision: 10, payload: 'initial' } });
    await ticket.presented;
    h.events!.push({ coversThroughSeq: 1, revision: 20, baseRevision: 19, payload: 'delta' });
    for (let index = 0; index < 32; index += 1) await Promise.resolve();
    expect(requests).toBe(4);
    expect(h.frames.map((frame) => frame.payload)).toEqual([
      'initial', 'insufficient', 'insufficient', 'insufficient', 'insufficient',
    ]);
    expect(h.disposals).toBe(1);
    expect(h.events!.listeners).toBe(0);
    await expect(h.session.flush()).rejects.toMatchObject({ code: 'message' });
    await h.session.dispose();
  });

  it('completes default full-frame recovery after greater coverage with a lower revision and a late response', async () => {
    const firstFull = deferred<Ack>();
    let requests = 0;
    const h = harness({ requestFullFrame(connection) {
      requests += 1;
      return requests === 1 ? firstFull.promise : Promise.resolve({ ...connection.initial,
        throughSeq: 2, version: 2, frame: { coversThroughSeq: 2, revision: 1, payload: 'retry' } });
    } }, true);
    const first = h.session.submit('insert', 'a');
    const second = h.session.submit('insert', 'b');
    h.pump();
    h.ack(0);
    await first.ack;
    h.events!.push({ coversThroughSeq: 1, revision: 10, baseRevision: 9, payload: 'delta' });
    expect(requests).toBe(1);
    h.ack(1, undefined, { frame: { coversThroughSeq: 2, revision: 1, payload: 'intervening full' } });
    expect(await second.presented).toBe('presented');
    firstFull.resolve({ ...h.connection.initial, throughSeq: 1, version: 1,
      frame: { coversThroughSeq: 1, revision: 10, payload: 'late full' } });
    for (let index = 0; index < 32; index += 1) await Promise.resolve();
    expect(requests).toBe(1);
    expect(h.disposals).toBe(0);
    expect(await first.presented).toBe('presented');
    expect(h.frames.map((frame) => frame.payload)).toEqual(['intervening full']);
    await h.session.flush();
    await h.session.dispose();
  });

  it('completes opted-out full-frame recovery by coverage despite lower and stale equal-coverage revisions', async () => {
    let requests = 0;
    const h = harness({ sameCoverageFrames: false, requestFullFrame(connection) {
      requests += 1;
      return Promise.resolve({ ...connection.initial, throughSeq: 1, version: 1,
        frame: { coversThroughSeq: 1, revision: 10, payload: 'full' } });
    } }, true);
    const ticket = h.session.submit('insert', 'a');
    h.pump();
    h.ack(0);
    await ticket.ack;
    h.events!.push({ coversThroughSeq: 1, revision: 20, baseRevision: 19, payload: 'delta' });
    for (let index = 0; index < 32; index += 1) await Promise.resolve();
    expect(requests).toBe(1);
    expect(h.disposals).toBe(0);
    expect(await ticket.presented).toBe('presented');
    expect(h.frames.map((frame) => frame.payload)).toEqual(['full']);
    await h.session.flush();
    await h.session.dispose();
  });

  it('merges pending full-frame targets by coverage before revision', async () => {
    const firstFull = deferred<Ack>();
    let requests = 0;
    const h = harness({ sameCoverageFrames: true, requestFullFrame(connection) {
      requests += 1;
      return requests === 1 ? firstFull.promise : Promise.resolve({ ...connection.initial,
        throughSeq: 2, version: 2, frame: { coversThroughSeq: 2, revision: 1, payload: 'retry' } });
    } }, true);
    const first = h.session.submit('insert', 'a');
    const second = h.session.submit('insert', 'b');
    h.pump();
    h.ack(0);
    await first.ack;
    h.ack(1);
    await second.ack;
    h.events!.push({ coversThroughSeq: 1, revision: 20, baseRevision: 19, payload: 'first delta' });
    h.events!.push({ coversThroughSeq: 2, revision: 1, baseRevision: 0, payload: 'second delta' });
    h.events!.push({ coversThroughSeq: 1, revision: 30, baseRevision: 29, payload: 'late delta' });
    expect(requests).toBe(1);
    firstFull.resolve({ ...h.connection.initial, throughSeq: 2, version: 2,
      frame: { coversThroughSeq: 2, revision: 1, payload: 'full' } });
    for (let index = 0; index < 32; index += 1) await Promise.resolve();
    expect(requests).toBe(1);
    expect(h.disposals).toBe(0);
    expect(await first.presented).toBe('presented');
    expect(await second.presented).toBe('presented');
    await h.session.flush();
    await h.session.dispose();
  });

  it('drops the outer ticket when its hook throws after a successful nested same-coverage push', async () => {
    let events!: ReturnType<typeof frameEvents>;
    const revisions: number[] = [];
    const h = harness({ sameCoverageFrames: true, echo: {
      paint() {}, drop() {},
      present(candidate) {
        revisions.push(candidate.revision);
        if (candidate.revision !== 11) return;
        events.push({ coversThroughSeq: 1, revision: 13, payload: 'nested' });
        throw new Error('present');
      },
    } }, true);
    events = h.events!;
    const ticket = h.session.submit('insert', 'a');
    h.pump();
    h.ack(0);
    await ticket.ack;
    events.push({ coversThroughSeq: 1, revision: 11, payload: 'outer' });
    expect(revisions).toEqual([11, 13]);
    expect(await ticket.presented).toBe('dropped');
    expect(await h.session.whenPresented(1)).toBe('dropped');
    expect(h.session.state.presentedThroughSeq).toBe(1);
    expect(h.disposals).toBe(0);
    await h.session.dispose();
  });

  it('stops an older state publication after a listener publishes newer coverage', async () => {
    const h = harness({}, true);
    const first = h.session.submit('insert', 'a');
    const second = h.session.submit('insert', 'b');
    h.pump();
    h.ack(0);
    await first.ack;
    h.ack(1);
    await second.ack;
    const earlier: number[] = [];
    const later: number[] = [];
    h.session.subscribe((state) => {
      earlier.push(state.presentedThroughSeq);
      if (state.presentedThroughSeq === 1) {
        h.events!.push({ coversThroughSeq: 2, revision: 2, payload: 'nested' });
      }
    });
    h.session.subscribe((state) => { later.push(state.presentedThroughSeq); });
    h.events!.push({ coversThroughSeq: 1, revision: 1, payload: 'outer' });
    expect(earlier).toEqual([1, 2]);
    expect(later).toEqual([2]);
    expect(h.session.state.presentedThroughSeq).toBe(2);
    expect(await first.presented).toBe('presented');
    expect(await second.presented).toBe('presented');
    await h.session.dispose();
  });

  it('removes pushed-frame listeners and ignores replaced, closed, recovering, and disposed connections', async () => {
    const reopened = deferred<CommandConnection<string, string, string, string>>();
    const replayStarted = deferred<void>();
    const replay = deferred<Ack>();
    const discarded: string[] = [];
    const h = harness({ sameCoverageFrames: true, recovery: { reopen: () => reopened.promise },
      discardFrame(candidate) { discarded.push(candidate.payload); } }, true);
    const ticket = h.session.submit('insert', 'a');
    h.pump();
    h.events!.push({ coversThroughSeq: 0, revision: 1, payload: 'first' });
    expect(h.frames.map((frame) => frame.payload)).toEqual(['first']);
    h.fail(new SessionFailure('crash', 'restart'));
    expect(h.events!.unsubscribes).toBe(1);
    expect(h.events!.listeners).toBe(0);
    h.events!.push({ coversThroughSeq: 0, revision: 2, payload: 'closed' }, true);
    let nextFailure: ((failure: SessionFailure) => void) | undefined;
    const next: CommandConnection<string, string, string, string> = {
      initial: { ...h.connection.initial, incarnation: 1,
        frame: { coversThroughSeq: 0, revision: 1, payload: 'reopen' } },
      client: {
        call: { edit() { replayStarted.resolve(); return replay.promise; }, checkpoint: h.connection.client.call.checkpoint },
        failure: undefined,
        onFailure(listener) { nextFailure = listener; return () => { nextFailure = undefined; }; },
        dispose: () => Promise.resolve(),
      },
    };
    const events = frameEvents(next);
    reopened.resolve(next);
    await replayStarted.promise;
    expect(events.listeners).toBe(1);
    events.push({ coversThroughSeq: 0, revision: 2, payload: 'recovering' });
    h.events!.push({ coversThroughSeq: 0, revision: 3, payload: 'replaced' }, true);
    expect(h.frames.map((frame) => frame.payload)).toEqual(['first', 'reopen']);
    replay.resolve({ ...next.initial, throughSeq: 1, version: 1, outcomes: [{ seq: 1, status: 'applied' }],
      frame: { coversThroughSeq: 1, revision: 2, payload: 'replayed' } });
    expect(await ticket.presented).toBe('presented');
    events.push({ coversThroughSeq: 1, revision: 3, payload: 'live' });
    expect(h.frames.map((frame) => frame.payload)).toEqual(['first', 'reopen', 'replayed', 'live']);
    nextFailure!(new SessionFailure('message', 'terminal'));
    expect(events.unsubscribes).toBe(1);
    expect(events.listeners).toBe(0);
    events.push({ coversThroughSeq: 1, revision: 4, payload: 'terminal' }, true);
    await h.session.dispose();
    events.push({ coversThroughSeq: 1, revision: 5, payload: 'disposed' }, true);
    expect(events.unsubscribes).toBe(1);
    expect(discarded).toEqual(['closed', 'recovering', 'replaced', 'terminal', 'disposed']);
    expect(h.frames.map((frame) => frame.payload)).toEqual(['first', 'reopen', 'replayed', 'live']);
    const disposed = harness({ sameCoverageFrames: true }, true);
    await disposed.session.dispose();
    expect(disposed.events!.unsubscribes).toBe(1);
    disposed.events!.push({ coversThroughSeq: 0, revision: 1, payload: 'late' }, true);
    expect(disposed.frames).toEqual([]);
  });

  it('marks unsent cancellations before reentrant echo drop submissions', async () => {
    let reentrant: CommandTicket | undefined;
    let session!: ReturnType<typeof harness>['session'];
    const h = harness({ maxInFlight: 1, echo: {
      paint() {}, present() {},
      drop(entries) {
        if (entries[0]?.value === 'child') {
          reentrant = session.submit('insert', 'late', { after: child });
        }
      },
    } });
    session = h.session;
    const first = h.session.submit('insert', 'first');
    const child = h.session.submit('insert', 'child', { after: first, echo: 'child' });
    h.pump();
    h.ack(0, [{ seq: 1, status: 'refused', reason: 'locked' }], { version: 0 });
    await child.ack;
    expect(reentrant).toBeDefined();
    expect(await reentrant!.ack).toMatchObject({ status: 'cancelled', reason: 'dependency-refused' });
    expect(await reentrant!.presented).toBe('dropped');
    await h.session.flush();
    await h.session.dispose();
  });

  it('shares seq and settlement promises across an unbounded merged group', async () => {
    const h = harness({ coalesce: (_previous, next) => next });
    const tickets = Array.from({ length: 2048 }, () => h.session.submit('insert', 'a', { coalesceKey: 'typing' }));
    const first = tickets[0]!;
    expect(tickets.every((ticket) => ticket.ack === first.ack && ticket.presented === first.presented)).toBe(true);
    let reads = 0;
    for (const ticket of tickets) {
      const id = ticket.id;
      Object.defineProperty(ticket, 'id', { get() { reads += 1; return id; } });
    }
    h.pump();
    expect(reads).toBeLessThan(16);
    expect(tickets.every((ticket) => ticket.seq === 1)).toBe(true);
    reads = 0;
    h.ack(0, undefined, { frame: { coversThroughSeq: 1, revision: 1, payload: 'a' } });
    expect(await first.ack).toMatchObject({ seq: 1, status: 'applied' });
    expect(await first.presented).toBe('presented');
    expect(reads).toBeLessThan(16);
    await h.session.dispose();
  });

  it('bounds ticket presentation groups and drops overflow before a covering frame', async () => {
    let boundary = 0;
    const h = harness({ checkpointPolicy: () => true, recovery: {
      reopen: () => Promise.reject(new Error('unexpected recovery')),
      checkpoint: () => Promise.resolve({ state: new Uint8Array(), epoch: 1, version: boundary, throughSeq: boundary }),
    } });
    const tickets: CommandTicket[] = [];
    for (let index = 0; index < 1025; index += 1) {
      const ticket = h.session.submit('insert', 'a');
      tickets.push(ticket);
      h.pump();
      h.ack(index);
      await ticket.ack;
      boundary = index + 1;
      h.pump();
      await Promise.resolve();
    }
    expect(await tickets[0]!.presented).toBe('dropped');
    expect(await h.session.whenPresented(1)).toBe('dropped');
    const later = h.session.submit('insert', 'b');
    h.pump();
    h.ack(1025, undefined, { frame: { coversThroughSeq: 1026, revision: 1, payload: 'full' } });
    expect(await later.presented).toBe('presented');
    expect(await tickets[0]!.presented).toBe('dropped');
    expect(await tickets[1024]!.presented).toBe('presented');
    await h.session.dispose();
  });

  it('fails loudly after bounded immediately resolved insufficient full frames', async () => {
    let requests = 0;
    const h = harness({ requestFullFrame: (connection) => {
      requests += 1;
      return Promise.resolve(connection.initial);
    } });
    try {
      const ticket = h.session.submit('insert', 'a');
      h.pump();
      h.ack(0, undefined, { frame: { coversThroughSeq: 1, revision: 1, baseRevision: 99, payload: 'delta' } });
      await ticket.ack;
      for (let index = 0; index < 32; index += 1) await Promise.resolve();
      expect(requests).toBe(4);
      expect(h.disposals).toBe(1);
      expect(await ticket.presented).toBe('dropped');
      await expect(h.session.flush()).rejects.toMatchObject({ code: 'message' });
    } finally {
      await h.session.dispose();
    }
  });

  it('assigns seq only at send, caps in-flight calls, and never guesses future versions', async () => {
    const h = harness({ maxInFlight: 2 });
    const first = h.session.submit('insert', 'a');
    const second = h.session.submit('insert', 'b', { after: first });
    const third = h.session.submit('insert', 'c', { after: second, expectVersion: true });
    expect(first.seq).toBeUndefined();
    h.pump();
    expect(h.calls.map((call) => call.envelope.seq)).toEqual([1, 2]);
    expect(h.calls.map((call) => call.envelope.basis.version)).toEqual([0, 0]);
    expect(h.calls[1]!.envelope.afterSeq).toBe(1);
    expect(third.seq).toBeUndefined();
    h.ack(0);
    await first.ack;
    h.pump();
    expect(third.seq).toBe(3);
    expect(h.calls[2]!.envelope).toMatchObject({ basis: { epoch: 1, version: 1 }, afterSeq: 2, expectVersion: true });
    await h.session.dispose();
  });

  it('settles ticket outcomes once, publishes state, and flushes a captured submission fence', async () => {
    const h = harness();
    const states: number[] = [];
    const off = h.session.subscribe((state) => { states.push(state.throughSeq); });
    const first = h.session.submit('insert', 'a');
    const second = h.session.submit('insert', 'b');
    const fence = h.session.flush();
    const later = h.session.submit('insert', 'c');
    h.pump();
    h.ack(1, [{ seq: 2, status: 'noop' }], { version: 1, selection: 'new', editState: 'toolbar' });
    await second.ack;
    h.ack(0, [{ seq: 1, status: 'applied' }], { selection: 'old', editState: 'old' });
    await fence;
    expect(await first.ack).toMatchObject({ seq: 1, status: 'applied' });
    expect(h.session.state).toMatchObject({ throughSeq: 2, version: 1, selection: 'new', editState: 'toolbar' });
    expect(states).toContain(2);
    expect(later.seq).toBe(3);
    off();
    await h.session.dispose();
  });

  it('drops refused echo, cancels unsent descendants, and keeps independent drafts', async () => {
    const h = harness({ maxInFlight: 1 });
    const first = h.session.submit('insert', 'a', { echo: 'a' });
    const child = h.session.submit('insert', 'b', { after: first, echo: 'b' });
    const grandchild = h.session.submit('insert', 'c', { after: child, echo: 'c' });
    const independent = h.session.submit('insert', 'd', { echo: 'd' });
    h.pump();
    h.ack(0, [{ seq: 1, status: 'refused', reason: 'locked' }], { version: 0 });
    expect(await first.presented).toBe('dropped');
    expect(await child.ack).toMatchObject({ status: 'cancelled', reason: 'dependency-refused' });
    expect(await grandchild.presented).toBe('dropped');
    expect(child.seq).toBeUndefined();
    h.pump();
    expect(independent.seq).toBe(2);
    expect(h.calls[1]!.envelope.payload).toBe('d');
    expect(h.drops.flatMap((drop) => drop.values)).toEqual(['a', 'b', 'c']);
    const lateChild = h.session.submit('insert', 'e', { after: first });
    expect(await lateChild.ack).toMatchObject({ status: 'cancelled', reason: 'dependency-refused' });
    await h.session.dispose();
  });

  it('queued echo attachment preserves coalescing', async () => {
    const h = harness({ coalesce(previous, next) {
      if (!!previous.options.echo !== !!next.options.echo) return undefined;
      return { ...previous, payload: previous.payload + next.payload, options: { ...previous.options,
        echo: previous.options.echo === undefined ? undefined : previous.options.echo + (next.options.echo ?? '') } };
    } }, true);
    const first = h.session.submit('insert', 'a', { coalesceKey: 'typing' });
    const alias = h.session.submit('insert', 'b', { coalesceKey: 'typing', after: first });
    const previews: Array<{ payload: string; count: number }> = [];
    expect(h.session.attachEcho!([first, alias], (draft, count) => {
      previews.push({ payload: draft.payload, count });
      return draft.payload;
    })).toBe(true);
    expect(previews).toEqual([{ payload: 'ab', count: 2 }]);
    const later = h.session.submit('insert', 'c', { coalesceKey: 'typing', after: alias, echo: 'c' });
    h.pump();
    expect(h.calls.map(({ envelope }) => envelope.payload)).toEqual(['abc']);
    h.ack(0);
    await later.ack;
    h.events!.push({ coversThroughSeq: 1, revision: 1, payload: 'covered' });
    expect(h.frames.at(-1)?.retired).toEqual(['abc']);
    expect(await first.presented).toBe('presented');
    await h.session.dispose();
  });

  it('in-flight echo attachment retires on coverage', async () => {
    const h = harness({}, true);
    const ticket = h.session.submit('insert', 'a');
    h.pump();
    expect(h.session.attachEcho!([ticket], (draft) => draft.payload)).toBe(true);
    h.ack(0);
    await ticket.ack;
    h.events!.push({ coversThroughSeq: 1, revision: 1, payload: 'covered' });
    expect(h.frames.at(-1)).toEqual({ payload: 'covered', retired: ['a'], remaining: [] });
    expect(await ticket.presented).toBe('presented');
    expect(h.session.attachEcho!([ticket], () => 'again')).toBe(false);
    await h.session.dispose();
  });

  it('in-flight echo attachment drops on failure', async () => {
    const h = harness({ maxRecoveryAttempts: 0 });
    const ticket = h.session.submit('insert', 'a');
    h.pump();
    expect(h.session.attachEcho!([ticket], (draft) => draft.payload)).toBe(true);
    h.fail(new SessionFailure('crash', 'worker stopped'));
    expect(await ticket.presented).toBe('dropped');
    expect(h.drops.flatMap(({ values }) => values)).toEqual(['a']);
    await h.session.dispose();
  });

  it('echo attachment declines partial batches', async () => {
    const paints: string[][] = [];
    const h = harness({ echo: { paint(entries) { paints.push(entries.map(({ value }) => value)); }, drop() {}, present() {} } });
    const first = h.session.submit('insert', 'a');
    const second = h.session.submit('insert', 'b');
    h.pump();
    expect(h.session.attachEcho!([first, second], (draft) => draft.payload === 'a' ? 'a' : undefined)).toBe(false);
    expect(paints).toEqual([]);
    await h.session.dispose();
  });

  it('echo attachment rejects the whole batch when a preview re-merges a draft', async () => {
    const paints: string[][] = [];
    const h = harness({
      coalesce: (previous, next) => ({ ...previous, payload: previous.payload + next.payload }),
      echo: { paint(entries) { paints.push(entries.map(({ value }) => value)); }, drop() {}, present() {} },
    });
    const first = h.session.submit('insert', 'a');
    const second = h.session.submit('insert', 'b', { coalesceKey: 'typing' });
    const previewed: Array<{ echo?: string }> = [];
    expect(h.session.attachEcho!([first, second], (draft) => {
      previewed.push(draft.options);
      if (draft.payload === 'b') h.session.submit('insert', 'c', { coalesceKey: 'typing', after: second });
      return draft.payload;
    })).toBe(false);
    expect(previewed).toHaveLength(2);
    expect(previewed.every((options) => options.echo === undefined)).toBe(true);
    expect(paints).toEqual([]);
    expect(h.session.attachEcho!([first, second], (draft) => draft.payload)).toBe(true);
    expect(paints).toEqual([['a', 'bc']]);
    await h.session.dispose();
  });

  it('echo attachment past the echo cap is declined without evicting', async () => {
    const paints: string[][] = [];
    const h = harness({
      maxEchoEntries: 2, echoOverflow: 'evict-oldest',
      echo: { paint(entries) { paints.push(entries.map(({ value }) => value)); }, drop() {}, present() {} },
    });
    const first = h.session.submit('insert', 'a');
    h.session.submit('insert', 'b', { echo: 'b' });
    h.session.submit('insert', 'c', { echo: 'c' });
    const painted = paints.length;
    expect(h.session.attachEcho!([first], (draft) => draft.payload)).toBe(false);
    expect(paints.length).toBe(painted);
    await h.session.dispose();
  });

  it('coalesces alias tickets at send and keeps flush boundaries separate', async () => {
    const h = harness({ coalesce: (prev, next) => ({ ...next, payload: prev.payload + next.payload }) });
    const first = h.session.submit('insert', 'a', { coalesceKey: 'typing', echo: 'a' });
    const alias = h.session.submit('insert', 'b', { coalesceKey: 'typing', echo: 'ab' });
    const flushed = h.session.flush();
    const later = h.session.submit('insert', 'c', { coalesceKey: 'typing' });
    h.pump();
    expect(first.seq).toBe(1);
    expect(alias.seq).toBe(1);
    expect(later.seq).toBe(2);
    h.ack(0, undefined, { frame: { coversThroughSeq: 1, revision: 1, payload: 'ab' } });
    await flushed;
    expect(await alias.ack).toEqual(await first.ack);
    expect(await alias.presented).toBe('presented');
    expect(h.frames[0]!.retired).toEqual(['ab']);
    await h.session.dispose();
  });

  it('holds coalescing drafts while one is in flight and sends the rest of the burst as one', async () => {
    const h = harness({
      coalesce: (prev, next) => ({ ...next, payload: prev.payload + next.payload }),
      holdForCoalescing: (draft) => draft.kind === 'insert',
    });
    const first = h.session.submit('insert', 'a', { coalesceKey: 'typing' });
    h.pump();
    const burst = [...'bcd'].map((text) => h.session.submit('insert', text, { coalesceKey: 'typing' }));
    const remote = h.session.submit('remote', 'r', { lane: 'collab' });
    h.pump();
    expect(h.calls.map(({ envelope }) => envelope.payload)).toEqual(['a', 'r']);
    h.ack(0);
    await first.ack;
    h.pump();
    expect(h.calls.map(({ envelope }) => envelope.payload)).toEqual(['a', 'r', 'bcd']);
    expect(burst.map((ticket) => ticket.seq)).toEqual([3, 3, 3]);
    h.ack(1);
    h.ack(2);
    expect((await Promise.all(burst.map((ticket) => ticket.ack))).every(({ status }) => status === 'applied')).toBe(true);
    await remote.ack;
    await h.session.dispose();
  });

  it('replaces merged echo by group identity without scanning aliases', async () => {
    const h = harness({ coalesce: (_previous, next) => next });
    let reads = 0;
    for (let index = 0; index < 128; index += 1) {
      const ticket = h.session.submit('insert', 'a', { coalesceKey: 'typing', echo: 'a' });
      const id = ticket.id;
      Object.defineProperty(ticket, 'id', { get() { reads += 1; return id; } });
    }
    expect(reads).toBeLessThan(128 * 8);
    await h.session.dispose();
  });

  it('reconciles frames by coverage, waits for installation, and requests full frames', async () => {
    const h = harness();
    const first = h.session.submit('insert', 'a', { echo: 'a' });
    const second = h.session.submit('remote', 'b', { lane: 'collab' });
    h.pump();
    h.ack(1, [{ seq: 2, status: 'applied' }], { frame: { coversThroughSeq: 1, revision: 1, payload: 'first' } });
    await second.ack;
    let firstPresented = false;
    void first.presented.then(() => { firstPresented = true; });
    await Promise.resolve();
    expect(firstPresented).toBe(false);
    let presented = false;
    void second.presented.then(() => { presented = true; });
    h.ack(0, [{ seq: 1, status: 'applied' }], { frame: { coversThroughSeq: 2, revision: 3, baseRevision: 2, payload: 'delta' } });
    await first.ack;
    expect(h.fullRequests).toBe(1);
    expect(presented).toBe(false);
    h.full.resolve({ ...h.connection.initial, throughSeq: 2, version: 2, outcomes: [],
      frame: { coversThroughSeq: 2, revision: 4, payload: 'full' } });
    expect(await h.session.whenPresented(2)).toBe('presented');
    expect(await second.presented).toBe('presented');
    expect(h.session.state.presentedThroughSeq).toBe(2);
    await h.session.dispose();
  });

  it('disposes idempotently and ignores late replies/listener reentrancy', async () => {
    const h = harness({ maxInFlight: 2 });
    const first = h.session.submit('insert', 'a');
    const sent = h.session.submit('insert', 'b');
    const queued = h.session.submit('insert', 'c');
    h.pump();
    let notifications = 0;
    h.session.subscribe(() => { notifications += 1; void h.session.dispose(); });
    h.ack(0, undefined, { frame: { coversThroughSeq: 1, revision: 1, payload: 'a' } });
    await first.ack;
    const disposal = h.session.dispose();
    expect(h.session.dispose()).toBe(disposal);
    await disposal;
    expect(await first.presented).toBe('presented');
    expect(await sent.ack).toMatchObject({ status: 'cancelled', reason: 'disposed' });
    expect(await queued.ack).toMatchObject({ status: 'cancelled', reason: 'disposed' });
    expect(await queued.presented).toBe('dropped');
    h.ack(1, undefined, { version: 99, selection: 'late' });
    await Promise.resolve();
    h.pump();
    expect(h.calls).toHaveLength(2);
    expect(h.session.state.version).toBe(1);
    expect(h.session.state.selection).toBeUndefined();
    expect(h.disposals).toBe(1);
    expect(notifications).toBe(1);
    expect(await h.session.submit('insert', 'late').ack).toMatchObject({ status: 'cancelled', reason: 'disposed' });
  });

  it('requests missing coverage after a second mismatch during a full-frame request', async () => {
    const requests: Array<ReturnType<typeof deferred<Ack>>> = [];
    const h = harness({ requestFullFrame() {
      const reply = deferred<Ack>();
      requests.push(reply);
      return reply.promise;
    } });
    const first = h.session.submit('insert', 'a');
    const second = h.session.submit('insert', 'b');
    h.pump();
    h.ack(0, undefined, { frame: { coversThroughSeq: 1, revision: 1, baseRevision: 9, payload: 'a' } });
    await first.ack;
    h.ack(1, undefined, { frame: { coversThroughSeq: 2, revision: 2, baseRevision: 9, payload: 'b' } });
    await second.ack;
    expect(requests).toHaveLength(1);
    let secondStatus: string | undefined;
    void second.presented.then((status) => { secondStatus = status; });
    requests[0]!.resolve({ ...h.connection.initial, throughSeq: 1, version: 1,
      frame: { coversThroughSeq: 1, revision: 3, payload: 'first full' } });
    expect(await first.presented).toBe('presented');
    expect(requests).toHaveLength(2);
    expect(secondStatus).toBeUndefined();
    requests[1]!.resolve({ ...h.connection.initial, throughSeq: 2, version: 2,
      frame: { coversThroughSeq: 2, revision: 4, payload: 'second full' } });
    expect(await h.session.whenPresented(2)).toBe('presented');
    expect(await second.presented).toBe('presented');
    expect(h.session.state.presentedThroughSeq).toBe(2);
    await h.session.dispose();
  });

  it('avoids a redundant full-frame request when intervening coverage resolves the mismatch', async () => {
    const h = harness();
    const first = h.session.submit('insert', 'a');
    const second = h.session.submit('insert', 'b');
    h.pump();
    h.ack(0, undefined, { frame: { coversThroughSeq: 1, revision: 1, baseRevision: 9, payload: 'a' } });
    await first.ack;
    h.ack(1, undefined, { frame: { coversThroughSeq: 2, revision: 2, payload: 'full' } });
    expect(await second.presented).toBe('presented');
    h.full.resolve({ ...h.connection.initial, throughSeq: 1, version: 1,
      frame: { coversThroughSeq: 1, revision: 3, payload: 'obsolete full' } });
    await Promise.resolve();
    expect(h.fullRequests).toBe(1);
    expect(await first.presented).toBe('presented');
    await h.session.dispose();
  });

  it('seeds a replacement epoch from the previous sequence high-water mark', async () => {
    const h = harness({ connection: {
      initial: { session: 'replacement', incarnation: 0, epoch: 2, version: 0, throughSeq: 41, outcomes: [] },
      client: {
        call: { edit: () => new Promise<Ack>(() => {}), checkpoint: () => new Promise<Checkpoint>(() => {}) },
        failure: undefined, onFailure: () => () => {}, dispose: () => Promise.resolve(),
      },
    } });
    const ticket = h.session.submit('insert', 'new');
    h.pump();
    expect(ticket.seq).toBe(42);
    expect(h.session.state.epoch).toBe(2);
    await h.session.dispose();
  });

  it('waits for outcomes after covering frames and drops late refusals', async () => {
    for (const status of ['applied', 'refused'] as const) {
      const h = harness();
      const future = h.session.whenPresented(1);
      const first = h.session.submit('insert', 'a');
      const second = h.session.submit('insert', 'b');
      h.pump();
      let presented = false;
      let covered = false;
      void first.presented.then(() => { presented = true; });
      const waiter = h.session.whenPresented(1);
      void waiter.then(() => { covered = true; });
      h.ack(1, undefined, { frame: { coversThroughSeq: 2, revision: 2, payload: 'frame' } });
      await second.ack;
      await Promise.resolve();
      expect(presented).toBe(false);
      expect(covered).toBe(false);
      h.ack(0, [{ seq: 1, status }], { version: status === 'applied' ? 1 : 0 });
      expect(await first.presented).toBe(status === 'applied' ? 'presented' : 'dropped');
      expect(await waiter).toBe(status === 'applied' ? 'presented' : 'dropped');
      expect(await future).toBe(status === 'applied' ? 'presented' : 'dropped');
      await h.session.dispose();
    }
  });

  it('reconciles selection and edit state with independent ack coverage', async () => {
    const h = harness();
    const first = h.session.submit('insert', 'a');
    const second = h.session.submit('insert', 'b');
    const third = h.session.submit('insert', 'c');
    h.pump();
    h.ack(2, undefined, { editState: 'latest' });
    await third.ack;
    h.ack(1, undefined, { selection: 'selected' });
    await second.ack;
    h.ack(0, undefined, { selection: 'old', editState: 'old' });
    await first.ack;
    expect(h.session.state).toMatchObject({ throughSeq: 3, version: 3, selection: 'selected', editState: 'latest' });
    await h.session.dispose();
  });

  it('dispatches captured flush submissions before later input', async () => {
    const h = harness({ maxInFlight: 1 });
    const first = h.session.submit('background', 'a', { lane: 'background' });
    const flushed = h.session.flush();
    const second = h.session.submit('background', 'b', { lane: 'background' });
    const flushedAgain = h.session.flush();
    const later = h.session.submit('input', 'c');
    h.pump();
    expect(h.calls.map((call) => call.envelope.payload)).toEqual(['a']);
    h.ack(0);
    await flushed;
    h.pump();
    expect(second.seq).toBe(2);
    expect(later.seq).toBeUndefined();
    h.ack(1);
    await flushedAgain;
    h.pump();
    expect(first.seq).toBe(1);
    expect(later.seq).toBe(3);
    await h.session.dispose();
  });

  it('settles tickets when frame and cleanup hooks throw', async () => {
    let clock = 0;
    const samples: Array<{ hook: string; elapsedMs: number; eventTotalMs: number }> = [];
    const h = harness({ echo: {
      paint() { clock += 3; throw new Error('paint'); },
      present() { clock += 3; throw new Error('present'); },
      drop() { clock += 3; throw new Error('drop'); },
    }, now: () => clock, onEchoOverrun: (sample) => { samples.push(sample); } });
    const first = h.session.submit('insert', 'a', { echo: 'a' });
    const second = h.session.submit('insert', 'b', { echo: 'b' });
    h.pump();
    h.ack(0, undefined, { frame: { coversThroughSeq: 1, revision: 1, payload: 'a' } });
    expect(await first.ack).toMatchObject({ status: 'applied' });
    expect(await first.presented).toBe('dropped');
    expect(h.disposals).toBe(0);
    h.ack(1, [{ seq: 2, status: 'refused', reason: 'locked' }], { version: 1 });
    expect(await second.ack).toMatchObject({ status: 'refused', reason: 'locked' });
    expect(await second.presented).toBe('dropped');
    expect(samples.map((sample) => sample.hook)).toEqual(['paint', 'paint', 'present', 'drop']);
    expect(samples.every((sample) => sample.elapsedMs === 3 && sample.eventTotalMs === 3)).toBe(true);
    expect(h.session.state.presentedThroughSeq).toBe(1);
    const third = h.session.submit('insert', 'c');
    h.pump();
    expect(third.seq).toBe(3);
    h.ack(2, undefined, { frame: { coversThroughSeq: 3, revision: 3, payload: 'c' } });
    expect(await third.ack).toMatchObject({ status: 'applied' });
    expect(await third.presented).toBe('dropped');
    await h.session.dispose();
  });

  it('drops late outcomes and coalesced aliases covered by a throwing present hook', async () => {
    const h = harness({
      coalesce: (_previous, next) => next,
      echo: { paint() {}, drop() {}, present() { throw new Error('present'); } },
    });
    const future = h.session.whenPresented(1);
    const first = h.session.submit('insert', 'a', { coalesceKey: 'typing' });
    const alias = h.session.submit('insert', 'b', { coalesceKey: 'typing' });
    const second = h.session.submit('insert', 'c');
    h.pump();
    h.ack(1, undefined, { frame: { coversThroughSeq: 2, revision: 2, payload: 'full' } });
    expect(await second.presented).toBe('dropped');
    h.ack(0);
    expect(await first.ack).toMatchObject({ status: 'applied' });
    expect(await first.presented).toBe('dropped');
    expect(await alias.presented).toBe('dropped');
    expect(await future).toBe('dropped');
    expect(await h.session.whenPresented(1)).toBe('dropped');
    expect(h.disposals).toBe(0);
    await h.session.dispose();
  });
});
