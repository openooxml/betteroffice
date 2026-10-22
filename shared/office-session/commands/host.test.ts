import { describe, expect, it } from 'bun:test';
import { isTransferResult, transferable, type TransferResult } from '../protocol';
import { createResidentScheduler } from '../scheduler';
import { SessionFailure } from '../types';
import { createCommandHost } from './host';
import {
  COMMAND_OUTCOME_UNKNOWN, isCommandOutcomeUnknownError,
  type ApplyContext, type CommandAck, type CommandEnvelope,
  type CommandHost, type CommandHostOptions, type FrameEnvelope,
} from './types';

type Payload = { delta: number };
type Ack = CommandAck<string, boolean, Uint8Array>;
type Host = CommandHost<Payload, string, boolean, Uint8Array>;
type Options = CommandHostOptions<Payload, string, boolean, Uint8Array>;

function deferred<T>() {
  let resolve!: (value: T | PromiseLike<T>) => void;
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}

function envelope(seq: number, overrides: Partial<CommandEnvelope<Payload>> = {}): CommandEnvelope<Payload> {
  return {
    session: 'document', incarnation: 2, seq, basis: { epoch: 3, version: 10 },
    kind: 'insert', payload: { delta: 1 }, ...overrides,
  };
}

function value<T>(reply: T | TransferResult<T>): T {
  return isTransferResult(reply) ? (reply as TransferResult<T>).value : reply;
}

function cloneTransfer<T>(reply: TransferResult<T>): T {
  const runtime = globalThis as typeof globalThis & {
    structuredClone?: <V>(input: V, options: { transfer: Transferable[] }) => V;
  };
  if (typeof runtime.structuredClone !== 'function') throw new Error('Missing native structuredClone');
  return runtime.structuredClone(reply.value, { transfer: reply.transfer });
}

function model(overrides: Partial<Options> = {}) {
  let version = 10;
  const calls: number[] = [];
  const contexts: ApplyContext[] = [];
  const frames: number[] = [];
  const options: Options = {
    session: 'document', incarnation: 2, epoch: 3,
    getVersion: () => version,
    apply(command, context) {
      calls.push(command.seq);
      contexts.push(context);
      if (command.kind === 'refuse') return { status: 'refused', reason: 'locked', affectedSeqs: [99] };
      if (command.kind === 'noop') return { status: 'noop' };
      version += command.payload.delta;
      return { status: 'applied' };
    },
    frame(input) {
      frames.push(input.coversThroughSeq);
      return { coversThroughSeq: input.coversThroughSeq, revision: input.version, payload: new Uint8Array([version]) };
    },
    snapshot: () => new Uint8Array([version]),
    ...overrides,
  };
  return {
    host: createCommandHost(options), calls, contexts, frames,
    get version() { return version; },
    setVersion(next: number) { version = next; },
  };
}

async function edit(host: Host, command: CommandEnvelope<Payload>): Promise<Ack> {
  host.policy(command);
  return value(await host.handler(null, command));
}

describe('command host', () => {
  it('defers opted-in selection noop frames until requested after the acknowledgement', async () => {
    const held = deferred<FrameEnvelope<Uint8Array>>();
    let frameCalls = 0;
    const s = model({
      apply: () => ({ status: 'noop', selection: 'paragraph-2', editState: true }),
      deferFrame: (command) => command.kind === 'select',
      frame() { frameCalls += 1; return held.promise; },
    });
    let acknowledged = false;
    const pending = edit(s.host, envelope(1, { kind: 'select' })).then((ack) => { acknowledged = true; return ack; });
    for (let turn = 0; turn < 12; turn++) await Promise.resolve();
    expect(acknowledged).toBe(true);
    expect(frameCalls).toBe(0);
    expect(await pending).toMatchObject({ throughSeq: 1, version: 10, selection: 'paragraph-2', editState: true, outcomes: [{ status: 'noop' }] });
    expect((await pending).frame).toBeUndefined();
    const pushed = s.host.frame();
    held.resolve({ coversThroughSeq: 1, revision: 10, payload: new Uint8Array([10]) });
    expect(value(await pushed)).toMatchObject({ coversThroughSeq: 1, revision: 10 });
    expect(frameCalls).toBe(1);
  });

  it('keeps non-selection commands and opted-out hosts on the covering frame path', async () => {
    for (const optIn of [false, true]) {
      const s = model({ ...(optIn ? { deferFrame: (command) => command.kind === 'select' } satisfies Partial<Options> : {}) });
      const ack = await edit(s.host, envelope(1, { kind: optIn ? 'noop' : 'select' }));
      expect(ack.frame?.coversThroughSeq).toBe(1);
      expect(s.frames).toEqual([1]);
    }
    const s = model({ deferFrame: () => true });
    expect((await edit(s.host, envelope(1))).frame?.coversThroughSeq).toBe(1);
  });

  it('acknowledges edits without a frame when the hook returns undefined and keeps later operations usable', async () => {
    for (const asynchronous of [false, true]) {
      let skip = true;
      const s = model({ frame(input) {
        const candidate = skip ? undefined : {
          coversThroughSeq: input.coversThroughSeq, revision: input.version, payload: new Uint8Array([input.version]),
        };
        return asynchronous ? Promise.resolve(candidate) : candidate;
      } });
      const first = await edit(s.host, envelope(1));
      expect(first).toMatchObject({ throughSeq: 1, version: 11, outcomes: [{ seq: 1, status: 'applied' }] });
      expect(Object.prototype.hasOwnProperty.call(first, 'frame')).toBe(false);
      skip = false;
      const next = await edit(s.host, envelope(2));
      expect(next.frame).toMatchObject({ coversThroughSeq: 2, revision: 12, payload: new Uint8Array([12]) });
      expect(value(await s.host.checkpoint(null))).toMatchObject({ throughSeq: 2, version: 12 });
      expect(value(await s.host.frame())).toMatchObject({ coversThroughSeq: 2, revision: 12 });
      expect(s.calls).toEqual([1, 2]);
    }
  });

  it('rejects an explicit missing frame with a plain Error and permits a later frame request', async () => {
    for (const asynchronous of [false, true]) {
      let skip = true;
      const s = model({ frame(input) {
        const candidate = skip ? undefined : {
          coversThroughSeq: input.coversThroughSeq, revision: input.version, payload: new Uint8Array([input.version]),
        };
        return asynchronous ? Promise.resolve(candidate) : candidate;
      } });
      const error = await s.host.frame().catch((cause: unknown) => cause);
      expect(error).toBeInstanceOf(Error);
      expect(error).not.toBeInstanceOf(SessionFailure);
      expect((error as Error).constructor).toBe(Error);
      skip = false;
      expect(value(await s.host.frame())).toMatchObject({ coversThroughSeq: 0, revision: 10 });
      const ack = await edit(s.host, envelope(1));
      expect(ack.frame).toMatchObject({ coversThroughSeq: 1, revision: 11 });
      expect(s.calls).toEqual([1]);
    }
  });

  it('builds frames after queued and admitted edits at their settled boundary', async () => {
    const entered = deferred<void>();
    const held = deferred<void>();
    let version = 10;
    const s = model({
      getVersion: () => version,
      async apply(command) {
        if (command.seq === 1) { entered.resolve(); await held.promise; }
        version += 1;
        return { status: 'applied' };
      },
    });
    const first = edit(s.host, envelope(1));
    await entered.promise;
    const second = edit(s.host, envelope(2));
    s.host.policy(envelope(4));
    let installed = false;
    const pushed = s.host.frame().then((reply) => { installed = true; return value(reply); });
    await Promise.resolve();
    expect(installed).toBe(false);
    held.resolve();
    await first;
    await second;
    expect(installed).toBe(false);
    const last = s.host.handler(null, envelope(4));
    const candidate = await pushed;
    await last;
    expect(candidate).toMatchObject({ coversThroughSeq: 4, revision: 13 });
    expect(s.frames).toEqual([4, 4]);
  });

  it('preserves frame transfers and rejects frame requests while poisoned', async () => {
    const bytes = new Uint8Array([4]);
    const s = model({ frame: (input) => transferable({
      coversThroughSeq: input.coversThroughSeq, revision: input.version, payload: bytes,
    }, [bytes.buffer]) });
    const reply = await s.host.frame();
    expect(isTransferResult(reply)).toBe(true);
    if (!isTransferResult(reply)) throw new Error('Missing frame transfer');
    expect(reply.transfer).toEqual([bytes.buffer]);
    expect(reply.value).toMatchObject({ coversThroughSeq: 0, revision: 10 });
    const poisoned = model({ apply() { throw new Error('apply failed'); } });
    const error = await edit(poisoned.host, envelope(1)).catch((cause: unknown) => cause);
    expect(await poisoned.host.frame().catch((cause: unknown) => cause)).toBe(error);
  });

  it('recovers more than 4096 alternating applied and noop replay sequences without a checkpoint', async () => {
    const first = model();
    const commands: CommandEnvelope<Payload>[] = [];
    const applied: number[] = [];
    for (let seq = 1; seq <= 8194; seq += 1) {
      const command = envelope(seq, {
        kind: seq % 2 === 1 ? 'insert' : 'noop',
        basis: { epoch: 3, version: first.version }, expectVersion: true,
      });
      commands.push(command);
      const ack = await edit(first.host, command);
      if (ack.outcomes[0]!.status === 'applied') applied.push(seq);
    }
    expect(applied).toHaveLength(4097);
    const recovered = model({ replayAppliedSeqs: applied });
    for (const command of commands) {
      const ack = await edit(recovered.host, { ...command, replay: true });
      expect(ack.outcomes[0]!.status).toBe(command.kind === 'noop' ? 'noop' : 'applied');
    }
    expect(recovered.version).toBe(first.version);
    expect(recovered.calls).toHaveLength(8194);
  });

  it('uses the supplied selection equality hook', async () => {
    const comparisons: Array<[string | undefined, string | undefined]> = [];
    const s = model({
      apply: () => ({ status: 'noop', selection: 'selected' }),
      selectionEquals(previous, next) { comparisons.push([previous, next]); return false; },
    });
    expect((await edit(s.host, envelope(1))).selection).toBe('selected');
    expect((await edit(s.host, envelope(2))).selection).toBe('selected');
    expect(comparisons).toEqual([['selected', 'selected']]);
  });

  it('uses the supplied edit state equality hook', async () => {
    const comparisons: Array<[boolean | undefined, boolean | undefined]> = [];
    const s = model({
      apply: (command) => ({ status: 'noop', editState: command.seq === 1 }),
      editStateEquals(previous, next) { comparisons.push([previous, next]); return true; },
    });
    expect((await edit(s.host, envelope(1))).editState).toBe(true);
    expect((await edit(s.host, envelope(2)))).not.toHaveProperty('editState');
    expect(comparisons).toEqual([[true, false]]);
  });

  it('applies increasing seqs serially, permits gaps, and dedupes retained or evicted outcomes', async () => {
    const s = model({ retainedOutcomes: 2 });
    const first = await edit(s.host, envelope(1));
    await edit(s.host, envelope(4));
    const retained = await edit(s.host, envelope(1));
    expect(retained.outcomes).toEqual(first.outcomes);
    await edit(s.host, envelope(8));
    expect((await edit(s.host, envelope(1))).outcomes).toEqual([{ seq: 1, status: 'duplicate' }]);
    expect((await edit(s.host, envelope(6))).outcomes).toEqual([{ seq: 6, status: 'duplicate' }]);
    expect(s.calls).toEqual([1, 4, 8]);
    expect(s.version).toBe(13);
  });

  it('returns duplicate for a pruned outcome when prunedDuplicateReason is absent', async () => {
    const s = model({ retainedOutcomes: 1 });
    await edit(s.host, envelope(1));
    await edit(s.host, envelope(2));
    const duplicate = await edit(s.host, envelope(1));
    expect(duplicate.outcomes).toEqual([{ seq: 1, status: 'duplicate' }]);
    expect(duplicate.throughSeq).toBe(2);
    expect(duplicate.frame?.coversThroughSeq).toBe(2);
    expect(s.calls).toEqual([1, 2]);
  });

  it('refuses pruned duplicates with a frame without changing the watermark, cached outcomes, or dependencies', async () => {
    const s = model({ retainedOutcomes: 1, prunedDuplicateReason: 'stale-basis' });
    await edit(s.host, envelope(1));
    const second = await edit(s.host, envelope(2));
    expect(second.outcomes).toEqual([{ seq: 2, status: 'applied' }]);
    const pruned = await edit(s.host, envelope(1));
    expect(pruned.outcomes).toEqual([
      { seq: 1, status: 'refused', reason: 'stale-basis', committed: false },
    ]);
    expect(pruned.throughSeq).toBe(2);
    expect(pruned.version).toBe(12);
    expect(pruned.frame).toEqual({ coversThroughSeq: 2, revision: 12, payload: new Uint8Array([12]) });
    expect((await edit(s.host, envelope(2))).outcomes).toEqual(second.outcomes);
    const third = await edit(s.host, envelope(3, { afterSeq: 1 }));
    expect(third.outcomes).toEqual([{ seq: 3, status: 'applied' }]);
    expect(third.throughSeq).toBe(3);
    expect(third.version).toBe(13);
    expect(s.calls).toEqual([1, 2, 3]);
  });

  it('rejects empty or non-string prunedDuplicateReason at construction', () => {
    for (const reason of ['', 7]) {
      let calls = 0;
      expect(() => model({
        prunedDuplicateReason: reason as string,
        apply() { calls += 1; return { status: 'noop' }; },
      })).toThrow(TypeError);
      expect(calls).toBe(0);
    }
  });

  it('acknowledges non-committing applied state without a version bump and admits its dependent', async () => {
    let version = 10;
    const calls: number[] = [];
    const s = model({
      getVersion: () => version,
      apply(command) {
        calls.push(command.seq);
        if (command.seq === 1) {
          return { status: 'applied', committed: false, selection: 'selected', editState: true };
        }
        version += 1;
        return { status: 'applied' };
      },
    });
    const first = await edit(s.host, envelope(1));
    expect(first.outcomes).toEqual([{ seq: 1, status: 'applied', committed: false }]);
    expect(first.throughSeq).toBe(1);
    expect(first.version).toBe(10);
    expect(first.selection).toBe('selected');
    expect(first.editState).toBe(true);
    const dependent = await edit(s.host, envelope(2, { afterSeq: 1 }));
    expect(dependent.outcomes).toEqual([{ seq: 2, status: 'applied' }]);
    expect(dependent.throughSeq).toBe(2);
    expect(dependent.version).toBe(11);
    expect(calls).toEqual([1, 2]);
  });

  it('poisons non-committing applied with a version bump and plain applied without a bump', async () => {
    for (const nonCommitting of [false, true]) {
      let version = 10;
      let calls = 0;
      const s = model({
        getVersion: () => version,
        apply() {
          calls += 1;
          if (nonCommitting) {
            version += 1;
            return { status: 'applied', committed: false };
          }
          return { status: 'applied' };
        },
      });
      const error = await edit(s.host, envelope(1)).catch((error: unknown) => error);
      expect(isCommandOutcomeUnknownError(error)).toBe(true);
      if (!isCommandOutcomeUnknownError(error)) throw new Error('Missing unknown-outcome error');
      expect(error.name).toBe(COMMAND_OUTCOME_UNKNOWN);
      expect(error.refusal).toEqual({ seq: 1 });
      expect(error.message).toContain('violated the command version contract');
      expect(await edit(s.host, envelope(2)).catch((error: unknown) => error)).toBe(error);
      expect(await s.host.checkpoint(null).catch((error: unknown) => error)).toBe(error);
      expect(calls).toBe(1);
      expect(s.frames).toEqual([]);
    }
  });

  it('acknowledges cancelled state with or without a reason and cancels its dependent', async () => {
    for (const reason of [undefined, 'empty-composition']) {
      let calls = 0;
      const s = model({
        apply() {
          calls += 1;
          return { status: 'cancelled', ...(reason === undefined ? {} : { reason }), selection: 'selected', editState: false };
        },
      });
      const cancelled = await edit(s.host, envelope(1));
      expect(cancelled.outcomes).toEqual([
        { seq: 1, status: 'cancelled', committed: false, ...(reason === undefined ? {} : { reason }) },
      ]);
      expect(cancelled.throughSeq).toBe(1);
      expect(cancelled.version).toBe(10);
      expect(cancelled.selection).toBe('selected');
      expect(cancelled.editState).toBe(false);
      const dependent = await edit(s.host, envelope(2, { afterSeq: 1 }));
      expect(dependent.outcomes).toEqual([
        { seq: 2, status: 'cancelled', reason: 'dependency-refused', committed: false },
      ]);
      expect(dependent.throughSeq).toBe(2);
      expect(dependent.version).toBe(10);
      expect(calls).toBe(1);
    }
  });

  it('poisons cancelled results with a version bump or a non-string reason after a valid cancellation', async () => {
    for (const bump of [true, false]) {
      let version = 10;
      const calls: number[] = [];
      const s = model({
        getVersion: () => version,
        apply(command) {
          calls.push(command.seq);
          if (command.seq === 1) return { status: 'cancelled' };
          if (bump) {
            version += 1;
            return { status: 'cancelled' };
          }
          return { status: 'cancelled', reason: 7 as unknown as string };
        },
      });
      const valid = await edit(s.host, envelope(1));
      expect(valid.outcomes).toEqual([{ seq: 1, status: 'cancelled', committed: false }]);
      const error = await edit(s.host, envelope(2)).catch((error: unknown) => error);
      expect(isCommandOutcomeUnknownError(error)).toBe(true);
      if (!isCommandOutcomeUnknownError(error)) throw new Error('Missing unknown-outcome error');
      expect(error.name).toBe(COMMAND_OUTCOME_UNKNOWN);
      expect(error.refusal).toEqual({ seq: 2 });
      expect(error.message).toContain(bump ? 'violated the command version contract' : 'Invalid format apply result');
      expect(await edit(s.host, envelope(3)).catch((error: unknown) => error)).toBe(error);
      expect(await s.host.checkpoint(null).catch((error: unknown) => error)).toBe(error);
      expect(calls).toEqual([1, 2]);
      expect(s.frames).toEqual([1]);
    }
  });

  it('maps applied/noop/refused/cancelled outcomes to throughSeq and worker version', async () => {
    const s = model();
    for (const [seq, kind, afterSeq, status, version] of [
      [1, 'insert', undefined, 'applied', 11],
      [2, 'noop', undefined, 'noop', 11],
      [3, 'refuse', undefined, 'refused', 11],
      [4, 'insert', 3, 'cancelled', 11],
      [5, 'insert', undefined, 'applied', 12],
    ] as const) {
      const ack = await edit(s.host, envelope(seq, { kind, afterSeq }));
      expect(ack).toMatchObject({ session: 'document', incarnation: 2, epoch: 3, throughSeq: seq, version });
      expect(ack.outcomes[0]).toMatchObject({ seq, status });
      expect(ack.frame?.coversThroughSeq).toBe(seq);
      if (status === 'refused') expect(ack.outcomes[0]).toEqual({ seq, status, reason: 'locked', committed: false, affectedSeqs: [99] });
      if (status === 'cancelled') expect(ack.outcomes[0]?.reason).toBe('dependency-refused');
    }
    expect(s.calls).toEqual([1, 2, 3, 5]);
  });

  it('refuses wrong identity, stale epoch, and strict-version mismatch before format apply', async () => {
    const s = model();
    for (const [command, reason, throughSeq] of [
      [envelope(100, { session: 'other' }), 'wrong-session', 0],
      [envelope(100, { incarnation: 1 }), 'stale-incarnation', 0],
      [envelope(1, { basis: { epoch: 2, version: 10 } }), 'stale-epoch', 1],
      [envelope(2, { basis: { epoch: 3, version: 9 }, expectVersion: true }), 'stale-version', 2],
    ] as const) {
      const ack = await edit(s.host, command);
      expect(ack.throughSeq).toBe(throughSeq);
      expect(ack.outcomes).toEqual([{ seq: command.seq, status: 'refused', reason, committed: false }]);
      if (reason !== 'stale-version') expect(ack.frame).toBeUndefined();
    }
    expect(s.calls).toEqual([]);
    expect((await edit(s.host, envelope(3))).outcomes[0]?.status).toBe('applied');
    expect((await edit(s.host, envelope(4, { afterSeq: 1 }))).outcomes[0]?.status).toBe('cancelled');
  });

  it('passes confirmed basis/current version for anchored rebasing and pushes changed selection/editState', async () => {
    let version = 10;
    const contexts: ApplyContext[] = [];
    const s = model({
      getVersion: () => version,
      apply(_, context) {
        contexts.push(context);
        version += 1;
        return { status: 'applied', selection: version < 13 ? 'anchor-a' : 'anchor-b', editState: version < 12 };
      },
    });
    const first = await edit(s.host, envelope(1));
    const second = await edit(s.host, envelope(2, { afterSeq: 1 }));
    const third = await edit(s.host, envelope(3));
    expect(contexts).toEqual([10, 11, 12].map((currentVersion) => ({
      basis: { epoch: 3, version: 10 }, currentVersion, replayPreviouslyApplied: false,
    })));
    expect(first.selection).toBe('anchor-a');
    expect(first.editState).toBe(true);
    expect(second.selection).toBeUndefined();
    expect(second.editState).toBe(false);
    expect(third.selection).toBe('anchor-b');
    expect(third.editState).toBeUndefined();
    const duplicate = await edit(s.host, envelope(1));
    expect(duplicate.selection).toBeUndefined();
    expect(duplicate.editState).toBeUndefined();
  });

  it('pushes object state only when its value changes, including reused mutable state', async () => {
    let version = 10;
    const selection = { anchor: 1 };
    const host = createCommandHost<Payload, { anchor: number }, { canUndo: boolean }, null>({
      session: 'document', incarnation: 2, epoch: 3, getVersion: () => version,
      apply(command) {
        version += 1;
        if (command.seq === 3) selection.anchor = 2;
        return { status: 'applied', selection, editState: { canUndo: true } };
      },
      frame: (input) => ({ coversThroughSeq: input.coversThroughSeq, revision: version, payload: null }),
      snapshot: () => new Uint8Array(),
    });
    host.policy(envelope(1));
    const first = value(await host.handler(null, envelope(1)));
    host.policy(envelope(2));
    const second = value(await host.handler(null, envelope(2)));
    host.policy(envelope(3));
    const third = value(await host.handler(null, envelope(3)));
    expect(first.selection).toEqual({ anchor: 1 });
    expect(first.editState).toEqual({ canUndo: true });
    expect(second.selection).toBeUndefined();
    expect(second.editState).toBeUndefined();
    expect(third.selection).toEqual({ anchor: 2 });
    expect(third.editState).toBeUndefined();
  });

  it('asserts refusal before mutation and transitively cancels dependencies without format calls', async () => {
    const s = model({ retainedOutcomes: 0 });
    await edit(s.host, envelope(1, { kind: 'refuse' }));
    await edit(s.host, envelope(2, { afterSeq: 1 }));
    await edit(s.host, envelope(3));
    expect((await edit(s.host, envelope(4, { afterSeq: 2 }))).outcomes).toEqual([
      { seq: 4, status: 'cancelled', reason: 'dependency-refused', committed: false },
    ]);
    expect(s.calls).toEqual([1, 3]);
    let version = 10;
    const broken = model({
      getVersion: () => version,
      apply() { version += 1; return { status: 'refused', reason: 'locked' }; },
    });
    const error = await edit(broken.host, envelope(1)).catch((error: unknown) => error);
    expect(isCommandOutcomeUnknownError(error)).toBe(true);
    expect(await edit(broken.host, envelope(2)).catch((error: unknown) => error)).toBe(error);
    expect(broken.frames).toEqual([]);
  });

  it('poisons throwing apply handlers and preserves unknown-outcome diagnostics', async () => {
    let calls = 0;
    const s = model({ apply() { calls += 1; throw new Error('mutation failed'); } });
    const error = await edit(s.host, envelope(7)).catch((error: unknown) => error);
    expect(isCommandOutcomeUnknownError(error)).toBe(true);
    if (!isCommandOutcomeUnknownError(error)) throw new Error('Missing unknown-outcome error');
    expect(error.name).toBe(COMMAND_OUTCOME_UNKNOWN);
    expect(error.refusal).toEqual({ seq: 7 });
    expect(error.message).toContain('mutation failed');
    expect(await edit(s.host, envelope(8)).catch((error: unknown) => error)).toBe(error);
    expect(await s.host.checkpoint(null).catch((error: unknown) => error)).toBe(error);
    expect(calls).toBe(1);
    expect(s.frames).toEqual([]);
  });

  for (const mode of ['throws', 'rejects'] as const) {
    it(`poisons an apply that ${mode} undefined with an unknown outcome`, async () => {
      const s = model({ apply: mode === 'throws' ? () => { throw undefined; } : () => Promise.reject(undefined) });
      const error = await edit(s.host, envelope(1)).catch((error: unknown) => error);
      expect(isCommandOutcomeUnknownError(error)).toBe(true);
      if (!isCommandOutcomeUnknownError(error)) throw new Error('Missing unknown-outcome error');
      expect(error.refusal).toEqual({ seq: 1 });
      expect(error.message).toBe('Command 1 has an unknown outcome: undefined');
      expect(await edit(s.host, envelope(2)).catch((error: unknown) => error)).toBe(error);
      expect(await s.host.checkpoint(null).catch((error: unknown) => error)).toBe(error);
    });
  }

  it('requires applied to advance version and noop to preserve it', async () => {
    for (const status of ['applied', 'noop'] as const) {
      let version = 10;
      const s = model({
        getVersion: () => version,
        apply() { if (status === 'noop') version += 1; return { status }; },
      });
      const error = await edit(s.host, envelope(1)).catch((error: unknown) => error);
      expect(isCommandOutcomeUnknownError(error)).toBe(true);
      expect(await edit(s.host, envelope(2)).catch((error: unknown) => error)).toBe(error);
    }
  });

  it('restores refused tombstones and bypasses strict version only for recorded applied replay', async () => {
    const s = model({ throughSeq: 5, retainedOutcomes: 0, refusedSeqs: [2, 6], replayAppliedSeqs: [7, 9] });
    expect((await edit(s.host, envelope(2, { replay: true }))).outcomes[0]?.status).toBe('duplicate');
    expect((await edit(s.host, envelope(6, { replay: true }))).outcomes[0]?.reason).toBe('unknown-outcome');
    expect((await edit(s.host, envelope(7, {
      replay: true, expectVersion: true, basis: { epoch: 3, version: 0 },
    }))).outcomes[0]?.status).toBe('applied');
    expect((await edit(s.host, envelope(8, { afterSeq: 2 }))).outcomes[0]?.status).toBe('cancelled');
    expect((await edit(s.host, envelope(9, {
      expectVersion: true, basis: { epoch: 3, version: 0 },
    }))).outcomes[0]?.reason).toBe('stale-version');
    expect((await edit(s.host, envelope(10, {
      replay: true, expectVersion: true, basis: { epoch: 3, version: 0 },
    }))).outcomes[0]?.reason).toBe('stale-version');
    expect(s.calls).toEqual([7]);
    expect(s.contexts[0]?.replayPreviouslyApplied).toBe(true);
  });

  it('validates envelope metadata before format hooks', async () => {
    const s = model();
    for (const command of [
      null, {}, envelope(0), envelope(Number.NaN), envelope(1, { incarnation: -1 }),
      envelope(1, { basis: { epoch: 3, version: Number.NaN } }),
      envelope(1, { afterSeq: 0 }), { ...envelope(1), replay: false },
    ]) {
      await expect(s.host.handler(null, command as CommandEnvelope<Payload>)).rejects.toBeInstanceOf(Error);
    }
    expect(s.calls).toEqual([]);
    expect(s.frames).toEqual([]);
    expect((await edit(s.host, envelope(1))).throughSeq).toBe(1);
  });

  it('rejects stale-epoch duplicates without calling format hooks or changing retained outcomes', async () => {
    const s = model();
    const first = await edit(s.host, envelope(1));
    const frames = s.frames.length;
    const stale = await edit(s.host, envelope(1, { basis: { epoch: 2, version: 10 } }));
    expect(stale.outcomes[0]?.reason).toBe('stale-epoch');
    expect(stale.frame).toBeUndefined();
    expect(s.frames.length).toBe(frames);
    expect((await edit(s.host, envelope(1))).outcomes).toEqual(first.outcomes);
    expect(s.calls).toEqual([1]);
  });

  it('bounds sparse replay tombstones without forgetting dependency failures', () => {
    expect(() => model({ refusedSeqs: Array.from({ length: 4097 }, (_, index) => index * 2 + 1) })).toThrow(SessionFailure);
    expect(() => model({ refusedSeqs: Array.from({ length: 4097 }, (_, index) => index * 2 + 1) }))
      .toThrow('Command replay ledger exhausted: 4096 disjoint ranges per epoch');
    expect(() => model({ refusedSeqs: Array.from({ length: 8192 }, (_, index) => index + 1) })).not.toThrow();
    const sparse = Array.from({ length: 4096 }, (_, index) => index * 2 + 1);
    expect(() => model({ refusedSeqs: [...sparse, 2, 8193] })).not.toThrow();
  });

  it('keeps host ledger exhaustion terminal for later commands', async () => {
    const s = model({
      throughSeq: 8191, refusedSeqs: Array.from({ length: 4096 }, (_, index) => index * 2 + 1),
    });
    const failure: unknown = await edit(s.host, envelope(8193, { kind: 'refuse' })).catch((error: unknown) => error);
    expect(failure).toBeInstanceOf(SessionFailure);
    expect(failure).toMatchObject({
      code: 'out-of-memory', message: 'Command replay ledger exhausted: 4096 disjoint ranges per epoch',
    });
    const later: unknown = await edit(s.host, envelope(8194)).catch((error: unknown) => error);
    expect(later).toBe(failure);
    expect(s.calls).toEqual([8193]);
    expect(s.version).toBe(10);
  });

  it('bounds direct handler queues and rejects before mutation when capacity is exhausted', async () => {
    const s = model();
    const replies = Array.from({ length: 4097 }, (_, index) =>
      s.host.handler(null, envelope(index + 1)).catch((error: unknown) => error));
    const errors = await Promise.all(replies);
    expect(errors[0]).toBeInstanceOf(SessionFailure);
    expect(errors.every((error) => error === errors[0])).toBe(true);
    expect(s.calls).toEqual([]);
    expect(s.frames).toEqual([]);
  });

  it('reports an in-flight apply failure as unknown after capacity poisons the host', async () => {
    const started = deferred<void>();
    const held = deferred<void>();
    const calls: number[] = [];
    const s = model({
      async apply(command) {
        calls.push(command.seq);
        started.resolve();
        await held.promise;
        throw new Error('mutation failed');
      },
    });
    const first = s.host.handler(null, envelope(1)).catch((error: unknown) => error);
    await started.promise;
    const queued = Array.from({ length: 4096 }, (_, index) =>
      s.host.handler(null, envelope(index + 2)).catch((error: unknown) => error));
    const capacity = await queued[4095];
    expect(capacity).toBeInstanceOf(SessionFailure);
    expect(capacity).toMatchObject({ code: 'out-of-memory' });
    held.resolve();
    const error = await first;
    expect(isCommandOutcomeUnknownError(error)).toBe(true);
    if (!isCommandOutcomeUnknownError(error)) throw new Error('Missing unknown-outcome error');
    expect(error.refusal).toEqual({ seq: 1 });
    expect(error.message).toContain('mutation failed');
    await Promise.all(queued);
    expect(calls).toEqual([1]);
  });

  it('counts queued admissions and returns a frame only with the last settled edit', async () => {
    const started = deferred<void>();
    const held = deferred<void>();
    let version = 10;
    const order: number[] = [];
    const s = model({
      getVersion: () => version,
      async apply(command) {
        order.push(command.seq);
        if (command.seq === 1) { started.resolve(); await held.promise; }
        if (command.kind === 'noop') return { status: 'noop' };
        version += 1;
        return { status: 'applied' };
      },
    });
    const first = edit(s.host, envelope(1));
    await started.promise;
    const second = edit(s.host, envelope(2));
    const third = edit(s.host, envelope(3, { kind: 'noop' }));
    expect(order).toEqual([1]);
    held.resolve();
    const acks = await Promise.all([first, second, third]);
    expect(order).toEqual([1, 2, 3]);
    expect(acks.map((ack) => ack.frame?.coversThroughSeq)).toEqual([undefined, undefined, 3]);
    expect(s.frames).toEqual([3]);
    expect(s.host.policy(envelope(4))).toEqual({ lane: 'input', mutates: true, reframes: true, reorderable: false });
    await s.host.handler(null, envelope(4));
  });

  it('does not let foreign admissions suppress the final live command frame', async () => {
    const s = model();
    const live = edit(s.host, envelope(1));
    const foreign = edit(s.host, envelope(2, { session: 'other' }));
    expect((await live).frame?.coversThroughSeq).toBe(1);
    expect((await foreign).frame).toBeUndefined();
    expect(s.calls).toEqual([1]);
    expect(s.frames).toEqual([1]);
  });

  it('transfers frame buffers in the covering acknowledgement', async () => {
    const bytes = new Uint8Array([4, 5, 6]);
    const s = model({
      frame: (input) => transferable<FrameEnvelope<Uint8Array>>({
        coversThroughSeq: input.coversThroughSeq, revision: 42, baseRevision: 41, payload: bytes,
      }, [bytes.buffer]),
    });
    s.host.policy(envelope(1));
    const reply = await s.host.handler(null, envelope(1));
    expect(isTransferResult(reply)).toBe(true);
    if (!isTransferResult(reply)) throw new Error('Missing transfer result');
    expect(reply.transfer).toEqual([bytes.buffer]);
    const ack = cloneTransfer(reply) as Ack;
    expect(ack.frame).toMatchObject({ coversThroughSeq: 1, revision: 42, baseRevision: 41 });
    expect([...ack.frame!.payload]).toEqual([4, 5, 6]);
    expect(bytes.byteLength).toBe(0);
  });

  it('orders the checkpoint barrier after earlier edits and before later edits through the resident scheduler', async () => {
    const entered = deferred<void>();
    const held = deferred<void>();
    let version = 10;
    const order: string[] = [];
    const s = model({
      getVersion: () => version,
      async apply(command) {
        order.push(`edit-${command.seq}`);
        if (command.seq === 1) { entered.resolve(); await held.promise; }
        version += 1;
        return { status: 'applied' };
      },
      snapshot() { order.push('snapshot'); return new Uint8Array([version]); },
    });
    const scheduler = createResidentScheduler({
      now: () => 0, turn: (callback) => callback(), timer: () => () => {},
      failed(error) { throw error; },
    });
    const first = deferred<Ack>();
    const checkpoint = deferred<Awaited<ReturnType<Host['checkpoint']>>>();
    const last = deferred<Ack>();
    scheduler.submit({ ...s.host.policy(envelope(1)), run: async () => { first.resolve(value(await s.host.handler(null, envelope(1)))); } });
    await entered.promise;
    scheduler.submit({ ...s.host.checkpointPolicy(), run: async () => { checkpoint.resolve(await s.host.checkpoint(null)); } });
    scheduler.submit({ ...s.host.policy(envelope(2)), run: async () => { last.resolve(value(await s.host.handler(null, envelope(2)))); } });
    held.resolve();
    await first.promise;
    const reply = await checkpoint.promise;
    expect(isTransferResult(reply)).toBe(true);
    if (!isTransferResult(reply)) throw new Error('Missing checkpoint transfer');
    expect(reply.value).toMatchObject({ epoch: 3, version: 11, throughSeq: 1 });
    const copied = cloneTransfer(reply);
    expect([...copied.state]).toEqual([11]);
    await last.promise;
    expect(order).toEqual(['edit-1', 'snapshot', 'edit-2']);
    expect(s.host.checkpointPolicy()).toMatchObject({ reorderable: false });
  });

  it('holds edits until an asynchronous checkpoint snapshot settles and transfers owned bytes', async () => {
    const entered = deferred<void>();
    const held = deferred<Uint8Array>();
    const backing = new Uint8Array([1, 10, 3]);
    const s = model({ snapshot() { entered.resolve(); return held.promise; } });
    await edit(s.host, envelope(1));
    const checkpoint = s.host.checkpoint(null);
    await entered.promise;
    const later = edit(s.host, envelope(2));
    expect(s.calls).toEqual([1]);
    held.resolve(backing.subarray(1, 2));
    const reply = await checkpoint;
    expect(isTransferResult(reply)).toBe(true);
    if (!isTransferResult(reply)) throw new Error('Missing checkpoint transfer');
    const copied = cloneTransfer(reply);
    expect([...copied.state]).toEqual([10]);
    expect(copied).toMatchObject({ epoch: 3, version: 11, throughSeq: 1 });
    expect([...backing]).toEqual([1, 10, 3]);
    await later;
    expect(s.calls).toEqual([1, 2]);
  });

  it('keeps frame failures terminal without marking an applied edit unknown', async () => {
    const s = model({ frame() { throw new Error('paint failed'); } });
    const error = await edit(s.host, envelope(1)).catch((error: unknown) => error);
    expect(error).toBeInstanceOf(SessionFailure);
    expect(isCommandOutcomeUnknownError(error)).toBe(false);
    expect(s.version).toBe(11);
    expect(await edit(s.host, envelope(2)).catch((error: unknown) => error)).toBe(error);
    expect(s.calls).toEqual([1]);
  });

  it('rejects a final frame that does not cover settled edits', async () => {
    const s = model({ frame: () => ({ coversThroughSeq: 0, revision: 1, payload: new Uint8Array() }) });
    const error = await edit(s.host, envelope(1)).catch((error: unknown) => error);
    expect(error).toBeInstanceOf(SessionFailure);
    expect(isCommandOutcomeUnknownError(error)).toBe(false);
    expect(await edit(s.host, envelope(2)).catch((error: unknown) => error)).toBe(error);
  });

  it('poisons a snapshot that changes the version', async () => {
    let version = 10;
    const s = model({
      getVersion: () => version,
      snapshot() { version += 1; return new Uint8Array([version]); },
    });
    const error = await s.host.checkpoint(null).catch((error: unknown) => error);
    expect(error).toBeInstanceOf(SessionFailure);
    expect(await edit(s.host, envelope(1)).catch((error: unknown) => error)).toBe(error);
    expect(s.calls).toEqual([]);
  });

  it('refuses a failed snapshot without poisoning later edits or checkpoints', async () => {
    let failing = true;
    const s = model({ snapshot() {
      if (failing) throw new Error('malformed request: Recovery projection cannot be reconstructed');
      return new Uint8Array([1]);
    } });
    const error = await s.host.checkpoint(null).catch((error: unknown) => error);
    expect(error).toBeInstanceOf(Error);
    expect(error).not.toBeInstanceOf(SessionFailure);
    expect(await edit(s.host, envelope(1))).toMatchObject({ throughSeq: 1, outcomes: [{ seq: 1, status: 'applied' }] });
    failing = false;
    expect(value(await s.host.checkpoint(null))).toMatchObject({ throughSeq: 1, version: 11 });
  });

  it('poisons a snapshot that traps', async () => {
    const s = model({ snapshot() { throw new WebAssembly.RuntimeError('unreachable'); } });
    const error = await s.host.checkpoint(null).catch((error: unknown) => error);
    expect(error).toBeInstanceOf(SessionFailure);
    expect(error).toMatchObject({ code: 'trap' });
    expect(await edit(s.host, envelope(1)).catch((error: unknown) => error)).toBe(error);
  });
});
