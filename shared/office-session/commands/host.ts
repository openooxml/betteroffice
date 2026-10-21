import { deferReply, hasOwn, isTransferResult, transferable, type DeferredReply, type TransferResult } from '../protocol';
import { SessionFailure, type MethodPolicy } from '../types';
import { cloneData } from './clone';
import { BASE_REFUSALS, isRefusalResult, type QueryFence, type QueryReply } from './spec';
import {
  COMMAND_OUTCOME_UNKNOWN, isCommandAck, isCommandEnvelope,
  type ApplyResult, type CommandAck, type CommandEnvelope, type CommandHost,
  type CommandHostOptions, type CommandOutcome, type CommandOutcomeUnknownError, type FrameEnvelope,
} from './types';

const MAX_LEDGER_RANGES = 4096;
const MAX_PENDING_HANDLERS = 4096;

function sameState(left: unknown, right: unknown): boolean {
  const seen = new WeakMap<object, object>();
  let remaining = 4096;

  function equal(a: unknown, b: unknown, depth: number): boolean {
    if (--remaining < 0 || depth > 32) return false;
    if (Object.is(a, b)) return true;
    if (a === null || b === null ||
      typeof a !== 'object' || typeof b !== 'object') return false;
    if (seen.has(a)) return seen.get(a) === b;
    seen.set(a, b);
    if (a instanceof Date || b instanceof Date) {
      return a instanceof Date && b instanceof Date && Object.is(a.getTime(), b.getTime());
    }
    if (a instanceof RegExp || b instanceof RegExp) {
      return a instanceof RegExp && b instanceof RegExp && a.source === b.source && a.flags === b.flags;
    }
    if (a instanceof Number || b instanceof Number || a instanceof Boolean || b instanceof Boolean ||
      a instanceof String || b instanceof String) {
      return a.constructor === b.constructor && Object.is(a.valueOf(), b.valueOf());
    }
    if (a instanceof ArrayBuffer || b instanceof ArrayBuffer || ArrayBuffer.isView(a) || ArrayBuffer.isView(b)) {
      const bytes = (value: object): Uint8Array | undefined => value instanceof ArrayBuffer
        ? new Uint8Array(value)
        : ArrayBuffer.isView(value) ? new Uint8Array(value.buffer, value.byteOffset, value.byteLength) : undefined;
      const x = bytes(a);
      const y = bytes(b);
      if (!x || !y || a.constructor !== b.constructor || x.byteLength !== y.byteLength) return false;
      if (x.byteLength > remaining) return false;
      remaining -= x.byteLength;
      for (let index = 0; index < x.length; index += 1) if (x[index] !== y[index]) return false;
      return true;
    }
    if (a instanceof Map || b instanceof Map || a instanceof Set || b instanceof Set) {
      if (a.constructor !== b.constructor || !(a instanceof Map || a instanceof Set) ||
        !(b instanceof Map || b instanceof Set) || a.size !== b.size || a.size > remaining) return false;
      const x = a.entries();
      const y = b.entries();
      for (let index = 0; index < a.size; index += 1) {
        if (!equal(x.next().value, y.next().value, depth + 1)) return false;
      }
      return true;
    }
    if (Array.isArray(a) !== Array.isArray(b)) return false;
    if (Array.isArray(a) && Array.isArray(b) && a.length !== b.length) return false;
    if (Object.getPrototypeOf(a) !== Object.getPrototypeOf(b)) return false;
    if (!Array.isArray(a) && Object.getPrototypeOf(a) !== Object.prototype && Object.getPrototypeOf(a) !== null) return false;
    const keys = Object.keys(a);
    if (keys.length > remaining || keys.length !== Object.keys(b).length) return false;
    const x = a as Record<string, unknown>;
    const y = b as Record<string, unknown>;
    return keys.every((key) => hasOwn(b, key) && equal(x[key], y[key], depth + 1));
  }

  return equal(left, right, 0);
}

function sequenceSet(initial: readonly number[] = []) {
  const ranges: [number, number][] = [];

  function contains(seq: number): boolean {
    let low = 0;
    let high = ranges.length;
    while (low < high) {
      const middle = Math.floor((low + high) / 2);
      const range = ranges[middle]!;
      if (seq < range[0]) high = middle;
      else if (seq > range[1]) low = middle + 1;
      else return true;
    }
    return false;
  }

  function add(seq: number): void {
    if (!Number.isSafeInteger(seq) || seq <= 0) throw new TypeError('Invalid replay sequence');
    let index = 0;
    while (index < ranges.length && ranges[index]![1] < seq - 1) index += 1;
    const next = ranges[index];
    if (next && next[0] <= seq + 1) {
      next[0] = Math.min(next[0], seq);
      next[1] = Math.max(next[1], seq);
      const following = ranges[index + 1];
      if (following && following[0] <= next[1] + 1) {
        next[1] = Math.max(next[1], following[1]);
        ranges.splice(index + 1, 1);
      }
      return;
    }
    if (ranges.length === MAX_LEDGER_RANGES) {
      throw new SessionFailure('out-of-memory', 'Command replay ledger exhausted: 4096 disjoint ranges per epoch');
    }
    ranges.splice(index, 0, [seq, seq]);
  }

  for (const seq of initial) add(seq);
  return { contains, add };
}

function copyOutcome(outcome: CommandOutcome): CommandOutcome {
  return {
    ...outcome,
    ...(outcome.affectedSeqs === undefined ? {} : { affectedSeqs: [...outcome.affectedSeqs] }),
  };
}

function unknownOutcome(seq: number, cause: unknown): CommandOutcomeUnknownError {
  const message = cause instanceof Error ? cause.message : String(cause);
  const data: Pick<CommandOutcomeUnknownError, 'name' | 'refusal'> = {
    name: COMMAND_OUTCOME_UNKNOWN, refusal: { seq },
  };
  return Object.assign(new Error(`Command ${seq} has an unknown outcome: ${message}`), data);
}

function terminalFailure(cause: unknown): SessionFailure {
  if (cause instanceof SessionFailure) return cause;
  const trap = typeof WebAssembly !== 'undefined' && cause instanceof WebAssembly.RuntimeError;
  return new SessionFailure(trap ? 'trap' : 'crash', cause instanceof Error ? cause.message : String(cause));
}

/** @internal Serial worker command handlers with replay dedupe and covering frame replies. */
export function createCommandHost<P, S, E, F>(
  options: CommandHostOptions<P, S, E, F>
): Required<CommandHost<P, S, E, F>> {
  const retainedLimit = options.retainedOutcomes ?? 256;
  const { session, incarnation, epoch, getVersion, apply, frame, snapshot, prunedDuplicateReason } = options;
  let throughSeq = options.throughSeq ?? 0;
  let version = getVersion();
  if (!Number.isSafeInteger(retainedLimit) || retainedLimit < 0 ||
    !Number.isSafeInteger(throughSeq) || throughSeq < 0 || !Number.isFinite(version) ||
    (prunedDuplicateReason !== undefined &&
      (typeof prunedDuplicateReason !== 'string' || prunedDuplicateReason.length === 0))) {
    throw new TypeError('Invalid command host ledger');
  }
  const outcomes = new Map<number, CommandOutcome>();
  const refused = sequenceSet(options.refusedSeqs);
  const replayApplied = new Set(options.replayAppliedSeqs);
  for (const seq of replayApplied) {
    if (!Number.isSafeInteger(seq) || seq <= 0) throw new TypeError('Invalid replay sequence');
  }
  let poison: Error | undefined;
  let tail: Promise<void> = Promise.resolve();
  let admitted = 0;
  let queuedEdits = 0;
  let pendingHandlers = 0;
  let enteredAdmissions = 0;
  let pendingFrames = 0;
  const frameBarriers = new Map<number, Array<() => void>>();
  let selection: S | undefined;
  let editState: E | undefined;
  let hasSelection = false;
  let hasEditState = false;
  let highestSeen = throughSeq;
  let selectionRevision = 0;
  let ackLayout = options.queries ? layoutRevision() : 0;
  const queryBarriers = new Map<number, Array<{ start(): void; reject(error: Error): void }>>();

  function layoutRevision(): number {
    const revision = options.layoutRevision?.() ?? 0;
    if (!Number.isSafeInteger(revision) || revision < 0) throw new TypeError('Invalid layout revision');
    return revision;
  }

  function checkRefusal(code: string, declared: readonly string[] = []): void {
    if (!options.checkRefusals || BASE_REFUSALS.includes(code as typeof BASE_REFUSALS[number]) ||
      declared.includes(code) || ['stale-epoch', 'stale-incarnation', 'wrong-session', 'stale-version',
        'dependency-refused', 'unknown-outcome', 'disposed'].includes(code)) return;
    throw poisonWith(new TypeError(`Undeclared refusal: ${code}`));
  }

  function poisonWith(error: Error): Error {
    poison = error;
    releasePoisonedFrames();
    return error;
  }

  function serialize<T>(run: () => Promise<T>): Promise<T> {
    if (pendingHandlers === MAX_PENDING_HANDLERS && !poison) {
      poisonWith(new SessionFailure('out-of-memory', 'Command handler capacity exceeded'));
    }
    if (poison) return Promise.reject(poison);
    pendingHandlers += 1;
    const result = tail.then(() => {
      if (poison) throw poison;
      return run();
    });
    const settled = result.finally(() => {
      pendingHandlers -= 1;
      releasePoisonedFrames();
    });
    tail = settled.then(() => {}, () => {});
    return settled;
  }

  function releaseFrames(boundary: number): void {
    const waiting = frameBarriers.get(boundary);
    if (!waiting) return;
    frameBarriers.delete(boundary);
    for (const start of waiting) start();
  }

  function releasePoisonedFrames(): void {
    if (!poison) return;
    const error = poison;
    const frames = [...frameBarriers.values()];
    const queries = [...queryBarriers.values()];
    frameBarriers.clear();
    queryBarriers.clear();
    for (const waiting of frames) for (const start of waiting) start();
    for (const waiting of queries) for (const waiter of waiting) waiter.reject(error);
  }

  async function buildFrame(): Promise<FrameEnvelope<F> | TransferResult<FrameEnvelope<F>> | undefined> {
    try {
      const rendered = await frame({ coversThroughSeq: throughSeq, version });
      if (rendered === undefined) return undefined;
      const candidate = isTransferResult(rendered) ? (rendered as TransferResult<FrameEnvelope<F>>).value : rendered;
      if (!isCommandAck({ session, incarnation, epoch, throughSeq, version, outcomes: [], frame: candidate }) ||
        candidate.coversThroughSeq !== throughSeq) {
        throw new Error('Frame does not cover the settled command boundary');
      }
      return rendered;
    } catch (error) {
      throw poisonWith(terminalFailure(error));
    }
  }

  async function requiredFrame(): Promise<FrameEnvelope<F> | TransferResult<FrameEnvelope<F>>> {
    const rendered = await buildFrame();
    if (rendered === undefined) throw new Error('No frame is available at the settled command boundary');
    return rendered;
  }

  function settle(outcome: CommandOutcome): CommandOutcome {
    if (outcome.status === 'refused' || outcome.status === 'cancelled') {
      try { refused.add(outcome.seq); }
      catch (error) { throw poisonWith(terminalFailure(error)); }
    }
    throughSeq = outcome.seq;
    if (retainedLimit > 0) {
      outcomes.set(outcome.seq, copyOutcome(outcome));
      if (outcomes.size > retainedLimit) outcomes.delete(outcomes.keys().next().value!);
    }
    if (queryBarriers.size) {
      for (const boundary of [...queryBarriers.keys()].sort((a, b) => a - b)) {
        if (boundary > throughSeq) break;
        const waiting = queryBarriers.get(boundary);
        if (!waiting) continue;
        queryBarriers.delete(boundary);
        for (const waiter of waiting) waiter.start();
      }
    }
    return outcome;
  }

  function refusal(seq: number, reason: string, cancelled = false): CommandOutcome {
    return { seq, status: cancelled ? 'cancelled' : 'refused', reason, committed: false };
  }

  function ack(outcome?: CommandOutcome, result?: { selection?: S; editState?: E }): CommandAck<S, E, F> {
    const reply: CommandAck<S, E, F> = {
      session, incarnation, epoch,
      throughSeq, version, outcomes: outcome ? [copyOutcome(outcome)] : [],
    };
    try {
      if (result && hasOwn(result, 'selection') &&
        (!hasSelection || !(options.selectionEquals ?? sameState)(selection, result.selection))) {
        selection = cloneData(result.selection);
        hasSelection = true;
        reply.selection = cloneData(selection);
        if (options.queries) selectionRevision += 1;
      }
      if (result && hasOwn(result, 'editState') &&
        (!hasEditState || !(options.editStateEquals ?? sameState)(editState, result.editState))) {
        editState = cloneData(result.editState);
        hasEditState = true;
        reply.editState = cloneData(editState);
      }
      if (options.queries) {
        ackLayout = layoutRevision();
        reply.revisions = { selection: selectionRevision, layout: ackLayout };
      }
    } catch (error) {
      throw poisonWith(terminalFailure(error));
    }
    return reply;
  }

  async function withFrame(reply: CommandAck<S, E, F>): Promise<CommandAck<S, E, F> | TransferResult<CommandAck<S, E, F>>> {
    if (admitted > 0 || queuedEdits > 1) return reply;
    try {
      const rendered = await buildFrame();
      if (rendered === undefined) return reply;
      if (isTransferResult(rendered)) {
        reply.frame = rendered.value;
        if (!isCommandAck(reply) || reply.frame.coversThroughSeq !== throughSeq) {
          throw new Error('Frame does not cover the settled command boundary');
        }
        return transferable(reply, rendered.transfer);
      }
      reply.frame = rendered;
      if (!isCommandAck(reply) || rendered.coversThroughSeq !== throughSeq) {
        throw new Error('Frame does not cover the settled command boundary');
      }
      return reply;
    } catch (error) {
      throw poisonWith(terminalFailure(error));
    }
  }

  function checkedVersion(seq: number): number {
    try {
      const current = getVersion();
      if (!Number.isFinite(current) || current !== version) throw new Error('Version changed outside a command');
      return current;
    } catch (error) {
      throw poisonWith(unknownOutcome(seq, error));
    }
  }

  function liveCommand(command: CommandEnvelope<P>): boolean {
    return isCommandEnvelope(command) && command.session === session &&
      command.incarnation === incarnation && command.basis.epoch === epoch;
  }

  async function run(command: CommandEnvelope<P>): Promise<CommandAck<S, E, F> | TransferResult<CommandAck<S, E, F>>> {
    if (!isCommandEnvelope(command)) throw new TypeError('Invalid command envelope');
    if (command.session !== session) return ack(refusal(command.seq, 'wrong-session'));
    if (command.incarnation !== incarnation) return ack(refusal(command.seq, 'stale-incarnation'));
    if (command.basis.epoch !== epoch) {
      const outcome = refusal(command.seq, 'stale-epoch');
      return ack(command.seq > throughSeq ? settle(outcome) : outcome);
    }
    if (command.seq <= throughSeq) {
      return withFrame(ack(outcomes.get(command.seq) ?? (prunedDuplicateReason === undefined
        ? { seq: command.seq, status: 'duplicate' }
        : refusal(command.seq, prunedDuplicateReason))));
    }
    if (refused.contains(command.seq)) return withFrame(ack(settle(refusal(command.seq, 'unknown-outcome'))));
    if (command.afterSeq !== undefined && refused.contains(command.afterSeq)) {
      return withFrame(ack(settle(refusal(command.seq, 'dependency-refused', true))));
    }
    const replayPreviouslyApplied = command.replay === true && replayApplied.has(command.seq);
    if (options.commands) {
      if (!hasOwn(options.commands, command.kind)) {
        return withFrame(ack(settle(refusal(command.seq, 'unsupported'))));
      }
      if (options.commands[command.kind]!.admission === 'local-write' &&
        !replayPreviouslyApplied && options.readOnly?.()) {
        return withFrame(ack(settle(refusal(command.seq, 'read-only'))));
      }
    }
    const before = checkedVersion(command.seq);
    if (command.expectVersion && !replayPreviouslyApplied && command.basis.version !== before) {
      return withFrame(ack(settle(refusal(command.seq, 'stale-version'))));
    }
    let result: ApplyResult<S, E>;
    try {
      result = await apply(command, {
        basis: command.basis, currentVersion: before, replayPreviouslyApplied,
      });
      if (!result || typeof result !== 'object' ||
        (result.status !== 'applied' && result.status !== 'noop' &&
          result.status !== 'cancelled' && result.status !== 'refused') ||
        (result.status === 'applied' && result.committed !== undefined && result.committed !== false) ||
        (result.status === 'cancelled' && result.reason !== undefined && typeof result.reason !== 'string')) {
        throw new Error('Invalid format apply result');
      }
      const after = getVersion();
      if (!Number.isFinite(after) ||
        (result.status === 'applied' && result.committed !== false ? after <= before : after !== before)) {
        throw new Error(`${result.status} violated the command version contract`);
      }
      version = after;
    } catch (error) {
      throw poisonWith(unknownOutcome(command.seq, error));
    }
    if (result.status === 'refused') checkRefusal(result.reason, options.commands?.[command.kind]?.refusals);
    const outcome: CommandOutcome = result.status === 'refused'
      ? { ...refusal(command.seq, result.reason), ...(result.affectedSeqs === undefined ? {} : { affectedSeqs: result.affectedSeqs }) }
      : result.status === 'cancelled'
        ? { seq: command.seq, status: 'cancelled', committed: false, ...(result.reason === undefined ? {} : { reason: result.reason }) }
        : result.status === 'applied' && result.committed === false
          ? { seq: command.seq, status: 'applied', committed: false }
          : { seq: command.seq, status: result.status };
    const reply = ack(settle(outcome), result);
    return result.status === 'noop' && options.deferFrame?.(command, result) ? reply : withFrame(reply);
  }

  const editPolicy: MethodPolicy = { lane: 'input', mutates: true, reframes: true, reorderable: false };

  function queryRefusal(code: string, detail?: string): QueryReply {
    return { ok: false, refusal: { code, ...(detail === undefined ? {} : { detail }) } };
  }

  function validFence(fence: QueryFence): boolean {
    return !!fence && typeof fence === 'object' && typeof fence.session === 'string' &&
      [fence.incarnation, fence.epoch, fence.afterSeq].every((value) => Number.isSafeInteger(value) && value >= 0);
  }

  function runQuery(context: unknown, kind: string, args: unknown[]): Promise<QueryReply | TransferResult<QueryReply>> {
    return serialize(async () => {
      let result: unknown;
      try {
        result = await options.queries!.handlers[kind]!(args, context);
      } catch (error) {
        if (error instanceof SessionFailure ||
          (typeof WebAssembly !== 'undefined' && error instanceof WebAssembly.RuntimeError)) {
          throw poisonWith(terminalFailure(error));
        }
        return queryRefusal('invalid-payload', error instanceof Error ? error.message : String(error));
      }
      if (isRefusalResult(result)) {
        checkRefusal(result.refusal.code, options.queries!.specs[kind]!.refusals);
        return { ok: false, refusal: result.refusal };
      }
      const reply: QueryReply = { ok: true, value: isTransferResult(result) ? result.value : result,
        at: { incarnation, epoch, throughSeq, doc: version, selection: selectionRevision, layout: layoutRevision() } };
      return isTransferResult(result) ? transferable(reply, result.transfer) : reply;
    });
  }

  return {
    handler(_context, command) {
      const live = liveCommand(command);
      const admission = live && admitted > 0;
      if (live) {
        highestSeen = Math.max(highestSeen, command.seq);
        admitted = Math.max(0, admitted - 1);
        queuedEdits += 1;
      }
      const result = serialize(() => run(command));
      if (admission) releaseFrames(++enteredAdmissions);
      releasePoisonedFrames();
      return result.finally(() => {
        if (live) queuedEdits -= 1;
        releasePoisonedFrames();
      });
    },
    policy(command) {
      if (liveCommand(command)) {
        admitted += 1;
        highestSeen = Math.max(highestSeen, command.seq);
      }
      return { ...editPolicy };
    },
    query(context, fence, kind, args): Promise<QueryReply | TransferResult<QueryReply>> | DeferredReply<QueryReply> {
      if (poison) return Promise.reject(poison);
      if (!validFence(fence) || typeof kind !== 'string' || !Array.isArray(args)) {
        return Promise.resolve(queryRefusal('invalid-payload'));
      }
      if (fence.session !== session || fence.incarnation !== incarnation || fence.epoch !== epoch ||
        fence.afterSeq > highestSeen) return Promise.resolve(queryRefusal('stale'));
      if (!options.queries || !hasOwn(options.queries.specs, kind) || !hasOwn(options.queries.handlers, kind)) {
        return Promise.resolve(queryRefusal('unsupported'));
      }
      if (throughSeq >= fence.afterSeq) return runQuery(context, kind, args);
      return deferReply(new Promise<QueryReply | TransferResult<QueryReply>>((resolve, reject) => {
        let waiting = queryBarriers.get(fence.afterSeq);
        if (!waiting) queryBarriers.set(fence.afterSeq, waiting = []);
        waiting.push({ start: () => { void runQuery(context, kind, args).then(resolve, reject); }, reject });
      }));
    },
    queryPolicy(_fence, kind, _args) {
      const spec = options.queries && hasOwn(options.queries.specs, kind) ? options.queries.specs[kind] : undefined;
      return { lane: 'interactive', reorderable: spec?.stale !== 'answer-at-fence' };
    },
    revise(change) {
      if (poison) throw poison;
      const previousVersion = version;
      const previousLayout = ackLayout;
      version = getVersion();
      const reply = ack(undefined, change);
      return version !== previousVersion || ackLayout !== previousLayout ||
        hasOwn(reply, 'selection') || hasOwn(reply, 'editState') ? reply : undefined;
    },
    checkpoint(_context) {
      return serialize(async () => {
        try {
          const before = getVersion();
          if (before !== version) throw new Error('Version changed before checkpoint');
          const state = (await snapshot()).slice();
          if (getVersion() !== before) throw new Error('Snapshot changed the command version');
          return transferable({ state, epoch, version, throughSeq }, [state.buffer]);
        } catch (error) {
          throw poisonWith(terminalFailure(error));
        }
      });
    },
    checkpointPolicy() {
      return { lane: 'input', reorderable: false };
    },
    frame() {
      if (pendingFrames === MAX_PENDING_HANDLERS && !poison) {
        poisonWith(new SessionFailure('out-of-memory', 'Command frame capacity exceeded'));
      }
      if (poison) {
        releasePoisonedFrames();
        return Promise.reject(poison);
      }
      if (admitted === 0) return serialize(requiredFrame);
      pendingFrames += 1;
      return new Promise<FrameEnvelope<F> | TransferResult<FrameEnvelope<F>>>((resolve, reject) => {
        const boundary = enteredAdmissions + admitted;
        let waiting = frameBarriers.get(boundary);
        if (!waiting) frameBarriers.set(boundary, waiting = []);
        waiting.push(() => {
          pendingFrames -= 1;
          void serialize(requiredFrame).then(resolve, reject);
        });
      });
    },
  };
}

/** @internal RPC registrations for a command host. */
export function commandSessionMethods<P, S, E, F>(host: Required<CommandHost<P, S, E, F>>) {
  return {
    handlers: { edit: host.handler, checkpoint: host.checkpoint, query: host.query },
    policies: { edit: host.policy, checkpoint: host.checkpointPolicy, query: host.queryPolicy },
  };
}
