import { SessionFailure } from '../types';
import { cancelOutboxAfter, createOutbox, fenceOutbox } from './outbox';
import { createPresenter } from './presenter';
import { cloneData } from './clone';
import { createBlockedLedger } from './ledger';
import { isCommandAck, isCommandOutcomeUnknownError } from './types';
import { hasOwn } from '../protocol';
import { isQueryReply, type ArgsOf, type Hold, type HoldReason, type QueryReply, type QueryTable,
  type Read, type ReadOptions, type RefusalOf, type ResultOf, type WorkerEditSession } from './spec';
import type {
  Checkpoint, CommandAck, CommandConnection, CommandDraft, CommandEnvelope,
  CommandOutcome, CommandTicket, EditSessionState, OutboxEntry,
  FrameEnvelope, PresentationStatus, RecoveryEvent, SubmitOptions, TicketOutcome, WorkerEditSessionOptions,
} from './types';

type Deferred<T> = {
  promise: Promise<T>;
  resolve(value: T): void;
  reject(error: unknown): void;
};

type TicketRecord = {
  ticket: CommandTicket;
  seq?: number;
  outcome?: TicketOutcome;
  ack: Deferred<TicketOutcome>;
  presentation: Deferred<PresentationStatus>;
  ackSettled: boolean;
  presentationSettled: boolean;
  released?: boolean;
};

type LogEntry<P> = {
  envelope: CommandEnvelope<P>;
  ticket: CommandTicket;
  group: TicketRecord;
  outcome?: CommandOutcome;
  confirmedVersion?: number;
  presentationStatus?: PresentationStatus;
  rpcSettled: boolean;
  bytes: number;
};

type ConnectionRecord<P, S, E, F> = {
  connection: CommandConnection<P, S, E, F>;
  active: boolean;
  failure?: SessionFailure;
  cancelReplay?: (failure: SessionFailure) => void;
  fullFrameBusy: boolean;
  fullFrameThroughSeq: number;
  fullFrameRevision?: number;
  fullFrameRetries: number;
  /** Highest seq this connection's host had seen at attach or replay end, and highSeq at that moment. */
  seenSeq: number;
  seenMark: number;
  unsubscribe(): void;
  unsubscribeFrame?: () => void;
  unsubscribeAck?: () => void;
};

function deferred<T>(): Deferred<T> {
  let resolve!: (value: T) => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<T>((yes, no) => { resolve = yes; reject = no; });
  void promise.catch(() => {});
  return { promise, resolve, reject };
}

function failureFor(error: unknown): SessionFailure {
  return error instanceof SessionFailure ? error : new SessionFailure(
    'crash', error instanceof Error ? error.message : String(error)
  );
}

function recoverable(failure: SessionFailure): boolean {
  return failure.code === 'crash' || failure.code === 'trap' ||
    failure.code === 'out-of-memory' || failure.code === 'silence';
}

function blocked(outcome: TicketOutcome | undefined): boolean {
  return outcome?.status === 'refused' || outcome?.status === 'cancelled';
}

function isCommittedApplied(outcome: TicketOutcome | undefined): boolean {
  return outcome?.status === 'applied' && outcome.committed !== false;
}

function limit(value: number | undefined, fallback: number): number {
  return value !== undefined && Number.isSafeInteger(value) && value > 0 ? value : fallback;
}

/** @internal Creates a worker-backed edit session with coverage reconciliation and ordered recovery. */
export function createWorkerEditSession<O, P, S, E, F, X, Q extends QueryTable = {}>(
  options: WorkerEditSessionOptions<O, P, S, E, F, X, Q>
): WorkerEditSession<P, S, E, X, Q> {
  const initial = options.connection.initial;
  const maxAttempts = options.maxRecoveryAttempts === 0 ? 0 : limit(options.maxRecoveryAttempts, 2);
  const source = maxAttempts === 0 ? undefined : cloneData(options.source);
  const sessionId = initial.session;
  const epoch = initial.epoch;
  const maxInFlight = limit(options.maxInFlight, 4);
  const enqueue = options.enqueue ?? queueMicrotask;
  const records = new WeakMap<CommandTicket, TicketRecord>();
  const echoDrafts = new WeakMap<TicketRecord, { draft: CommandDraft<P, X>; originalIntentCount: number }>();
  const pending = new Set<TicketRecord>();
  const listeners = new Set<(state: EditSessionState<S, E>) => void>();
  const log = new Map<number, LogEntry<P>>();
  const dependencies = new Map<number, Set<number>>();
  const cancellations = new Set<IterableIterator<number>>();
  let cancellationQueued = false;
  let cancellationDone: Deferred<void> | undefined;
  const presentations = new Map<number, LogEntry<P>>();
  const uncoveredPresentations = new Map<number, LogEntry<P>>();
  const futurePresentations = new Map<number, Deferred<PresentationStatus>>();
  const blockedSeqs = createBlockedLedger();
  const inFlight = new Map<number, boolean>();
  let ticketId = 0;
  let highSeq = initial.throughSeq;
  let incarnation = initial.incarnation;
  let checkpoint: Checkpoint | undefined;
  let connection: ConnectionRecord<P, S, E, F> | undefined;
  let state: EditSessionState<S, E> = {
    epoch, version: initial.version, throughSeq: initial.throughSeq,
    presentedThroughSeq: -1, selection: initial.selection, editState: initial.editState,
  };
  let disposed = false;
  let terminal: SessionFailure | undefined;
  let recovering = false;
  let pumpQueued = false;
  let checkpointBusy = false;
  let checkpointAttemptedAt = -1;
  let framePainted = true;
  let frameCoverage = initial.throughSeq;
  let frameRevision = 0;
  let disposal: Promise<void> | undefined;
  let recoveryAttempt = 0;
  let selectionThroughSeq = initial.throughSeq;
  let editStateThroughSeq = initial.throughSeq;
  let revisionsThroughSeq = initial.throughSeq;
  let knownSelection = initial.revisions?.selection ?? 0;
  let knownLayout = initial.revisions?.layout ?? 0;
  let appliedCount = 0;
  let retainedBytes = 0;
  let unknownCount = 0;
  const diverged = new Set<number>();
  const holdTimeoutMs = options.holdTimeoutMs ?? 2000;
  if (!Number.isFinite(holdTimeoutMs) || holdTimeoutMs < 0) throw new TypeError('Invalid hold timeout');
  const now = options.now ?? Date.now;
  const timer = options.timer ?? ((callback: () => void, ms: number) => {
    const id = setTimeout(callback, ms);
    return () => { clearTimeout(id); };
  });
  type Held = { hold: Hold; opened: number; waiters: Set<() => void> };
  type PendingRead = {
    kind: string; args: unknown[]; watermark: number; attempt: number; done: boolean;
    signal?: AbortSignal;
    connection?: ConnectionRecord<P, S, E, F>;
    result: Deferred<QueryReply>;
    cleanup(): void;
  };
  const holds = new Map<Hold, Held>();
  const reads = new Set<PendingRead>();
  const waitingReads = new Set<PendingRead>();
  const outbox = createOutbox<P, X>({
    inputBurstCap: options.inputBurstCap,
    coalesce: options.coalesce,
  });
  const presenter = createPresenter<X, F>({
    echo: options.echo && {
      paint: (entries) => options.echo!.paint(entries),
      drop: (entries, reason) => options.echo!.drop(entries, reason),
      present(frame, retired, remaining) {
        framePainted = false;
        const result = options.echo!.present(frame, retired, remaining);
        framePainted = !options.transactionalPresentation || result !== false;
        return result;
      },
    },
    maxEchoEntries: options.maxEchoEntries, now: options.now,
    sameCoverageFrames: options.sameCoverageFrames,
    transactionalPresentation: options.transactionalPresentation,
    echoOverflow: options.echoOverflow,
    discardFrame,
    onEchoOverrun: options.onEchoOverrun,
    requestFullFrame: () => requestFullFrame(frameCoverage, frameRevision),
  });

  function live(record: ConnectionRecord<P, S, E, F>): boolean {
    return !disposed && !terminal && record.active && connection === record;
  }

  function finishRead(request: PendingRead, result: QueryReply | SessionFailure): void {
    if (request.done) return;
    request.done = true;
    reads.delete(request);
    waitingReads.delete(request);
    request.cleanup();
    if (result instanceof SessionFailure) request.result.reject(result);
    else request.result.resolve(result);
  }

  function readRefusal(code: string, detail?: string): QueryReply {
    return { ok: false, refusal: { code, ...(detail === undefined ? {} : { detail }) } };
  }

  function checkReads(): void {
    if (disposed || terminal || recovering) return;
    const record = connection;
    if (!record || !live(record)) return;
    for (const request of waitingReads) {
      if (!live(record) || recovering) return;
      if (request.signal?.aborted) {
        finishRead(request, readRefusal('superseded', 'aborted'));
        continue;
      }
      let queued = false;
      for (const ticket of pending) {
        if (ticket.ticket.id <= request.watermark && ticket.seq === undefined && !ticket.ackSettled) {
          queued = true;
          break;
        }
      }
      if (queued) continue;
      waitingReads.delete(request);
      const query = record.connection.client.call.query;
      if (!query) { finishRead(request, readRefusal('unsupported')); continue; }
      const fence = { session: sessionId, incarnation, epoch,
        afterSeq: highSeq > record.seenMark ? highSeq : record.seenSeq };
      const attempt = ++request.attempt;
      request.connection = record;
      const current = () => !request.done && request.attempt === attempt && live(record);
      try {
        void query.call(record.connection.client.call, fence, request.kind, request.args).then((reply) => {
          if (!current()) return;
          if (!isQueryReply(reply) || (reply.ok && (reply.at.incarnation !== fence.incarnation ||
            reply.at.epoch !== fence.epoch || reply.at.throughSeq < fence.afterSeq))) {
            fail(new SessionFailure('message', 'Malformed or unfenced query reply'));
            return;
          }
          const stale = options.queries![request.kind]!.stale;
          if (reply.ok && stale !== 'answer-at-fence' && stale.reject.some((dependency) =>
            (dependency === 'doc' ? state.version : dependency === 'selection' ? knownSelection : knownLayout) > reply.at[dependency])) {
            finishRead(request, readRefusal('superseded'));
          } else finishRead(request, reply);
        }, (error) => { if (current()) fail(failureFor(error)); })
          .catch((error) => { if (current()) fail(failureFor(error)); });
      } catch (error) { if (current()) fail(failureFor(error)); }
    }
  }

  function hold(reason: HoldReason, holdOptions: { signal?: AbortSignal } = {}): Hold {
    let released = disposed || terminal !== undefined;
    const record: Held = { hold: undefined!, opened: now(), waiters: new Set() };
    const signal = holdOptions.signal;
    const owner: Hold = {
      reason, get released() { return released; },
      release() {
        if (released) return;
        released = true;
        holds.delete(owner);
        signal?.removeEventListener('abort', owner.release);
        for (const waiter of record.waiters) waiter();
        record.waiters.clear();
      },
    };
    record.hold = owner;
    if (!released) {
      holds.set(owner, record);
      if (signal?.aborted) owner.release();
      else signal?.addEventListener('abort', owner.release, { once: true });
    }
    return owner;
  }

  function read<K extends keyof Q & string>(kind: K, args: ArgsOf<Q[K]>, readOptions: ReadOptions = {}):
    Promise<Read<ResultOf<Q[K]>, RefusalOf<Q[K]>>> {
    const result = deferred<QueryReply>();
    const request: PendingRead = { kind, args, watermark: ticketId, attempt: 0, done: false,
      signal: readOptions.signal, result, cleanup: () => {} };
    const captured = [...holds.values()].filter((record) => record.hold !== readOptions.within);
    const waiting = new Map<Held, { release(): void; cancel(): void }>();
    const signal = readOptions.signal;
    const abort = () => finishRead(request, readRefusal('superseded', 'aborted'));
    request.cleanup = () => {
      signal?.removeEventListener('abort', abort);
      for (const [record, waiter] of waiting) { record.waiters.delete(waiter.release); waiter.cancel(); }
      waiting.clear();
    };
    reads.add(request);
    if (disposed) finishRead(request, readRefusal('disposed'));
    else if (terminal) finishRead(request, terminal);
    else if (!options.queries || !hasOwn(options.queries, kind) || !connection?.connection.client.call.query) {
      finishRead(request, readRefusal('unsupported'));
    } else if (signal?.aborted) abort();
    else if (captured.some((record) => now() - record.opened >= holdTimeoutMs)) finishRead(request, readRefusal('busy'));
    else {
      signal?.addEventListener('abort', abort, { once: true });
      if (captured.length === 0) { waitingReads.add(request); checkReads(); }
      else for (const record of captured) {
        const release = () => {
          if (signal?.aborted) { abort(); return; }
          if (now() - record.opened >= holdTimeoutMs) { finishRead(request, readRefusal('busy')); return; }
          waiting.get(record)!.cancel();
          waiting.delete(record);
          record.waiters.delete(release);
          if (waiting.size === 0 && !request.done) {
            request.watermark = ticketId;
            waitingReads.add(request);
            checkReads();
          }
        };
        const waiter = { release, cancel: () => {} };
        waiting.set(record, waiter);
        record.waiters.add(release);
        waiter.cancel = timer(() => finishRead(request, readRefusal('busy')), Math.max(0, holdTimeoutMs - (now() - record.opened)));
      }
    }
    return result.promise as Promise<Read<ResultOf<Q[K]>, RefusalOf<Q[K]>>>;
  }

  function releaseHolds(): void {
    for (const owner of holds.keys()) owner.release();
  }

  function discardFrame(frame: FrameEnvelope<F> | undefined): void {
    if (frame === undefined) return;
    try { options.discardFrame?.(frame); } catch {}
  }

  function publish(): void {
    state = { ...state, presentedThroughSeq: presenter.presentedThroughSeq };
    const published = state;
    for (const listener of [...listeners]) {
      if (disposed || terminal || state !== published) break;
      if (!listeners.has(listener)) continue;
      try { listener(published); } catch {}
    }
  }

  function release(record: TicketRecord): void {
    if (!record.ackSettled || !record.presentationSettled || record.released) return;
    record.released = true;
    pending.delete(record);
  }

  function settlePresentation(record: TicketRecord, status: PresentationStatus): void {
    if (record.presentationSettled) return;
    record.presentationSettled = true;
    echoDrafts.delete(record);
    record.presentation.resolve(status);
    if (record.seq !== undefined) {
      presentations.delete(record.seq);
      uncoveredPresentations.delete(record.seq);
      futurePresentations.get(record.seq)?.resolve(status);
      futurePresentations.delete(record.seq);
    }
    release(record);
  }

  function settleAck(record: TicketRecord, outcome: TicketOutcome): void {
    if (record.ackSettled) return;
    record.ackSettled = true;
    record.outcome = outcome;
    record.ack.resolve(outcome);
    release(record);
  }

  function makeTicket(): TicketRecord {
    const ack = deferred<TicketOutcome>();
    const presentation = deferred<PresentationStatus>();
    const record: TicketRecord = {
      ticket: {
        id: ++ticketId, get seq() { return records.get(record.ticket)!.seq; },
        get ack() { return records.get(record.ticket)!.ack.promise; },
        get presented() { return records.get(record.ticket)!.presentation.promise; },
      },
      ack, presentation,
      ackSettled: false, presentationSettled: false,
    };
    records.set(record.ticket, record);
    pending.add(record);
    return record;
  }

  function cancelTickets(tickets: readonly CommandTicket[], reason: string): void {
    for (const ticket of tickets) {
      const record = records.get(ticket)!;
      settleAck(record, { seq: record.seq, status: 'cancelled', reason });
      settlePresentation(record, 'dropped');
    }
    presenter.drop(tickets, reason);
  }

  function cancelQueued(tickets: readonly CommandTicket[], synchronous = false): void {
    let removed = cancelOutboxAfter(outbox, tickets);
    while (true) {
      cancelTickets(removed.map((entry) => entry.tickets[0]!), 'dependency-refused');
      if (removed.length < 64) {
        if (waitingReads.size) checkReads();
        return;
      }
      if (!synchronous) {
        enqueueCancellation(() => cancelQueued([]));
        return;
      }
      removed = cancelOutboxAfter(outbox, []);
    }
  }

  function enqueueCancellation(callback: () => void): void {
    try {
      enqueue(() => {
        if (disposed || terminal) return;
        try { presenter.event(callback); } catch (error) { terminate(failureFor(error)); }
      });
    } catch (error) { terminate(failureFor(error)); }
  }

  function dependencyBlocked(ticket: CommandTicket): boolean {
    const record = records.get(ticket);
    if (record?.seq !== undefined && blockedSeqs.contains(record.seq)) return true;
    const logged = record?.seq === undefined ? undefined : log.get(record.seq);
    return blocked(logged?.outcome ?? record?.outcome);
  }

  function blockSequence(seq: number): void {
    try { blockedSeqs.add(seq); }
    catch (error) { terminate(failureFor(error)); throw error; }
  }

  function report(phase: RecoveryEvent['phase'], failure?: SessionFailure): void {
    try {
      options.recovery.onRecovery?.({
        phase, incarnation, attempt: recoveryAttempt,
        diverged: [...diverged].sort((a, b) => a - b), historyReset: checkpoint !== undefined, failure,
      });
    } catch {}
  }

  function recordOutcome(entry: LogEntry<P>, outcome: CommandOutcome, replay: boolean, force = false): void {
    if (entry.outcome && !replay && !force) return;
    if (entry.outcome && outcome.status === 'duplicate') return;
    const previous = entry.outcome;
    const group = entry.group;
    if (!previous) unknownCount -= 1;
    if (isCommittedApplied(previous)) appliedCount -= 1;
    entry.outcome = outcome;
    if (isCommittedApplied(outcome)) appliedCount += 1;
    inFlight.delete(outcome.seq);
    if (blocked(outcome)) blockSequence(outcome.seq);
    group.outcome = outcome;
    settleAck(group, outcome);
    if (!blocked(outcome)) {
      if (entry.presentationStatus !== undefined) {
        settlePresentation(group, entry.presentationStatus);
      }
      return;
    }
    settlePresentation(group, 'dropped');
    presenter.drop([entry.ticket], outcome.reason ?? outcome.status);
    cancelQueued([entry.ticket]);
    if (replay && !blocked(previous)) {
      diverged.add(entry.envelope.seq);
      cancelLoggedDescendants(entry.envelope.seq);
      report('diverged');
    }
  }

  function cancelLoggedDescendants(seq: number): void {
    const children = dependencies.get(seq);
    if (children) {
      cancellations.add(children.values());
      dependencies.delete(seq);
    }
    if (cancellationQueued) return;
    drainCancellations();
  }

  function drainCancellations(synchronous = false): void {
    cancellationQueued = true;
    if (disposed || terminal) return;
    let remaining = synchronous ? Infinity : 64;
    while (cancellations.size && remaining > 0) {
      remaining -= 1;
      const iterator = cancellations.values().next().value!;
      const next = iterator.next();
      if (next.done) { cancellations.delete(iterator); continue; }
      const children = dependencies.get(next.value);
      if (children) {
        cancellations.add(children.values());
        dependencies.delete(next.value);
      }
      const entry = log.get(next.value);
      if (!entry || blocked(entry.outcome)) continue;
      blockSequence(entry.envelope.seq);
      if (!entry.outcome) unknownCount -= 1;
      if (isCommittedApplied(entry.outcome)) appliedCount -= 1;
      entry.outcome = { seq: entry.envelope.seq, status: 'cancelled', reason: 'dependency-refused' };
      inFlight.delete(entry.envelope.seq);
      if (recovering) diverged.add(entry.envelope.seq);
      entry.group.outcome = entry.outcome;
      settleAck(entry.group, entry.outcome);
      settlePresentation(entry.group, 'dropped');
      presenter.drop([entry.ticket], 'dependency-refused');
      cancelQueued([entry.ticket], synchronous);
    }
    if (cancellations.size) {
      cancellationDone ??= deferred<void>();
      enqueueCancellation(drainCancellations);
    } else {
      cancellationQueued = false;
      cancellationDone?.resolve();
      cancellationDone = undefined;
    }
  }

  function acceptAck(record: ConnectionRecord<P, S, E, F>, ack: CommandAck<S, E, F>, replay = false): boolean {
    const frame = ack?.frame;
    if (!live(record)) { discardFrame(frame); return false; }
    if (!isCommandAck(ack)) {
      discardFrame(frame);
      fail(new SessionFailure('message', 'Malformed command acknowledgement'));
      return false;
    }
    if (ack.session !== sessionId || ack.incarnation !== incarnation || ack.epoch !== epoch) {
      discardFrame(ack.frame);
      return false;
    }
    if (ack.throughSeq > highSeq || (ack.frame && ack.frame.coversThroughSeq > highSeq) ||
      ack.outcomes.some((outcome) => outcome.seq > ack.throughSeq)) {
      discardFrame(ack.frame);
      fail(new SessionFailure('message', 'Command acknowledgement exceeds sent coverage'));
      return false;
    }
    let frameRejected = false;
    let frameHandled = false;
    try {
      presenter.event(() => {
        for (const outcome of ack.outcomes) {
          const entry = log.get(outcome.seq);
          if (entry) recordOutcome(entry, outcome, replay);
        }
        const boundary = log.get(ack.throughSeq);
        if (boundary) boundary.confirmedVersion = ack.version;
        if (ack.throughSeq > state.throughSeq ||
          (ack.throughSeq === state.throughSeq && ack.version >= state.version)) {
          state = { ...state, throughSeq: ack.throughSeq, version: ack.version };
        }
        if ('selection' in ack && ack.throughSeq >= selectionThroughSeq) {
          state = { ...state, selection: ack.selection };
          selectionThroughSeq = ack.throughSeq;
        }
        if ('editState' in ack && ack.throughSeq >= editStateThroughSeq) {
          state = { ...state, editState: ack.editState };
          editStateThroughSeq = ack.throughSeq;
        }
        if (ack.revisions && ack.throughSeq >= revisionsThroughSeq) {
          knownSelection = ack.throughSeq === revisionsThroughSeq
            ? Math.max(knownSelection, ack.revisions.selection) : ack.revisions.selection;
          knownLayout = ack.throughSeq === revisionsThroughSeq
            ? Math.max(knownLayout, ack.revisions.layout) : ack.revisions.layout;
          revisionsThroughSeq = ack.throughSeq;
        }
        if (ack.frame) {
          frameHandled = true;
          frameRejected = presentFrame(ack.frame);
        }
        if (maxAttempts === 0 && cancellations.size === 0) pruneLog(state.throughSeq, true);
        publish();
        options.onAck?.(ack);
      });
    } catch (error) {
      if (!frameHandled) discardFrame(ack.frame);
      terminate(error instanceof SessionFailure ? error : new SessionFailure('message',
        error instanceof Error ? error.message : String(error)));
      return false;
    }
    wake();
    return frameRejected;
  }

  function presentFrame(frame: FrameEnvelope<F>): boolean {
    const previousCoverage = presenter.presentedThroughSeq;
    const previousPainted = framePainted;
    framePainted = true;
    frameCoverage = frame.coversThroughSeq;
    frameRevision = frame.revision;
    try {
      const disposition = presenter.present(frame);
      const rejected = !!options.transactionalPresentation && !framePainted;
      if (disposition !== 'presented' || presenter.presentedThroughSeq <= previousCoverage) return rejected;
      const status = framePainted ? 'presented' : 'dropped';
      const throughSeq = Math.min(presenter.presentedThroughSeq, frame.coversThroughSeq);
      for (const [seq, entry] of uncoveredPresentations) {
        if (seq > throughSeq) break;
        uncoveredPresentations.delete(seq);
        entry.presentationStatus = status;
        if (!entry.outcome || blocked(entry.outcome)) continue;
        settlePresentation(entry.group, status);
      }
      return rejected;
    } finally {
      framePainted = previousPainted;
    }
  }

  function acceptFrame(record: ConnectionRecord<P, S, E, F>, frame: FrameEnvelope<F>): void {
    if (!live(record) || recovering) { discardFrame(frame); return; }
    if (frame === undefined || !isCommandAck({ session: sessionId, incarnation, epoch,
      throughSeq: state.throughSeq, version: state.version, outcomes: [], frame })) {
      discardFrame(frame);
      fail(new SessionFailure('message', 'Malformed command frame'));
      return;
    }
    if (frame.coversThroughSeq > state.throughSeq) { discardFrame(frame); return; }
    try {
      presenter.event(() => { presentFrame(frame); publish(); });
    } catch (error) {
      terminate(error instanceof SessionFailure ? error : new SessionFailure('message',
        error instanceof Error ? error.message : String(error)));
    }
  }

  function close(record: ConnectionRecord<P, S, E, F>, failure: SessionFailure): Promise<void> {
    if (!record.active) return Promise.resolve();
    record.active = false;
    record.failure = failure;
    try { record.unsubscribe(); } catch {}
    try { record.unsubscribeFrame?.(); } catch {}
    try { record.unsubscribeAck?.(); } catch {}
    for (const request of reads) if (request.connection === record) {
      request.connection = undefined;
      request.attempt += 1;
      waitingReads.add(request);
    }
    record.cancelReplay?.(failure);
    record.cancelReplay = undefined;
    try { return Promise.resolve(record.connection.client.dispose()).catch(() => {}); }
    catch { return Promise.resolve(); }
  }

  function attach(next: CommandConnection<P, S, E, F>): ConnectionRecord<P, S, E, F> {
    const record: ConnectionRecord<P, S, E, F> = {
      connection: next, active: true, fullFrameBusy: false, fullFrameThroughSeq: -1, fullFrameRetries: 0,
      seenSeq: highSeq, seenMark: highSeq, unsubscribe: () => {},
    };
    connection = record;
    const unsubscribe = next.client.onFailure((failure) => {
      if (!live(record)) return;
      if (recovering) {
        void close(record, failure);
      } else fail(failure);
    });
    if (record.active) record.unsubscribe = unsubscribe;
    else unsubscribe();
    if (record.active && next.client.onFrame) {
      const unsubscribeFrame = next.client.onFrame((frame) => acceptFrame(record, frame));
      if (record.active) record.unsubscribeFrame = unsubscribeFrame;
      else unsubscribeFrame();
    }
    if (record.active && next.client.onAck) {
      const unsubscribeAck = next.client.onAck((ack) => acceptAck(record, ack));
      if (record.active) record.unsubscribeAck = unsubscribeAck;
      else unsubscribeAck();
    }
    return record;
  }

  function terminate(failure: SessionFailure): void {
    if (disposed || terminal) return;
    terminal = failure;
    recovering = false;
    for (const request of reads) finishRead(request, failure);
    releaseHolds();
    if (connection) void close(connection, failure);
    outbox.clear();
    inFlight.clear();
    presenter.dispose();
    for (const record of [...pending]) {
      if (!record.ackSettled) {
        record.ackSettled = true;
        record.ack.reject(failure);
      }
      settlePresentation(record, 'dropped');
    }
    log.clear();
    dependencies.clear();
    cancellations.clear();
    cancellationDone?.resolve();
    cancellationDone = undefined;
    presentations.clear();
    uncoveredPresentations.clear();
    for (const waiter of futurePresentations.values()) waiter.resolve('dropped');
    futurePresentations.clear();
    blockedSeqs.clear();
    listeners.clear();
    report('failed', failure);
    diverged.clear();
  }

  function fail(failure: SessionFailure): void {
    if (disposed || terminal) return;
    if (recovering) {
      if (connection) void close(connection, failure);
      return;
    }
    if (maxAttempts === 0 || !recoverable(failure)) {
      if (maxAttempts === 0 && recoverable(failure)) {
        try {
          presenter.event(() => { cancelQueued([], true); drainCancellations(true); });
        } catch (error) {
          terminate(failureFor(error));
          return;
        }
      }
      terminate(failure);
      return;
    }
    recovering = true;
    recoveryAttempt = 0;
    diverged.clear();
    if (connection) void close(connection, failure);
    inFlight.clear();
    checkpointBusy = false;
    void recover(failure).catch((error) => terminate(failureFor(error)));
  }

  function unknown(entry: LogEntry<P>, error: unknown, replay: boolean): SessionFailure {
    const seq = isCommandOutcomeUnknownError(error) ? error.refusal.seq : entry.envelope.seq;
    const target = log.get(seq) ?? entry;
    presenter.event(() => {
      recordOutcome(target, { seq: target.envelope.seq, status: 'refused', reason: 'unknown-outcome' }, replay, true);
      cancelLoggedDescendants(target.envelope.seq);
    });
    return failureFor(error);
  }

  function replayCall(record: ConnectionRecord<P, S, E, F>, envelope: CommandEnvelope<P>): Promise<CommandAck<S, E, F>> {
    return new Promise((resolve, reject) => {
      let settled = false;
      function finish(callback: () => void): void {
        if (settled) return;
        settled = true;
        if (record.cancelReplay === cancel) record.cancelReplay = undefined;
        callback();
      }
      const cancel = (failure: SessionFailure) => finish(() => reject(failure));
      record.cancelReplay = cancel;
      if (!live(record)) {
        cancel(record.failure ?? new SessionFailure('disposed', 'Command connection is closed'));
        return;
      }
      try {
        void record.connection.client.call.edit(options.copyCommandForDispatch?.(envelope) ?? cloneData(envelope)).then(
          (ack) => {
            if (settled) discardFrame(ack?.frame);
            else finish(() => resolve(ack));
          }, (error) => finish(() => reject(error))
        );
      } catch (error) { finish(() => reject(error)); }
    });
  }

  async function recover(firstFailure: SessionFailure): Promise<void> {
    let failure = firstFailure;
    while (!disposed && !terminal && recoveryAttempt < maxAttempts) {
      recoveryAttempt += 1;
      incarnation += 1;
      report('recovering', failure);
      if (disposed || terminal) return;
      let initialFrame: FrameEnvelope<F> | undefined;
      try {
        if (cancellationDone) await cancellationDone.promise;
        if (disposed || terminal) return;
        const retainedCheckpoint = checkpoint && { ...checkpoint, state: checkpoint.state.slice() };
        const next = await options.recovery.reopen({
          source: cloneData(source!), checkpoint: retainedCheckpoint,
          session: sessionId, incarnation, epoch,
          refusedSeqs: blockedSeqs.values(),
          replayAppliedSeqs: [...log.values()].filter((entry) => isCommittedApplied(entry.outcome)).map((entry) => entry.envelope.seq),
        });
        const opened = next.initial;
        initialFrame = opened?.frame;
        if (disposed || terminal) {
          discardFrame(initialFrame);
          initialFrame = undefined;
          try { await next.client.dispose(); } catch {}
          return;
        }
        const record = attach(next);
        if (!isCommandAck(opened) || opened.session !== sessionId || opened.incarnation !== incarnation ||
          opened.epoch !== epoch ||
          opened.throughSeq !== (checkpoint?.throughSeq ?? initial.throughSeq) ||
          opened.version !== (checkpoint?.version ?? initial.version)) {
          throw new SessionFailure('message', 'Reopened command connection has invalid initial state');
        }
        if (next.client.failure) throw next.client.failure;
        presenter.reset();
        uncoveredPresentations.clear();
        for (const [seq, entry] of presentations) {
          entry.presentationStatus = undefined;
          uncoveredPresentations.set(seq, entry);
        }
        state = {
          epoch, version: opened.version, throughSeq: opened.throughSeq,
          presentedThroughSeq: -1, selection: opened.selection, editState: opened.editState,
        };
        selectionThroughSeq = opened.throughSeq;
        editStateThroughSeq = opened.throughSeq;
        revisionsThroughSeq = opened.throughSeq;
        knownSelection = opened.revisions?.selection ?? 0;
        knownLayout = opened.revisions?.layout ?? 0;
        initialFrame = undefined;
        acceptAck(record, opened);
        const entries = [...log.values()].sort((a, b) => a.envelope.seq - b.envelope.seq);
        let replayed = opened.throughSeq;
        for (const entry of entries) {
          if (!live(record)) throw record.failure;
          if (blocked(entry.outcome)) continue;
          const envelope = { ...entry.envelope, incarnation, replay: true as const };
          let ack: CommandAck<S, E, F>;
          try {
            ack = await replayCall(record, envelope);
          } catch (error) {
            if (error instanceof SessionFailure) throw error;
            throw unknown(entry, error, true);
          }
          replayed = Math.max(replayed, entry.envelope.seq);
          acceptAck(record, ack, true);
          if (cancellationDone) await cancellationDone.promise;
          if (!live(record)) {
            if (disposed || terminal) return;
            throw record.failure;
          }
          if (!ack.outcomes.some((outcome) => outcome.seq === entry.envelope.seq)) {
            throw new SessionFailure('message', 'Replay acknowledgement omitted command outcome');
          }
        }
        if (!live(record)) throw record.failure;
        record.seenSeq = replayed;
        record.seenMark = highSeq;
        recovering = false;
        checkpointAttemptedAt = -1;
        report('recovered');
        diverged.clear();
        if (waitingReads.size) checkReads();
        wake();
        return;
      } catch (error) {
        failure = failureFor(error);
        if (connection) void close(connection, failure);
        if (disposed || terminal) return;
        if (!recoverable(failure)) break;
      } finally {
        discardFrame(initialFrame);
      }
    }
    if (!disposed && !terminal) terminate(failure);
  }

  function requestFullFrame(throughSeq: number, revision: number): void {
    const record = connection;
    if (!record || !live(record)) return;
    if (record.fullFrameRevision === undefined || throughSeq > record.fullFrameThroughSeq ||
      (throughSeq === record.fullFrameThroughSeq && revision > record.fullFrameRevision)) {
      record.fullFrameThroughSeq = throughSeq;
      record.fullFrameRevision = revision;
    }
    fetchFullFrame(record);
  }

  function fullFrameSatisfied(record: ConnectionRecord<P, S, E, F>): boolean {
    return presenter.presentedThroughSeq > record.fullFrameThroughSeq ||
      (presenter.presentedThroughSeq === record.fullFrameThroughSeq &&
        (!options.sameCoverageFrames ||
          (presenter.revision !== undefined && presenter.revision >= record.fullFrameRevision!)));
  }

  function fetchFullFrame(record: ConnectionRecord<P, S, E, F>): void {
    if (!live(record) || record.fullFrameBusy) return;
    record.fullFrameBusy = true;
    const requestedThroughSeq = record.fullFrameThroughSeq;
    const requestedRevision = record.fullFrameRevision!;
    try {
      void options.requestFullFrame(record.connection).then((ack) => {
        if (!live(record)) { discardFrame(ack?.frame); return; }
        const previousCoverage = presenter.presentedThroughSeq;
        const frameRejected = acceptAck(record, ack);
        record.fullFrameBusy = false;
        if (!live(record)) return;
        if (frameRejected) {
          record.fullFrameRetries = 0;
          if ((record.fullFrameThroughSeq > requestedThroughSeq ||
            (record.fullFrameThroughSeq === requestedThroughSeq && record.fullFrameRevision! > requestedRevision)) &&
            !fullFrameSatisfied(record)) {
            fetchFullFrame(record);
            return;
          }
          record.fullFrameThroughSeq = -1;
          record.fullFrameRevision = undefined;
          return;
        }
        if (fullFrameSatisfied(record)) {
          record.fullFrameRetries = 0;
          record.fullFrameThroughSeq = -1;
          record.fullFrameRevision = undefined;
          return;
        }
        record.fullFrameRetries = presenter.presentedThroughSeq > previousCoverage ? 0 : record.fullFrameRetries + 1;
        if (record.fullFrameRetries >= 4) {
          const failure = new SessionFailure('message', 'Full frame responses did not satisfy requested coverage and revision');
          if (recovering) void close(record, failure);
          else terminate(failure);
          return;
        }
        fetchFullFrame(record);
      }, (error) => {
        if (!live(record)) return;
        record.fullFrameBusy = false;
        if (recovering) void close(record, failureFor(error));
        else fail(failureFor(error));
      }).catch((error) => {
        if (!live(record)) return;
        record.fullFrameBusy = false;
        fail(failureFor(error));
      });
    } catch (error) {
      record.fullFrameBusy = false;
      if (recovering) void close(record, failureFor(error));
      else fail(failureFor(error));
    }
  }

  function checkpointDue(): boolean {
    if (maxAttempts === 0 || !options.recovery.checkpoint || state.throughSeq === checkpointAttemptedAt) return false;
    const applied = appliedCount;
    const bytes = retainedBytes;
    if (applied === 0 && bytes === 0) return false;
    return options.checkpointPolicy ? options.checkpointPolicy({ applied, bytes }) :
      applied >= 128 || bytes >= 1024 * 1024;
  }

  function validCheckpoint(value: Checkpoint): boolean {
    if (!value || typeof value !== 'object' || !(value.state instanceof Uint8Array) || value.epoch !== epoch ||
      !Number.isSafeInteger(value.throughSeq) || value.throughSeq < 0 ||
      value.throughSeq > state.throughSeq || !Number.isFinite(value.version) ||
      value.version > state.version || (checkpoint &&
        (value.throughSeq < checkpoint.throughSeq || value.version < checkpoint.version))) return false;
    if (value.throughSeq === state.throughSeq && value.version !== state.version) return false;
    const boundary = log.get(value.throughSeq);
    const version = checkpoint?.throughSeq === value.throughSeq ? checkpoint.version :
      initial.throughSeq === value.throughSeq ? initial.version : boundary?.confirmedVersion;
    if (version === undefined || version !== value.version) return false;
    for (const entry of log.values()) {
      if (entry.envelope.seq <= value.throughSeq && !entry.outcome) return false;
    }
    return true;
  }

  function takeCheckpoint(record: ConnectionRecord<P, S, E, F>): void {
    checkpointBusy = true;
    checkpointAttemptedAt = state.throughSeq;
    const request = options.recovery.checkpoint!;
    let snapshot: Promise<Checkpoint>;
    try { snapshot = request(record.connection); }
    catch (error) {
      checkpointBusy = false;
      if (error instanceof SessionFailure) fail(error);
      else wake();
      return;
    }
    void snapshot.then((value) => {
      if (!live(record)) return;
      if (validCheckpoint(value)) {
        checkpoint = { ...value, state: value.state.slice() };
        pruneLog(value.throughSeq);
      }
      checkpointBusy = false;
      wake();
    }, (error) => {
      if (!live(record)) return;
      checkpointBusy = false;
      if (error instanceof SessionFailure) fail(error);
      else wake();
    }).catch((error) => {
      if (!live(record)) return;
      checkpointBusy = false;
      fail(failureFor(error));
    });
  }

  function pruneLog(throughSeq: number, knownPrefix = false): void {
    for (const [seq, entry] of log) {
      if (seq > throughSeq) break;
      if (knownPrefix && (!entry.outcome || !entry.rpcSettled)) break;
      if (isCommittedApplied(entry.outcome)) appliedCount -= 1;
      retainedBytes -= entry.bytes;
      const afterSeq = entry.envelope.afterSeq;
      if (afterSeq !== undefined) {
        const children = dependencies.get(afterSeq);
        children?.delete(seq);
        if (children?.size === 0) dependencies.delete(afterSeq);
      }
      log.delete(seq);
    }
  }

  function settleRpc(entry: LogEntry<P>): void {
    entry.rpcSettled = true;
    if (maxAttempts === 0 && cancellations.size === 0) pruneLog(state.throughSeq, true);
  }

  function held(entry: OutboxEntry<P, X>): boolean {
    if (options.holdForCoalescing?.(entry.draft) !== true) return false;
    const after = entry.draft.options.after;
    if (after) return after.seq !== undefined && inFlight.get(after.seq) === true;
    for (const coalescing of inFlight.values()) if (coalescing) return true;
    return false;
  }

  function eligible(entry: OutboxEntry<P, X>): boolean {
    const after = entry.draft.options.after;
    if (!after) return true;
    const dependency = records.get(after);
    return dependency !== undefined && dependency.seq !== undefined && !dependencyBlocked(after);
  }

  function send(entry: OutboxEntry<P, X>, record: ConnectionRecord<P, S, E, F>): void {
    if (highSeq >= Number.MAX_SAFE_INTEGER) {
      terminate(new SessionFailure('message', 'Command sequence space is exhausted'));
      return;
    }
    const seq = ++highSeq;
    const draft = entry.draft;
    const metadata = draft.options;
    const envelope: CommandEnvelope<P> = {
      session: sessionId, incarnation, seq, basis: { epoch, version: state.version },
      kind: draft.kind, payload: draft.payload,
      ...(metadata.after ? { afterSeq: metadata.after.seq } : {}),
      ...(metadata.expectVersion !== undefined ? { expectVersion: metadata.expectVersion } : {}),
      ...(metadata.inputTime !== undefined ? { inputTime: metadata.inputTime } : {}),
      ...(metadata.lastInputTime !== undefined ? { lastInputTime: metadata.lastInputTime } : {}),
      ...(metadata.group !== undefined ? { group: metadata.group } : {}),
    };
    const group = records.get(entry.tickets[0]!)!;
    const logged: LogEntry<P> = {
      envelope, ticket: entry.tickets[0]!, group, rpcSettled: false, bytes: 0,
    };
    log.set(seq, logged);
    presentations.set(seq, logged);
    uncoveredPresentations.set(seq, logged);
    unknownCount += 1;
    group.seq = seq;
    if (envelope.afterSeq !== undefined) {
      let children = dependencies.get(envelope.afterSeq);
      if (!children) dependencies.set(envelope.afterSeq, children = new Set());
      children.add(seq);
    }
    if (presentations.size > 1024) {
      const oldest = presentations.values().next().value!;
      oldest.presentationStatus = 'dropped';
      settlePresentation(oldest.group, 'dropped');
      presenter.drop([oldest.ticket], 'presentation-overflow');
    }
    if (!live(record)) return;
    inFlight.set(seq, options.holdForCoalescing?.(draft) === true);
    try {
      logged.envelope = cloneData(envelope);
      const bytes = options.commandBytes?.(logged.envelope) ?? 0;
      logged.bytes = Number.isFinite(bytes) && bytes >= 0 ? bytes : 0;
      retainedBytes += logged.bytes;
      const outgoing = options.copyCommandForDispatch?.(envelope) ?? cloneData(logged.envelope);
      void record.connection.client.call.edit(outgoing).then((ack) => {
        try {
          acceptAck(record, ack);
        } finally { settleRpc(logged); }
      }, (error) => {
        try {
          if (live(record)) fail(error instanceof SessionFailure ? error : unknown(logged, error, false));
        } finally { settleRpc(logged); }
      }).catch((error) => { if (live(record)) fail(failureFor(error)); });
    } catch (error) {
      try {
        if (live(record)) fail(error instanceof SessionFailure ? error : unknown(logged, error, false));
      } finally { settleRpc(logged); }
    }
    if (waitingReads.size) checkReads();
  }

  function pump(): void {
    const record = connection;
    if (!record || !live(record) || recovering || checkpointBusy) return;
    if (checkpointDue()) {
      if (inFlight.size === 0 && unknownCount === 0) {
        takeCheckpoint(record);
      }
      return;
    }
    while (live(record) && !recovering && inFlight.size < maxInFlight) {
      const entry = outbox.takeNext(eligible, options.holdForCoalescing ? held : undefined);
      if (!entry) break;
      send(entry, record);
    }
  }

  function wake(): void {
    if (disposed || terminal || recovering || pumpQueued) return;
    pumpQueued = true;
    try {
      enqueue(() => {
        pumpQueued = false;
        try { pump(); } catch (error) { terminate(failureFor(error)); }
      });
    } catch (error) { pumpQueued = false; terminate(failureFor(error)); }
  }

  function submit(kind: string, payload: P, submitOptions: SubmitOptions<X> = {}): CommandTicket {
    const record = makeTicket();
    if (disposed) {
      settleAck(record, { status: 'cancelled', reason: 'disposed' });
      settlePresentation(record, 'dropped');
      return record.ticket;
    }
    if (terminal) {
      record.ackSettled = true;
      record.ack.reject(terminal);
      settlePresentation(record, 'dropped');
      return record.ticket;
    }
    const dependency = submitOptions.after && records.get(submitOptions.after);
    if (submitOptions.after && (!dependency || dependencyBlocked(submitOptions.after))) {
      settleAck(record, { status: 'cancelled', reason: 'dependency-refused' });
      settlePresentation(record, 'dropped');
      return record.ticket;
    }
    const draft: CommandDraft<P, X> = { kind, payload, options: { ...submitOptions } };
    try {
      presenter.event(() => {
        const entry = outbox.enqueue(draft, record.ticket);
        if (entry.tickets.length > 1) {
          records.set(record.ticket, records.get(entry.tickets[0]!)!);
          pending.delete(record);
        }
        echoDrafts.set(records.get(entry.tickets[0]!)!, {
          draft: entry.draft, originalIntentCount: entry.tickets.length,
        });
        const echo = entry.draft.options.echo;
        const painted = echo === undefined ? undefined : { ticket: entry.tickets[0]!, value: echo };
        if (entry.tickets.length > 1) presenter.replace([entry.tickets[0]!], painted);
        else if (painted) presenter.add(painted);
      });
    } catch (error) { terminate(failureFor(error)); }
    wake();
    return record.ticket;
  }

  function attachEcho(
    tickets: readonly CommandTicket[],
    preview: (draft: CommandDraft<P, X>, originalIntentCount: number) => X | undefined
  ): boolean {
    if (disposed || terminal || recovering) return false;
    const groups = new Set<TicketRecord>();
    const attachments: Array<{ record: TicketRecord; entry: NonNullable<ReturnType<typeof echoDrafts.get>>; value: X }> = [];
    for (const ticket of tickets) {
      const record = records.get(ticket);
      if (!record || record.presentationSettled || blocked(record.outcome) ||
        (record.seq !== undefined && record.seq <= presenter.presentedThroughSeq)) return false;
      groups.add(record);
    }
    for (const record of [...groups].sort((a, b) => a.ticket.id - b.ticket.id)) {
      const entry = echoDrafts.get(record);
      if (!entry || entry.draft.options.echo !== undefined) return false;
      const value = preview(entry.draft, entry.originalIntentCount);
      if (value === undefined) return false;
      attachments.push({ record, entry, value });
    }
    return presenter.event(() => {
      if (disposed || terminal || recovering || attachments.some(({ record, entry }) =>
        record.presentationSettled || blocked(record.outcome) ||
        (record.seq !== undefined && record.seq <= presenter.presentedThroughSeq) ||
        echoDrafts.get(record) !== entry || entry.draft.options.echo !== undefined)) return false;
      for (const { entry, value } of attachments) entry.draft.options.echo = value;
      const attached = presenter.attach(attachments.map(({ record, value }) => ({ ticket: record.ticket, value })));
      if (!attached) for (const { entry } of attachments) entry.draft.options.echo = undefined;
      return attached;
    });
  }

  const session: WorkerEditSession<P, S, E, X, Q> = {
    backend: 'worker', get state() { return state; }, submit, attachEcho, read, hold,
    get stage() { return disposed ? 'disposed' : terminal ? 'failed' : recovering ? 'recovering' : 'ready'; },
    whenPresented(seq) {
      if (disposed || terminal) return Promise.resolve('dropped');
      const entry = log.get(seq) ?? presentations.get(seq);
      if (blockedSeqs.contains(seq)) return Promise.resolve('dropped');
      const ticket = entry && records.get(entry.ticket);
      if (ticket) return ticket.presentation.promise;
      if (seq <= highSeq) return presenter.whenPresented(seq);
      const existing = futurePresentations.get(seq);
      if (existing) return existing.promise;
      if (futurePresentations.size >= 1024) {
        const oldest = futurePresentations.keys().next().value!;
        futurePresentations.get(oldest)!.resolve('dropped');
        futurePresentations.delete(oldest);
      }
      const waiter = deferred<PresentationStatus>();
      futurePresentations.set(seq, waiter);
      return waiter.promise;
    },
    subscribe(listener) {
      if (disposed || terminal) return () => {};
      listeners.add(listener);
      return () => { listeners.delete(listener); };
    },
    flush() {
      fenceOutbox(outbox);
      if (terminal) return Promise.reject(terminal);
      const acknowledgements = [...pending].filter((record) => !record.ackSettled).map((record) => record.ack.promise);
      wake();
      const flushed = Promise.all(acknowledgements).then(() => {});
      void flushed.catch(() => {});
      return flushed;
    },
    dispose() {
      if (disposal) return disposal;
      const finished = deferred<void>();
      disposal = finished.promise;
      disposed = true;
      recovering = false;
      for (const request of reads) finishRead(request, readRefusal('disposed'));
      releaseHolds();
      const closing = connection ? close(connection, new SessionFailure('disposed', 'Edit session was disposed')) : Promise.resolve();
      outbox.clear();
      inFlight.clear();
      presenter.dispose();
      for (const record of [...pending]) {
        settleAck(record, { seq: record.seq, status: 'cancelled', reason: 'disposed' });
        settlePresentation(record, 'dropped');
      }
      log.clear();
      dependencies.clear();
      cancellations.clear();
      cancellationDone?.resolve();
      cancellationDone = undefined;
      presentations.clear();
      uncoveredPresentations.clear();
      for (const waiter of futurePresentations.values()) waiter.resolve('dropped');
      futurePresentations.clear();
      blockedSeqs.clear();
      listeners.clear();
      diverged.clear();
      void closing.then(() => finished.resolve());
      return disposal;
    },
  };
  const first = attach(options.connection);
  acceptAck(first, initial);
  if (options.connection.client.failure) fail(options.connection.client.failure);
  return session;
}
