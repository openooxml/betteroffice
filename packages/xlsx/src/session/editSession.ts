import { createWorkerEditSession } from '../../../../shared/office-session/commands/session';
import type {
  EchoHooks, EditSession, FrameEnvelope, RecoveryHooks, WorkerEditSessionOptions,
} from '../../../../shared/office-session/commands/types';
import { SessionFailure } from '../../../../shared/office-session/types';
import {
  validateWorkbookCommandEnvelope,
  type CommandIdentity, type WorkbookAck, type WorkbookCommand, type WorkbookCommandFrame,
  type WorkbookCommandSelection, type WorkbookEditConnection, type WorkbookEditPayload,
  type WorkbookEditState, type WorkbookSharedAck, type WorkbookSharedCommand,
  type WorkbookCommandTicket, type WorkbookSubmitOptions,
} from './commandProtocol';

type SessionOptions<X> = WorkerEditSessionOptions<unknown, WorkbookEditPayload,
  WorkbookCommandSelection, Partial<WorkbookEditState>, WorkbookCommandFrame, X>;

export interface WorkbookEditSession<X = unknown> extends EditSession<WorkbookEditPayload,
  WorkbookCommandSelection, Partial<WorkbookEditState>, X> {
  readonly engineVersion: string;
  readonly failure: SessionFailure | undefined;
  readonly presentationRevision: number | undefined;
  submit(kind: string, payload: WorkbookEditPayload, options?: WorkbookSubmitOptions<X>): WorkbookCommandTicket;
  command(input: WorkbookCommand): Promise<WorkbookAck>;
  runAfterEdits<T>(call: () => Promise<T>): Promise<T>;
  onAck(listener: (ack: WorkbookAck) => void): () => void;
  onFailure(listener: (failure: SessionFailure) => void): () => void;
  bindPresenter(echo: EchoHooks<X, WorkbookCommandFrame>): () => void;
  present(frame: FrameEnvelope<WorkbookCommandFrame>): void;
  markPresented(frame: Pick<WorkbookCommandFrame, 'session' | 'incarnation' | 'coversThroughSeq'> &
    Partial<WorkbookCommandFrame>): boolean;
}

export function createWorkbookEditSession<X = unknown>(options: Omit<SessionOptions<X>,
  'connection' | 'source' | 'recovery' | 'maxRecoveryAttempts' | 'transactionalPresentation' |
  'echoOverflow' | 'sameCoverageFrames'> & {
  connection: WorkbookEditConnection;
  identity?: CommandIdentity;
  source?: unknown;
  recovery?: RecoveryHooks<unknown, WorkbookEditPayload, WorkbookCommandSelection,
    Partial<WorkbookEditState>, WorkbookCommandFrame>;
  onAdmit?(command: WorkbookSharedCommand): void;
  normalizePayload?(payload: WorkbookCommand['payload']): WorkbookCommand['payload'];
}): WorkbookEditSession<X> {
  const initial = options.connection.initial;
  const identity = options.identity ?? { session: initial.session, incarnation: String(initial.incarnation) };
  const frames = new Set<(frame: FrameEnvelope<WorkbookCommandFrame>) => void>();
  const ackListeners = new Set<(ack: WorkbookAck) => void>();
  const failureListeners = new Set<(failure: SessionFailure) => void>();
  const replies: WorkbookSharedAck[] = [];
  const waiting = new Map<number, { resolve(ack: WorkbookAck): void; reject(error: SessionFailure): void }>();
  const duplicates = new Set<{ reject(error: SessionFailure): void }>();
  const nativeTickets = new Set<{ ticket?: WorkbookCommandTicket }>();
  const barriers = new Set<{ order: number; done: Promise<void> }>();
  const acceptedFrames: FrameEnvelope<WorkbookCommandFrame>[] = [];
  let engineVersion = initial.engineVersion;
  let engineThroughSeq = initial.throughSeq;
  let presenter = options.echo;
  let candidate = initial.frame;
  let manual = false;
  let revision: number | undefined;
  let revisionThroughSeq = -1;
  let nextSeq = initial.throughSeq + 1;
  let dispatchOrder = 0;
  let waitingDispatch: Promise<void> | undefined;
  let editState = initial.editState;
  let editStateThroughSeq = initial.throughSeq;
  let failure: SessionFailure | undefined;
  let disposed = false;

  function receipt(ack: WorkbookSharedAck): WorkbookAck {
    return { ...(ack.compatibility ?? { throughSeq: ack.throughSeq, version: ack.engineVersion,
      statuses: ack.outcomes.map((outcome) => outcome.workbook ?? {
        ...outcome, status: outcome.status === 'duplicate' ? 'noop' : outcome.status,
        committed: outcome.status === 'applied' && outcome.committed !== false,
      }), ...(ack.selection ? { selection: ack.selection } : {}),
      ...(ack.editState ? { editState: ack.editState } : {}) }),
      ...(ack.frame ? { frame: ack.frame.payload } : {}) };
  }

  function publishAck(ack: WorkbookSharedAck): void {
    if (failure || disposed) return;
    if (ack.throughSeq >= engineThroughSeq) {
      engineThroughSeq = ack.throughSeq;
      engineVersion = ack.engineVersion;
    }
    const decoded = receipt(ack);
    for (const outcome of ack.outcomes) {
      waiting.get(outcome.seq)?.resolve(decoded);
      waiting.delete(outcome.seq);
    }
    for (const listener of [...ackListeners]) {
      if (!ackListeners.has(listener)) continue;
      try { listener(decoded); } catch {}
    }
  }

  function rejectWaiting(error: SessionFailure): void {
    for (const waiter of waiting.values()) waiter.reject(error);
    waiting.clear();
    for (const duplicate of duplicates) duplicate.reject(error);
    duplicates.clear();
  }

  function checkActive(): void {
    if (failure) throw failure;
    if (disposed) throw new SessionFailure('disposed', 'Session was disposed');
  }

  async function dispatch(input: WorkbookSharedCommand): Promise<WorkbookSharedAck> {
    const blocked = [...barriers].filter((barrier) => (input.payload.dispatchOrder ?? 0) > barrier.order);
    const earlier = waitingDispatch;
    if (blocked.length || earlier) {
      const ready = Promise.all([earlier, ...blocked.map((barrier) => barrier.done)]).then(() => {});
      waitingDispatch = ready;
      try { await ready; } finally { if (waitingDispatch === ready) waitingDispatch = undefined; }
    }
    checkActive();
    const compatibility = input.payload.compatibility;
    if (compatibility && compatibility.seq !== input.seq) {
      throw new SessionFailure('message', 'Workbook compatibility sequence differs from shared dispatch');
    }
    const ack = await options.connection.client.call.edit(compatibility ? {
      ...input, ...(compatibility.afterSeq === undefined ? {} : { afterSeq: compatibility.afterSeq }),
    } : input);
    checkActive();
    if (ack.session === initial.session && ack.incarnation === initial.incarnation && ack.epoch === initial.epoch &&
      ack.editState && ack.throughSeq >= editStateThroughSeq) {
      editState = { ...editState, ...ack.editState };
      editStateThroughSeq = ack.throughSeq;
      return { ...ack, editState };
    }
    return ack;
  }

  const connection: WorkbookEditConnection = {
    initial,
    client: {
      call: {
        edit: (input) => {
          options.onAdmit?.(input);
          nextSeq = Math.max(nextSeq, input.seq + 1);
          const reply = dispatch(input).catch((error: unknown) => {
            throw error instanceof SessionFailure ? error : new SessionFailure('crash',
              error instanceof Error ? error.message : String(error));
          });
          void reply.then((ack) => {
            if (!failure && !disposed && ack.session === initial.session && ack.incarnation === initial.incarnation &&
              ack.epoch === initial.epoch) replies.push(ack);
          }, () => {});
          return reply;
        },
        checkpoint: () => options.connection.client.call.checkpoint(),
      },
      get failure() { return options.connection.client.failure; },
      onFailure: (listener) => options.connection.client.onFailure(listener),
      dispose: () => options.connection.client.dispose(),
      onFrame(listener) {
        frames.add(listener);
        const off = options.connection.client.onFrame?.(listener);
        return () => { frames.delete(listener); off?.(); };
      },
    },
  };

  const shared = createWorkerEditSession<unknown, WorkbookEditPayload, WorkbookCommandSelection,
    Partial<WorkbookEditState>, WorkbookCommandFrame, X>({
    source: undefined, connection,
    recovery: {
      reopen: options.recovery?.reopen ?? (async () => { throw new Error('Workbook command recovery is disabled'); }),
      checkpoint: options.recovery?.checkpoint,
      onRecovery(event) {
        if (event.phase === 'failed' && event.failure) {
          failure = event.failure;
          candidate = undefined;
          replies.length = 0;
          acceptedFrames.length = 0;
          rejectWaiting(failure);
          nativeTickets.clear();
          for (const listener of [...failureListeners]) {
            if (!failureListeners.has(listener)) continue;
            try { listener(failure); } catch {}
          }
        }
        options.recovery?.onRecovery?.(event);
      },
    },
    maxRecoveryAttempts: 0, transactionalPresentation: true, echoOverflow: 'evict-oldest',
    sameCoverageFrames: true, maxInFlight: options.maxInFlight ?? 64,
    maxEchoEntries: options.maxEchoEntries ?? 256, inputBurstCap: options.inputBurstCap,
    coalesce: options.coalesce && ((previous, next) => {
      const merged = options.coalesce!(previous, next);
      return merged && { ...merged, payload: { ...merged.payload,
        dispatchOrder: Math.max(previous.payload.dispatchOrder ?? 0, next.payload.dispatchOrder ?? 0) } };
    }), enqueue: options.enqueue, now: options.now,
    commandBytes: options.commandBytes, checkpointPolicy: options.checkpointPolicy,
    onEchoOverrun: options.onEchoOverrun, requestFullFrame: options.requestFullFrame,
    echo: {
      paint: (entries) => presenter?.paint(entries),
      drop: (entries, reason) => presenter?.drop(entries, reason),
      present(frame, retired, remaining) {
        if (frame.payload.coversThroughSeq !== frame.coversThroughSeq || options.identity &&
          (frame.payload.session !== identity.session || frame.payload.incarnation !== identity.incarnation)) return false;
        candidate = frame;
        if (!manual && !presenter) return false;
        const accepted = manual || presenter!.present(frame, retired, remaining) !== false;
        if (accepted) {
          acceptedFrames.push(frame);
          if (candidate === frame) candidate = undefined;
        }
        return accepted;
      },
    },
  });
  function commitPresentation(): void {
    for (const frame of acceptedFrames.splice(0)) {
      if (frame.coversThroughSeq <= shared.state.presentedThroughSeq) {
        if (frame.coversThroughSeq > revisionThroughSeq) revision = frame.revision;
        else if (frame.coversThroughSeq === revisionThroughSeq) revision = Math.max(revision ?? frame.revision, frame.revision);
        revisionThroughSeq = Math.max(revisionThroughSeq, frame.coversThroughSeq);
      }
    }
  }
  commitPresentation();
  shared.subscribe(() => {
    commitPresentation();
    const ack = replies[0];
    if (ack && ack.throughSeq <= shared.state.throughSeq) {
      replies.shift();
      publishAck(ack);
    }
  });

  function present(frame: FrameEnvelope<WorkbookCommandFrame>): void {
    if (failure || disposed) return;
    for (const listener of [...frames]) listener(frame);
  }

  return {
    backend: shared.backend,
    get state() { return failure ? { ...shared.state,
      editState: { ...shared.state.editState, readOnly: true } } : shared.state; },
    get engineVersion() { return engineVersion; },
    get failure() { return failure; },
    get presentationRevision() { return revision; },
    submit(kind, payload, metadata) {
      const native: { ticket?: WorkbookCommandTicket } = {};
      nativeTickets.add(native);
      try {
        if (options.normalizePayload) {
          const normalized = options.normalizePayload(payload.payload);
          payload = { ...payload, payload: normalized, ...(payload.compatibility ? {
            compatibility: { ...payload.compatibility, payload: normalized } as WorkbookCommand,
          } : {}) };
        }
        const ticket = shared.submit(kind, { ...payload, dispatchOrder: ++dispatchOrder }, metadata);
        native.ticket = ticket;
        void ticket.ack.then(() => { nativeTickets.delete(native); }, () => { nativeTickets.delete(native); });
        if (ticket.seq !== undefined) nextSeq = Math.max(nextSeq, ticket.seq + 1);
        return ticket;
      } catch (error) { nativeTickets.delete(native); throw error; }
    },
    subscribe: (listener) => shared.subscribe(listener),
    flush: () => shared.flush(),
    whenPresented: (seq) => shared.whenPresented(seq),
    runAfterEdits(call) {
      if (failure) return Promise.reject(failure);
      if (disposed) return Promise.reject(new SessionFailure('disposed', 'Session was disposed'));
      const flushed = dispatchOrder === 0 ? undefined : shared.flush();
      const invoke = () => {
        if (failure) throw failure;
        if (disposed) throw new SessionFailure('disposed', 'Session was disposed');
        return call();
      };
      const result = flushed ? flushed.then(invoke) : invoke();
      const barrier = { order: dispatchOrder, done: result.then(() => {}, () => {}) };
      barriers.add(barrier);
      void barrier.done.then(() => { barriers.delete(barrier); });
      return result;
    },
    async command(input) {
      validateWorkbookCommandEnvelope(input);
      if (options.onAdmit) input = structuredClone(input);
      checkActive();
      if (input.session !== identity.session || input.incarnation !== identity.incarnation ||
        input.basis.epoch !== initial.epoch) throw new Error('Workbook command identity or epoch mismatch');
      if (options.normalizePayload) input.payload = options.normalizePayload(input.payload);
      const payload: WorkbookEditPayload = { payload: input.payload, compatibility: input, frame: input.frame };
      if (input.seq < nextSeq) {
        let resolve!: (ack: WorkbookSharedAck) => void;
        let reject!: (error: SessionFailure) => void;
        const reply = new Promise<WorkbookSharedAck>((done, fail) => { resolve = done; reject = fail; });
        const duplicate = { reject };
        duplicates.add(duplicate);
        void dispatch({ session: initial.session, incarnation: initial.incarnation, seq: input.seq,
          basis: { epoch: initial.epoch, version: shared.state.version }, afterSeq: input.afterSeq,
          kind: input.kind, payload: { ...payload, frame: undefined } }).then(resolve, reject);
        try {
          const ack = await reply;
          checkActive();
          publishAck(ack);
          checkActive();
          return receipt(ack);
        } finally { duplicates.delete(duplicate); }
      }
      if ([...nativeTickets].some((native) => native.ticket?.seq === undefined)) {
        throw new Error('Workbook command sequence is awaiting native dispatch');
      }
      if (input.seq !== nextSeq) throw new Error(`Expected workbook command sequence ${nextSeq}`);
      options.onAdmit?.({ session: initial.session, incarnation: initial.incarnation, seq: input.seq,
        basis: { epoch: initial.epoch, version: shared.state.version }, afterSeq: input.afterSeq,
        kind: input.kind, payload });
      nextSeq += 1;
      let resolve!: (ack: WorkbookAck) => void;
      let reject!: (error: SessionFailure) => void;
      const reply = new Promise<WorkbookAck>((done, fail) => { resolve = done; reject = fail; });
      void reply.catch(() => {});
      waiting.set(input.seq, { resolve, reject });
      const metadata = input.payload && typeof input.payload === 'object' ? input.payload : undefined;
      const inputTime = metadata && 'inputTime' in metadata ? metadata.inputTime : undefined;
      const group = metadata && 'undoGroupId' in metadata ? metadata.undoGroupId : undefined;
      const ticket: WorkbookCommandTicket = shared.submit(input.kind, { ...payload, dispatchOrder: ++dispatchOrder }, {
        lane: input.kind === 'applyUpdate' || input.kind === 'collabPipe' ? 'collab' : 'input',
        fence: true,
        ...(typeof inputTime === 'number' && Number.isFinite(inputTime) ? { inputTime } : {}),
        ...(typeof group === 'string' ? { group } : {}) });
      try {
        const outcome = await ticket.ack;
        checkActive();
        if (outcome.seq === undefined || outcome.status === 'cancelled') {
          waiting.delete(input.seq);
          const { workbook, ...status } = outcome;
          return { throughSeq: shared.state.throughSeq, version: engineVersion,
            selection: shared.state.selection, editState: shared.state.editState,
            statuses: [workbook ?? { ...status, seq: input.seq,
              status: outcome.status === 'duplicate' ? 'noop' : outcome.status,
              committed: false, affectedSeqs: [input.seq] }] };
        }
        const ack = await reply;
        checkActive();
        return ack;
      } catch (error) { waiting.delete(input.seq); throw error; }
    },
    onAck(listener) { ackListeners.add(listener); return () => { ackListeners.delete(listener); }; },
    onFailure(listener) {
      if (failure) { listener(failure); return () => {}; }
      failureListeners.add(listener);
      return () => { failureListeners.delete(listener); };
    },
    bindPresenter(echo) {
      presenter = echo;
      if (candidate) present(candidate);
      return () => { if (presenter === echo) presenter = undefined; };
    },
    present,
    markPresented(paint) {
      if (failure || disposed || paint.session !== identity.session || paint.incarnation !== identity.incarnation ||
        !Number.isSafeInteger(paint.coversThroughSeq) || paint.coversThroughSeq < 0 ||
        paint.coversThroughSeq < shared.state.presentedThroughSeq || paint.coversThroughSeq > shared.state.throughSeq) return false;
      manual = true;
      try {
        present({ coversThroughSeq: paint.coversThroughSeq,
          revision: Math.max(paint.epoch ?? 0, (revision ?? 0) + 1), payload: paint as WorkbookCommandFrame });
      } finally { manual = false; }
      return shared.state.presentedThroughSeq >= paint.coversThroughSeq;
    },
    dispose() {
      disposed = true;
      candidate = undefined;
      rejectWaiting(new SessionFailure('disposed', 'Session was disposed'));
      nativeTickets.clear();
      replies.length = 0;
      acceptedFrames.length = 0;
      ackListeners.clear();
      failureListeners.clear();
      return shared.dispose();
    },
  };
}
