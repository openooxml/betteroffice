import { XlsxCommandAdmissionError } from './createXlsxCommandStore';
import type { InputCoordinatorHooks, InputDraft, InputSeal } from './inputCoordinator';
import type {
  CommandStatus, WorkbookAck, WorkbookCommand, WorkbookCommandFrameRequest,
  WorkbookCommandSelection, WorkbookSession, XlsxRangeTarget,
} from '@betteroffice/xlsx';

import { WorkerInputRefusal } from './workerInputErrors';

function deferred<T>() {
  let resolve!: (value: T | PromiseLike<T>) => void;
  let rejectPromise!: (error: unknown) => void;
  let rejection: { error: unknown } | null = null;
  const cancellations = new Set<(error: unknown) => void>();
  const promise = new Promise<T>((done, failed) => {
    resolve = done;
    rejectPromise = failed;
  });
  void promise.catch(() => {});
  return {
    promise, resolve,
    reject(error: unknown) {
      rejection = { error };
      rejectPromise(error);
      for (const cancel of cancellations) cancel(error);
      cancellations.clear();
    },
    wait<R>(value: R | Promise<R>): Promise<R> {
      if (rejection) {
        void Promise.resolve(value).catch(() => {});
        return Promise.reject(rejection.error);
      }
      return new Promise<R>((done, failed) => {
        const cancel = (error: unknown) => { cancellations.delete(cancel); failed(error); };
        cancellations.add(cancel);
        void Promise.resolve(value).then((result) => {
          cancellations.delete(cancel);
          done(result);
        }, cancel);
      });
    },
  };
}

function sameCell(a: InputDraft, b: InputDraft): boolean {
  return (
    a.generation === b.generation &&
    a.source === b.source &&
    a.sheet === b.sheet &&
    a.row === b.row &&
    a.col === b.col
  );
}

function sameDraft(a: InputDraft, b: InputDraft): boolean {
  return sameCell(a, b) && a.value === b.value &&
    (!('revision' in a && 'revision' in b) || a.revision === b.revision);
}

function isPromise<T>(value: T | Promise<T>): value is Promise<T> {
  return typeof (value as Promise<T> | null)?.then === 'function';
}

export type PeerlessToolbarPayload = Extract<WorkbookCommand, { kind: 'applyOps' }>['payload'] extends infer P
  ? P extends { toolbar: infer T } ? T : never : never;

export type PeerlessWorkerCommand = WorkbookCommand;

export type PeerlessWorkerCommandInput = PeerlessWorkerCommand extends infer C
  ? C extends PeerlessWorkerCommand ? Pick<C, 'kind' | 'payload'> : never : never;

export interface PeerlessWorkerCommandSession extends Pick<WorkbookSession,
  'commandIdentity' | 'confirmedBasis' | 'throughSeq' | 'onAck' | 'save' | 'whenPresented'> {
  submit(command: PeerlessWorkerCommand): Promise<CommandStatus>;
  onFailure?: WorkbookSession['onFailure'];
  readonly failure?: unknown;
}

export interface PeerlessWorkerCommandTicket {
  readonly seq: number | null;
  readonly result: Promise<CommandStatus>;
}

export interface PeerlessWorkerCommandOptions {
  dispatch?<T>(write: () => T): T;
  acknowledge?<T>(publish: () => T): T;
  after?: PeerlessWorkerCommandTicket | number | null;
  compositionId?: string;
  frame?: WorkbookCommandFrameRequest;
  target?: XlsxRangeTarget;
  selection?: WorkbookCommandSelection;
  selectionRevision?: number;
}

export type PeerlessWorkerInputEvent = {
  type: 'dispatched' | 'accepted' | 'refused' | 'cancelled';
  seq: number | null;
  target: XlsxRangeTarget | null;
  draft: InputDraft | null;
  command?: PeerlessWorkerCommand;
  status?: CommandStatus;
  selectionRevision?: number;
} | {
  type: 'coverage';
  throughSeq: number;
  coversThroughSeq?: number;
  ack: WorkbookAck;
};

export interface PeerlessWorkerInputSeal extends InputSeal {
  compositionDraft?(): InputDraft | null;
  pendingInput?: PeerlessWorkerPendingInput;
}

export interface PeerlessWorkerPendingInput {
  ready: Promise<boolean>;
  target: InputDraft;
  draft(): InputDraft | null;
  options?: PeerlessWorkerCommandOptions;
  ticket?: PeerlessWorkerCommandTicket;
}

export interface PeerlessWorkerInputCoordinatorHooks
  extends Pick<InputCoordinatorHooks, 'generation' | 'seal' | 'sync'> {
  seal(): PeerlessWorkerInputSeal;
  session(): PeerlessWorkerCommandSession | null;
  isReady(): boolean;
  compositionId?(): string | undefined;
  dependency?(draft: InputDraft): PeerlessWorkerCommandOptions['after'];
  target(sheet: number, row: number, col: number): XlsxRangeTarget;
  draftTarget?(draft: InputDraft): XlsxRangeTarget | undefined;
  frame?(): WorkbookCommandFrameRequest | undefined;
  selectionRevision?(): number;
  onEvent?(event: PeerlessWorkerInputEvent, publish?: () => void): void;
  onError?(error: unknown): void;
  maxOutstanding?: number;
  maxQueued?: number;
  readOnly?(): boolean;
}

export interface PeerlessWorkerInputCoordinator {
  readonly draft: InputDraft | null;
  readonly committed: readonly InputDraft[];
  readonly rejected: readonly InputDraft[];
  readonly pending: boolean;
  readonly error: unknown;
  readonly readOnly: boolean;
  readonly throughSeq: number;
  readonly submittedThroughSeq: number;
  readonly recoveryCommands: readonly PeerlessWorkerCommandInput[];
  setDraft(draft: InputDraft | null): void;
  discardRejected(draft: InputDraft): void;
  submit(draft: InputDraft, options?: PeerlessWorkerCommandOptions): PeerlessWorkerCommandTicket;
  reserveInput(input: PeerlessWorkerPendingInput): PeerlessWorkerCommandTicket;
  command(input: PeerlessWorkerCommandInput, options?: PeerlessWorkerCommandOptions): PeerlessWorkerCommandTicket;
  clipboard<T>(
    capture: () => T | Promise<T>,
    command: (input: T) => PeerlessWorkerCommandInput,
    options?: PeerlessWorkerCommandOptions
  ): PeerlessWorkerCommandTicket;
  setReadOnly(readOnly: boolean): PeerlessWorkerCommandTicket;
  runAfterPendingInput<T>(operation: () => T | Promise<T>, fence?: boolean | 'posted' | 'settled'): Promise<T>;
  flush(): Promise<void>;
  whenApplied(seq?: number): Promise<void>;
  save(): Promise<Uint8Array>;
  whenPresented(seq?: number): Promise<void>;
  subscribe(listener: (event: PeerlessWorkerInputEvent) => void): () => void;
  fail(error: unknown): void;
  reset(): void;
  retire(): void;
}

interface PeerlessTicket extends PeerlessWorkerCommandTicket {
  acknowledge?: PeerlessWorkerCommandOptions['acknowledge'];
  lease?: object;
  outcome?: CommandStatus;
  failure?: { error: unknown };
}

interface PeerlessEntry {
  dispatch?: PeerlessWorkerCommandOptions['dispatch'];
  input: PeerlessWorkerCommandInput | null;
  draft: InputDraft | null;
  target: XlsxRangeTarget | null;
  after?: PeerlessWorkerCommandTicket | number;
  ticket: PeerlessTicket;
  result: ReturnType<typeof deferred<CommandStatus>>;
  frame?: WorkbookCommandFrameRequest;
  selection?: WorkbookCommandSelection;
  selectionRevision?: number;
  fence: boolean;
  prepare?: Promise<PeerlessWorkerCommandInput>;
}

interface PeerlessGate {
  run(): void | Promise<void>;
  reject(error: unknown): void;
  recoveryInput?(): PeerlessWorkerCommandInput | null;
}

interface PeerlessInputBarrier {
  ticket?: PeerlessWorkerCommandTicket;
  composition?: Promise<boolean>;
  compositionDraft?(): InputDraft | null;
  pendingTarget?: InputDraft;
}

function commandTarget(input: PeerlessWorkerCommandInput): XlsxRangeTarget | null {
  if ('target' in input.payload) return structuredClone(input.payload.target ?? null);
  if ('toolbar' in input.payload) return structuredClone(input.payload.toolbar.target ?? null);
  if ('selection' in input.payload) return structuredClone(input.payload.selection ?? null);
  return null;
}

function commandSelection(input: PeerlessWorkerCommandInput, selection?: WorkbookCommandSelection): PeerlessWorkerCommandInput {
  return selection && input.kind !== 'setReadOnly' && input.kind !== 'applyUpdate' && input.kind !== 'collabPipe'
    ? { ...input, payload: { ...input.payload, selection } } as PeerlessWorkerCommandInput : input;
}

export function createPeerlessWorkerInputCoordinator(
  hooks: PeerlessWorkerInputCoordinatorHooks
): PeerlessWorkerInputCoordinator {
  const limit = (value: number | undefined, fallback: number) =>
    Number.isSafeInteger(value) && value! > 0 ? Math.min(value!, 128) : fallback;
  const maxOutstanding = limit(hooks.maxOutstanding, 64);
  const maxQueued = limit(hooks.maxQueued, 128);
  const listeners = new Set<(event: PeerlessWorkerInputEvent) => void>();
  const pendingGates = new Set<object>();
  let session: PeerlessWorkerCommandSession | null = null;
  let identity = '';
  let generation = hooks.generation();
  let cycle = deferred<never>();
  let queue: (PeerlessEntry | PeerlessGate)[] = [];
  let recoveryCommands: PeerlessWorkerCommandInput[] = [];
  let outstanding = new Map<number, PeerlessEntry>();
  const acknowledgements = new Map<number, PeerlessTicket>();
  let draft: InputDraft | null = null;
  let rejected: readonly InputDraft[] = [];
  let failure: { error: unknown } | null = null;
  let retired = false;
  let suspended = false;
  let nextSeq = 1;
  let throughSeq = 0;
  let presentationSeq = 0;
  let inputBarrier: PeerlessInputBarrier | undefined;
  let boundary: PeerlessWorkerCommandTicket | undefined;
  let last: PeerlessWorkerCommandTicket | undefined;
  let compositions = new Map<string, PeerlessWorkerCommandTicket>();
  let submitted = new WeakMap<InputDraft, PeerlessWorkerCommandTicket>();
  let compositionIds = new WeakMap<Promise<boolean>, string>();
  let nextCompositionId = 1;
  let localReadOnly = hooks.readOnly?.() ?? false;
  let offAck = () => {};
  let offFailure = () => {};
  let pumping = false;
  let blocked: PeerlessEntry | PeerlessGate | null = null;
  let insertion: number | null = null;
  let admittedReadOnly: boolean | null = null;
  let admittedAfter: { last: PeerlessWorkerCommandTicket | null;
    boundary: PeerlessWorkerCommandTicket | null } | undefined;

  const emit = (event: PeerlessWorkerInputEvent) => {
    let published = false;
    const publish = () => {
      if (published) return;
      published = true;
      for (const listener of listeners) { try { listener(event); } catch {} }
    };
    try { hooks.onEvent?.(event, publish); } catch {}
    publish();
  };
  const cancel = (entry: PeerlessEntry, error: unknown) => {
    if (entry.ticket.outcome || entry.ticket.failure) return;
    entry.ticket.failure = { error };
    if (boundary === entry.ticket) boundary = undefined;
    if (last === entry.ticket) last = undefined;
    entry.result.reject(error);
    emit({ type: 'cancelled', seq: entry.ticket.seq, target: entry.target, draft: entry.draft });
  };
  const invalidate = (error: unknown) => {
    offAck(); offFailure();
    cycle.reject(error);
    for (const entry of [...queue, ...outstanding.values()]) {
      if ('ticket' in entry) cancel(entry, error);
      else entry.reject(error);
    }
    queue = [];
    outstanding = new Map();
    acknowledgements.clear();
    recoveryCommands = [];
    blocked = null;
    cycle = deferred<never>();
    boundary = undefined;
    inputBarrier = undefined;
    last = undefined;
    draft = null;
    rejected = [];
    compositions = new Map();
    compositionIds = new WeakMap();
    submitted = new WeakMap();
    pendingGates.clear();
  };
  const current = () => {
    if (retired) return;
    const next = hooks.session();
    const key = next ? `${next.commandIdentity.session}:${next.commandIdentity.incarnation}:${next.confirmedBasis.epoch}` : '';
    if (next === session && key === identity && generation === hooks.generation()) return;
    invalidate(new XlsxCommandAdmissionError('document-replaced'));
    session = next;
    identity = key;
    generation = hooks.generation();
    suspended = false;
    failure = null;
    localReadOnly = hooks.readOnly?.() ?? false;
    throughSeq = session?.throughSeq ?? 0;
    presentationSeq = throughSeq;
    nextSeq = throughSeq + 1;
    const lease = cycle;
    offAck = session?.onAck((ack) => {
      current();
      if (lease !== cycle || retired || failure) return;
      if (ack.frame && (ack.frame.session !== session?.commandIdentity.session ||
        ack.frame.incarnation !== session?.commandIdentity.incarnation)) return;
      const committed = ack.statuses.filter((status) => status.committed && status.status === 'applied');
      const acknowledge = committed.length ? acknowledgements.get(committed[0].seq)?.acknowledge : undefined;
      for (const status of ack.statuses) receive(status);
      if (ack.throughSeq >= throughSeq) {
        throughSeq = ack.throughSeq;
        const publish = () => emit({ type: 'coverage', throughSeq, coversThroughSeq: ack.frame?.coversThroughSeq, ack });
        if (acknowledge && committed.every((status) => acknowledgements.get(status.seq)?.acknowledge === acknowledge)) {
          acknowledge(publish);
        } else publish();
        for (const seq of acknowledgements.keys()) if (seq <= throughSeq) acknowledgements.delete(seq);
      }
      pump();
    }) ?? (() => {});
    offFailure = session?.onFailure?.((error) => {
      current();
      if (lease === cycle) fail(error);
    }) ?? (() => {});
    if (session?.failure) fail(session.failure);
  };
  const fail = (error: unknown) => {
    if (failure || retired) return;
    const pending = blocked && !('ticket' in blocked) ? [blocked, ...queue] : queue;
    recoveryCommands = pending.flatMap((entry) => {
      if ('ticket' in entry && (entry.ticket.failure || dependencyFailed(entry))) return [];
      const input = 'ticket' in entry ? entry.input : entry.recoveryInput?.();
      return input ? [structuredClone(input)] : [];
    });
    failure = { error };
    localReadOnly = true;
    cycle.reject(error);
    for (const entry of [...queue, ...outstanding.values()]) {
      if ('ticket' in entry) cancel(entry, error);
      else entry.reject(error);
    }
    queue = [];
    outstanding.clear();
    acknowledgements.clear();
    pendingGates.clear();
    try { hooks.onError?.(error); } catch {}
  };
  const dependencyFailed = (entry: PeerlessEntry) => {
    const after = entry.after;
    return typeof after === 'object' && (after as PeerlessTicket).failure;
  };
  const receive = (status: CommandStatus) => {
    const entry = outstanding.get(status.seq);
    if (!entry) return;
    outstanding.delete(status.seq);
    entry.ticket.outcome = status;
    if (status.status === 'applied' || status.status === 'noop') {
      if (entry.draft) rejected = rejected.filter((value) => !sameCell(value, entry.draft!));
      entry.result.resolve(status);
      emit({ type: 'accepted', seq: status.seq, target: entry.target, draft: entry.draft, status });
    } else {
      const error = new PeerlessWorkerCommandRefusal(status, entry.target, entry.draft);
      entry.ticket.failure = { error };
      if (boundary === entry.ticket) boundary = undefined;
      if (last === entry.ticket) last = undefined;
      if (status.status === 'refused' && entry.draft) {
        rejected = [...rejected.filter((value) => !sameCell(value, entry.draft!)), entry.draft].slice(-maxQueued);
        if (!draft) draft = entry.draft;
      }
      entry.result.reject(error);
      emit({ type: status.status, seq: status.seq, target: entry.target, draft: entry.draft, status });
      for (const queued of queue) {
        if ('ticket' in queued && dependencyFailed(queued)) cancel(queued, error);
      }
    }
  };
  const check = () => {
    current();
    if (retired) throw new XlsxCommandAdmissionError('document-replaced');
    if (suspended) throw new XlsxCommandAdmissionError('document-replaced');
    if (failure) throw failure.error;
    if (!hooks.isReady() || !session?.commandIdentity.session || !session.commandIdentity.incarnation) {
      throw new XlsxCommandAdmissionError('editor-unavailable');
    }
  };
  const enqueue = (entry: PeerlessEntry | PeerlessGate) => {
    if (queue.length >= maxQueued) throw new WorkerInputRefusal('Worker input queue is full');
    if (insertion === null) queue.push(entry);
    else queue.splice(insertion++, 0, entry);
  };
  const pump = () => {
    current();
    if (pumping || blocked || failure || retired || suspended) return;
    pumping = true;
    try {
      while (queue.length && !failure && !retired) {
        const entry = queue[0];
        if ('ticket' in entry) {
          if (entry.ticket.failure) { queue.shift(); continue; }
          const failed = dependencyFailed(entry);
          if (failed) { queue.shift(); cancel(entry, failed.error); continue; }
          if (entry.prepare) {
            const lease = cycle;
            blocked = entry;
            void lease.wait(entry.prepare).then((input) => {
              if (lease !== cycle || blocked !== entry) return;
              try {
                entry.input = structuredClone(commandSelection(input, entry.selection));
                entry.target = commandTarget(entry.input);
                entry.prepare = undefined;
              } catch (error) { cancel(entry, error); }
              blocked = null;
              pump();
            }, (error) => {
              if (lease !== cycle || blocked !== entry) return;
              cancel(entry, error);
              blocked = null;
              pump();
            });
            break;
          }
          if (outstanding.size >= maxOutstanding) break;
          const dependency = typeof entry.after === 'object' ? entry.after as PeerlessTicket : undefined;
          if (dependency?.lease && dependency.lease !== cycle) {
            queue.shift(); cancel(entry, new XlsxCommandAdmissionError('document-replaced')); continue;
          }
          const afterSeq = dependency?.outcome ? undefined :
            typeof entry.after === 'number' ? entry.after : entry.after?.seq ?? undefined;
          if (typeof entry.after === 'number' && (!Number.isSafeInteger(entry.after) || entry.after <= 0 || entry.after >= nextSeq)) {
            queue.shift(); cancel(entry, new WorkerInputRefusal('Invalid command dependency')); continue;
          }
          if (entry.after && !dependency?.outcome && afterSeq === undefined) break;
          queue.shift();
          const dispatch = () => {
            const seq = nextSeq++;
            if (entry.frame || !(entry.input?.kind === 'applyOps' && 'toolbar' in entry.input.payload &&
              entry.input.payload.toolbar.id === 'refreshProposals')) presentationSeq = seq;
            Object.defineProperty(entry.ticket, 'seq', { value: seq, configurable: true });
            const command = { ...structuredClone(entry.input!), ...session!.commandIdentity, seq,
              basis: { ...session!.confirmedBasis }, ...(afterSeq ? { afterSeq } : {}),
              ...(entry.frame ? { frame: entry.frame } : {}) } as PeerlessWorkerCommand;
            outstanding.set(seq, entry);
            if (entry.ticket.acknowledge) acknowledgements.set(seq, entry.ticket);
            const lease = cycle;
            emit({ type: 'dispatched', seq, target: entry.target, draft: entry.draft, command,
              selectionRevision: entry.selectionRevision });
            if (lease !== cycle || failure || retired) return;
            try {
              void session!.submit(command).then((status) => {
                current();
                if (lease !== cycle || failure || retired) return;
                receive(status);
                pump();
              }, (error) => {
                current();
                if (lease === cycle) fail(error);
              });
            } catch (error) { fail(error); }
          };
          try { if (entry.dispatch) entry.dispatch(dispatch); else dispatch(); }
          catch (error) { cancel(entry, error); }
          if (failure || retired) break;
        } else {
          queue.shift();
          insertion = 0;
          try {
            const result = entry.run();
            if (isPromise(result)) {
              blocked = entry;
              const lease = cycle;
              void lease.wait(result).then(() => {
                if (lease !== cycle || blocked !== entry) return;
                blocked = null; pump();
              }, (error) => {
                entry.reject(error);
                if (lease !== cycle || blocked !== entry) return;
                blocked = null; pump();
              });
              break;
            }
          } catch (error) { entry.reject(error); }
          finally { insertion = null; }
        }
      }
    } finally { pumping = false; }
  };
  const accept = (
    input: PeerlessWorkerCommandInput | null, value: InputDraft | null,
    options: PeerlessWorkerCommandOptions = {}, fence = false,
    prepare?: Promise<PeerlessWorkerCommandInput>
  ): PeerlessWorkerCommandTicket => {
    const result = deferred<CommandStatus>();
    const ticket: PeerlessTicket = { seq: null, result: result.promise, lease: cycle, acknowledge: options.acknowledge };
    let target: XlsxRangeTarget | null = null;
    try {
      target = input ? commandTarget(input) : null;
      check();
      ticket.lease = cycle;
      if ((admittedReadOnly ?? localReadOnly) && (prepare || input &&
        input.kind !== 'select' && input.kind !== 'setActiveSheet' && input.kind !== 'setReadOnly' && input.kind !== 'applyUpdate' && input.kind !== 'collabPipe' &&
        !(input.kind === 'applyOps' && 'toolbar' in input.payload && input.payload.toolbar.id === 'refreshProposals'))) {
        throw new XlsxCommandAdmissionError('input-failed');
      }
      const entry: PeerlessEntry = {
        input: input ? structuredClone(commandSelection(input, options.selection)) : null,
        draft: value ? Object.freeze({ ...value }) : null,
        target,
        after: options.after === null ? undefined : options.after ??
          (admittedAfter !== undefined ? (fence ? admittedAfter.last : admittedAfter.boundary) ?? undefined : fence ? last : boundary),
        ticket, result, fence, prepare, dispatch: options.dispatch,
        frame: structuredClone('frame' in options ? options.frame : hooks.frame?.()),
        selection: structuredClone(options.selection),
        selectionRevision: 'selectionRevision' in options ? options.selectionRevision : hooks.selectionRevision?.(),
      };
      const previous = queue[queue.length - 1];
      if (!fence && value && previous && 'ticket' in previous && !previous.fence &&
        !previous.ticket.failure && !previous.prepare && previous.draft &&
        sameCell(previous.draft, value) && previous.after === entry.after &&
        JSON.stringify(previous.target) === JSON.stringify(entry.target)) {
        previous.input = entry.input;
        previous.draft = entry.draft;
        previous.frame = entry.frame;
        previous.selectionRevision = entry.selectionRevision;
        submitted.set(value, previous.ticket);
        return previous.ticket;
      }
      enqueue(entry);
      last = ticket;
      if (fence) boundary = ticket;
      pump();
    } catch (error) {
      ticket.failure = { error };
      result.reject(error);
      const saved = value ? Object.freeze({ ...value }) : null;
      if (saved) {
        rejected = [...rejected.filter((value) => !sameCell(value, saved)), saved].slice(-maxQueued);
        if (!draft) draft = saved;
      }
      emit({ type: 'refused', seq: null, target, draft: saved });
    }
    return ticket;
  };
  const submit = (value: InputDraft, options: PeerlessWorkerCommandOptions = {}) => {
    current();
    const existing = options.compositionId ? compositions.get(options.compositionId) : submitted.get(value);
    if (existing && (options.compositionId || !(existing as PeerlessTicket).failure)) return existing;
    let ticket: PeerlessWorkerCommandTicket;
    if (('modified' in value && value.modified === false) || options.compositionId && value.value === '') {
      const outcome: CommandStatus = { seq: 0, status: 'noop', committed: false };
      ticket = { seq: null, result: Promise.resolve(outcome), outcome, lease: cycle } as PeerlessTicket;
    } else if (value.generation !== generation) {
      const result = deferred<CommandStatus>();
      const error = new XlsxCommandAdmissionError('document-replaced');
      result.reject(error);
      ticket = { seq: null, result: result.promise, failure: { error } } as PeerlessTicket;
    } else {
      try {
        const target = options.target ?? hooks.draftTarget?.(value) ?? hooks.target(value.sheet, value.row, value.col);
        ticket = accept({ kind: 'editCell', payload: {
          args: [value.sheet, value.row, value.col, value.value], target,
      } }, value, { ...options, after: options.after === undefined ? hooks.dependency?.(value) ?? null : options.after },
      Boolean(options.compositionId || options.selection));
      } catch (error) {
        const result = deferred<CommandStatus>();
        result.reject(error);
        ticket = { seq: null, result: result.promise, failure: { error } } as PeerlessTicket;
        const saved = Object.freeze({ ...value });
        rejected = [...rejected.filter((entry) => !sameCell(entry, saved)), saved].slice(-maxQueued);
        if (!draft) draft = saved;
        emit({ type: 'refused', seq: null, target: null, draft: saved });
      }
    }
    submitted.set(value, ticket);
    if (!(ticket as PeerlessTicket).failure && draft && sameCell(draft, value) && draft.value === value.value) draft = null;
    if (options.compositionId) {
      compositions.set(options.compositionId, ticket);
      if (compositions.size > maxQueued) compositions.delete(compositions.keys().next().value!);
    }
    return ticket;
  };
  const runAfterPendingInput = <T>(operation: () => T | Promise<T>, fence: boolean | 'posted' | 'settled' = false,
    reservedInput?: PeerlessWorkerPendingInput): Promise<T> => {
    const result = deferred<T>();
    try {
      check();
      if (pendingGates.size >= maxQueued) throw new WorkerInputRefusal('Worker input queue is full');
      pendingGates.add(result);
      void result.promise.then(() => pendingGates.delete(result), () => pendingGates.delete(result));
      const lease = cycle;
      const readOnly = localReadOnly;
      const after = { last: last ?? null, boundary: boundary ?? null };
      const seal = reservedInput ? { pendingInput: reservedInput } as PeerlessWorkerInputSeal : hooks.seal();
      if (seal.refused) throw new XlsxCommandAdmissionError(seal.refused);
      const previousInput = inputBarrier;
      const waiting = seal.pendingInput?.ready ?? seal.composition;
      const snapshot = seal.pendingInput?.draft ?? seal.compositionDraft;
      const existing = seal.pendingInput?.ticket;
      const composition = !existing && waiting !== previousInput?.composition ? waiting : undefined;
      if (!waiting) hooks.sync();
      const completed = previousInput?.compositionDraft?.() ?? previousInput?.pendingTarget;
      const changed = completed && draft && !sameDraft(completed, draft);
      const input: PeerlessInputBarrier = !composition && previousInput?.composition && !changed ? previousInput :
        existing ? { ticket: previousInput?.ticket ?? existing } : {};
      const reserved = composition ? deferred<CommandStatus>() : undefined;
      const retainInput = (ticket?: PeerlessWorkerCommandTicket) => {
        input.ticket = ticket;
        const clear = () => { if (inputBarrier === input) inputBarrier = undefined; };
        if (ticket) void ticket.result.then(clear, clear);
        else clear();
      };
      const composedTarget = composition && !seal.pendingInput && draft ? { ...draft } : null;
      const captureComposition = (finished: InputDraft) => composedTarget
        ? { ...finished, generation: composedTarget.generation, sheet: composedTarget.sheet,
          row: composedTarget.row, col: composedTarget.col, source: composedTarget.source } : finished;
      const compositionId = composition && !seal.pendingInput ? hooks.compositionId?.() ?? compositionIds.get(composition) ??
        `local-composition:${nextCompositionId++}` : undefined;
      if (composition && compositionId) compositionIds.set(composition, compositionId);
      let sealed: PeerlessWorkerCommandTicket | undefined;
      let compositionSubmitted = false;
      if (reserved) {
        const ticket: PeerlessTicket = { seq: null, result: reserved.promise, lease: cycle };
        Object.defineProperty(ticket, 'seq', { get: () => sealed?.seq ?? null });
        void reserved.promise.then((status) => { ticket.outcome = status; pump(); }, (error) => {
          ticket.failure = { error };
          if (boundary === ticket) boundary = undefined;
          if (last === ticket) last = undefined;
          pump();
        });
        void result.promise.catch(reserved.reject);
        input.composition = composition;
        input.compositionDraft = snapshot;
        input.pendingTarget = seal.pendingInput?.target;
        inputBarrier = input;
        retainInput(ticket);
        if (seal.pendingInput) { seal.pendingInput.ticket = ticket; last = boundary = ticket; }
      } else if (!input.composition) {
        if (waiting && !existing) hooks.sync();
        if (draft && (!existing || !sameDraft(draft, seal.pendingInput!.target))) {
          inputBarrier = input;
          sealed = submit(draft, { after: previousInput?.ticket ?? existing });
          retainInput(sealed);
        }
      }
      const ahead = [...outstanding.values(), ...queue.slice(0, insertion ?? queue.length)
        .filter((entry): entry is PeerlessEntry => 'ticket' in entry)]
        .map((entry) => entry.ticket.result);
      if (input.ticket) ahead.push(input.ticket.result);
      const invoke = () => {
        if (lease !== cycle) throw new XlsxCommandAdmissionError('document-replaced');
        check();
        const previous = admittedReadOnly;
        const previousAfter = admittedAfter;
        admittedReadOnly = readOnly;
        const dependency = sealed ?? input.ticket ?? previousInput?.ticket;
        admittedAfter = dependency ? { last: dependency, boundary: dependency } : after;
        try {
          const value = operation();
          void lease.wait(value).then((value) => {
            seal.restoreFocus?.();
            result.resolve(value);
          }, result.reject);
          return fence === true || fence === 'settled' ? Promise.resolve(value).then(() => {}) : undefined;
        } finally { admittedReadOnly = previous; admittedAfter = previousAfter; }
      };
      const fencedInvoke = async () => {
        const earlier = [...ahead, ...[...outstanding.values()].map((entry) => entry.ticket.result)];
        if (fence === 'settled') await lease.wait(Promise.allSettled(earlier));
        else {
          await lease.wait(Promise.all(earlier));
          if (rejected.length) throw new XlsxCommandAdmissionError('input-failed');
        }
        const previousPumping = pumping;
        const previousInsertion = insertion;
        pumping = true;
        insertion = 0;
        let completion: void | Promise<void>;
        try { completion = invoke(); }
        finally { insertion = previousInsertion; pumping = previousPumping; }
        await completion;
      };
      enqueue({ reject: result.reject, recoveryInput: () => {
        if (!composition || compositionSubmitted || readOnly ||
          (input.ticket as PeerlessTicket | undefined)?.failure ||
          (previousInput?.ticket as PeerlessTicket | undefined)?.failure) return null;
        const finished = snapshot?.();
        if (!finished || finished.generation !== generation || finished.value === '' && !seal.pendingInput ||
          'modified' in finished && finished.modified === false) return null;
        const captured = captureComposition(finished);
        const target = hooks.draftTarget?.(captured) ?? hooks.target(captured.sheet, captured.row, captured.col);
        return { kind: 'editCell', payload: {
          args: [captured.sheet, captured.row, captured.col, captured.value], target,
        } };
      }, run: () => {
        if (composition || fence) return (async () => {
          if (composition) {
            const ended = await lease.wait(composition);
            if (lease !== cycle) throw new XlsxCommandAdmissionError('document-replaced');
            pumping = true;
            insertion = 0;
            try {
              if (ended) {
                let finished = snapshot?.();
                if (!snapshot) { hooks.sync(); finished = draft; }
                const previous = admittedReadOnly;
                admittedReadOnly = readOnly;
                try {
                  if (finished) {
                    compositionSubmitted = true;
                    sealed = submit(captureComposition(finished), { ...seal.pendingInput?.options,
                      compositionId, after: previousInput?.ticket });
                  }
                }
                finally { admittedReadOnly = previous; }
              }
              sealed ??= previousInput?.ticket;
              reserved!.resolve(sealed?.result ?? { seq: 0, status: 'noop', committed: false });
              input.composition = undefined;
              if (fence) enqueue({ run: fencedInvoke, reject: result.reject });
              else invoke();
            } finally {
              insertion = null;
              pumping = false;
              blocked = null;
              pump();
            }
            return;
          }
          await fencedInvoke();
        })();
        return invoke();
      } });
      pump();
    } catch (error) { result.reject(error); }
    return result.promise;
  };
  current();
  return {
    get draft() { current(); return draft; },
    get committed() {
      current();
      return [...outstanding.values(), ...queue.filter((entry): entry is PeerlessEntry => 'ticket' in entry)]
        .flatMap((entry) => entry.draft && !entry.ticket.failure ? [entry.draft] : []);
    },
    get rejected() { current(); return rejected; },
    get pending() { current(); return queue.length > 0 || outstanding.size > 0 || blocked !== null; },
    get error() { current(); return failure?.error ?? null; },
    get readOnly() { current(); return localReadOnly || failure !== null || retired || suspended; },
    get throughSeq() { current(); return throughSeq; },
    get submittedThroughSeq() { current(); return nextSeq - 1; },
    get recoveryCommands() { return structuredClone(recoveryCommands); },
    setDraft(value) {
      current();
      if (!inputBarrier?.pendingTarget && (!value || (!draft && !inputBarrier?.composition))) inputBarrier = undefined;
      draft = value ? Object.freeze({ ...value }) : null;
    },
    discardRejected(value) {
      current();
      rejected = rejected.filter((entry) => !sameCell(entry, value));
      if (draft && sameCell(draft, value)) draft = null;
    },
    submit(value, options) {
      current();
      return submit(value, { ...options, after: inputBarrier?.pendingTarget ? inputBarrier.ticket ?? options?.after : options?.after });
    },
    reserveInput(input) {
      current();
      if (input.ticket) return input.ticket;
      const admitted = runAfterPendingInput(() => {}, false, input);
      void admitted.catch(() => {});
      return inputBarrier?.composition === input.ready && inputBarrier.ticket ? inputBarrier.ticket : {
        seq: null, result: admitted.then((): CommandStatus => ({ seq: 0, status: 'noop', committed: false })),
      };
    },
    command: (input, options) => accept(input, null, options, input.kind === 'setReadOnly'),
    clipboard(capture, command, options) {
      try {
        check();
        if (queue.length >= maxQueued) throw new WorkerInputRefusal('Worker input queue is full');
        const captured = { ...options, target: structuredClone(options?.target), selection: structuredClone(options?.selection),
          frame: structuredClone(options?.frame ?? hooks.frame?.()),
          selectionRevision: options?.selectionRevision ?? hooks.selectionRevision?.() };
        const lease = cycle;
        const data = capture();
        const target = captured.target;
        const prepare = lease.wait(data).then((value) => {
          check();
          if (lease !== cycle) throw new XlsxCommandAdmissionError('document-replaced');
          const input = command(value);
          return target ? { ...input, payload: { ...input.payload, target } } as PeerlessWorkerCommandInput : input;
        });
        void prepare.catch(() => {});
        const result = deferred<CommandStatus>();
        const ticket: PeerlessTicket = { seq: null, result: result.promise, lease: cycle };
        const admitted = runAfterPendingInput(() => {
          const pasted = accept(null, null, captured, true, prepare);
          Object.defineProperty(ticket, 'seq', { get: () => pasted.seq });
          return pasted.result;
        });
        void admitted.then((status) => { ticket.outcome = status; result.resolve(status); }, (error) => {
          ticket.failure = { error }; result.reject(error);
        });
        return ticket;
      } catch (error) {
        const result = deferred<CommandStatus>();
        result.reject(error);
        return { seq: null, result: result.promise };
      }
    },
    setReadOnly(value) {
      current();
      const result = deferred<CommandStatus>();
      const ticket: PeerlessTicket = { seq: null, result: result.promise, lease: cycle };
      const admitted = runAfterPendingInput(() => {
        const mode = accept({ kind: 'setReadOnly', payload: { readOnly: value } }, null, { after: null }, true);
        Object.defineProperty(ticket, 'seq', { get: () => mode.seq });
        return mode.result;
      });
      void admitted.then((status) => { ticket.outcome = status; result.resolve(status); }, (error) => {
        if (!value && ticket.lease === cycle) localReadOnly = true;
        ticket.failure = { error }; result.reject(error);
      });
      localReadOnly = value;
      return ticket;
    },
    runAfterPendingInput,
    flush: () => runAfterPendingInput(() => {}, true),
    whenApplied(seq = nextSeq - 1) {
      try {
        check();
        if (!Number.isSafeInteger(seq) || seq < 0 || seq >= nextSeq) throw new WorkerInputRefusal('Invalid command sequence');
        const lease = cycle;
        return lease.wait(Promise.all([...outstanding.entries()].filter(([pending]) => pending <= seq)
          .map(([, entry]) => entry.ticket.result.catch((error) => {
            if (!(error instanceof PeerlessWorkerCommandRefusal)) throw error;
          })))).then(() => {
          if (lease !== cycle) throw new XlsxCommandAdmissionError('document-replaced');
          check();
        });
      } catch (error) { return Promise.reject(error); }
    },
    save: () => runAfterPendingInput(() => session!.save(), true),
    whenPresented: (seq) => runAfterPendingInput(() => session!.whenPresented(seq ?? presentationSeq), true),
    subscribe(listener) { listeners.add(listener); return () => { listeners.delete(listener); }; },
    fail,
    reset() { current(); invalidate(new XlsxCommandAdmissionError('document-replaced')); suspended = true; },
    retire() { retired = true; invalidate(new XlsxCommandAdmissionError('document-replaced')); listeners.clear(); },
  };
}

export class PeerlessWorkerCommandRefusal extends WorkerInputRefusal {
  constructor(
    readonly status: CommandStatus, readonly target: XlsxRangeTarget | null,
    readonly draft: InputDraft | null
  ) {
    super(status.reason ?? 'Worker command refused');
    this.name = 'PeerlessWorkerCommandRefusal';
  }
}
