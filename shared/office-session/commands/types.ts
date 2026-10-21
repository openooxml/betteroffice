import { hasOwn, type DeferredReply, type TransferResult } from '../protocol';
import type { MethodPolicy, SessionFailure } from '../types';
import type { CommandTable, QueryFence, QueryHandlers, QueryReply, QueryTable, WorkerEditSession } from './spec';

/** @internal Reply error name for an indeterminate command outcome. */
export const COMMAND_OUTCOME_UNKNOWN = 'CommandOutcomeUnknown';

/** @internal Shared refusal and cancellation reasons; formats may supply other strings. */
export type CommandRefusalReason =
  | 'stale-epoch'
  | 'stale-incarnation'
  | 'wrong-session'
  | 'stale-version'
  | 'dependency-refused'
  | 'unknown-outcome'
  | 'disposed';

/** @internal Error data preserved by the existing RPC reply protocol. */
export type CommandOutcomeUnknownError = Error & {
  name: typeof COMMAND_OUTCOME_UNKNOWN;
  refusal: { seq: number };
};

/** @internal Command with a confirmed basis and a sequence assigned at send time. */
export type CommandEnvelope<P> = {
  session: string;
  incarnation: number;
  seq: number;
  basis: { epoch: number; version: number };
  afterSeq?: number;
  expectVersion?: boolean;
  kind: string;
  payload: P;
  inputTime?: number;
  lastInputTime?: number;
  group?: string;
  replay?: true;
};

/** @internal Settled worker outcome for one command sequence. */
export type CommandOutcome = {
  seq: number;
  status: 'applied' | 'noop' | 'refused' | 'cancelled' | 'duplicate';
  reason?: string;
  committed?: false;
  affectedSeqs?: number[];
};

/** @internal Full or incremental frame identified by command coverage and revision. */
export type FrameEnvelope<F> = {
  coversThroughSeq: number;
  revision: number;
  baseRevision?: number;
  payload: F;
};

/** @internal Acknowledgement with authoritative state and an optional covering frame. */
export type CommandAck<S, E, F> = {
  session: string;
  incarnation: number;
  epoch: number;
  throughSeq: number;
  version: number;
  outcomes: CommandOutcome[];
  selection?: S;
  editState?: E;
  frame?: FrameEnvelope<F>;
  revisions?: { selection: number; layout: number };
};

/** @internal Snapshot captured at a settled command boundary. */
export type Checkpoint = {
  state: Uint8Array;
  epoch: number;
  version: number;
  throughSeq: number;
};

/** @internal Methods registered on the existing session RPC. */
export type CommandMethods<P, S, E, F> = {
  edit(envelope: CommandEnvelope<P>): CommandAck<S, E, F>;
  checkpoint(): Checkpoint;
};

/** @internal Local outcome; unsent cancellations have no sequence. */
export type TicketOutcome = Omit<CommandOutcome, 'seq'> & { seq?: number };

/** @internal Non-rejecting presentation settlement for tickets and coverage waiters. */
export type PresentationStatus = 'presented' | 'dropped';

/** @internal Immediate local identity with a live send-time sequence getter. */
export type CommandTicket = {
  readonly id: number;
  readonly seq: number | undefined;
  readonly ack: Promise<TicketOutcome>;
  readonly presented: Promise<PresentationStatus>;
};

/** @internal Main-side command lane, distinct from resident scheduler lanes. */
export type CommandLane = 'input' | 'collab' | 'background';

/** @internal Submission metadata, causal ticket dependency, and optional local echo. */
export type SubmitOptions<X> = {
  lane?: CommandLane;
  after?: CommandTicket;
  expectVersion?: boolean;
  inputTime?: number;
  lastInputTime?: number;
  group?: string;
  coalesceKey?: string;
  fence?: boolean;
  echo?: X;
};

/** @internal Unsent command accepted by the outbox and coalescing hook. */
export type CommandDraft<P, X> = {
  kind: string;
  payload: P;
  options: SubmitOptions<X>;
};

/** @internal Format-owned echo associated with a local ticket and optional sent sequence. */
export type EchoEntry<X> = {
  ticket: CommandTicket;
  seq?: number;
  value: X;
};

/** @internal Synchronous painters; present installs, retires, and redraws atomically. */
export type EchoHooks<X, F> = {
  paint(pending: readonly EchoEntry<X>[]): void;
  drop(entries: readonly EchoEntry<X>[], reason: string): void;
  present(
    frame: FrameEnvelope<F>,
    retired: readonly EchoEntry<X>[],
    remaining: readonly EchoEntry<X>[]
  ): void | boolean;
};

/** @internal Painter timing within one synchronous submit, acknowledgement, or frame event. */
export type EchoOverrunSample = {
  hook: keyof EchoHooks<unknown, unknown>;
  elapsedMs: number;
  eventTotalMs: number;
};

/** @internal Latest confirmed worker state and installed frame coverage. */
export type EditSessionState<S, E> = {
  epoch: number;
  version: number;
  throughSeq: number;
  presentedThroughSeq: number;
  selection: S | undefined;
  editState: E | undefined;
};

/** @internal Shared asynchronous editor contract for worker and legacy backends. */
export interface EditSession<P, S, E, X> {
  readonly backend: 'worker' | 'legacy-sync' | 'legacy-in-thread';
  readonly state: EditSessionState<S, E>;
  submit(kind: string, payload: P, options?: SubmitOptions<X>): CommandTicket;
  attachEcho?(
    tickets: readonly CommandTicket[],
    preview: (draft: CommandDraft<P, X>, originalIntentCount: number) => X | undefined
  ): boolean;
  whenPresented(seq: number): Promise<PresentationStatus>;
  subscribe(listener: (state: EditSessionState<S, E>) => void): () => void;
  flush(): Promise<void>;
  dispose(): Promise<void>;
}

/** @internal Initial acknowledgement and structural subset of the existing RPC client. */
export type CommandConnection<P, S, E, F> = {
  initial: CommandAck<S, E, F>;
  client: {
    readonly call: {
      edit(envelope: CommandEnvelope<P>): Promise<CommandAck<S, E, F>>;
      checkpoint(): Promise<Checkpoint>;
      /** @internal Posts queries after all edits previously handed to this connection. */
      query?(fence: QueryFence, kind: string, args: unknown[]): Promise<QueryReply>;
    };
    readonly failure: SessionFailure | undefined;
    onFailure(listener: (failure: SessionFailure) => void): () => void;
    /** @internal Subscribes to frames posted by this connection's worker. */
    onFrame?(listener: (frame: FrameEnvelope<F>) => void): () => void;
    onAck?(listener: (ack: CommandAck<S, E, F>) => void): () => void;
    dispose(): Promise<void>;
  };
};

/** @internal Confirmed basis and live version provided for format-owned rebasing. */
export type ApplyContext = {
  basis: CommandEnvelope<unknown>['basis'];
  currentVersion: number;
  replayPreviouslyApplied: boolean;
};

/** @internal Format apply result; noop, refusal, cancellation, and applied with committed: false must preserve the version. */
export type ApplyResult<S, E> =
  | { status: 'applied'; committed?: false; selection?: S; editState?: E }
  | { status: 'noop'; selection?: S; editState?: E }
  | { status: 'cancelled'; reason?: string; selection?: S; editState?: E }
  | {
    status: 'refused';
    reason: string;
    selection?: S;
    editState?: E;
    affectedSeqs?: number[];
  };

/** @internal Worker identity, replay ledger, version access, and format hooks. */
export type CommandHostOptions<P, S, E, F> = {
  session: string;
  incarnation: number;
  epoch: number;
  throughSeq?: number;
  retainedOutcomes?: number;
  /** @internal Refuses pruned duplicates with this reason; absent preserves duplicate outcomes. */
  prunedDuplicateReason?: string;
  refusedSeqs?: readonly number[];
  replayAppliedSeqs?: readonly number[];
  commands?: CommandTable;
  queries?: { specs: QueryTable; handlers: QueryHandlers<QueryTable> };
  readOnly?(): boolean;
  layoutRevision?(): number;
  checkRefusals?: boolean;
  selectionEquals?(previous: S | undefined, next: S | undefined): boolean;
  editStateEquals?(previous: E | undefined, next: E | undefined): boolean;
  getVersion(): number;
  apply(
    envelope: CommandEnvelope<P>, context: ApplyContext
  ): ApplyResult<S, E> | Promise<ApplyResult<S, E>>;
  deferFrame?(envelope: CommandEnvelope<P>, result: ApplyResult<S, E>): boolean;
  frame(input: { coversThroughSeq: number; version: number }):
    | FrameEnvelope<F>
    | TransferResult<FrameEnvelope<F>>
    | undefined
    | Promise<FrameEnvelope<F> | TransferResult<FrameEnvelope<F>> | undefined>;
  snapshot(): Uint8Array | Promise<Uint8Array>;
};

/** @internal Edit and checkpoint handlers with admission-time scheduler policies. */
export interface CommandHost<P, S, E, F> {
  handler(
    context: unknown, envelope: CommandEnvelope<P>
  ): Promise<CommandAck<S, E, F> | TransferResult<CommandAck<S, E, F>>>;
  policy(envelope: CommandEnvelope<P>): MethodPolicy;
  checkpoint(context: unknown): Promise<Checkpoint | TransferResult<Checkpoint>>;
  checkpointPolicy(): MethodPolicy;
  query?(context: unknown, fence: QueryFence, kind: string, args: unknown[]):
    Promise<QueryReply | TransferResult<QueryReply>> | DeferredReply<QueryReply>;
  queryPolicy?(fence: QueryFence, kind: string, args: unknown[]): MethodPolicy;
  revise?(change?: { selection?: S; editState?: E }): CommandAck<S, E, F> | undefined;
  /** @internal Builds a frame at a command boundary after admitted edits. */
  frame(): Promise<FrameEnvelope<F> | TransferResult<FrameEnvelope<F>>>;
}

/** @internal Factory contract implemented by host.ts. */
export type CreateCommandHost = <P, S, E, F>(
  options: CommandHostOptions<P, S, E, F>
) => Required<CommandHost<P, S, E, F>>;

/** @internal Queued merged draft and all tickets sharing its eventual sequence and outcome. */
export type OutboxEntry<P, X> = {
  readonly id: number;
  readonly draft: CommandDraft<P, X>;
  readonly tickets: readonly CommandTicket[];
};

/** @internal Pure outbox coalescing and input fairness configuration. */
export type OutboxOptions<P, X> = {
  coalesce?(prev: CommandDraft<P, X>, next: CommandDraft<P, X>): CommandDraft<P, X> | undefined;
  inputBurstCap?: number;
};

/** @internal Unsent queue; takeNext removes an eligible entry subject to dependencies and fences. */
export interface Outbox<P, X> {
  readonly size: number;
  readonly entries: readonly OutboxEntry<P, X>[];
  enqueue(draft: CommandDraft<P, X>, ticket: CommandTicket): OutboxEntry<P, X>;
  takeNext(
    eligible: (entry: OutboxEntry<P, X>) => boolean, hold?: (entry: OutboxEntry<P, X>) => boolean
  ): OutboxEntry<P, X> | undefined;
  remove(predicate: (entry: OutboxEntry<P, X>) => boolean): readonly OutboxEntry<P, X>[];
  clear(): readonly OutboxEntry<P, X>[];
}

/** @internal Factory contract implemented by outbox.ts; never assigns command sequences. */
export type CreateOutbox = <P, X>(options: OutboxOptions<P, X>) => Outbox<P, X>;

/** @internal Result of synchronously reconciling a candidate frame. */
export type FrameDisposition = 'presented' | 'stale' | 'base-mismatch';

/** @internal Pure presenter hooks, echo bound, and injectable painter clock. */
export type PresenterOptions<X, F> = {
  echo?: EchoHooks<X, F>;
  /** @internal Releases rejected frame resources; accepted frames belong to the painter. */
  discardFrame?(frame: FrameEnvelope<F>): void;
  sameCoverageFrames?: boolean;
  /** @internal Commits presentation only after the echo hook accepts it; off by default. */
  transactionalPresentation?: boolean;
  /** @internal Handles echo capacity overflow; defaults to declining new entries. */
  echoOverflow?: 'decline' | 'evict-oldest';
  maxEchoEntries?: number;
  now?(): number;
  requestFullFrame(): void;
  onEchoOverrun?(sample: EchoOverrunSample): void;
};

/** @internal Coverage and echo reconciler; reset invalidates coverage while retaining echo and waiters. */
export interface Presenter<X, F> {
  readonly presentedThroughSeq: number;
  readonly revision: number | undefined;
  readonly pending: readonly EchoEntry<X>[];
  event<T>(callback: () => T): T;
  add(entry: EchoEntry<X>): boolean;
  attach(entries: readonly EchoEntry<X>[]): boolean;
  replace(tickets: readonly CommandTicket[], entry?: EchoEntry<X>): void;
  drop(tickets: readonly CommandTicket[], reason: string): void;
  present(frame: FrameEnvelope<F>): FrameDisposition;
  whenPresented(seq: number): Promise<PresentationStatus>;
  reset(): void;
  dispose(): void;
}

/** @internal Factory contract implemented by presenter.ts. */
export type CreatePresenter = <X, F>(options: PresenterOptions<X, F>) => Presenter<X, F>;

/** @internal Recovery progress, replay divergence, and format history reset signal. */
export type RecoveryEvent = {
  phase: 'recovering' | 'recovered' | 'diverged' | 'failed';
  incarnation: number;
  attempt: number;
  diverged: number[];
  historyReset: boolean;
  failure?: SessionFailure;
};

/** @internal Format reopen and checkpoint hooks used by main-side recovery. */
export type RecoveryHooks<O, P, S, E, F> = {
  reopen(input: {
    source: O;
    checkpoint?: Checkpoint;
    session: string;
    incarnation: number;
    epoch: number;
    refusedSeqs: readonly number[];
    replayAppliedSeqs: readonly number[];
  }): Promise<CommandConnection<P, S, E, F>>;
  checkpoint?(connection: CommandConnection<P, S, E, F>): Promise<Checkpoint>;
  onRecovery?(event: RecoveryEvent): void;
};

/** @internal Worker-backed session limits, recovery, reconciliation, and injected scheduling. */
export type WorkerEditSessionOptions<O, P, S, E, F, X, Q extends QueryTable = {}> = {
  source: O;
  connection: CommandConnection<P, S, E, F>;
  queries?: Q;
  holdTimeoutMs?: number;
  timer?(callback: () => void, ms: number): () => void;
  recovery: RecoveryHooks<O, P, S, E, F>;
  echo?: EchoHooks<X, F>;
  /** @internal Releases rejected frame resources; accepted frames belong to the painter. */
  discardFrame?(frame: FrameEnvelope<F>): void;
  sameCoverageFrames?: boolean;
  /** @internal Commits presentation only after the echo hook accepts it; off by default. */
  transactionalPresentation?: boolean;
  /** @internal Handles echo capacity overflow; defaults to declining new entries. */
  echoOverflow?: 'decline' | 'evict-oldest';
  coalesce?(prev: CommandDraft<P, X>, next: CommandDraft<P, X>): CommandDraft<P, X> | undefined;
  /** @internal Keeps a matching draft queued while the matching command it follows is in flight, so a burst coalesces into one send. */
  holdForCoalescing?(draft: CommandDraft<P, X>): boolean;
  requestFullFrame(connection: CommandConnection<P, S, E, F>): Promise<CommandAck<S, E, F>>;
  maxInFlight?: number;
  inputBurstCap?: number;
  maxEchoEntries?: number;
  maxRecoveryAttempts?: number;
  checkpointPolicy?(stats: { applied: number; bytes: number }): boolean;
  commandBytes?(envelope: CommandEnvelope<P>): number;
  /** @internal May share caller data only with a transport that copies before returning. */
  copyCommandForDispatch?(envelope: CommandEnvelope<P>): CommandEnvelope<P>;
  now?(): number;
  enqueue?(callback: () => void): void;
  onEchoOverrun?(sample: EchoOverrunSample): void;
  onAck?(ack: CommandAck<S, E, F>): void;
};

/** @internal Factory contract implemented by session.ts. */
export type CreateWorkerEditSession = <O, P, S, E, F, X, Q extends QueryTable = {}>(
  options: WorkerEditSessionOptions<O, P, S, E, F, X, Q>
) => WorkerEditSession<P, S, E, X, Q>;

/** @internal Immediate legacy mutation with host-confirmed rendering and unchanged backend choice. */
export type LegacyEditSessionOptions<P, S, E, X> = {
  backend: 'legacy-sync' | 'legacy-in-thread';
  initial: EditSessionState<S, E>;
  apply(kind: string, payload: P, options?: SubmitOptions<X>): {
    outcome: Omit<CommandOutcome, 'seq'>;
    state: EditSessionState<S, E>;
  };
  render(throughSeq: number, presented: (coversThroughSeq: number) => void): void;
  dispose?(): void | Promise<void>;
};

/** @internal Factory contract implemented by legacy.ts. */
export type CreateLegacyEditSession = <P, S, E, X>(
  options: LegacyEditSessionOptions<P, S, E, X>
) => EditSession<P, S, E, X>;

/** @internal Validates envelope metadata without asserting the format payload type. */
export function isCommandEnvelope(value: unknown): value is CommandEnvelope<unknown> {
  return record(value) && typeof value.session === 'string' &&
    counter(value.incarnation) && sequence(value.seq) && record(value.basis) &&
    counter(value.basis.epoch) && finite(value.basis.version) &&
    (value.afterSeq === undefined || sequence(value.afterSeq)) &&
    (value.expectVersion === undefined || typeof value.expectVersion === 'boolean') &&
    typeof value.kind === 'string' && hasOwn(value, 'payload') &&
    (value.inputTime === undefined || finite(value.inputTime)) &&
    (value.lastInputTime === undefined || finite(value.lastInputTime)) &&
    (value.group === undefined || typeof value.group === 'string') &&
    (value.replay === undefined || value.replay === true);
}

/** @internal Validates acknowledgement metadata while leaving format-owned values unknown. */
export function isCommandAck(value: unknown): value is CommandAck<unknown, unknown, unknown> {
  return record(value) && typeof value.session === 'string' &&
    counter(value.incarnation) && counter(value.epoch) && counter(value.throughSeq) &&
    finite(value.version) && Array.isArray(value.outcomes) && value.outcomes.every(outcome) &&
    (value.frame === undefined || frame(value.frame)) &&
    (value.revisions === undefined || (record(value.revisions) &&
      counter(value.revisions.selection) && counter(value.revisions.layout)));
}

/** @internal Identifies the unknown-outcome reply error and its preserved sequence. */
export function isCommandOutcomeUnknownError(value: unknown): value is CommandOutcomeUnknownError {
  return value instanceof Error && value.name === COMMAND_OUTCOME_UNKNOWN &&
    record(value) && record(value.refusal) && sequence(value.refusal.seq);
}

function record(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object';
}

function finite(value: unknown): value is number {
  return typeof value === 'number' && Number.isFinite(value);
}

function counter(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0;
}

function sequence(value: unknown): value is number {
  return counter(value) && value > 0;
}

function outcome(value: unknown): value is CommandOutcome {
  return record(value) && sequence(value.seq) &&
    (value.status === 'applied' || value.status === 'noop' || value.status === 'refused' ||
      value.status === 'cancelled' || value.status === 'duplicate') &&
    (value.reason === undefined || typeof value.reason === 'string') &&
    (value.committed === undefined || value.committed === false) &&
    (value.affectedSeqs === undefined ||
      (Array.isArray(value.affectedSeqs) && value.affectedSeqs.every(sequence)));
}

function frame(value: unknown): value is FrameEnvelope<unknown> {
  return record(value) && counter(value.coversThroughSeq) && finite(value.revision) &&
    (value.baseRevision === undefined || finite(value.baseRevision)) && hasOwn(value, 'payload');
}
