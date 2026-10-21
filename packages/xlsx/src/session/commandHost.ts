import type { PipeCommandPayload } from '../../../../shared/collaboration-pipe/protocol';
import type { AwarenessCursor } from '../collaboration/types';
import { createCommandHost } from '../../../../shared/office-session/commands/host';
import { isCommandOutcomeUnknownError, type CommandOutcome } from '../../../../shared/office-session/commands/types';
import { isTransferResult, transferable, type TransferResult } from '../../../../shared/office-session/protocol';
import { SessionFailure } from '../../../../shared/office-session/types';
import { workbookRecoveryCheckpoint, type WorkbookHandle } from '../wasm/loader';
import { encodeWorkbookRecovery } from './recovery';
import { applyWorkbookCommandIntent, type WorkbookCommandState } from './commandHandlers';
import type { WorkbookCheckpoint } from './methods';
import {
  isWorkbookCommandFrameRequest, validateWorkbookCommandEnvelope,
  type WorkbookCommand, type WorkbookCommandFrameRequest, type WorkbookCommandSelection,
  type WorkbookCommandWireFrame, type WorkbookEditPayload, type WorkbookEditState,
  type WorkbookSharedCommand, type WorkbookSharedWireAck, type WorkbookWireAck,
  type CommandStatus, type WorkbookSharedOutcome,
} from './commandProtocol';

function projectionEquals(
  previous: WorkbookCommandSelection | Partial<WorkbookEditState> | undefined,
  next: WorkbookCommandSelection | Partial<WorkbookEditState> | undefined
): boolean {
  return previous === next || JSON.stringify(previous) === JSON.stringify(next);
}

export function createWorkbookCommandHost(options: {
  handle: WorkbookHandle;
  state: WorkbookCommandState;
  frame(request: WorkbookCommandFrameRequest, coversThroughSeq: number): WorkbookCommandWireFrame;
  onCommit?(): void;
  applyPipe?(payload: PipeCommandPayload<AwarenessCursor>): void;
  revision?: number;
  versions?: [number, string][];
  refusedSeqs?: number[];
  refusedRanges?: [number, number][];
  refusedOutcomes?: [number, CommandOutcome][];
  captureRecovery?(): Uint8Array;
  sessionState?(): { version: number; dirty: boolean };
  now?(): number;
}) {
  const { handle, state } = options;
  const now = options.now ?? (() => performance.now());
  const receipts = new Map<number, WorkbookWireAck>();
  let revision = options.revision ?? 0;
  let directRevisions = 0;
  const versions = new Map<number, string>(options.versions ?? [[revision, handle.version()]]);
  const refusedSeqs = new Set(options.refusedSeqs);
  const refusedRanges: [number, number][] = options.refusedRanges?.map(([start, end]): [number, number] => [start, end]) ?? [];
  const refusedOutcomes = new Map(options.refusedOutcomes);
  if (!options.refusedRanges) {
    for (const seq of [...refusedSeqs].sort((left, right) => left - right)) retainRefusal(seq);
  }
  const requested = new Map<number, { frame: WorkbookCommandFrameRequest | undefined; pipe: boolean }>();
  let represented: { version: string; request: WorkbookCommandFrameRequest } | undefined;
  let pendingVisual = false;
  let frameRequest: WorkbookCommandFrameRequest | undefined;
  let latestFrameRequest: WorkbookCommandFrameRequest | undefined;
  let frameTail: Promise<void> = Promise.resolve();
  let failure: SessionFailure | undefined;
  let publishRecovery = state.historyReset === true;
  const settlements = new Set<Promise<void>>();
  const checkpointSettlements: { settled: Promise<void>[]; preserveHistory: boolean; encodeMs: number }[] = [];

  function retainRefusal(seq: number): void {
    const last = refusedRanges[refusedRanges.length - 1];
    if (last && seq <= last[1] + 1) last[1] = Math.max(last[1], seq);
    else {
      if (refusedRanges.length === 4096) {
        throw new SessionFailure('out-of-memory', 'Workbook refusal ledger exhausted');
      }
      refusedRanges.push([seq, seq]);
    }
  }

  function retainStatus(seq: number, status: CommandStatus): void {
    state.statuses.set(seq, structuredClone(status));
    if (state.statuses.size > 256) state.statuses.delete(state.statuses.keys().next().value!);
    for (const retained of refusedSeqs) if (!state.statuses.has(retained)) refusedSeqs.delete(retained);
    for (const retained of receipts.keys()) if (!state.statuses.has(retained)) receipts.delete(retained);
    for (const retained of refusedOutcomes.keys()) if (!state.statuses.has(retained)) refusedOutcomes.delete(retained);
  }

  function command(input: WorkbookSharedCommand): WorkbookCommand {
    if (input.payload.compatibility) return input.payload.compatibility;
    return { ...state.identity, seq: input.seq, basis: { epoch: input.basis.epoch,
      version: versions.get(input.basis.version) ?? '' }, afterSeq: input.afterSeq,
      kind: input.kind, payload: input.payload.payload,
      ...(isWorkbookCommandFrameRequest(input.payload.frame) ? { frame: input.payload.frame } : {}) } as WorkbookCommand;
  }

  function compatible(input: WorkbookCommand): void {
    validateWorkbookCommandEnvelope(input);
    if (input.session !== state.identity.session || input.incarnation !== state.identity.incarnation ||
      input.basis.epoch !== state.epoch) throw new Error('Workbook command identity or epoch mismatch');
  }

  const shared = createCommandHost<WorkbookEditPayload, WorkbookCommandSelection,
    Partial<WorkbookEditState>, WorkbookCommandWireFrame>({
    session: state.identity.session, incarnation: 1, epoch: state.epoch,
    throughSeq: state.throughSeq, refusedSeqs: refusedRanges.map(([, end]) => end),
    getVersion: () => revision - directRevisions,
    selectionEquals: projectionEquals, editStateEquals: projectionEquals,
    apply(input) {
      const intent = command({ ...input, basis: { ...input.basis, version: input.basis.version + directRevisions } });
      let refusal: string | undefined;
      if (input.payload.compatibility) {
        compatible(intent);
        if (intent.seq !== input.seq || intent.seq !== state.throughSeq + 1) {
          throw new Error(`Expected workbook command sequence ${state.throughSeq + 1}`);
        }
        if (intent.afterSeq !== undefined && !state.statuses.has(intent.afterSeq)) refusal = 'dependency-expired';
      }
      const before = handle.version();
      let reply: WorkbookWireAck;
      try { reply = applyWorkbookCommandIntent(handle, state, intent, refusal, false, options.applyPipe); }
      catch (error) {
        failure = error instanceof SessionFailure ? error : new SessionFailure(
          error instanceof WebAssembly.RuntimeError ? 'trap' : 'crash', 'Workbook command failed',
          error instanceof Error ? error.message : String(error));
        throw failure;
      }
      const status = reply.statuses[0];
      if (input.kind !== 'collabPipe' && isWorkbookCommandFrameRequest(requested.get(input.seq)?.frame)) pendingVisual = true;
      if (publishRecovery) {
        reply.selection = structuredClone(state.selection);
        reply.editState = state.editState;
        publishRecovery = false;
      }
      if (handle.version() !== before) {
        revision += 1;
        versions.set(revision, handle.version());
        if (versions.size > 256) versions.delete(versions.keys().next().value!);
        options.onCommit?.();
      }
      receipts.set(input.seq, reply);
      retainStatus(input.seq, status);
      state.throughSeq = input.seq;
      const projection = { ...(reply.selection ? { selection: reply.selection } : {}),
        ...(reply.editState ? { editState: reply.editState } : {}) };
      return status.status === 'refused'
        ? { status: 'refused', reason: status.reason!, affectedSeqs: status.affectedSeqs,
          ...projection }
        : status.status === 'cancelled'
          ? { status: 'cancelled', reason: status.reason, ...projection }
          : status.status === 'applied'
            ? { status: 'applied', ...(status.committed ? {} : { committed: false as const }), ...projection }
            : { status: 'noop', ...projection };
    },
    frame({ coversThroughSeq }) {
      const pending = requested.get(coversThroughSeq);
      const request = frameRequest ?? pending?.frame;
      if (!isWorkbookCommandFrameRequest(request) || request.sheet >= handle.sheetCount()) return undefined;
      const previous = represented?.request;
      if (!frameRequest && !pendingVisual && pending?.pipe && represented?.version === handle.version() && previous &&
        previous.sheet === request.sheet && previous.zoom === request.zoom && previous.dpr === request.dpr &&
        previous.viewport.x === request.viewport.x && previous.viewport.y === request.viewport.y &&
        previous.viewport.width === request.viewport.width && previous.viewport.height === request.viewport.height) return undefined;
      const paint = options.frame(request, coversThroughSeq);
      represented = { version: paint.version, request: structuredClone(request) };
      pendingVisual = false;
      return transferable({ coversThroughSeq, revision: paint.epoch, payload: paint }, [paint.displayList]);
    },
    snapshot: async () => {
      const checkpoint = checkpointSettlements.shift();
      await Promise.all(checkpoint?.settled ?? []);
      const preserveHistory = checkpoint?.preserveHistory === true;
      const started = now();
      try {
        return encodeWorkbookRecovery(
          options.captureRecovery?.() ?? workbookRecoveryCheckpoint(handle, state.throughSeq, preserveHistory),
          state, revision, [...versions], [...refusedSeqs], options.sessionState?.(), undefined,
          { refusedRanges, refusedOutcomes: [...refusedOutcomes],
            ...(preserveHistory ? { historySelections: state.historySelections } : {}) });
      } finally {
        if (checkpoint) checkpoint.encodeMs = Math.max(0, now() - started);
      }
    },
  });

  async function handler(context: unknown, input: WorkbookSharedCommand): Promise<
    WorkbookSharedWireAck | TransferResult<WorkbookSharedWireAck>
  > {
    if (failure) throw failure;
    if (input.payload.compatibility) compatible(input.payload.compatibility);
    const previousThroughSeq = state.throughSeq;
    let settle!: () => void;
    const settled = new Promise<void>((resolve) => { settle = resolve; });
    settlements.add(settled);
    requested.set(input.seq, { frame: input.payload.frame ??
      ((input.kind === 'applyUpdate' || input.kind === 'collabPipe') ? latestFrameRequest : undefined),
      pipe: input.kind === 'collabPipe' });
    if (input.payload.frame) latestFrameRequest = input.payload.frame;
    if (requested.size > 256) requested.delete(requested.keys().next().value!);
    try {
      const afterSeq = input.afterSeq;
      const dependency = afterSeq === undefined ? undefined :
        refusedRanges.find(([start, end]) => afterSeq >= start && afterSeq <= end);
      const result = await shared.handler(context, { ...input,
        basis: { ...input.basis, version: input.basis.version - directRevisions },
        ...(dependency ? { afterSeq: dependency[1] } : {}) });
      const raw = isTransferResult(result) ? result.value : result;
      state.throughSeq = Math.max(state.throughSeq, raw.throughSeq);
      const outcomes: WorkbookSharedOutcome[] = raw.outcomes.map((outcome) => {
        const settledRefusal = (outcome.status === 'refused' || outcome.status === 'cancelled') &&
          input.session === state.identity.session && input.incarnation === 1 &&
          (input.basis.epoch === state.epoch ||
            (outcome.seq === raw.throughSeq && !state.statuses.has(outcome.seq) && !receipts.has(outcome.seq))) &&
          outcome.seq > previousThroughSeq && outcome.seq <= raw.throughSeq;
        if (settledRefusal) {
          refusedSeqs.add(outcome.seq);
          retainRefusal(outcome.seq);
          refusedOutcomes.set(outcome.seq, { ...outcome,
            ...(outcome.affectedSeqs ? { affectedSeqs: [...outcome.affectedSeqs] } : {}) });
        }
        const validIdentity = input.session === state.identity.session && input.incarnation === 1 && input.basis.epoch === state.epoch;
        const receipt = validIdentity ? receipts.get(outcome.seq) : undefined;
        const retained = validIdentity ? state.statuses.get(outcome.seq) : undefined;
        const workbook: CommandStatus = receipt?.statuses[0] ?? retained ?? { ...outcome, status: outcome.status === 'duplicate' ? 'noop' : outcome.status,
          committed: false, affectedSeqs: outcome.affectedSeqs ?? [outcome.seq] };
        if (!receipt && settledRefusal) {
          retainStatus(outcome.seq, workbook);
        }
        const recovered: CommandOutcome | undefined = validIdentity && outcome.status === 'duplicate'
          ? refusedOutcomes.get(outcome.seq) ?? (retained && (retained.status === 'refused' || retained.status === 'cancelled')
            ? { seq: retained.seq, status: retained.status, committed: false,
              ...(retained.reason ? { reason: retained.reason } : {}),
              ...(retained.status === 'refused' && retained.affectedSeqs ? { affectedSeqs: retained.affectedSeqs } : {}) }
            : undefined) : undefined;
        return { ...(recovered ?? outcome),
          ...(recovered?.affectedSeqs ? { affectedSeqs: [...recovered.affectedSeqs] } : {}),
          workbook: !receipt && retained ? structuredClone(workbook) : workbook };
      });
      const receipt = receipts.get(input.seq);
      const version = raw.version + directRevisions;
      const engineVersion = versions.get(version) ?? handle.version();
      const compatibility: WorkbookWireAck = { throughSeq: raw.throughSeq, version: engineVersion,
        statuses: outcomes.map((outcome) => outcome.workbook!),
        ...(receipt?.selection ? { selection: receipt.selection } : {}),
        ...(receipt?.editState ? { editState: receipt.editState } : {}) };
      if (!receipt && outcomes.some((outcome) => outcome.status === 'refused' || outcome.status === 'cancelled')) {
        compatibility.selection = structuredClone(state.selection);
        compatibility.editState = state.editState;
      }
      const ack: WorkbookSharedWireAck = { ...raw, version, engineVersion, outcomes, compatibility };
      return isTransferResult(result) ? transferable(ack, result.transfer) : ack;
    } catch (error) {
      if (failure) throw failure;
      if (isCommandOutcomeUnknownError(error)) {
        failure = new SessionFailure('crash', error.message);
        throw failure;
      }
      throw error;
    } finally {
      settlements.delete(settled);
      settle();
    }
  }

  function requestFrame(request?: WorkbookCommandFrameRequest): ReturnType<typeof shared.frame> {
    const result = frameTail.then(async () => {
      frameRequest = request ?? latestFrameRequest;
      if (request) latestFrameRequest = request;
      try { return await shared.frame(); }
      finally { frameRequest = undefined; }
    });
    frameTail = result.then(() => {}, () => {});
    return result;
  }

  return {
    validateDirectEdits(): void {
      if (state.throughSeq !== 0 || settlements.size !== 0) throw new Error('Direct edits require an unused command host');
      if (failure) throw failure;
    },
    synchronizeDirectEdits(engineVersion = handle.version()): { revision: number; versions: [number, string][]; refusedSeqs: number[] } {
      revision += 1;
      directRevisions += 1;
      versions.set(revision, engineVersion);
      if (versions.size > 256) versions.delete(versions.keys().next().value!);
      state.bases.set(engineVersion, state.sheetRevision);
      if (state.bases.size > 256) state.bases.delete(state.bases.keys().next().value!);
      state.formattingDirty = true;
      return { revision, versions: [...versions], refusedSeqs: [...refusedSeqs] };
    },
    handler, policy: (input: WorkbookSharedCommand) => ({ ...shared.policy(input),
      lane: (input.kind === 'applyUpdate' || input.kind === 'collabPipe') ? 'collab' as const : 'input' as const,
      userInput: input.kind !== 'applyUpdate' && input.kind !== 'collabPipe' }),
    async checkpoint(context: unknown, preserveHistory = false): Promise<WorkbookCheckpoint | TransferResult<WorkbookCheckpoint>> {
      const capture = { settled: [...settlements], preserveHistory, encodeMs: 0 };
      checkpointSettlements.push(capture);
      const result = await shared.checkpoint(context);
      const checkpoint = isTransferResult(result) ? result.value : result;
      const value = { ...checkpoint, version: checkpoint.version + directRevisions, encodeMs: capture.encodeMs };
      return isTransferResult(result) ? transferable(value, result.transfer) : value;
    },
    checkpointPolicy: shared.checkpointPolicy,
    frame: () => requestFrame(),
    rememberFrame(request: WorkbookCommandFrameRequest): void {
      latestFrameRequest = request;
      represented = { version: handle.version(), request: structuredClone(request) };
    },
    requestFrame,
    async command(input: WorkbookCommand): Promise<WorkbookWireAck | TransferResult<WorkbookWireAck>> {
      compatible(input);
      const duplicate = input.seq <= state.throughSeq;
      if (duplicate) {
        const status = state.statuses.get(input.seq);
        if (!status) throw new Error(`Expected workbook command sequence ${state.throughSeq + 1}`);
        const compatibility: WorkbookWireAck = { throughSeq: state.throughSeq, version: handle.version(),
          statuses: [structuredClone(status)] };
        return { ...compatibility, shared: { session: state.identity.session, incarnation: 1, epoch: state.epoch,
          throughSeq: state.throughSeq, version: revision, engineVersion: handle.version(), compatibility,
          outcomes: [{ seq: status.seq, status: status.status,
            workbook: structuredClone(status), ...(status.committed ? {} : { committed: false }) }] } };
      }
      const result = await handler(null, { session: input.session, incarnation: 1, seq: input.seq,
        basis: { epoch: input.basis.epoch, version: revision }, afterSeq: input.afterSeq,
        kind: input.kind, payload: { payload: input.payload, compatibility: input, frame: duplicate ? undefined : input.frame } });
      const ack = isTransferResult(result) ? result.value : result;
      const reply: WorkbookWireAck = { ...ack.compatibility!, shared: ack,
        ...(ack.frame ? { frame: ack.frame.payload } : {}) };
      return isTransferResult(result) ? transferable(reply, result.transfer) : reply;
    },
  };
}
