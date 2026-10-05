import { SessionFailure } from '../../../../shared/office-session';
import {
  adoptWorkbookPeerVersion, StaleProposalError, type WorkbookCalculationContext, type WorkbookHandle,
} from '../wasm/loader';
import type { WorkbookSession } from './client';
import { WorkbookRecoveryRefusal, workbookEditPeerInternals } from './editPeerInternals';
import { WorkbookPeerHydrationError } from './peerHydrationError';
import {
  applyWorkbookReplayOp,
  validateWorkbookReplayEnvelope,
  WORKBOOK_REPLAY_MUTATORS,
  workbookReplayRefused,
  workbookSessionInternals,
  type WorkbookReplayEnvelope,
  type WorkbookReplayMethod,
  type WorkbookReplayOp,
  type WorkbookReplayReply,
} from './replay';

export interface WorkbookEditPeerOptions {
  session: WorkbookSession;
  peer: WorkbookHandle;
  now?(): number;
  randomSeed?(): number;
  onError?(error: Error): void;
}

export type WorkbookEditPeer = Pick<WorkbookHandle, WorkbookReplayMethod> & {
  readonly state: 'ready' | 'failed';
  readonly error: Error | undefined;
  readonly acknowledgedSequence: number;
  readonly sentSequence: number;
  flush(): Promise<void>;
  save(): Promise<ArrayBuffer>;
  recoverySave(): { bytes: ArrayBuffer; recovery: true };
  dispose(): void;
};

export class WorkbookEditPeerFailedError extends Error {
  constructor(readonly failure: Error) {
    super(`Workbook edit peer failed: ${failure.message}`);
    this.name = 'WorkbookEditPeerFailedError';
  }
}

function randomSeed(): number {
  return globalThis.crypto.getRandomValues(new Uint32Array(1))[0];
}

function buffer(bytes: Uint8Array): ArrayBuffer {
  return new Uint8Array(bytes).buffer;
}

/** @internal */
export function createWorkbookEditPeer(options: WorkbookEditPeerOptions): WorkbookEditPeer {
  const { session, peer } = options;
  const internal = workbookSessionInternals.get(session);
  if (!internal) throw new TypeError('Workbook session does not support edit replay');
  if (internal.editPeerAttached) throw new Error('Workbook session already has an attached edit peer');
  if (internal.initialVersion !== undefined && peer.version() !== internal.initialVersion) {
    if (internal.attachPeer) {
      throw new WorkbookPeerHydrationError('version-mismatch',
        'Workbook peer version differs from retained worker hydration');
    }
    adoptWorkbookPeerVersion(peer, internal.initialVersion);
  }
  const replay = internal.replay;
  const now = options.now ?? Date.now;
  const seed = options.randomSeed ?? randomSeed;
  let error: Error | undefined;
  let precedingCalculation = internal.attachPeer ? internal.initialCalculation : undefined;
  const trackCalculation = Boolean(internal.attachPeer) && precedingCalculation !== undefined;
  let acknowledgedSequence = 0;
  let sentSequence = 0;
  let recovered = false;
  let disposed = false;
  let tail = Promise.resolve();
  const attachment = internal.attachPeer?.(peer.version());
  if (attachment) tail = attachment.catch((cause: unknown) => { fail(cause); });
  let offFailure = () => {};
  let rejectFailure!: (cause: WorkbookEditPeerFailedError) => void;
  const failed = new Promise<never>((_, reject) => { rejectFailure = reject; });
  void failed.catch(() => {});
  const outcomes = new WeakMap<WorkbookReplayOp,
    { result: WorkbookReplayReply['result'] } | { error: unknown }
  >();
  const applying = new WeakSet<WorkbookReplayOp>();
  let activeApplications = 0;
  const pending: { resolved: boolean; envelope?: Omit<WorkbookReplayEnvelope, 'sequence'> }[] = [];
  const drainWaiters: (() => void)[] = [];
  const acknowledgements = new Set<{
    sequence: number; resolve(): void; reject(error: unknown): void;
  }>();
  let dispatching = false;

  function resolveDrains(): void {
    for (const resolve of drainWaiters.splice(0)) resolve();
  }

  function fail(cause: unknown, notify = true): void {
    if (error) return;
    error = cause instanceof Error ? cause : new Error(String(cause));
    rejectFailure(new WorkbookEditPeerFailedError(error));
    for (const waiter of acknowledgements) waiter.reject(new WorkbookEditPeerFailedError(error));
    acknowledgements.clear();
    pending.length = 0;
    resolveDrains();
    offFailure();
    if (notify) {
      try { options.onError?.(error); } catch {}
    }
  }

  function synchronizeFailure(): void {
    if (session.failure) fail(session.failure);
  }

  function assertReady(): void {
    synchronizeFailure();
    if (error) throw new WorkbookEditPeerFailedError(error);
  }

  function enqueue(envelope: WorkbookReplayEnvelope): void {
    const sent = replay(envelope);
    tail = Promise.all([tail, sent]).then(([, reply]) => {
      if (reply.sequence !== envelope.sequence) {
        throw new Error(`Workbook replay acknowledgement mismatch at sequence ${envelope.sequence} (${envelope.op.method})`);
      }
      assertReady();
      acknowledgedSequence = reply.sequence;
      for (const waiter of acknowledgements) {
        if (waiter.sequence > acknowledgedSequence) continue;
        acknowledgements.delete(waiter);
        waiter.resolve();
      }
    }).catch((cause: unknown) => { fail(cause); });
  }

  function dispatch(): void {
    if (dispatching) return;
    dispatching = true;
    try {
      synchronizeFailure();
      while (!error && pending[0]?.resolved) {
        const envelope = pending.shift()?.envelope;
        if (!envelope) continue;
        sentSequence += 1;
        try {
          enqueue({ sequence: sentSequence, ...envelope });
        } catch (cause) { fail(cause); }
        synchronizeFailure();
      }
    } finally {
      dispatching = false;
      if (!pending.length) resolveDrains();
    }
  }

  function recoveryResult(result: WorkbookReplayReply['result']): WorkbookReplayReply['result'] {
    if (workbookReplayRefused(result)) {
      throw new WorkbookRecoveryRefusal(result);
    }
    return result;
  }

  function apply(op: WorkbookReplayOp, recovery: boolean): WorkbookReplayReply['result'] {
    const outcome = outcomes.get(op);
    if (outcome) {
      if ('error' in outcome) throw outcome.error;
      return recovery ? recoveryResult(outcome.result) : outcome.result;
    }
    if (applying.has(op)) throw new Error('Workbook queued operation is already applying');
    const previousCalculation = precedingCalculation;
    const calculation: WorkbookCalculationContext = op.calculation ?? {
      nowSerial: now() / 86_400_000 + 25_569,
      randSeed: seed(),
    };
    const slot: typeof pending[number] = { resolved: false };
    if (!recovery) pending.push(slot);
    applying.add(op);
    activeApplications += 1;
    let envelope: typeof slot.envelope;
    let snapshotFailure: unknown;
    let proposals: string | undefined;
    try {
      if (op.method === 'acceptProposal') proposals = JSON.stringify(peer.listProposals());
      try {
        envelope = structuredClone({ calculation, op });
      } catch (cause) { snapshotFailure = cause; }
      if (recovery) {
        if (!envelope) throw snapshotFailure;
        validateWorkbookReplayEnvelope({ sequence: 1, ...envelope });
      }
      peer.setCalculationContext(calculation);
      if (trackCalculation) precedingCalculation = calculation;
      const result = applyWorkbookReplayOp(peer, op);
      outcomes.set(op, { result });
      if (recovery) return recoveryResult(result);
      if (!workbookReplayRefused(result)) {
        if (envelope) slot.envelope = envelope;
        else fail(snapshotFailure);
      } else if (trackCalculation && precedingCalculation === calculation) {
        peer.setCalculationContext(previousCalculation ?? null);
        precedingCalculation = previousCalculation;
      }
      return result;
    } catch (cause) {
      if (!recovery && cause instanceof StaleProposalError && proposals !== undefined &&
        proposals !== JSON.stringify(peer.listProposals())) {
        if (envelope) {
          slot.envelope = {
            ...envelope, staleProposal: structuredClone({ cells: cause.cells, targets: cause.targets }),
          };
        } else fail(snapshotFailure);
      }
      if (!slot.envelope && trackCalculation && precedingCalculation === calculation) {
        peer.setCalculationContext(previousCalculation ?? null);
        precedingCalculation = previousCalculation;
      }
      if (!outcomes.has(op)) outcomes.set(op, { error: cause });
      throw cause;
    } finally {
      applying.delete(op);
      activeApplications -= 1;
      slot.resolved = true;
      if (!recovery) dispatch();
    }
  }

  function applyQueuedOp(op: WorkbookReplayOp): WorkbookReplayReply['result'] {
    assertReady();
    return apply(op, false);
  }

  function assertRecovery(): void {
    synchronizeFailure();
    if (!error) throw new Error('Recovery requires a failed workbook edit peer');
    if (disposed && internal?.attachPeer) throw new Error('Workbook edit peer was disposed');
    if (recovered) throw new Error('Workbook edit peer recovery was already saved');
    if (activeApplications) throw new Error('Workbook queued operation is still applying');
  }

  const mutators = {} as Pick<WorkbookHandle, WorkbookReplayMethod>;
  for (const method of Object.keys(WORKBOOK_REPLAY_MUTATORS) as WorkbookReplayMethod[]) {
    Object.defineProperty(mutators, method, { enumerable: true, value: (
      ...args: Parameters<WorkbookHandle[typeof method]>
    ) => applyQueuedOp({ method, args } as WorkbookReplayOp) });
  }

  internal.editPeerAttached = true;
  offFailure = session.onFailure(fail);
  synchronizeFailure();

  async function flush(): Promise<void> {
    do {
      assertReady();
      if (pending.length || dispatching) {
        await new Promise<void>((resolve) => { drainWaiters.push(resolve); });
        assertReady();
      }
      const watermark = sentSequence;
      await Promise.race([tail, failed]);
      assertReady();
      if (!pending.length && !dispatching && sentSequence === watermark &&
        acknowledgedSequence >= watermark) return;
    } while (true);
  }

  const edits: WorkbookEditPeer = {
    ...mutators,
    get state() { synchronizeFailure(); return error ? 'failed' : 'ready'; },
    get error() { synchronizeFailure(); return error; },
    get acknowledgedSequence() { return acknowledgedSequence; },
    get sentSequence() { return sentSequence; },
    flush,
    async save() {
      await flush();
      try {
        const bytes = await Promise.race([session.save(), failed]);
        assertReady();
        return buffer(bytes);
      } catch (cause) {
        fail(cause);
        throw cause;
      }
    },
    recoverySave() {
      assertRecovery();
      const bytes = buffer(peer.save());
      recovered = true;
      return { bytes, recovery: true };
    },
    dispose() {
      if (disposed) return;
      disposed = true;
      fail(new SessionFailure('disposed', 'Workbook edit peer was disposed'), false);
      offFailure();
      internal.editPeerAttached = false;
      void internal.detachPeer?.().catch(() => {});
    },
  };
  workbookEditPeerInternals.set(edits, {
    fail,
    whenAcknowledged() {
      try { assertReady(); } catch (cause) { return Promise.reject(cause); }
      if (acknowledgedSequence >= sentSequence) return Promise.resolve();
      return new Promise<void>((resolve, reject) => {
        acknowledgements.add({ sequence: sentSequence, resolve, reject });
      });
    },
    applyQueuedOp,
    applyRecoveryOp(op) {
      assertRecovery();
      return apply(op, true);
    },
  });
  return edits;
}
