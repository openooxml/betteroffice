import { SessionFailure } from '../../../../shared/office-session';
import type { WorkbookCalculationContext, WorkbookHandle } from '../wasm/loader';
import type { WorkbookSession } from './client';
import {
  applyWorkbookReplayOp,
  WORKBOOK_REPLAY_MUTATORS,
  workbookReplayRefused,
  workbookSessionInternals,
  type WorkbookReplayEnvelope,
  type WorkbookReplayMethod,
  type WorkbookReplayOp,
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
  const replay = internal.replay;
  const now = options.now ?? Date.now;
  const seed = options.randomSeed ?? randomSeed;
  let error: Error | undefined;
  let acknowledgedSequence = 0;
  let sentSequence = 0;
  let recovered = false;
  let disposed = false;
  let tail = Promise.resolve();
  let offFailure = () => {};
  const pending: { resolved: boolean; envelope?: Omit<WorkbookReplayEnvelope, 'sequence'> }[] = [];
  const drainWaiters: (() => void)[] = [];
  let dispatching = false;

  function resolveDrains(): void {
    for (const resolve of drainWaiters.splice(0)) resolve();
  }

  function fail(cause: unknown, notify = true): void {
    if (error) return;
    error = cause instanceof Error ? cause : new Error(String(cause));
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
    tail = Promise.all([tail, replay(envelope)]).then(([, reply]) => {
      if (reply.sequence !== envelope.sequence) {
        throw new Error(`Workbook replay acknowledgement mismatch at sequence ${envelope.sequence} (${envelope.op.method})`);
      }
      assertReady();
      acknowledgedSequence = reply.sequence;
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

  const mutators = {} as Pick<WorkbookHandle, WorkbookReplayMethod>;
  for (const method of Object.keys(WORKBOOK_REPLAY_MUTATORS) as WorkbookReplayMethod[]) {
    Object.defineProperty(mutators, method, { enumerable: true, value: (
      ...args: Parameters<WorkbookHandle[typeof method]>
    ) => {
      assertReady();
      const calculation: WorkbookCalculationContext = {
        nowSerial: now() / 86_400_000 + 25_569,
        randSeed: seed(),
      };
      const op = { method, args } as WorkbookReplayOp;
      const slot: typeof pending[number] = { resolved: false };
      pending.push(slot);
      try {
        let envelope: typeof slot.envelope;
        let snapshotFailure: unknown;
        try {
          envelope = structuredClone({ calculation, op });
        } catch (cause) { snapshotFailure = cause; }
        peer.setCalculationContext(calculation);
        const result = applyWorkbookReplayOp(peer, op);
        if (!workbookReplayRefused(result)) {
          if (envelope) slot.envelope = envelope;
          else fail(snapshotFailure);
        }
        return result;
      } finally {
        slot.resolved = true;
        dispatch();
      }
    } });
  }

  internal.editPeerAttached = true;
  offFailure = session.onFailure(fail);
  synchronizeFailure();

  async function flush(): Promise<void> {
    assertReady();
    if (pending.length || dispatching) {
      await new Promise<void>((resolve) => { drainWaiters.push(resolve); });
      assertReady();
    }
    await tail;
    assertReady();
  }

  return {
    ...mutators,
    get state() { synchronizeFailure(); return error ? 'failed' : 'ready'; },
    get error() { synchronizeFailure(); return error; },
    get acknowledgedSequence() { return acknowledgedSequence; },
    get sentSequence() { return sentSequence; },
    flush,
    async save() {
      await flush();
      try {
        const bytes = await session.save();
        assertReady();
        return buffer(bytes);
      } catch (cause) {
        fail(cause);
        throw cause;
      }
    },
    recoverySave() {
      synchronizeFailure();
      if (!error) throw new Error('Recovery save requires a failed workbook edit peer');
      if (recovered) throw new Error('Workbook edit peer recovery was already saved');
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
    },
  };
}
