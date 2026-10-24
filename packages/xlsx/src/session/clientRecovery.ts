import { createSessionClient } from '../../../../shared/office-session/client';
import type { Checkpoint } from '../../../../shared/office-session/commands/types';
import { createWorkerTransport, type SessionTransport } from '../../../../shared/office-session/transport';
import { SessionFailure } from '../../../../shared/office-session/types';
import type { OpenWorkbookSessionOptions, WorkbookSession } from './client';
import type { CommandStatus, WorkbookAck, WorkbookCommand, WorkbookCommandContext, WorkbookSharedCommand } from './commandProtocol';
import { localNowSerial } from './calculationClock';
import { WORKBOOK_SESSION_METHODS, type WorkbookCheckpoint, type WorkbookSessionMethods } from './methods';

const CHECKPOINT_MARK = 256;
const RETAINED_LIMIT = 4096;

/** @experimental */
export class WorkbookRecoveryReplayError extends Error {
  readonly code = 'workbook-recovery-refused';
  constructor(readonly seq: number, readonly reason: string) {
    super(`Workbook recovery refused sequence ${seq}: ${reason}`);
    this.name = 'WorkbookRecoveryReplayError';
  }
}

/** @experimental */
export interface WorkbookSessionRecovery {
  readonly throughSeq: number;
  readonly retainedCommands: number;
  readonly checkpointIdleMs: number;
  capture(allowed: () => boolean): Promise<boolean>;
  save(assertCurrent: () => void, pending?: readonly Pick<WorkbookCommand, 'kind' | 'payload'>[]): Promise<Uint8Array>;
}

interface RecoverySession {
  replay(command: WorkbookSharedCommand, pending?: boolean): Promise<CommandStatus>;
  save(): Promise<Uint8Array>;
  dispose(): Promise<void>;
}

interface RecoveryEntry {
  command: WorkbookSharedCommand;
  outcome?: CommandStatus;
  identifier?: string;
  pending?: boolean;
}

/** @experimental */
export const workbookSessionRecovery = new WeakMap<WorkbookSession, WorkbookSessionRecovery>();

export const workbookRecoveryBackend = {
  open(bytes: Uint8Array, options: OpenWorkbookSessionOptions, checkpoint?: Checkpoint, version?: string): Promise<RecoverySession> {
    if (options.collaborative) return Promise.reject(new Error('Recovery is unavailable for collaborative workbooks'));
    const transport = createWorkerTransport(options.worker ? options.worker() :
      new Worker(new URL('./xlsxSessionWorker.mjs', import.meta.url), { type: 'module' }));
    return createWorkbookRecoverySession(bytes, options, transport, checkpoint, version);
  },
};

export async function createWorkbookRecoverySession(
  bytes: Uint8Array, options: OpenWorkbookSessionOptions, transport: SessionTransport, checkpoint?: Checkpoint, version?: string
): Promise<RecoverySession> {
  const client = createSessionClient<WorkbookSessionMethods, {}>(transport, { methods: WORKBOOK_SESSION_METHODS });
  try {
    if (options.collaborative) throw new Error('Recovery is unavailable for collaborative workbooks');
    const wasm = options.wasm instanceof ArrayBuffer ? options.wasm.slice(0) : options.wasm;
    const opened = checkpoint
      ? await client.call.openRecovery({ ...checkpoint, state: checkpoint.state.slice() }, { wasm })
      : await client.call.open(bytes.slice().buffer, { collaborative: options.collaborative,
        clientId: options.clientId, calculation: options.calculation, wasm,
        initialViewport: options.initialViewport, includeEditState: true, recoveryVersion: version });
    const context: WorkbookCommandContext | undefined = opened.commandContext;
    if (!context) throw new Error('Workbook recovery omitted command context');
    let engineVersion = context.basis.version;
    let revision = checkpoint?.version ?? 0;
    return {
      async replay(command, pending = false) {
        const compatibility = command.payload.compatibility;
        const reply = await client.call.edit({ ...command, session: context.session, incarnation: 1, replay: true,
          ...(pending ? { basis: { epoch: context.basis.epoch, version: revision } } : {}),
          payload: { ...command.payload, ...(compatibility ? { compatibility: { ...compatibility,
            session: context.session, incarnation: context.incarnation,
            ...(pending ? { basis: { epoch: context.basis.epoch, version: engineVersion } } : {}) } } : {}) } });
        engineVersion = reply.engineVersion;
        revision = reply.version;
        const outcome = reply.outcomes.find((entry) => entry.seq === command.seq);
        if (!outcome || reply.throughSeq < command.seq) throw new WorkbookRecoveryReplayError(command.seq, 'missing-outcome');
        return outcome.workbook ?? { ...outcome, status: outcome.status === 'duplicate' ? 'noop' : outcome.status };
      },
      save: async () => new Uint8Array(await client.call.save()),
      dispose: () => client.dispose(),
    };
  } catch (error) {
    await client.dispose();
    throw error;
  }
}

function identifier(status: CommandStatus): string | undefined {
  const result = status.result;
  if (!result || typeof result !== 'object') return undefined;
  if ('formatToken' in result && typeof result.formatToken === 'string') return result.formatToken;
  if ('id' in result && typeof result.id === 'string') return result.id;
  return undefined;
}

export function createWorkbookSessionRecovery(
  bytes: Uint8Array, options: OpenWorkbookSessionOptions,
  checkpoint: (allowed: () => boolean) => Promise<WorkbookCheckpoint | undefined>, epoch: number, liveCalculation = false,
  version?: string, fail?: (error: unknown) => void
) {
  let source: Uint8Array | undefined = bytes;
  const settings = { ...options, signal: undefined, retainPeerHydration: false,
    calculation: options.calculation ? { ...options.calculation } : undefined,
    initialViewport: options.initialViewport ? { ...options.initialViewport } : undefined,
    wasm: options.wasm instanceof ArrayBuffer ? options.wasm.slice(0) : options.wasm };
  const log = new Map<number, RecoveryEntry>();
  let captured: WorkbookCheckpoint | undefined;
  let capture: Promise<boolean> | undefined;
  let saving: Promise<Uint8Array> | undefined;
  let highSeq = 0;
  let forcing = false;
  let retryStep = 0;
  let captureMark = CHECKPOINT_MARK;
  const active = () => { if (!source) throw new SessionFailure('disposed', 'Workbook recovery was disposed'); };
  const normalizePayload = (payload: WorkbookCommand['payload']): WorkbookCommand['payload'] => {
    if (!liveCalculation || !('args' in payload) || payload.calculation) return payload;
    return { ...payload, calculation: { nowSerial: localNowSerial(Date.now()),
      randSeed: globalThis.crypto.getRandomValues(new Uint32Array(1))[0] >>> 0 } };
  };
  const recovery: WorkbookSessionRecovery = {
    get throughSeq() { return captured?.throughSeq ?? 0; },
    get retainedCommands() { return log.size; },
    get checkpointIdleMs() { return Math.max(150, 2 * (captured?.encodeMs ?? 0)); },
    async capture(allowed) {
      active();
      if (capture) return capture;
      if (retryStep && log.size < captureMark) return false;
      capture = (async () => {
        if (!allowed()) return false;
        const value = await checkpoint(allowed);
        active();
        if (!value) return false;
        if (!(value.state instanceof Uint8Array) || value.epoch !== epoch ||
          !Number.isSafeInteger(value.throughSeq) || value.throughSeq < (captured?.throughSeq ?? 0) ||
          value.throughSeq > highSeq || !Number.isSafeInteger(value.version) || value.version < (captured?.version ?? 0) ||
          value.encodeMs !== undefined && (!Number.isFinite(value.encodeMs) || value.encodeMs < 0)) {
          throw new WorkbookRecoveryReplayError(value.throughSeq, 'invalid-checkpoint');
        }
        captured = { ...value, state: value.state.slice() };
        for (const seq of log.keys()) if (seq <= value.throughSeq) log.delete(seq);
        retryStep = 0;
        captureMark = CHECKPOINT_MARK;
        return true;
      })().catch((error: unknown) => {
        if (!source || error instanceof SessionFailure || error instanceof WorkbookRecoveryReplayError) {
          if (source) fail?.(error);
          throw error;
        }
        const message = error instanceof Error ? error.message : String(error);
        if (log.size >= RETAINED_LIMIT) {
          const refused = Object.assign(new WorkbookRecoveryReplayError(highSeq, 'log-full'), { cause: error });
          fail?.(refused);
          throw refused;
        }
        retryStep = Math.min(1024, retryStep * 2 || 64);
        captureMark = Math.min(RETAINED_LIMIT, log.size + retryStep);
        console.warn(`xlsx recovery: checkpoint failed, retaining ${log.size} commands until a later capture: ${message}`);
        return false;
      }).finally(() => { capture = undefined; });
      return capture;
    },
    async save(assertCurrent, pending = []) {
      active();
      assertCurrent();
      if (saving) return saving;
      saving = (async () => {
        const entries = [...log.values()].sort((left, right) => left.command.seq - right.command.seq);
        for (const [index, input] of pending.entries()) {
          const seq = highSeq + index + 1;
          const payload = normalizePayload(structuredClone(input.payload));
          entries.push({ pending: true, command: { session: '', incarnation: 1, seq, basis: { epoch, version: 0 },
            kind: input.kind, payload: { payload, compatibility: {
              session: '', incarnation: '', seq, basis: { epoch, version: '' },
              kind: input.kind, payload,
            } as WorkbookCommand } } });
        }
        const opened = await workbookRecoveryBackend.open(source!, settings, captured, version);
        try {
          assertCurrent(); active();
          const identifiers = new Map<string, string>();
          for (const entry of entries) {
            const command = structuredClone(entry.command);
            const payload = command.payload.payload;
            if ('formatToken' in payload) payload.formatToken = identifiers.get(payload.formatToken) ?? payload.formatToken;
            if ((command.kind === 'acceptProposal' || command.kind === 'rejectProposal') && 'args' in payload) {
              const args = payload.args as [string];
              args[0] = identifiers.get(args[0]) ?? args[0];
            }
            if (command.payload.compatibility) command.payload.compatibility.payload = payload;
            const status = await opened.replay(command, entry.pending);
            assertCurrent(); active();
            const refused = status.status === 'refused' || status.status === 'cancelled';
            const expected = entry.outcome;
            if (refused && !expected || expected && (expected.status !== status.status ||
              expected.reason !== status.reason || expected.committed !== status.committed ||
              JSON.stringify(expected.affectedSeqs) !== JSON.stringify(status.affectedSeqs) ||
              refused && JSON.stringify(expected.result) !== JSON.stringify(status.result))) {
              throw new WorkbookRecoveryReplayError(command.seq, status.reason ?? 'outcome-mismatch');
            }
            const replacement = identifier(status);
            if (entry.identifier && replacement) identifiers.set(entry.identifier, replacement);
          }
          const result = await opened.save();
          assertCurrent(); active();
          return result;
        } finally { await opened.dispose(); }
      })().finally(() => { saving = undefined; });
      return saving;
    },
  };
  const checkpointAboveMark = () => {
    if (forcing || !source || log.size < captureMark) return;
    forcing = true;
    void (capture ?? Promise.resolve(true)).then(() => {
      forcing = false;
      if (source && log.size >= captureMark) return recovery.capture(() => source !== undefined);
    }, () => { forcing = false; }).catch(() => {});
  };
  return {
    recovery,
    normalizePayload,
    admit(input: WorkbookSharedCommand) {
      active();
      if (input.seq <= highSeq) return;
      if (input.seq !== highSeq + 1) throw new WorkbookRecoveryReplayError(input.seq, 'sequence-gap');
      const command = structuredClone(input);
      command.payload.frame = undefined;
      if (command.payload.compatibility) command.payload.compatibility.frame = undefined;
      log.set(command.seq, { command });
      highSeq = command.seq;
      checkpointAboveMark();
    },
    acknowledge(ack: WorkbookAck) {
      for (const status of ack.statuses) {
        const entry = log.get(status.seq);
        if (!entry) continue;
        entry.outcome = structuredClone(status);
        entry.identifier = identifier(status);
      }
    },
    dispose() { source = undefined; captured = undefined; log.clear(); },
  };
}
