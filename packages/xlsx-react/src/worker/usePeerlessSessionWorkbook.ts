import { workbookSessionRecovery, WorkbookRecoveryReplayError } from '@betteroffice/xlsx';
import type { OpenWorkbookSessionOptions, WorkbookSession } from '@betteroffice/xlsx';
import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react';
import type { XlsxEditorCollaborationOptions } from '../XlsxEditor';
import { XlsxCommandAdmissionError } from '../commands/createXlsxCommandStore';
import type { XlsxCommandStore } from '../commands/types';
import type { PeerlessWorkerInputCoordinator } from '../commands/peerlessWorkerInputCoordinator';
import { createAsyncEditorApi, type XlsxAsyncEditorApi } from './createAsyncEditorApi';
import type { WorkerEditorApiBridge } from './createWorkerEditorApi';
import type { WorkerDraftController } from './WorkerDraftController';
import { asError, editableWorkbookSessionBackend, XlsxWorkerSessionFailedError } from './useEditableSessionWorkbook';
import type { EditableSessionWorkbookProps, EditableWorkbookSessionOptions } from './useEditableSessionWorkbook';

interface SessionWorkbookOptions {
  backend: 'legacy' | 'peerless';
  initialViewport?: () => OpenWorkbookSessionOptions['initialViewport'];
}

interface SessionWorkbookResult<R> {
  run: R | null;
  error: Error | null;
  loading: boolean;
  reportError(value: unknown): void;
}

interface CheckpointScheduler {
  now(): number;
  schedule(callback: () => void, delay: number): () => void;
}

export class PeerlessWorkbookSession {
  alive = true;
  retiring = false;
  failure: XlsxWorkerSessionFailedError | null = null;
  beforeRetire: (() => void) | null = null;
  readonly session: WorkbookSession;
  private painted = false;
  private cleanup: void | (() => void) = undefined;
  private readonly listeners = new Set<() => void>();
  private readonly pending = new Set<{ reject(error: unknown): void }>();
  private readonly readiness: Promise<void>;
  private resolveReady!: () => void;
  private rejectReady!: (error: unknown) => void;
  private offFailure = () => {};
  private offAck = () => {};
  private offDraft = () => {};
  private offInput = () => {};
  private recoveryInput: (Pick<PeerlessWorkerInputCoordinator, 'draft' | 'pending'> &
    Partial<Pick<PeerlessWorkerInputCoordinator, 'recoveryCommands' | 'fail'>>) | null = null;
  private recoveryDrafts: Pick<WorkerDraftController, 'compositionId' | 'subscribe'> | null = null;
  private checkpointTimer: (() => void) | undefined;
  private readonly checkpointScheduler: CheckpointScheduler;
  private checkpointRevision = 0;
  private checkpointBusy = false;
  private disposal: Promise<void> = Promise.resolve();
  private readonly disposeSession = () => this.dispose();

  constructor(
    private readonly workerSession: WorkbookSession,
    readonly generation: number,
    private readonly options: EditableWorkbookSessionOptions & { checkpointScheduler?: CheckpointScheduler }
  ) {
    this.checkpointScheduler = options.checkpointScheduler ?? {
      now: () => performance.now(),
      schedule(callback, delay) {
        const timer = setTimeout(callback, delay);
        return () => clearTimeout(timer);
      },
    };
    this.readiness = new Promise((resolve, reject) => {
      this.resolveReady = resolve;
      this.rejectReady = reject;
    });
    void this.readiness.catch(() => {});
    const methods = (target: object) => {
      const wrappers = new WeakMap<(...args: unknown[]) => Promise<unknown>, (...args: unknown[]) => Promise<unknown>>();
      return (key: PropertyKey) => {
        const method = Reflect.get(target, key, target);
        if (typeof method !== 'function') return method;
        let wrapper = wrappers.get(method);
        if (!wrapper) {
          wrapper = (...args: unknown[]) => this.track(() => Reflect.apply(method, target, args),
            key === 'submit' || key === 'command' || key === 'applyEdits');
          wrappers.set(method, wrapper);
        }
        return wrapper;
      };
    };
    const callMethod = methods(workerSession.call);
    const sessionMethod = methods(workerSession);
    const calls = new Proxy(workerSession.call, { get: (_, key) => callMethod(key) });
    this.session = new Proxy(workerSession, {
      get: (target, key) => {
        if (key === 'call') return calls;
        if (key === 'dispose') return this.disposeSession;
        if (key === 'submit' || key === 'save' || key === 'whenPresented') return sessionMethod(key);
        return Reflect.get(target, key, target);
      },
    });
    this.offFailure = workerSession.onFailure((error) => this.fail(error));
    this.offAck = workerSession.onAck(() => { this.notify(); this.scheduleCheckpoint(true); });
    if (workerSession.failure) this.fail(workerSession.failure);
  }

  get current(): boolean { return this.alive && this.options.isCurrent(); }
  get commandsAccepted(): boolean {
    const identity = this.workerSession.commandIdentity;
    return this.current && !this.retiring && !this.failure && this.workerSession.state.stage === 'ready' &&
      Boolean(identity.session && identity.incarnation);
  }
  get ready(): boolean { return this.commandsAccepted && this.painted; }
  get readOnly(): boolean { return !this.ready || this.workerSession.editState.readOnly === true; }

  subscribe(listener: () => void): () => void {
    this.listeners.add(listener);
    return () => { this.listeners.delete(listener); };
  }

  whenReady(): Promise<void> {
    if (!this.current || this.retiring) return Promise.reject(new XlsxCommandAdmissionError('document-replaced'));
    if (this.failure) return Promise.reject(this.failure);
    return this.readiness;
  }

  connectRecoveryInput(
    input: Pick<PeerlessWorkerInputCoordinator, 'draft' | 'pending'> & Partial<Pick<PeerlessWorkerInputCoordinator, 'subscribe' | 'recoveryCommands' | 'fail'>>,
    drafts: Pick<WorkerDraftController, 'compositionId' | 'subscribe'>
  ): void {
    this.offDraft();
    this.offInput();
    this.recoveryInput = input;
    if (this.failure) input.fail?.(this.failure);
    this.recoveryDrafts = drafts;
    this.offDraft = drafts.subscribe(() => this.scheduleCheckpoint());
    this.offInput = input.subscribe?.(() => this.scheduleCheckpoint()) ?? (() => {});
    this.scheduleCheckpoint();
  }

  recoverySave(): Promise<{ bytes: Uint8Array; recovery: true }> {
    const assertCurrent = () => {
      if (!this.current || this.retiring) throw new XlsxCommandAdmissionError('document-replaced');
    };
    return this.track(async () => {
      assertCurrent();
      if (this.workerSession.replica) throw new Error('Recovery is unavailable for collaborative workbooks');
      if (!this.failure) throw new Error('Recovery requires a failed workbook edit peer');
      try {
        const recovery = workbookSessionRecovery.get(this.workerSession);
        if (!recovery) throw new Error('Workbook recovery state is unavailable');
        const bytes = await recovery.save(assertCurrent, this.recoveryInput?.recoveryCommands);
        assertCurrent();
        return { bytes, recovery: true as const };
      } catch (error) {
        assertCurrent();
        if (error instanceof XlsxCommandAdmissionError) throw error;
        this.failure = error instanceof XlsxWorkerSessionFailedError ? error : new XlsxWorkerSessionFailedError(error);
        this.changed();
        try { this.options.onError(this.failure); } catch {}
        throw this.failure;
      }
    }, false, true);
  }

  firstPaint(): void {
    if (!this.commandsAccepted || this.painted) return;
    this.painted = true;
    try {
      const cleanup = this.options.onReady();
      if (!this.current && typeof cleanup === 'function') cleanup();
      else this.cleanup = cleanup;
    } catch (error) { this.fail(error); }
    if (this.ready) this.resolveReady();
    this.changed();
    this.scheduleCheckpoint();
  }

  fail(cause: unknown): void {
    if (!this.current || this.failure) return;
    this.failure = cause instanceof XlsxWorkerSessionFailedError ? cause : new XlsxWorkerSessionFailedError(cause);
    this.checkpointTimer?.();
    this.checkpointTimer = undefined;
    this.checkpointRevision += 1;
    this.recoveryInput?.fail?.(this.failure);
    this.rejectReady(this.failure);
    this.rejectPending(this.failure);
    this.changed();
    try { this.options.onError(this.failure); } catch {}
  }

  dispose(): Promise<void> {
    if (!this.alive || this.retiring) return this.disposal;
    this.retiring = true;
    try { this.beforeRetire?.(); }
    catch (error) { try { this.options.onError(asError(error)); } catch {} }
    this.alive = false;
    const error = new XlsxCommandAdmissionError('document-replaced');
    this.rejectReady(error);
    this.rejectPending(error);
    this.offFailure();
    this.offAck();
    this.offDraft();
    this.offInput();
    this.checkpointTimer?.();
    this.checkpointTimer = undefined;
    this.checkpointRevision += 1;
    this.recoveryInput = null;
    this.recoveryDrafts = null;
    this.changed();
    this.listeners.clear();
    try { if (typeof this.cleanup === 'function') this.cleanup(); }
    catch (error) { try { this.options.onError(asError(error)); } catch {} }
    finally {
      this.disposal = this.workerSession.dispose();
      void this.disposal.catch(() => {});
    }
    return this.disposal;
  }

  private changed(): void {
    this.options.changed();
    this.notify();
  }

  private notify(): void {
    for (const listener of this.listeners) { try { listener(); } catch {} }
  }

  private rejectPending(error: unknown): void {
    for (const request of this.pending) request.reject(error);
    this.pending.clear();
  }

  private scheduleCheckpoint(acknowledged = false): void {
    const revision = ++this.checkpointRevision;
    this.checkpointTimer?.();
    this.checkpointTimer = undefined;
    const recovery = workbookSessionRecovery.get(this.workerSession);
    const highWater = (recovery?.retainedCommands ?? 0) >= 192;
    const allowed = () => this.ready && Boolean(this.recoveryInput) &&
      !this.recoveryInput?.draft && !this.recoveryDrafts?.compositionId &&
      (highWater || revision === this.checkpointRevision && !this.recoveryInput?.pending);
    if (!recovery || recovery.retainedCommands < 64 || this.checkpointBusy || !allowed()) return;
    const capture = () => {
      this.checkpointTimer = undefined;
      if (!allowed()) return;
      this.checkpointBusy = true;
      const throughSeq = recovery.throughSeq;
      void recovery.capture(allowed).then((captured) => {
        this.checkpointBusy = false;
        if (captured && recovery.throughSeq > throughSeq || revision !== this.checkpointRevision) this.scheduleCheckpoint();
      }, (error) => {
        this.checkpointBusy = false;
        if (this.current && !this.retiring) this.fail(error);
      });
    };
    if (highWater) {
      if (acknowledged || !this.recoveryInput?.pending) capture();
      return;
    }
    const started = this.checkpointScheduler.now();
    const idleMs = recovery.checkpointIdleMs ?? 150;
    const idle = () => {
      this.checkpointTimer = undefined;
      if (!allowed()) return;
      const remaining = idleMs - (this.checkpointScheduler.now() - started);
      if (remaining > 0) this.checkpointTimer = this.checkpointScheduler.schedule(idle, remaining);
      else capture();
    };
    this.checkpointTimer = this.checkpointScheduler.schedule(idle, idleMs);
  }

  private track<T>(operation: () => Promise<T>, requiresReady = false, recovery = false): Promise<T> {
    if (!this.current || this.retiring) return Promise.reject(new XlsxCommandAdmissionError('document-replaced'));
    if (this.failure && !recovery) return Promise.reject(this.failure);
    if (requiresReady && !this.ready) return Promise.reject(new XlsxCommandAdmissionError('editor-unavailable'));
    return new Promise<T>((resolve, reject) => {
      const request = { reject };
      this.pending.add(request);
      try {
        void operation().then(resolve, (error) => {
          if (error instanceof WorkbookRecoveryReplayError) this.fail(error);
          reject(recovery ? error : this.failure ?? error);
        }).finally(() => this.pending.delete(request));
      } catch (error) {
        this.pending.delete(request);
        if (error instanceof WorkbookRecoveryReplayError) this.fail(error);
        reject(this.failure ?? error);
      }
    });
  }
}

export type PeerlessSessionWorkbookProps = Omit<EditableSessionWorkbookProps, 'onReady'> & {
  collaboration?: XlsxEditorCollaborationOptions;
  onReady?: (api: XlsxAsyncEditorApi) => void | (() => void);
};

export function usePeerlessSessionWorkbook(
  props: PeerlessSessionWorkbookProps, commands: XlsxCommandStore,
  bridge: WorkerEditorApiBridge, options: SessionWorkbookOptions
): SessionWorkbookResult<PeerlessWorkbookSession> {
  const latest = useRef({ props, bridge, options });
  latest.current = { props, bridge, options };
  const generation = useRef(0);
  const [run, setRun] = useState<PeerlessWorkbookSession | null>(null);
  const [, setRevision] = useState(0);
  const [error, setError] = useState<Error | null>(null);
  const [loading, setLoading] = useState(false);
  const collaboration = props.collaboration;
  const collaborationKey = collaboration !== undefined;
  const collaborationClientId = collaboration?.clientId;
  const collaborationInitialUpdate = collaboration?.initialUpdate;
  const reportError = useCallback((value: unknown) => {
    const error = asError(value);
    setError(error);
    try { latest.current.props.onError?.(error); } catch {}
  }, []);

  useEffect(() => {
    const token = ++generation.current;
    const controller = new AbortController();
    let disposed = false;
    let opened: PeerlessWorkbookSession | undefined;
    const current = () => !disposed && token === generation.current && latest.current.props.file === props.file;
    const changed = () => { if (current()) setRevision((revision) => revision + 1); };
    setRun(null);
    setError(null);
    setLoading(Boolean(props.file));
    if (props.file) void (async () => {
      try {
        const session = await editableWorkbookSessionBackend.open(props.file!, {
          signal: controller.signal,
          includeEditState: true,
          ...(collaboration ? { collaborative: true,
            clientId: collaborationClientId, initialUpdate: collaborationInitialUpdate?.slice() } : {}),
          ...(latest.current.options.initialViewport
            ? { initialViewport: latest.current.options.initialViewport() } : {}),
        });
        if (!current()) { void session.dispose().catch(() => {}); return; }
        const lifecycle = {
          changed, isCurrent: current,
          onError: (error: Error) => {
            if (current()) reportError(error);
            else { try { latest.current.props.onError?.(error); } catch {} }
          },
        };
        const peerless: PeerlessWorkbookSession = new PeerlessWorkbookSession(session, token, {
          ...lifecycle, onReady: () => current() ? latest.current.props.onReady?.(api) : undefined,
        });
        const api = createAsyncEditorApi(peerless, commands, () => latest.current.bridge);
        opened = peerless;
        setRun(opened);
        setLoading(false);
      } catch (error) {
        if (!current()) return;
        if (opened) opened.fail(error);
        else reportError(new XlsxWorkerSessionFailedError(error));
      } finally {
        if (current()) setLoading(false);
      }
    })();
    return () => {
      disposed = true;
      generation.current += 1;
      controller.abort();
      opened?.dispose();
    };
  }, [props.file, collaborationKey, collaborationClientId, collaborationInitialUpdate, commands, reportError]);

  useLayoutEffect(() => () => { run?.dispose(); }, [run, props.file, collaborationKey, collaborationClientId, collaborationInitialUpdate]);

  const onReplica = collaboration?.onReplica;
  const replicaReady = run instanceof PeerlessWorkbookSession && run.ready;
  useEffect(() => {
    if (!replicaReady || !(run instanceof PeerlessWorkbookSession) || !onReplica) return;
    onReplica(run.session.replica ?? null);
    return () => onReplica(null);
  }, [run, replicaReady, onReplica]);

  return { run: run?.current ? run : null, error, loading, reportError };
}
