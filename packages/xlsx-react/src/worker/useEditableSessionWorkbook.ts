import {
  createWorkbookEditPeer, failWorkbookEditPeer, hydratePeer, openWorkbookSession,
  type WorkbookHandle, type WorkbookSession, type WorkbookEditPeer,
} from '@betteroffice/xlsx';
import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react';
import type { XlsxEditorProps } from '../XlsxEditor';
import { XlsxCommandAdmissionError } from '../commands/createXlsxCommandStore';
import type { XlsxCommandStore } from '../commands/types';
import type { WorkerInputCoordinator } from '../commands/workerInputCoordinator';
import {
  createWorkerEditorApi, type WorkerEditorApiBridge, type WorkerEditorSessionAccess,
  type XlsxWorkerEditorApi, XlsxWorkerEditorCollaborationError,
} from './createWorkerEditorApi';

export const editableWorkbookSessionBackend = {
  open: openWorkbookSession,
  hydrate: hydratePeer,
  attach: createWorkbookEditPeer,
};

export type EditableSessionWorkbookProps = Omit<XlsxEditorProps, 'onReady' | 'collaboration'> & {
  onError?: (error: Error) => void;
  onReady?: (api: XlsxWorkerEditorApi) => void | (() => void);
};

export interface EditableWorkbookSessionOptions {
  changed(): void;
  onError(error: Error): void;
  onReady(): void | (() => void);
  isCurrent(): boolean;
}

function asError(value: unknown): Error {
  return value instanceof Error ? value : new Error(
    value && typeof value === 'object' && 'message' in value ? String(value.message) : String(value)
  );
}

export class EditableWorkbookSession implements WorkerEditorSessionAccess {
  alive = true;
  peer: WorkbookHandle | null = null;
  editPeer: WorkbookEditPeer | null = null;
  failure: Error | null = null;
  recovering = false;
  retiring = false;
  beforeRetire: (() => void) | null = null;
  private painted = false;
  private hydration: Promise<void> | null = null;
  private readonly hydrated: Promise<void>;
  private resolveHydrated!: () => void;
  private rejectHydrated!: (error: unknown) => void;
  private readonly hydrationWaiters = new Set<{ recovery: boolean; reject(error: unknown): void }>();
  private offFailure = () => {};
  private cleanup: void | (() => void) = undefined;
  private input: WorkerInputCoordinator | null = null;

  constructor(
    readonly session: WorkbookSession,
    readonly generation: number,
    private readonly options: EditableWorkbookSessionOptions
  ) {
    this.hydrated = new Promise((resolve, reject) => {
      this.resolveHydrated = resolve;
      this.rejectHydrated = reject;
    });
    void this.hydrated.catch(() => {});
    this.offFailure = session.onFailure((error) => this.fail(error));
    if (session.failure) this.fail(session.failure);
  }

  get current(): boolean { return this.alive && (this.retiring || this.options.isCurrent()); }

  get ready(): boolean {
    return this.current && (!this.failure || this.recovering) && this.peer !== null &&
      (this.editPeer?.state === 'ready' || this.recovering && this.editPeer !== null);
  }

  whenHydrated(): Promise<void> {
    if (!this.current) return Promise.reject(new XlsxCommandAdmissionError('document-replaced'));
    if (this.recovering && this.peer && this.editPeer) return Promise.resolve();
    if (this.failure) return Promise.reject(this.failure);
    return this.hydrated;
  }

  connectInput(input: WorkerInputCoordinator | null): void {
    if (!this.current || this.retiring || this.input === input) return;
    this.input = input;
    if (this.failure) input?.fail(this.failure);
  }

  firstPaint(): void {
    if (!this.current || this.failure || this.painted) return;
    this.painted = true;
    try {
      const cleanup = this.options.onReady();
      if (!this.current && typeof cleanup === 'function') cleanup();
      else this.cleanup = cleanup;
    } catch (error) { this.fail(error); }
    if (this.current && !this.failure) void this.requestHydration('first-paint').catch(() => {});
  }

  requestHydration(reason: string): Promise<void> {
    if (!this.current) return Promise.reject(new XlsxCommandAdmissionError('document-replaced'));
    if (this.failure && reason !== 'recovery') return Promise.reject(this.failure);
    if (this.editPeer) return Promise.resolve();
    if (this.hydration) return this.waitForHydration(this.hydration, reason);
    const pending = (async () => {
      let peer: WorkbookHandle | null = null;
      try {
        peer = await editableWorkbookSessionBackend.hydrate(this.session);
        if (!this.current) throw new XlsxCommandAdmissionError('document-replaced');
        const edits = editableWorkbookSessionBackend.attach({
          session: this.session, peer, onError: (error) => this.fail(error),
        });
        if (!this.current) {
          edits.dispose();
          throw new XlsxCommandAdmissionError('document-replaced');
        }
        this.peer = peer;
        this.editPeer = edits;
        if (this.failure) failWorkbookEditPeer(edits, this.failure);
        peer = null;
        if (!this.failure) this.resolveHydrated();
        this.options.changed();
      } catch (error) {
        peer?.dispose();
        if (this.current) this.fail(error);
        throw error;
      }
    })();
    this.hydration = pending;
    void pending.catch(() => {
      if (this.hydration === pending) this.hydration = null;
    });
    return this.waitForHydration(pending, reason);
  }

  fail(value: unknown): void {
    if (!this.current || this.failure) return;
    this.failure = asError(value);
    if (this.editPeer) failWorkbookEditPeer(this.editPeer, this.failure);
    this.rejectHydrated(this.failure);
    for (const waiter of this.hydrationWaiters) {
      if (waiter.recovery) continue;
      this.hydrationWaiters.delete(waiter);
      waiter.reject(this.failure);
    }
    this.input?.fail(this.failure);
    this.options.changed();
    try { this.options.onError(this.failure); } catch {}
  }

  reportRefusal(value: unknown): void {
    try { this.options.onError(asError(value)); } catch {}
  }

  dispose(): void {
    if (!this.alive || this.retiring) return;
    this.retiring = true;
    try { this.beforeRetire?.(); } catch (error) { this.reportRefusal(error); }
    if (this.input?.pending || this.input?.draft) {
      void this.input.drain().catch((error) => {
        this.reportRefusal(error);
      }).finally(() => this.destroy());
    } else this.destroy();
  }

  private destroy(): void {
    if (!this.alive) return;
    this.alive = false;
    const error = new XlsxCommandAdmissionError('document-replaced');
    this.rejectHydrated(error);
    for (const waiter of this.hydrationWaiters) waiter.reject(error);
    this.hydrationWaiters.clear();
    this.offFailure();
    this.input?.reset();
    this.input = null;
    try { if (typeof this.cleanup === 'function') this.cleanup(); }
    catch (error) {
      try { this.options.onError(asError(error)); } catch {}
    }
    finally {
      try { this.editPeer?.dispose(); }
      finally {
        try { this.peer?.dispose(); }
        finally {
          this.editPeer = null;
          this.peer = null;
          void this.session.dispose().catch(() => {});
        }
      }
    }
  }

  private waitForHydration(pending: Promise<void>, reason: string): Promise<void> {
    if (!this.current) return Promise.reject(new XlsxCommandAdmissionError('document-replaced'));
    if (this.failure && reason !== 'recovery') return Promise.reject(this.failure);
    return new Promise<void>((resolve, reject) => {
      const waiter = { recovery: reason === 'recovery', reject };
      this.hydrationWaiters.add(waiter);
      void pending.then(() => {
        this.hydrationWaiters.delete(waiter);
        resolve();
      }, (error) => {
        this.hydrationWaiters.delete(waiter);
        reject(error);
      });
    });
  }
}

export function useEditableSessionWorkbook(
  props: EditableSessionWorkbookProps, commands: XlsxCommandStore, bridge: WorkerEditorApiBridge
): {
  run: EditableWorkbookSession | null;
  error: Error | null;
  loading: boolean;
  reportError(value: unknown): void;
} {
  const latest = useRef({ props, bridge });
  latest.current = { props, bridge };
  const generation = useRef(0);
  const [run, setRun] = useState<EditableWorkbookSession | null>(null);
  const [, setRevision] = useState(0);
  const [error, setError] = useState<Error | null>(null);
  const [loading, setLoading] = useState(false);
  const collaboration = (props as EditableSessionWorkbookProps & Pick<XlsxEditorProps, 'collaboration'>).collaboration;
  const reportError = useCallback((value: unknown) => {
    const error = asError(value);
    setError(error);
    try { latest.current.props.onError?.(error); } catch {}
  }, []);

  useEffect(() => {
    const token = ++generation.current;
    const controller = new AbortController();
    let disposed = false;
    let opened: EditableWorkbookSession | undefined;
    const current = () => !disposed && token === generation.current && latest.current.props.file === props.file;
    const changed = () => { if (current()) setRevision((revision) => revision + 1); };
    setRun(null);
    setError(null);
    setLoading(Boolean(props.file));
    if (collaboration) {
      reportError(new XlsxWorkerEditorCollaborationError());
      setLoading(false);
    } else if (props.file) void (async () => {
      try {
        const session = await editableWorkbookSessionBackend.open(props.file!, {
          signal: controller.signal, retainPeerHydration: true,
        });
        if (!current()) { void session.dispose().catch(() => {}); return; }
        opened = new EditableWorkbookSession(session, token, {
          changed, isCurrent: current,
          onError: (error) => {
            if (current()) reportError(error);
            else { try { latest.current.props.onError?.(error); } catch {} }
          },
          onReady: () => current() ? latest.current.props.onReady?.(api) : undefined,
        });
        const api = createWorkerEditorApi(opened, commands, () => latest.current.bridge);
        setRun(opened);
        setLoading(false);
      } catch (error) {
        if (!current()) return;
        if (opened) opened.fail(error);
        else reportError(error);
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
  }, [props.file, collaboration, commands, reportError]);

  useLayoutEffect(() => {
    if (run?.current) run.connectInput(bridge.coordinator());
  });
  useLayoutEffect(() => () => run?.dispose(), [run, props.file, collaboration]);

  return { run: run?.current ? run : null, error, loading, reportError };
}
