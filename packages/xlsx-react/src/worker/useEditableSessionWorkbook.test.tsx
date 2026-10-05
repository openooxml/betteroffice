import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, describe, expect, mock, spyOn, test } from 'bun:test';
import * as xlsx from '@betteroffice/xlsx';
import type { WorkbookHandle, WorkbookSession } from '@betteroffice/xlsx';
import { StrictMode } from 'react';
import { SessionFailure } from '../../../../shared/office-session';
import type { WorkbookEditPeer } from '../../../xlsx/src/session/editPeer';
import { createXlsxCommandController, XlsxCommandAdmissionError } from '../commands/createXlsxCommandStore';
import { createWorkerInputCoordinator } from '../commands/workerInputCoordinator';
import type { WorkerEditorApiBridge, XlsxWorkerEditorApi } from './createWorkerEditorApi';
import {
  EditableWorkbookSession, editableWorkbookSessionBackend, useEditableSessionWorkbook,
  type EditableSessionWorkbookProps, type EditableWorkbookSessionOptions,
} from './useEditableSessionWorkbook';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook, waitFor } = await import('@testing-library/react');
const restorers: (() => void)[] = [];
const file = new Uint8Array([1, 2, 3]);

function deferred<T>() {
  let resolve!: (value: T | PromiseLike<T>) => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<T>((done, failed) => { resolve = done; reject = failed; });
  return { promise, resolve, reject };
}

function session(log: string[] = []) {
  let listener: Parameters<WorkbookSession['onFailure']>[0] | undefined;
  let failure: SessionFailure | undefined;
  const value: WorkbookSession = {
    state: { format: 'xlsx', stage: 'ready', version: 0, dirty: false, sheets: [], activeSheet: 0 },
    call: {} as WorkbookSession['call'],
    save: mock(async () => new Uint8Array([9])), on: () => () => {},
    onFailure: (next) => {
      listener = next;
      return () => { log.push('off'); listener = undefined; };
    },
    get failure() { return failure; },
    dispose: mock(async () => { log.push('session'); }),
  };
  return { value, fail: (error: SessionFailure) => { failure = error; listener?.(error); } };
}

function resources(log: string[] = []) {
  const peer = { dispose: mock(() => { log.push('peer'); }) } as unknown as WorkbookHandle;
  const edits = {
    state: 'ready', dispose: mock(() => { log.push('facade'); }),
  } as unknown as WorkbookEditPeer;
  const hydrate = spyOn(editableWorkbookSessionBackend, 'hydrate').mockResolvedValue(peer);
  const attach = spyOn(editableWorkbookSessionBackend, 'attach').mockImplementation(() => {
    log.push('attach');
    return edits;
  });
  restorers.push(() => hydrate.mockRestore(), () => attach.mockRestore());
  return { peer, edits, hydrate, attach };
}

function owner(value: WorkbookSession, options: Partial<EditableWorkbookSessionOptions> = {}) {
  return new EditableWorkbookSession(value, 1, {
    changed: () => {}, onError: () => {}, onReady: () => {}, isCurrent: () => true, ...options,
  });
}

function bridge(): WorkerEditorApiBridge {
  return {
    coordinator: () => null, readOnly: () => false,
    clearSelection: mock(() => {}), focus: mock(() => {}), refreshProposals: mock(() => {}),
    recoverInput: async () => {}, selectCells: () => true, selectCellsAsync: async () => true,
    apply: () => {},
  };
}

function open(value: WorkbookSession) {
  const spy = spyOn(editableWorkbookSessionBackend, 'open').mockResolvedValue(value);
  restorers.push(() => spy.mockRestore());
  return spy;
}

afterEach(() => {
  cleanup();
  for (const restore of restorers.reverse()) restore();
  restorers.length = 0;
});
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

describe('editable session workbook', () => {
  test('opens a retained session and starts eager hydration only after the first paint', async () => {
    const { value } = session();
    const opener = open(value);
    const { peer, hydrate, attach } = resources();
    const pending = deferred<WorkbookHandle>();
    hydrate.mockImplementation(() => pending.promise);
    const localOpen = spyOn(xlsx, 'openWorkbook').mockImplementation(() => { throw new Error('Unexpected local open'); });
    restorers.push(() => localOpen.mockRestore());
    const log: string[] = [];
    const ready = mock((api: XlsxWorkerEditorApi) => {
      log.push('ready');
      expect(api.handle).toBeNull();
      expect(api.hydrated).toBe(false);
    });
    hydrate.mockImplementation(() => { log.push('hydrate'); return pending.promise; });
    const controls = bridge();
    const commands = createXlsxCommandController().store;
    const { result } = renderHook(() => useEditableSessionWorkbook({ file, onReady: ready }, commands, controls));
    await waitFor(() => expect(result.current.run).not.toBeNull());
    expect(opener).toHaveBeenCalledWith(file, { signal: expect.any(AbortSignal), retainPeerHydration: true });
    expect(hydrate).not.toHaveBeenCalled();
    const run = result.current.run!;
    await act(async () => run.firstPaint());
    expect(log).toEqual(['ready', 'hydrate']);
    expect(run.peer).toBeNull();
    await act(async () => pending.resolve(peer));
    expect(run.ready).toBe(true);
    expect(attach).toHaveBeenCalledWith({ session: value, peer, onError: expect.any(Function) });
    await act(async () => run.firstPaint());
    expect(ready).toHaveBeenCalledTimes(1);
    expect(hydrate).toHaveBeenCalledTimes(1);
    expect(localOpen).not.toHaveBeenCalled();
    expect(controls.focus).not.toHaveBeenCalled();
    expect(controls.clearSelection).not.toHaveBeenCalled();
  });

  test('attaches the facade before releasing passive hydration waiters', async () => {
    const log: string[] = [];
    const { value } = session(log);
    const { edits } = resources(log);
    const run = owner(value);
    const waiting = run.whenHydrated().then(() => {
      log.push('released');
      expect(run.editPeer).toBe(edits);
    });
    run.firstPaint();
    await waiting;
    expect(log).toEqual(['attach', 'released']);
    run.dispose();
  });

  test('keeps readiness passive and coalesces lazy hydration demands', async () => {
    const { value } = session();
    const { hydrate } = resources();
    const run = owner(value, { hydration: 'lazy' });
    const waiting = run.whenHydrated();
    run.firstPaint();
    expect(hydrate).not.toHaveBeenCalled();
    await Promise.all([run.requestHydration('edit-intent'), run.requestHydration('flush'), waiting]);
    expect(hydrate).toHaveBeenCalledTimes(1);
    run.dispose();
  });

  test('surfaces typed hydration failure once and retains the failed run', async () => {
    const { value } = session();
    open(value);
    const { hydrate } = resources();
    const error = new xlsx.WorkbookPeerHydrationError('missing-hydration', 'Worker hydration is missing');
    hydrate.mockRejectedValue(error);
    const errors = mock((_error: Error) => {});
    let api!: XlsxWorkerEditorApi;
    const ready = mock((value: XlsxWorkerEditorApi) => { api = value; });
    const commands = createXlsxCommandController().store;
    const { result } = renderHook(() => useEditableSessionWorkbook({ file, onError: errors, onReady: ready }, commands, bridge()));
    await waitFor(() => expect(result.current.run).not.toBeNull());
    const run = result.current.run!;
    await act(async () => run.firstPaint());
    await waitFor(() => expect(result.current.error).toBe(error));
    expect(result.current.run).toBe(run);
    expect(api.failure).toBe(error);
    await expect(api.whenHydrated()).rejects.toBeInstanceOf(xlsx.WorkbookPeerHydrationError);
    await act(async () => { run.fail(new Error('Second failure')); run.firstPaint(); });
    expect(errors).toHaveBeenCalledTimes(1);
    expect(ready).toHaveBeenCalledTimes(1);
  });

  test('retains a peer hydrated after worker failure for recovery', async () => {
    const opened = session();
    const { hydrate, peer, edits } = resources();
    const pending = deferred<WorkbookHandle>();
    hydrate.mockImplementation(() => pending.promise);
    const run = owner(opened.value);
    const waiting = run.requestHydration('flush');
    const error = new SessionFailure('crash', 'Worker stopped');
    opened.fail(error);
    await expect(waiting).rejects.toBe(error);
    const recovery = run.requestHydration('recovery');
    pending.resolve(peer);
    await recovery;
    expect(run.peer).toBe(peer);
    expect(run.editPeer).toBe(edits);
    expect(run.current).toBe(true);
    expect(run.ready).toBe(false);
    await expect(run.whenHydrated()).rejects.toBe(error);
    run.dispose();
  });

  test('rejects cancelled waits and disposes a late hydrated peer without attaching', async () => {
    const { value } = session();
    const { hydrate, peer, attach } = resources();
    const pending = deferred<WorkbookHandle>();
    hydrate.mockImplementation(() => pending.promise);
    const run = owner(value);
    const waiting = run.whenHydrated();
    const demanded = run.requestHydration('flush');
    run.dispose();
    await expect(waiting).rejects.toBeInstanceOf(XlsxCommandAdmissionError);
    await expect(demanded).rejects.toMatchObject({ code: 'document-replaced' });
    pending.resolve(peer);
    await pending.promise;
    await Promise.resolve();
    expect(peer.dispose).toHaveBeenCalledTimes(1);
    expect(attach).not.toHaveBeenCalled();
    expect(run.peer).toBeNull();
  });

  test('runs ready cleanup before facade, peer and session disposal', async () => {
    const log: string[] = [];
    const { value } = session(log);
    resources(log);
    const run = owner(value, { onReady: () => () => { log.push('cleanup'); } });
    run.firstPaint();
    await run.whenHydrated();
    log.length = 0;
    run.dispose();
    run.dispose();
    expect(log).toEqual(['off', 'cleanup', 'facade', 'peer', 'session']);
  });

  test('forwards failure to linked input once and discards it only on disposal', async () => {
    const { value } = session();
    const errors = mock((_error: Error) => {});
    const run = owner(value, { onError: errors });
    const input = createWorkerInputCoordinator({
      generation: () => run.generation, capture: () => ({ sheet: 0, target: 'A1' }),
      isReady: () => run.ready, whenReady: () => run.whenHydrated(),
      seal: () => ({}), sync: () => {}, preview: async () => {}, write: () => true,
      requestHydration: (reason) => run.requestHydration(reason), flushEdits: async () => {},
      onError: (error) => run.fail(error),
    });
    run.connectInput(input);
    const pending = input.submitAsync({ generation: 1, sheet: 0, row: 0, col: 0, source: 'cell', value: 'retained' });
    const failure = new SessionFailure('crash', 'Worker stopped');
    run.fail(failure);
    run.fail(new Error('Second failure'));
    await expect(pending).rejects.toBe(failure);
    expect(input.error).toBe(failure);
    expect(input.unapplied).toHaveLength(1);
    expect(errors).toHaveBeenCalledTimes(1);
    run.dispose();
    expect(input.unapplied).toEqual([]);
    expect(input.draft).toBeNull();
  });

  test('aborts a pending open and disposes its session after replacement', async () => {
    const old = session();
    const next = session();
    const pending = deferred<WorkbookSession>();
    const opener = open(next.value).mockImplementationOnce(() => pending.promise);
    const commands = createXlsxCommandController().store;
    const controls = bridge();
    const { result, rerender } = renderHook(
      (props: EditableSessionWorkbookProps) => useEditableSessionWorkbook(props, commands, controls),
      { initialProps: { file } }
    );
    await waitFor(() => expect(opener).toHaveBeenCalledTimes(1));
    const signal = opener.mock.calls[0][1]!.signal!;
    rerender({ file: new Uint8Array([4]) });
    await waitFor(() => expect(result.current.run?.session).toBe(next.value));
    expect(signal.aborted).toBe(true);
    await act(async () => pending.resolve(old.value));
    expect(old.value.dispose).toHaveBeenCalledTimes(1);
    expect(result.current.run?.session).toBe(next.value);
  });

  test('replaces generations and disposes hydrated runs on unmount', async () => {
    const old = session();
    const next = session();
    const opener = open(old.value);
    const { peer, edits } = resources();
    const commands = createXlsxCommandController().store;
    const controls = bridge();
    const { result, rerender, unmount } = renderHook(
      (props: EditableSessionWorkbookProps) => useEditableSessionWorkbook(props, commands, controls),
      { initialProps: { file } }
    );
    await waitFor(() => expect(result.current.run).not.toBeNull());
    const first = result.current.run!;
    await act(async () => { first.firstPaint(); await first.whenHydrated(); });
    opener.mockResolvedValue(next.value);
    rerender({ file: new Uint8Array([4]) });
    await waitFor(() => expect(result.current.run?.session).toBe(next.value));
    expect(result.current.run!.generation).toBeGreaterThan(first.generation);
    expect(first.current).toBe(false);
    expect(edits.dispose).toHaveBeenCalledTimes(1);
    expect(peer.dispose).toHaveBeenCalledTimes(1);
    unmount();
    expect(next.value.dispose).toHaveBeenCalledTimes(1);
  });

  test('disposes stale StrictMode opens and publishes readiness once', async () => {
    const old = session();
    const next = session();
    const opener = open(next.value).mockResolvedValueOnce(old.value);
    resources();
    const ready = mock((_api: XlsxWorkerEditorApi) => {});
    const commands = createXlsxCommandController().store;
    const { result } = renderHook(() => useEditableSessionWorkbook({ file, onReady: ready }, commands, bridge()), {
      wrapper: ({ children }) => <StrictMode>{children}</StrictMode>,
    });
    await waitFor(() => expect(result.current.run?.session).toBe(next.value));
    await act(async () => { result.current.run!.firstPaint(); result.current.run!.firstPaint(); });
    expect(opener).toHaveBeenCalledTimes(2);
    expect(old.value.dispose).toHaveBeenCalledTimes(1);
    expect(ready).toHaveBeenCalledTimes(1);
  });

  test('disposes every resource when ready cleanup throws', async () => {
    const log: string[] = [];
    const { value } = session(log);
    resources(log);
    const run = owner(value, { onReady: () => () => { throw new Error('Cleanup failed'); } });
    run.firstPaint();
    await run.whenHydrated();
    log.length = 0;
    expect(() => run.dispose()).toThrow('Cleanup failed');
    expect(log).toEqual(['off', 'facade', 'peer', 'session']);
  });
});
