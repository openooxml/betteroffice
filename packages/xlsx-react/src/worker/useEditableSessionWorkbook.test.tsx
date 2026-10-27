import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, describe, expect, mock, spyOn, test } from 'bun:test';
import * as xlsx from '@betteroffice/xlsx';
import type { WorkbookHandle, WorkbookSession, WorkbookSheetView } from '@betteroffice/xlsx';
import type { WorkbookCommand, WorkbookCommandFrame, WorkbookEditState } from '@betteroffice/xlsx/worker';
import { SessionFailure } from '../../../../shared/office-session';
import type { WorkbookEditPeer } from '@betteroffice/xlsx/worker';
import { createXlsxCommandController, XlsxCommandAdmissionError } from '../commands/createXlsxCommandStore';
import { PLAIN_FORMATTING, testBinding, testSessionCommands } from '../commands/testing';
import type { InputDraft } from '../commands/inputCoordinator';
import { createWorkerInputCoordinator, type WorkerInputCoordinatorHooks } from '../commands/workerInputCoordinator';
import type { WorkerEditorApiBridge, XlsxWorkerEditorApi } from './createWorkerEditorApi';
import { createAsyncEditorApi, type XlsxAsyncEditorApi } from './createAsyncEditorApi';
import { PeerlessWorkerPaintSource as WorkerPaintSource, type WorkerPaintRequest } from './PeerlessWorkerPaintSource';
import {
  EditableWorkbookSession, editableWorkbookSessionBackend, useEditableSessionWorkbook,
  XlsxWorkerSessionFailedError, type EditableSessionWorkbookProps, type EditableWorkbookSessionOptions,
} from './useEditableSessionWorkbook';
import { PeerlessWorkbookSession, usePeerlessSessionWorkbook } from './usePeerlessSessionWorkbook';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, render, renderHook, waitFor } = await import('@testing-library/react');
const { XlsxWorkerEditor } = await import('./XlsxWorkerEditor');
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
    ...testSessionCommands(),
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

function firstFrame(value: WorkbookSession): WorkbookCommandFrame {
  return { ...value.commandIdentity, coversThroughSeq: 0, requestGeneration: 1,
    zoom: 1, dpr: 1, sheet: 0, sequence: 0, epoch: 1, version: 'v1',
    viewport: { x: 0, y: 0, width: 800, height: 600 },
    displayList: { width: 800, height: 600, commands: [] } };
}

function paintRequest(run: PeerlessWorkbookSession): WorkerPaintRequest {
  return { generation: run.generation, sheet: 0, navigation: 0,
    width: 800, height: 600, zoom: 1, dpr: 1, viewport: { x: 0, y: 0, width: 800, height: 600 } };
}

function editorHarness(value: WorkbookSession, width = 800, height = 600) {
  const dimensions = { width, height };
  const frames = new Map<number, FrameRequestCallback>();
  let nextFrame = 0;
  const originalRaf = globalThis.requestAnimationFrame;
  const originalCancel = globalThis.cancelAnimationFrame;
  const originalResize = globalThis.ResizeObserver;
  globalThis.requestAnimationFrame = (callback) => { const id = ++nextFrame; frames.set(id, callback); return id; };
  globalThis.cancelAnimationFrame = (id) => { frames.delete(id); };
  globalThis.ResizeObserver = class {
    observe() {}
    disconnect() {}
  } as unknown as typeof ResizeObserver;
  restorers.push(() => {
    globalThis.requestAnimationFrame = originalRaf;
    globalThis.cancelAnimationFrame = originalCancel;
    globalThis.ResizeObserver = originalResize;
  });
  for (const [name, get] of [
    ['clientWidth', () => dimensions.width], ['clientHeight', () => dimensions.height],
  ] as const) {
    const original = Object.getOwnPropertyDescriptor(HTMLElement.prototype, name);
    Object.defineProperty(HTMLElement.prototype, name, { configurable: true, get });
    restorers.push(() => {
      if (original) Object.defineProperty(HTMLElement.prototype, name, original);
      else Reflect.deleteProperty(HTMLElement.prototype, name);
    });
  }
  const context = spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue({} as CanvasRenderingContext2D);
  const paint = spyOn(xlsx, 'paintDisplayList').mockImplementation(() => {});
  const proposals = spyOn(xlsx, 'isProposalsAvailable').mockReturnValue(false);
  restorers.push(() => context.mockRestore(), () => paint.mockRestore(), () => proposals.mockRestore());
  const opener = open(value);
  const { hydrate, attach } = resources();
  hydrate.mockImplementation(() => deferred<WorkbookHandle>().promise);
  const view: WorkbookSheetView = { sheet: 0, version: 'v1', contentWidth: 4000, contentHeight: 6000,
    frozenRows: 0, frozenCols: 0, frozenWidth: 0, frozenHeight: 0, initialScrollX: 0, initialScrollY: 0 };
  const sheetView = mock(async () => view);
  const frame = mock(async (viewport: xlsx.Viewport) => ({ ...firstFrame(value), viewport,
    sequence: value.throughSeq, coversThroughSeq: value.throughSeq }));
  Object.assign(value.call, { sheetView, frame, cellInputs: async () => ({ sheet: 0, version: 'v1', cells: [[null]] }) });
  const ready = mock((_api: XlsxWorkerEditorApi | XlsxAsyncEditorApi) => {});
  const errors = mock((_error: Error) => {});
  return { dimensions, opener, view, sheetView, frame, paint, ready, errors, hydrate, attach,
    render: () => render(<XlsxWorkerEditor backend="peerless" file={file} showToolbar={false} onReady={ready} onError={errors} />),
    tick: async () => act(async () => {
      const callbacks = [...frames.values()];
      frames.clear();
      for (const callback of callbacks) callback(0);
    }),
  };
}

afterEach(() => {
  cleanup();
  for (const restore of restorers.reverse()) restore();
  restorers.length = 0;
});
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

describe('editable session workbook', () => {
  test('async recovery saving requires failure and preserves the failed editor', async () => {
    const opened = session();
    const save = mock(async () => new Uint8Array([7]));
    const recovery = { save, retainedCommands: 0, throughSeq: 0, capture: mock(async () => true) };
    xlsx.workbookSessionRecovery.set(opened.value, recovery as unknown as xlsx.WorkbookSessionRecovery);
    const errors = mock((_error: Error) => {});
    const run = new PeerlessWorkbookSession(opened.value, 1, {
      changed: () => {}, onError: errors, onReady: () => {}, isCurrent: () => true,
    });
    const api = createAsyncEditorApi(run, createXlsxCommandController().store, bridge);
    run.firstPaint();
    const healthy = await api.recoverySave().catch((error: unknown) => error);
    expect(healthy).toBeInstanceOf(Error);
    expect(Object.getPrototypeOf(healthy)).toBe(Error.prototype);
    expect((healthy as Error).message).toBe('Recovery requires a failed workbook edit peer');
    expect(save).not.toHaveBeenCalled();
    opened.fail(new SessionFailure('crash', 'Worker stopped'));
    expect(await api.recoverySave()).toEqual({ bytes: new Uint8Array([7]), recovery: true });
    expect(api.failure).toBe(run.failure);
    expect(api.hydrated).toBe(false);
    expect(run.readOnly).toBe(true);
    expect(errors).toHaveBeenCalledTimes(1);
    await run.dispose();
  });

  test.each(['replaced', 'disposed'] as const)('stale async recovery API rejects when %s', async (state) => {
    const opened = session();
    let current = true;
    const save = mock(async () => new Uint8Array([7]));
    xlsx.workbookSessionRecovery.set(opened.value, { save } as unknown as xlsx.WorkbookSessionRecovery);
    const run = new PeerlessWorkbookSession(opened.value, 1, {
      changed: () => {}, onError: () => {}, onReady: () => {}, isCurrent: () => current,
    });
    const api = createAsyncEditorApi(run, createXlsxCommandController().store, bridge);
    opened.fail(new SessionFailure('crash', 'Worker stopped'));
    if (state === 'disposed') await run.dispose();
    else current = false;
    await expect(api.recoverySave()).rejects.toBeInstanceOf(XlsxCommandAdmissionError);
    await expect(api.recoverySave()).rejects.toMatchObject({ code: 'document-replaced' });
    expect(save).not.toHaveBeenCalled();
    await run.dispose();
  });

  test('async recovery reports replay refusal through failure and onError', async () => {
    const opened = session();
    const refusal = new xlsx.WorkbookRecoveryReplayError(2, 'locked-target');
    xlsx.workbookSessionRecovery.set(opened.value, {
      save: mock(async () => { throw refusal; }),
    } as unknown as xlsx.WorkbookSessionRecovery);
    const errors = mock((_error: Error) => {});
    const run = new PeerlessWorkbookSession(opened.value, 1, {
      changed: () => {}, onError: errors, onReady: () => {}, isCurrent: () => true,
    });
    const api = createAsyncEditorApi(run, createXlsxCommandController().store, bridge);
    opened.fail(new SessionFailure('crash', 'Worker stopped'));
    await expect(api.recoverySave()).rejects.toBeInstanceOf(XlsxWorkerSessionFailedError);
    expect(run.failure?.cause).toBe(refusal);
    const failure = api.failure;
    if (!failure) throw new Error('Missing recovery failure');
    expect(errors.mock.calls[errors.mock.calls.length - 1]?.[0]).toBe(failure);
    await run.dispose();
  });

  test('async recovery rejects replacement during a save', async () => {
    const opened = session();
    const pending = deferred<Uint8Array>();
    let current = true;
    xlsx.workbookSessionRecovery.set(opened.value, {
      save: () => pending.promise,
    } as unknown as xlsx.WorkbookSessionRecovery);
    const run = new PeerlessWorkbookSession(opened.value, 1, {
      changed: () => {}, onError: () => {}, onReady: () => {}, isCurrent: () => current,
    });
    const api = createAsyncEditorApi(run, createXlsxCommandController().store, bridge);
    opened.fail(new SessionFailure('crash', 'Worker stopped'));
    const saved = api.recoverySave();
    current = false;
    pending.resolve(new Uint8Array([7]));
    await expect(saved).rejects.toMatchObject({ code: 'document-replaced' });
    await run.dispose();
  });

  test('checkpointing waits for closed drafts and compositions and batches acknowledgements', async () => {
    const timers = new Map<number, { callback(): void; delay: number }>();
    let nextTimer = 0;
    let now = 0;
    const checkpointScheduler = {
      now: () => now,
      schedule(callback: () => void, delay: number) {
        const id = ++nextTimer;
        timers.set(id, { callback, delay });
        return () => { timers.delete(id); };
      },
    };
    const tick = async () => {
      now += 150;
      const callbacks = [...timers.values()];
      timers.clear();
      for (const timer of callbacks) timer.callback();
      await Promise.resolve();
      await Promise.resolve();
    };
    const opened = session();
    let ack!: Parameters<WorkbookSession['onAck']>[0];
    opened.value.onAck = (listener) => { ack = listener; return () => {}; };
    const checkpoint = mock(async (allowed: () => boolean) => allowed());
    const recovery: xlsx.WorkbookSessionRecovery = { retainedCommands: 64, throughSeq: 0,
      checkpointIdleMs: 150, capture: checkpoint, save: async () => new Uint8Array([7]) };
    xlsx.workbookSessionRecovery.set(opened.value, recovery);
    const run = new PeerlessWorkbookSession(opened.value, 1, {
      changed: () => {}, onError: () => {}, onReady: () => {}, isCurrent: () => true, checkpointScheduler,
    });
    const draft: InputDraft = { generation: 1, sheet: 0, row: 0, col: 0, value: 'draft', source: 'cell' };
    const input: { draft: InputDraft | null; pending: boolean } = { draft, pending: false };
    let changed!: () => void;
    const drafts = { compositionId: undefined as string | undefined,
      subscribe(listener: () => void) { changed = listener; return () => {}; } };
    run.connectRecoveryInput(input, drafts);
    run.firstPaint();
    const receipt = { throughSeq: 64, version: 'v64', statuses: [] };
    ack(receipt);
    await Promise.resolve();
    expect(checkpoint).not.toHaveBeenCalled();
    input.draft = null;
    drafts.compositionId = 'composition';
    changed(); ack(receipt);
    await Promise.resolve();
    expect(checkpoint).not.toHaveBeenCalled();
    drafts.compositionId = undefined;
    changed(); ack(receipt); ack(receipt);
    expect(checkpoint).not.toHaveBeenCalled();
    await Promise.resolve();
    expect(checkpoint).not.toHaveBeenCalled();
    expect(timers.size).toBe(1);
    expect([...timers.values()].map((timer) => timer.delay)).toEqual([150]);
    input.draft = { ...draft, value: 'next' };
    changed();
    await tick();
    expect(checkpoint).not.toHaveBeenCalled();
    input.draft = null;
    input.pending = true;
    changed();
    await tick();
    expect(checkpoint).not.toHaveBeenCalled();
    input.pending = false;
    changed();
    await tick();
    expect(checkpoint).toHaveBeenCalledTimes(1);
    expect(checkpoint.mock.calls[0][0]()).toBe(true);
    input.draft = { ...draft, value: 'next' };
    expect(checkpoint.mock.calls[0][0]()).toBe(false);
    input.draft = null;
    changed();
    await run.dispose();
    await tick();
    expect(checkpoint).toHaveBeenCalledTimes(1);
  });

  test('ready follows first successful paint without hydration', async () => {
    const { value } = session();
    const opener = open(value);
    const { hydrate, attach } = resources();
    const readyCleanup = mock(() => {});
    const ready = mock((_api: XlsxAsyncEditorApi) => readyCleanup);
    const controller = createXlsxCommandController();
    const binding = testBinding();
    controller.attach(binding.binding);
    const commands = controller.store;
    const { result, unmount } = renderHook(() => usePeerlessSessionWorkbook({ file, onReady: ready }, commands, bridge(), {
      backend: 'peerless', initialViewport: () => ({ width: 800, height: 600, zoom: 1, dpr: 1, requestGeneration: 1 }),
    }));
    await waitFor(() => expect(result.current.run).toBeInstanceOf(PeerlessWorkbookSession));
    const run = result.current.run!;
    expect(run.commandsAccepted).toBe(true);
    expect(run.ready).toBe(false);
    expect(run.readOnly).toBe(true);
    let readyResolved = false;
    const readiness = run.whenReady().then(() => { readyResolved = true; });
    expect(readyResolved).toBe(false);
    expect(ready).not.toHaveBeenCalled();
    expect(opener.mock.calls[0][1]).toMatchObject({ initialViewport: { width: 800, height: 600, zoom: 1, dpr: 1, requestGeneration: 1 } });
    expect(opener.mock.calls[0][1]?.retainPeerHydration).toBeUndefined();
    expect(opener.mock.calls[0][1]?.includeEditState).toBe(true);
    const frame = firstFrame(value);
    const source = new WorkerPaintSource({ generation: run.generation, commandIdentity: value.commandIdentity,
      capture: () => paintRequest(run), sentSequence: () => 0,
      requestFrame: mock(async () => { throw new Error('Unexpected frame RPC'); }),
      publish: () => run.firstPaint(), onError: (error) => run.fail(error),
    });
    await act(async () => { expect(source.installInitial(frame)).toBe(true); run.firstPaint(); });
    await readiness;
    expect(readyResolved).toBe(true);
    expect(run.ready).toBe(true);
    expect(run.readOnly).toBe(false);
    expect(ready).toHaveBeenCalledTimes(1);
    expect(ready.mock.calls[0][0]).not.toBe(run);
    const api = ready.mock.calls[0][0];
    expect(api.hydrated).toBe(true);
    expect(api.commands).not.toBe(commands);
    expect(api.commands.getState('undo')).toBe(commands.getState('undo'));
    const changed = mock(() => {});
    const unsubscribe = api.commands.subscribe(changed);
    binding.update({ canUndo: false });
    controller.refresh();
    expect(changed).toHaveBeenCalledTimes(1);
    unsubscribe();
    binding.update({ canUndo: true });
    controller.refresh();
    expect(changed).toHaveBeenCalledTimes(1);
    expect(hydrate).not.toHaveBeenCalled();
    expect(attach).not.toHaveBeenCalled();
    expect('peer' in run).toBe(false);
    expect('editPeer' in run).toBe(false);
    source.dispose();
    unmount();
    expect(readyCleanup).toHaveBeenCalledTimes(1);
    expect(value.dispose).toHaveBeenCalledTimes(1);
    await expect(api.commands.execute('undo', null)).rejects.toMatchObject({ code: 'document-replaced' });
    expect(binding.calls).toEqual([]);
  });

  test('peerless refuses user mutations until the first successful paint while allowing initialization reads', async () => {
    const { value } = session();
    const submit = mock<WorkbookSession['submit']>(async (command) => ({ seq: command.seq, status: 'noop' }));
    const command = mock<WorkbookSession['call']['command']>(async () => ({ ...value.commandIdentity,
      throughSeq: 1, version: 'v1', statuses: [{ seq: 1, status: 'noop' }] }));
    const applyEdits = mock<WorkbookSession['call']['applyEdits']>(async () => ({ ok: false, version: 'v1',
      failure: { code: 'stale-version', message: 'Version changed' } }));
    value.submit = submit;
    const initial = firstFrame(value);
    Object.assign(value.call, { command, applyEdits, version: async () => 'v1', frame: async () => initial });
    open(value);
    const commands = createXlsxCommandController().store;
    const { result } = renderHook(() => usePeerlessSessionWorkbook({ file }, commands, bridge(), { backend: 'peerless' }));
    await waitFor(() => expect(result.current.run).not.toBeNull());
    const run = result.current.run!;
    const input: WorkbookCommand = { ...value.commandIdentity, seq: 1, basis: value.confirmedBasis,
      kind: 'editCell', payload: { args: [0, 0, 0, 'blocked'] } };
    const batch = { expectVersion: 'v1', steps: [] };
    expect(run.commandsAccepted).toBe(true);
    expect(run.readOnly).toBe(true);
    for (const mutation of [() => run.session.submit(input), () => run.session.call.command(input),
      () => run.session.call.applyEdits(batch)]) {
      const refused = mutation();
      await expect(refused).rejects.toBeInstanceOf(XlsxCommandAdmissionError);
      await expect(refused).rejects.toMatchObject({ code: 'editor-unavailable' });
    }
    expect(submit).not.toHaveBeenCalled();
    expect(command).not.toHaveBeenCalled();
    expect(applyEdits).not.toHaveBeenCalled();
    expect(await run.session.call.version()).toBe('v1');
    expect(await run.session.call.frame(initial.viewport)).toBe(initial);
    await act(async () => run.firstPaint());
    expect(run.readOnly).toBe(false);
    await run.session.submit(input);
    await run.session.call.command(input);
    await run.session.call.applyEdits(batch);
    expect(submit).toHaveBeenCalledTimes(1);
    expect(command).toHaveBeenCalledTimes(1);
    expect(applyEdits).toHaveBeenCalledTimes(1);
  });

  test.each([false, true])('peerless proxy disposal retires readiness and pending work (painted=%s)', async (painted) => {
    const log: string[] = [];
    const { value } = session(log);
    const pending = deferred<string>();
    const disposal = deferred<void>();
    value.call.version = () => pending.promise;
    const offAck = mock(() => { log.push('off-ack'); });
    value.onAck = () => offAck;
    value.dispose = mock(() => { log.push('session'); return disposal.promise; });
    const cleanup = mock(() => { log.push('cleanup'); });
    const errors = mock((_error: Error) => {});
    const run = new PeerlessWorkbookSession(value, 1, {
      changed: () => {}, onError: errors, onReady: () => cleanup, isCurrent: () => true,
    });
    const notify = mock(() => {});
    run.subscribe(notify);
    const readiness = run.whenReady();
    if (painted) { run.firstPaint(); await readiness; }
    const reading = run.session.call.version();
    const { dispose } = run.session;
    expect(run.session.dispose).toBe(dispose);
    const disposed = dispose();
    expect(disposed).toBe(disposal.promise);
    expect(dispose()).toBe(disposed);
    expect(run.dispose()).toBe(disposed);
    expect(run.current).toBe(false);
    expect(run.ready).toBe(false);
    expect(run.readOnly).toBe(true);
    if (!painted) await expect(readiness).rejects.toMatchObject({ code: 'document-replaced' });
    await expect(run.whenReady()).rejects.toMatchObject({ code: 'document-replaced' });
    await expect(reading).rejects.toMatchObject({ code: 'document-replaced' });
    await expect(run.session.call.version()).rejects.toMatchObject({ code: 'document-replaced' });
    expect(cleanup).toHaveBeenCalledTimes(painted ? 1 : 0);
    expect(log).toEqual(painted ? ['off', 'off-ack', 'cleanup', 'session'] : ['off', 'off-ack', 'session']);
    expect(offAck).toHaveBeenCalledTimes(1);
    expect(value.dispose).toHaveBeenCalledTimes(1);
    expect(notify).toHaveBeenCalled();
    expect(errors).not.toHaveBeenCalled();
    const failure = new Error('Disposal failed');
    disposal.reject(failure);
    await expect(disposed).rejects.toBe(failure);
    pending.resolve('v1');
    await pending.promise;
  });

  test('unknown viewport delays ready without reopening', async () => {
    const { value } = session();
    const opener = open(value);
    const { hydrate, attach } = resources();
    const ready = mock((_api: XlsxAsyncEditorApi) => {});
    const commands = createXlsxCommandController().store;
    let viewport: { width: number; height: number; zoom: number; dpr: number; requestGeneration: number } | undefined;
    const { result, rerender } = renderHook(() => usePeerlessSessionWorkbook({ file, onReady: ready }, commands, bridge(), {
      backend: 'peerless', initialViewport: () => viewport,
    }));
    await waitFor(() => expect(result.current.run).not.toBeNull());
    expect(result.current.run!.ready).toBe(false);
    expect(opener.mock.calls[0][1]?.initialViewport).toBeUndefined();
    const run = result.current.run!;
    const pending = deferred<WorkbookCommandFrame>();
    const requestFrame = mock(() => pending.promise);
    Object.assign(value.call, { frame: requestFrame });
    let scheduled: (() => void) | undefined;
    const source = new WorkerPaintSource({ generation: run.generation, commandIdentity: value.commandIdentity,
      capture: () => viewport ? paintRequest(run) : null, sentSequence: () => 0,
      requestFrame: (request) => run.session.call.frame(request.viewport, { sheet: request.sheet }),
      publish: () => run.firstPaint(), onError: (error) => run.fail(error),
      scheduler: { request: (callback) => { scheduled = callback; return 1; }, cancel: () => { scheduled = undefined; } },
    });
    source.schedule();
    scheduled?.();
    expect(requestFrame).not.toHaveBeenCalled();
    viewport = { width: 800, height: 600, zoom: 1, dpr: 1, requestGeneration: 2 };
    rerender();
    expect(opener).toHaveBeenCalledTimes(1);
    expect(ready).not.toHaveBeenCalled();
    source.schedule();
    scheduled?.();
    expect(requestFrame).toHaveBeenCalledTimes(1);
    expect(ready).not.toHaveBeenCalled();
    await act(async () => { pending.resolve(firstFrame(value)); await pending.promise; });
    expect(ready).toHaveBeenCalledTimes(1);
    expect(hydrate).not.toHaveBeenCalled();
    expect(attach).not.toHaveBeenCalled();
    source.dispose();
  });

  test('replacement disposes late open', async () => {
    const old = session();
    const next = session();
    const pending = deferred<WorkbookSession>();
    const opener = open(next.value).mockImplementationOnce(() => pending.promise);
    const commands = createXlsxCommandController().store;
    const controls = bridge();
    const { result, rerender } = renderHook((props: { file: Uint8Array }) =>
      usePeerlessSessionWorkbook(props, commands, controls, { backend: 'peerless' }), { initialProps: { file } });
    rerender({ file: new Uint8Array([4]) });
    await waitFor(() => expect(result.current.run?.session.state).toBe(next.value.state));
    await act(async () => pending.resolve(old.value));
    expect(old.value.dispose).toHaveBeenCalledTimes(1);
    expect(opener).toHaveBeenCalledTimes(2);
    expect(result.current.run).toBeInstanceOf(PeerlessWorkbookSession);
  });

  test('component installs the mounted viewport fused frame without a frame RPC', async () => {
    const { value } = session();
    const editor = editorHarness(value);
    const editState: WorkbookEditState = { canUndo: false, canRedo: false, formatting: PLAIN_FORMATTING,
      formattingAvailable: true, selectionEditable: true, mergedRanges: [], calculation: { limitedCells: [] },
      sheets: [], activeSheet: 0, readOnly: false };
    const frame = { ...firstFrame(value), requestGeneration: 0, dpr: window.devicePixelRatio || 1 };
    Object.assign(value, { initialPaint: { sheetView: editor.view, frame, editState } });
    const mounted = editor.render();
    await waitFor(() => expect(editor.ready).toHaveBeenCalledTimes(1));
    expect(editor.opener.mock.calls[0][1]?.initialViewport).toEqual({ width: 800, height: 600,
      zoom: 1, dpr: window.devicePixelRatio || 1, requestGeneration: 0 });
    expect(editor.opener.mock.calls[0][1]?.retainPeerHydration).toBeUndefined();
    expect(editor.opener.mock.calls[0][1]?.includeEditState).toBe(true);
    expect(editor.paint).toHaveBeenCalledWith(expect.anything(), frame.displayList, frame.dpr);
    await editor.tick();
    expect(editor.frame).not.toHaveBeenCalled();
    expect(editor.sheetView).not.toHaveBeenCalled();
    expect(editor.errors).not.toHaveBeenCalled();
    mounted.rerender(<XlsxWorkerEditor backend="legacy" file={file} showToolbar={false}
      onReady={editor.ready} onError={editor.errors} />);
    await editor.tick();
    expect(editor.opener).toHaveBeenCalledTimes(1);
    expect(editor.ready.mock.calls[0][0]).not.toBeInstanceOf(PeerlessWorkbookSession);
    expect(editor.ready.mock.calls[0][0].hydrated).toBe(true);
    expect(editor.hydrate).not.toHaveBeenCalled();
    expect(editor.attach).not.toHaveBeenCalled();
  });

  test.each(['unknown', 'resized'])('component presents a covering fallback frame after %s viewport open', async (viewport) => {
    const { value } = session();
    const editor = editorHarness(value, viewport === 'unknown' ? 0 : 800);
    const opening = deferred<WorkbookSession>();
    editor.opener.mockImplementation(() => opening.promise);
    if (viewport === 'resized') Object.assign(value, { initialPaint: {
      sheetView: editor.view, frame: { ...firstFrame(value), requestGeneration: 0 }, editState: {},
    } });
    const presented = deferred<void>();
    value.whenPresented = (seq) => seq === 1 ? presented.promise : Promise.resolve();
    value.markPresented = mock<WorkbookSession['markPresented']>((frame) => {
      expect(frame).toEqual({ ...value.commandIdentity, coversThroughSeq: 1 });
      presented.resolve();
      return true;
    });
    value.submit = async (command) => {
      Object.assign(value, { throughSeq: command.seq });
      return { seq: command.seq, status: 'noop' };
    };
    editor.render();
    await waitFor(() => expect(editor.opener).toHaveBeenCalledTimes(1));
    editor.dimensions.width = 900;
    await act(async () => opening.resolve(value));
    expect(editor.ready).not.toHaveBeenCalled();
    await value.submit({ ...value.commandIdentity, seq: 1, basis: value.confirmedBasis,
      kind: 'select', payload: { selection: { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1' } } } });
    const waiting = value.whenPresented(1);
    await editor.tick();
    await waitFor(() => expect(value.markPresented).toHaveBeenCalledTimes(1));
    await waiting;
    expect(editor.ready).toHaveBeenCalledTimes(1);
    expect(editor.frame.mock.calls[0][0].width).toBe(900);
    expect(editor.opener).toHaveBeenCalledTimes(1);
    expect(editor.errors).not.toHaveBeenCalled();
  });

  test('component does not mark a failed fallback paint as presented', async () => {
    const { value } = session();
    const editor = editorHarness(value);
    const markPresented = mock<WorkbookSession['markPresented']>(() => true);
    value.markPresented = markPresented;
    editor.paint.mockImplementation(() => { throw new Error('Paint failed'); });
    editor.render();
    await waitFor(() => expect(editor.sheetView).toHaveBeenCalledTimes(1));
    await editor.tick();
    expect(editor.errors).toHaveBeenCalledTimes(1);
    expect(markPresented).not.toHaveBeenCalled();
    expect(editor.ready).not.toHaveBeenCalled();
  });

  test('component sizes fetched geometry before assigning initial scroll offsets', async () => {
    const { value } = session();
    const editor = editorHarness(value, 0);
    editor.view.initialScrollX = 192;
    editor.view.initialScrollY = 120;
    const geometry = deferred<WorkbookSheetView>();
    editor.sheetView.mockImplementation(() => geometry.promise);
    const mounted = editor.render();
    await waitFor(() => expect(editor.sheetView).toHaveBeenCalledTimes(1));
    const scroll = mounted.getByTestId('xlsx-scroll');
    const content = scroll.firstElementChild as HTMLElement;
    for (const [name, dimension, viewport] of [
      ['scrollLeft', 'width', 800], ['scrollTop', 'height', 600],
    ] as const) {
      let offset = 0;
      Object.defineProperty(scroll, name, { configurable: true, get: () => offset,
        set: (value: number) => { offset = Math.min(value, Math.max(0, (parseFloat(content.style[dimension]) || 0) - viewport)); },
      });
    }
    editor.dimensions.width = 800;
    await act(async () => geometry.resolve(editor.view));
    expect(scroll.scrollLeft).toBe(192);
    expect(scroll.scrollTop).toBe(120);
    await editor.tick();
    expect(editor.frame.mock.calls[0][0]).toEqual({ x: 192, y: 120, width: 800, height: 600 });
    expect(editor.errors).not.toHaveBeenCalled();
  });

  test('peerless proxy methods preserve their receiver and cached identity', async () => {
    const { value } = session();
    value.call.version = async function (this: WorkbookSession['call']) { expect(this).toBe(value.call); return 'v1'; };
    value.save = async function (this: WorkbookSession) { expect(this).toBe(value); return new Uint8Array([9]); };
    value.submit = async function (this: WorkbookSession, command) { expect(this).toBe(value); return { seq: command.seq, status: 'noop' }; };
    value.whenPresented = async function (this: WorkbookSession, seq) { expect(this).toBe(value); expect(seq).toBe(1); };
    const run = new PeerlessWorkbookSession(value, 1, {
      changed: () => {}, onError: () => {}, onReady: () => {}, isCurrent: () => true,
    });
    const { version } = run.session.call;
    const { save, submit, whenPresented } = run.session;
    expect(run.session.call.version).toBe(version);
    expect(run.session.save).toBe(save);
    expect(run.session.submit).toBe(submit);
    expect(run.session.whenPresented).toBe(whenPresented);
    expect(await version()).toBe('v1');
    expect(await save()).toEqual(new Uint8Array([9]));
    run.firstPaint();
    expect(await submit({ ...value.commandIdentity, seq: 1, basis: value.confirmedBasis,
      kind: 'select', payload: { selection: { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1' } } } })).toEqual({ seq: 1, status: 'noop' });
    await whenPresented(1);
    run.dispose();
  });

  test('peerless StrictMode disposes stale opens and keeps the backend fixed at mount', async () => {
    const old = session();
    const next = session();
    const opener = open(next.value).mockResolvedValueOnce(old.value);
    const { hydrate, attach } = resources();
    const ready = mock((_api: XlsxAsyncEditorApi) => {});
    const commands = createXlsxCommandController().store;
    const { result, rerender } = renderHook((props: { peerless: boolean }) =>
      usePeerlessSessionWorkbook({ file, onReady: ready }, commands, bridge(), {
        backend: (props.peerless ? 'peerless' : 'legacy') as 'peerless',
      }), {
      initialProps: { peerless: true }, reactStrictMode: true,
    });
    await waitFor(() => expect(result.current.run).toBeInstanceOf(PeerlessWorkbookSession));
    expect(old.value.dispose).toHaveBeenCalledTimes(1);
    rerender({ peerless: false });
    expect(result.current.run).toBeInstanceOf(PeerlessWorkbookSession);
    await act(async () => { result.current.run!.firstPaint(); result.current.run!.firstPaint(); });
    expect(opener).toHaveBeenCalledTimes(2);
    expect(ready).toHaveBeenCalledTimes(1);
    expect(hydrate).not.toHaveBeenCalled();
    expect(attach).not.toHaveBeenCalled();
  });

  test('failed first paint never publishes peerless readiness', async () => {
    const { value } = session();
    const { hydrate, attach } = resources();
    const ready = mock(() => {});
    const errors = mock((_error: Error) => {});
    const run = new PeerlessWorkbookSession(value, 1, {
      changed: () => {}, onError: errors, onReady: ready, isCurrent: () => true,
    });
    const cause = new Error('Canvas unavailable');
    const source = new WorkerPaintSource({ generation: 1, commandIdentity: value.commandIdentity,
      capture: () => paintRequest(run), sentSequence: () => 0,
      requestFrame: async () => { throw new Error('Unexpected frame RPC'); },
      publish: () => { throw cause; }, onError: (error) => run.fail(error),
    });
    expect(source.installInitial(firstFrame(value))).toBe(false);
    expect(source.painted).toBeNull();
    expect(run.ready).toBe(false);
    expect(ready).not.toHaveBeenCalled();
    expect(run.failure?.cause).toBe(cause);
    await expect(run.whenReady()).rejects.toBe(run.failure);
    expect(hydrate).not.toHaveBeenCalled();
    expect(attach).not.toHaveBeenCalled();
    expect(errors).toHaveBeenCalledTimes(1);
    source.dispose();
    run.dispose();
  });

  test('legacy still hydrates and attaches', async () => {
    const { value } = session();
    open(value);
    const { hydrate, attach } = resources();
    const commands = createXlsxCommandController().store;
    const { result } = renderHook(() => useEditableSessionWorkbook({ file }, commands, bridge(), { backend: 'legacy' }));
    await waitFor(() => expect(result.current.run).toBeInstanceOf(EditableWorkbookSession));
    await act(async () => { result.current.run!.firstPaint(); await result.current.run!.whenHydrated(); });
    expect(result.current.run!.ready).toBe(true);
    expect(hydrate).toHaveBeenCalledTimes(1);
    expect(attach).toHaveBeenCalledTimes(1);
  });

  test('worker failure reports typed error and freezes read-only', async () => {
    const opened = session();
    const pending = deferred<never>();
    Object.assign(opened.value.call, { version: () => pending.promise, command: () => pending.promise });
    opened.value.save = () => pending.promise;
    const errors = mock((_error: Error) => {});
    const changed = mock(() => {});
    const run = new PeerlessWorkbookSession(opened.value, 1, {
      changed, onError: errors, onReady: () => {}, isCurrent: () => true,
    });
    const notified = mock(() => {});
    const off = run.subscribe(notified);
    const source = new WorkerPaintSource({ generation: run.generation, commandIdentity: opened.value.commandIdentity,
      capture: () => paintRequest(run), sentSequence: () => 0, isCurrent: () => run.current && !run.failure,
      requestFrame: mock(async () => { throw new Error('Unexpected frame RPC'); }),
      publish: () => run.firstPaint(), onError: (error) => run.fail(error),
    });
    expect(source.installInitial(firstFrame(opened.value))).toBe(true);
    const lastPaint = source.painted;
    const reading = run.session.call.version();
    const saving = run.session.save();
    const commanding = run.session.call.command({ ...opened.value.commandIdentity, seq: 1,
      basis: opened.value.confirmedBasis, kind: 'select', payload: { selection: { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1' } } } });
    const cause = new SessionFailure('crash', 'Worker stopped');
    opened.fail(cause);
    expect(run.failure).toBeInstanceOf(XlsxWorkerSessionFailedError);
    expect(run.failure?.cause).toBe(cause);
    await expect(reading).rejects.toBe(run.failure);
    await expect(saving).rejects.toBe(run.failure);
    await expect(commanding).rejects.toBe(run.failure);
    await expect(run.session.call.version()).rejects.toBe(run.failure);
    expect(run.readOnly).toBe(true);
    expect(run.ready).toBe(false);
    expect(run.current).toBe(true);
    expect(errors).toHaveBeenCalledTimes(1);
    expect(errors).toHaveBeenCalledWith(run.failure);
    expect(notified).toHaveBeenCalled();
    expect(opened.value.dispose).not.toHaveBeenCalled();
    expect(source.installInitial(firstFrame(opened.value))).toBe(false);
    expect(source.painted).toBe(lastPaint);
    source.dispose();
    off();
    run.dispose();
  });

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

  test('forwards failure to linked input once and recovers accepted edits before disposal', async () => {
    const log: string[] = [];
    const { value } = session(log);
    resources(log);
    const errors = mock((_error: Error) => {});
    const run = owner(value, { onError: errors });
    const write = mock((..._args: Parameters<WorkerInputCoordinatorHooks['write']>) => {
      log.push('write'); return true;
    });
    const input = createWorkerInputCoordinator({
      generation: () => run.generation, capture: () => ({ sheet: 0, target: 'A1' }),
      isReady: () => run.ready, whenReady: () => run.whenHydrated(),
      seal: () => ({}), sync: () => {}, preview: async () => {}, write,
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
    run.recovering = true;
    await run.requestHydration('recovery');
    await input.recover();
    expect(write.mock.calls).toEqual([[
      { generation: 1, sheet: 0, row: 0, col: 0, source: 'cell', value: 'retained' }, expect.any(Function),
    ]]);
    expect(input.unapplied).toEqual([]);
    run.dispose();
    expect(write).toHaveBeenCalledTimes(1);
    expect(log.indexOf('write')).toBeLessThan(log.indexOf('facade'));
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

  test('contains throwing readiness cleanup when the hook replaces its document', async () => {
    const log: string[] = [];
    const old = session(log);
    const next = session(log);
    const opener = open(old.value);
    const { peer, edits } = resources(log);
    const error = new Error('Cleanup failed');
    const errors = mock((_error: Error) => {});
    const readyCleanup = mock(() => { throw error; });
    const commands = createXlsxCommandController().store;
    const controls = bridge();
    const { result, rerender, unmount } = renderHook(
      (props: EditableSessionWorkbookProps) => useEditableSessionWorkbook(props, commands, controls),
      { initialProps: { file, onError: errors, onReady: () => readyCleanup } }
    );
    await waitFor(() => expect(result.current.run).not.toBeNull());
    const first = result.current.run!;
    await act(async () => { first.firstPaint(); await first.whenHydrated(); });
    opener.mockResolvedValue(next.value);
    expect(() => rerender({ file: new Uint8Array([4]), onError: errors, onReady: () => readyCleanup })).not.toThrow();
    await waitFor(() => expect(result.current.run?.session).toBe(next.value));
    expect(readyCleanup).toHaveBeenCalledTimes(1);
    expect(edits.dispose).toHaveBeenCalledTimes(1);
    expect(peer.dispose).toHaveBeenCalledTimes(1);
    expect(old.value.dispose).toHaveBeenCalledTimes(1);
    expect(first.current).toBe(false);
    expect(result.current.error).toBeNull();
    expect(errors).toHaveBeenCalledTimes(1);
    expect(errors).toHaveBeenCalledWith(error);
    unmount();
    expect(next.value.dispose).toHaveBeenCalledTimes(1);
  });

  test('contains a throwing open-error callback and clears loading before replacement', async () => {
    const next = session();
    const failure = new Error('Open failed');
    const opener = open(next.value).mockRejectedValueOnce(failure);
    const errors = mock((_error: Error) => { throw new Error('Callback failed'); });
    const commands = createXlsxCommandController().store;
    const controls = bridge();
    const { result, rerender } = renderHook(
      (props: EditableSessionWorkbookProps) => useEditableSessionWorkbook(props, commands, controls),
      { initialProps: { file, onError: errors } }
    );
    await waitFor(() => {
      expect(result.current.error).toBe(failure);
      expect(result.current.loading).toBe(false);
    });
    expect(result.current.run).toBeNull();
    expect(errors).toHaveBeenCalledTimes(1);
    expect(errors).toHaveBeenCalledWith(failure);
    rerender({ file: new Uint8Array([4]), onError: errors });
    await waitFor(() => expect(result.current.run?.session).toBe(next.value));
    expect(result.current.error).toBeNull();
    expect(result.current.loading).toBe(false);
    expect(opener).toHaveBeenCalledTimes(2);
  });

  test('disposes stale StrictMode opens and publishes readiness once', async () => {
    const old = session();
    const next = session();
    const opener = open(next.value).mockResolvedValueOnce(old.value);
    resources();
    const ready = mock((_api: XlsxWorkerEditorApi) => {});
    const commands = createXlsxCommandController().store;
    const { result } = renderHook(() => useEditableSessionWorkbook({ file, onReady: ready }, commands, bridge()), {
      reactStrictMode: true,
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
    const error = new Error('Cleanup failed');
    const errors = mock((_error: Error) => {});
    const run = owner(value, { onError: errors, onReady: () => () => { throw error; } });
    run.firstPaint();
    await run.whenHydrated();
    log.length = 0;
    expect(() => run.dispose()).not.toThrow();
    expect(log).toEqual(['off', 'facade', 'peer', 'session']);
    expect(errors).toHaveBeenCalledTimes(1);
    expect(errors).toHaveBeenCalledWith(error);
  });
});
