import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
import type { DocxCommandAck, DocxCommandFrame, ResidentCommandConnection } from '@betteroffice/docx/peerless';
import { SessionFailure } from '../../../../../../shared/office-session/types';
import { usePeerlessEditSession, type UsePeerlessEditSessionOptions } from './usePeerlessEditSession';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook, waitFor } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => {
  await act(async () => { await new Promise<void>((resolve) => setImmediate(resolve)); });
  if (ownsDom) await GlobalRegistrator.unregister();
});

function stagedFrame(): DocxCommandFrame {
  const bitmap = () => ({ width: 1, height: 1, close: mock(() => {}) }) as unknown as ImageBitmap;
  return { coversThroughSeq: 0, revision: 0, payload: {
    frame: new ArrayBuffer(0), layoutRevision: 1, caret: { frameEpoch: 0, caretRect: null },
    stagedPresentation: { frameEpoch: 0, pages: [{
      page: { pageId: '1', pageIndex: 0 }, frameEpoch: 0, bitmap: bitmap(),
      backgrounds: [{ id: 'background', rect: { x: 0, y: 0, width: 1, height: 1 }, bitmap: bitmap() }],
    }] },
  } };
}

function resident(incarnation = 1) {
  const initial: ResidentCommandConnection['initial'] = {
    session: 'lifecycle', incarnation, epoch: 1, version: 0, throughSeq: 0, outcomes: [], selection: null,
    editState: { readOnly: false, toolbar: null, effectiveFormatting: null, storedMarks: null,
      canUndo: false, canRedo: false, hyperlink: null, tableContext: null }, frame: stagedFrame(),
  };
  const dispose = mock(async () => {});
  const failures = new Set<(failure: SessionFailure) => void>();
  const connection: ResidentCommandConnection = {
    initial, requestFullFrame: async () => initial, requestViewport: mock(async () => {}),
    client: { failure: undefined, dispose,
      onFailure: (listener) => { failures.add(listener); return () => failures.delete(listener); },
      call: { edit: async () => initial,
        checkpoint: async () => ({ state: new Uint8Array(), epoch: 1, version: 0, throughSeq: 0 }) },
    },
  };
  const options: UsePeerlessEditSessionOptions = {
    source: Uint8Array.of(1), openConnection: async () => connection,
    canvasForPage: () => null, prepareGeometry: () => undefined,
  };
  const crash = (failure: SessionFailure) => { for (const listener of [...failures]) listener(failure); };
  return { connection, dispose, options, crash };
}

function expectClosed(frame: DocxCommandFrame) {
  const page = frame.payload.stagedPresentation!.pages[0]!;
  expect(page.bitmap.close).toHaveBeenCalledTimes(1);
  expect(page.backgrounds![0]!.bitmap.close).toHaveBeenCalledTimes(1);
}

test('opening reads the latest raster scale before command initialization', async () => {
  const { connection, options } = resident();
  const scales: (number | undefined)[] = [];
  let finish!: () => void;
  connection.requestFullFrame = () => new Promise(() => {});
  options.openConnection = (...args: unknown[]) => new Promise((resolve) => {
    finish = () => {
      scales.push(typeof args[2] === 'function' ? (args[2] as () => number)() : undefined);
      resolve(connection);
    };
  });
  const view = renderHook(({ rasterScale }) => usePeerlessEditSession({ ...options, rasterScale } as UsePeerlessEditSessionOptions),
    { initialProps: { rasterScale: 2 } });
  view.rerender({ rasterScale: 3 });
  await act(async () => { finish(); });
  expect(scales).toEqual([3]);
  expect(view.result.current.status).toBe('ready');
});

test('unmounting during connection open releases the unused initial frame', async () => {
  const { connection, dispose, options } = resident();
  let resolve!: (value: ResidentCommandConnection) => void;
  options.openConnection = () => new Promise((yes) => { resolve = yes; });
  const view = renderHook(() => usePeerlessEditSession(options));
  view.unmount();
  await act(async () => { resolve(connection); });
  expectClosed(connection.initial.frame);
  expect(dispose).toHaveBeenCalledTimes(1);
});

test('a presentation that keeps failing stops requesting frames, reports once without failing and re-arms on viewport changes', async () => {
  const { connection, options } = resident();
  let calls = 0;
  const request = mock(() => ++calls > 8 ? new Promise<DocxCommandAck>(() => {})
    : Promise.resolve({ ...connection.initial, frame: stagedFrame() }));
  connection.requestFullFrame = request;
  const onError = mock((_error: Error) => {});
  const onRecoverableError = mock((_error: Error) => {});
  const view = renderHook(() => usePeerlessEditSession({ ...options, onError, onRecoverableError }));
  const settle = () => act(async () => { await new Promise<void>((resolve) => setImmediate(resolve)); });
  await waitFor(() => expect(view.result.current.status).toBe('ready'));
  await settle();
  expect(request).toHaveBeenCalledTimes(2);
  expect(onError).not.toHaveBeenCalled();
  expect(onRecoverableError.mock.calls.map(([error]) => error.message))
    .toEqual(['Full frame responses did not satisfy requested coverage and revision']);
  const binding = view.result.current.binding!;
  await act(async () => { binding.setRasterScale!(2); });
  await settle();
  expect(request).toHaveBeenCalledTimes(3);
  await act(async () => { binding.setDisplayWindow!(0, 2); });
  await settle();
  expect(request).toHaveBeenCalledTimes(4);
  expect(onRecoverableError).toHaveBeenCalledTimes(1);
  expect(onError).not.toHaveBeenCalled();
  expect(view.result.current.status).toBe('ready');
});

test('a requested frame arriving after unmount releases its transferred bitmaps', async () => {
  const { connection, options } = resident();
  let resolve!: (value: DocxCommandAck) => void;
  const request = mock(() => new Promise<DocxCommandAck>((yes) => { resolve = yes; }));
  connection.requestFullFrame = request;
  const view = renderHook(() => usePeerlessEditSession(options));
  await waitFor(() => expect(view.result.current.status).toBe('ready'));
  expect(request).toHaveBeenCalledTimes(1);
  view.unmount();
  const frame = stagedFrame();
  await act(async () => { resolve({ ...connection.initial, frame }); });
  expectClosed(frame);
});

test('a crashed worker respawns once with the viewport re-sent and a second crash fails the editor through onError', async () => {
  const first = resident();
  let rejectViewport!: (error: Error) => void;
  first.connection.requestViewport = mock(() => new Promise<void>((_, reject) => { rejectViewport = reject; }));
  const respawns: ReturnType<typeof resident>[] = [];
  const reopen = mock(async (input: { incarnation: number }) => {
    const next = resident(input.incarnation);
    respawns.push(next);
    return next.connection;
  });
  const onError = mock((_error: Error) => {});
  const phases: string[] = [];
  const view = renderHook(() => usePeerlessEditSession({ ...first.options, onError,
    recovery: { reopen, onRecovery: (event) => { phases.push(event.phase); } } }));
  await waitFor(() => expect(view.result.current.status).toBe('ready'));
  const binding = view.result.current.binding!;
  await act(async () => { binding.setDisplayWindow!(4, 9); });
  const failure = new SessionFailure('crash', 'Worker crashed');
  await act(async () => { first.crash(failure); rejectViewport(failure); });
  await waitFor(() => expect(phases).toEqual(['recovering', 'recovered']));
  expect(onError).not.toHaveBeenCalled();
  expect(reopen).toHaveBeenCalledTimes(1);
  expect(reopen.mock.calls[0]![0].incarnation).toBe(2);
  expect(binding.viewStore.getSnapshot().identity.incarnation).toBe(2);
  expect(respawns[0]!.connection.requestViewport).toHaveBeenCalledWith(expect.objectContaining({ start: 4, end: 9 }));
  const second = new SessionFailure('crash', 'Worker crashed again');
  await act(async () => { respawns[0]!.crash(second); });
  await waitFor(() => expect(view.result.current.status).toBe('failed'));
  expect(reopen).toHaveBeenCalledTimes(1);
  expect(onError.mock.calls).toEqual([[second]]);
  expect(view.result.current.error).toBe(second);
  expect(phases).toEqual(['recovering', 'recovered', 'recovering', 'failed']);
});

for (const reason of ['stale-version', 'stale-anchor', 'read-only']) {
  test(`recovery fails loudly when replay refuses acknowledged work with ${reason}`, async () => {
    const first = resident();
    first.connection.client.call.edit = async command => ({ ...first.connection.initial, frame: undefined,
      version: 1, throughSeq: command.seq, outcomes: [{ seq: command.seq, status: 'applied' }] });
    const next = resident(2);
    next.connection.client.call.edit = async command => ({ ...next.connection.initial, frame: undefined,
      throughSeq: command.seq, outcomes: [{ seq: command.seq, status: 'refused', reason }] });
    const phases: string[] = [];
    const onError = mock((_error: Error) => {});
    const view = renderHook(() => usePeerlessEditSession({ ...first.options, onError,
      recovery: { reopen: async () => next.connection, onRecovery: event => { phases.push(event.phase); } } }));
    await waitFor(() => expect(view.result.current.status).toBe('ready'));
    await act(async () => { await view.result.current.binding!.session.submit('insertText', { text: 'kept' }).ack; });
    act(() => first.crash(new SessionFailure('crash', 'Worker stopped')));
    await waitFor(() => expect(onError).toHaveBeenCalledTimes(1));
    expect(onError.mock.calls[0]![0]).toBeInstanceOf(SessionFailure);
    expect(onError.mock.calls[0]![0].message).toContain(reason);
    expect(phases).not.toContain('recovered');
    expect(view.result.current.status).toBe('failed');
  });
}

test('replacement source OOM preserves its classification and memory diagnostics', async () => {
  const first = resident();
  const { ResidentWorkerOutOfMemoryError } = await import('@betteroffice/docx/yrs');
  const memory = [{ label: 'docx-edit', bufferBytes: 2048 }];
  const failure = new ResidentWorkerOutOfMemoryError('Replacement source allocation failed', memory);
  const onError = mock((_error: Error) => {});
  const phases: string[] = [];
  const view = renderHook(() => usePeerlessEditSession({ ...first.options, onError,
    recovery: { reopen: async () => { throw failure; }, onRecovery: event => { phases.push(event.phase); } } }));
  await waitFor(() => expect(view.result.current.status).toBe('ready'));
  act(() => first.crash(new SessionFailure('crash', 'Worker stopped')));
  await waitFor(() => expect(onError).toHaveBeenCalledTimes(1));
  expect(onError.mock.calls[0]![0]).toMatchObject({ code: 'out-of-memory', diagnostics: JSON.stringify(memory) });
  expect(phases).not.toContain('recovered');
});
