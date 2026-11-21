import { beforeAll, expect, spyOn, test } from 'bun:test';
import type { SessionTransport } from '../../../../shared/office-session/transport';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import * as workbookWasm from '../wasm/loader';
import { createWorkbookSession, type WorkbookSession } from './client';
import { workbookCollaborationClientLoader } from './collaborationLoader';
import { createWorkbookSessionHost, workbookCollaborationHostLoader } from './host';
import { loadWorkbookSessionFixtures } from './testHelpers';

let bytes: Uint8Array;
beforeAll(async () => { bytes = (await loadWorkbookSessionFixtures()).fixture; });

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}

function sessionTransport() {
  const pair = createInProcessPair();
  let receiveHost: ((message: unknown) => void) | undefined;
  const listeners = new Set<unknown>();
  let closed = false;
  createWorkbookSessionHost({ ...pair.host, listen(listener) {
    receiveHost = listener;
    return pair.host.listen(listener);
  } });
  const transport: SessionTransport = { ...pair.client,
    listen(listener) {
      listeners.add(listener);
      const off = pair.client.listen(listener);
      return () => { listeners.delete(listener); off(); };
    },
    onError(listener) {
      listeners.add(listener);
      const off = pair.client.onError(listener);
      return () => { listeners.delete(listener); off(); };
    },
    close() {
      receiveHost?.({ protocol: 1, kind: 'dispose' });
      pair.client.close();
      closed = true;
    },
  };
  return { transport, listeners, get closed() { return closed; } };
}

test('an ordinary open never imports worker or main collaboration support', async () => {
  const workerLoad = spyOn(workbookCollaborationHostLoader, 'load').mockImplementation(() => {
    throw new Error('Unexpected worker collaboration import');
  });
  const clientLoad = spyOn(workbookCollaborationClientLoader, 'load').mockImplementation(() => {
    throw new Error('Unexpected main collaboration import');
  });
  const pair = sessionTransport();
  let session: WorkbookSession | undefined;
  try {
    session = await createWorkbookSession(bytes, { includeEditState: true }, pair.transport);
    expect(session.state.stage).toBe('ready');
    expect(session.replica).toBeUndefined();
    expect((await session.call.sheets()).length).toBeGreaterThan(0);
    expect(workerLoad).not.toHaveBeenCalled();
    expect(clientLoad).not.toHaveBeenCalled();
  } finally {
    await session?.dispose();
    pair.transport.close();
    workerLoad.mockRestore(); clientLoad.mockRestore();
  }
});

test('collaborative open waits for both chunks and preserves initial and subsequent updates', async () => {
  const workerModule = await workbookCollaborationHostLoader.load();
  const clientModule = await workbookCollaborationClientLoader.load();
  const workerChunk = deferred<typeof workerModule>();
  const clientChunk = deferred<typeof clientModule>();
  const workerStarted = deferred<void>();
  const clientStarted = deferred<void>();
  const workerLoad = spyOn(workbookCollaborationHostLoader, 'load').mockImplementation(() => {
    workerStarted.resolve(); return workerChunk.promise;
  });
  const clientLoad = spyOn(workbookCollaborationClientLoader, 'load').mockImplementation(() => {
    clientStarted.resolve(); return clientChunk.promise;
  });
  const peer = workbookWasm.openWorkbook(bytes, { collaborative: true, clientId: 9101 });
  peer.editCell(0, 2, 1, 'initial');
  const initialUpdate = peer.encodeStateAsUpdate();
  const opened = spyOn(workbookWasm, 'openWorkbook');
  const pair = sessionTransport();
  let session: WorkbookSession | undefined;
  let ready = false;
  const opening = createWorkbookSession(bytes, { collaborative: true, clientId: 9102, initialUpdate,
    initialViewport: { width: 800, height: 600, zoom: 1, dpr: 1, requestGeneration: 0 } }, pair.transport)
    .then((value) => { ready = true; session = value; return value; });
  try {
    await workerStarted.promise;
    expect(opened).not.toHaveBeenCalled();
    expect(clientLoad).not.toHaveBeenCalled();
    workerChunk.resolve(workerModule);
    await clientStarted.promise;
    expect(opened).toHaveBeenCalledTimes(1);
    expect(ready).toBe(false);
    clientChunk.resolve(clientModule);
    session = await opening;
    expect(session.replica).toBeDefined();
    expect(session.initialPaint?.frame.displayList.commands.some((command) => command.op === 'text' && command.text === 'initial')).toBe(true);
    expect((await session.call.cellInputs(0, 'B3')).cells[0][0].input).toBe('initial');
    const vector = await session.replica!.async.encodeStateVector();
    peer.editCell(0, 2, 1, 'remote');
    await session.replica!.async.applyUpdate(peer.encodeStateAsUpdate(vector));
    expect((await session.call.cellInputs(0, 'B3')).cells[0][0].input).toBe('remote');
    expect(workerLoad).toHaveBeenCalledTimes(1);
    expect(clientLoad).toHaveBeenCalledTimes(1);
  } finally {
    workerChunk.resolve(workerModule); clientChunk.resolve(clientModule);
    await opening.catch(() => {});
    await session?.dispose(); pair.transport.close(); peer.dispose();
    opened.mockRestore(); workerLoad.mockRestore(); clientLoad.mockRestore();
  }
});

for (const side of ['worker', 'client'] as const) {
  test(`aborting during the ${side} collaboration import closes the session without late setup`, async () => {
    const workerModule = await workbookCollaborationHostLoader.load();
    const clientModule = await workbookCollaborationClientLoader.load();
    const workerChunk = deferred<typeof workerModule>();
    const clientChunk = deferred<typeof clientModule>();
    const started = deferred<void>();
    const workerLoad = spyOn(workbookCollaborationHostLoader, 'load').mockImplementation(() => {
      if (side === 'worker') { started.resolve(); return workerChunk.promise; }
      return Promise.resolve(workerModule);
    });
    const clientLoad = spyOn(workbookCollaborationClientLoader, 'load').mockImplementation(() => {
      started.resolve(); return clientChunk.promise;
    });
    const createPipe = spyOn(workerModule, 'createWorkbookCollaborationHost');
    const createReplica = spyOn(clientModule, 'createAsyncCollaborationReplica');
    const openWorkbook = workbookWasm.openWorkbook;
    let disposedWorkbooks = 0;
    const opened = spyOn(workbookWasm, 'openWorkbook').mockImplementation((...args) => {
      const handle = openWorkbook(...args);
      const dispose = handle.dispose.bind(handle);
      handle.dispose = () => { disposedWorkbooks++; dispose(); };
      return handle;
    });
    const pair = sessionTransport();
    const controller = new AbortController();
    const opening = createWorkbookSession(bytes, { collaborative: true, signal: controller.signal }, pair.transport);
    const outcome = opening.catch((error: unknown) => error);
    try {
      await started.promise;
      controller.abort();
      expect(await outcome).toMatchObject({ code: 'disposed' });
      expect(pair.closed).toBe(true);
      expect(pair.listeners.size).toBe(0);
      expect(disposedWorkbooks).toBe(side === 'worker' ? 0 : 1);
      const calls = createPipe.mock.calls.length;
      expect(calls).toBe(side === 'worker' ? 0 : 1);
      workerChunk.resolve(workerModule); clientChunk.resolve(clientModule);
      await new Promise<void>((resolve) => setTimeout(resolve, 0));
      expect(createPipe).toHaveBeenCalledTimes(calls);
      expect(createReplica).not.toHaveBeenCalled();
      expect(opened).toHaveBeenCalledTimes(side === 'worker' ? 0 : 1);
    } finally {
      workerChunk.resolve(workerModule); clientChunk.resolve(clientModule);
      await opening.then((session) => session.dispose(), () => {});
      pair.transport.close();
      createPipe.mockRestore(); createReplica.mockRestore(); opened.mockRestore();
      workerLoad.mockRestore(); clientLoad.mockRestore();
    }
  });
}

test('a failed client chunk disposes the opened worker and releases its listeners', async () => {
  const load = spyOn(workbookCollaborationClientLoader, 'load').mockRejectedValue(new Error('Chunk unavailable'));
  const pair = sessionTransport();
  try {
    await expect(createWorkbookSession(bytes, { collaborative: true }, pair.transport)).rejects.toThrow('Chunk unavailable');
    expect(pair.closed).toBe(true);
    expect(pair.listeners.size).toBe(0);
  } finally { pair.transport.close(); load.mockRestore(); }
});
