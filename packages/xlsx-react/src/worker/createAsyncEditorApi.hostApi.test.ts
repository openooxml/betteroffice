import { expect, mock, test } from 'bun:test';
import { createXlsxCommandController, XlsxCommandAdmissionError } from '../commands/createXlsxCommandStore';
import type { PeerlessWorkerInputCoordinator } from '../commands/peerlessWorkerInputCoordinator';
import { createPeerlessPluginDocument } from '../plugins/createPeerlessPluginDocument';
import { createAsyncEditorApi, type XlsxAsyncExportBridge } from './createAsyncEditorApi';
import type { WorkerEditorApiBridge } from './createWorkerEditorApi';
import type { PeerlessWorkbookSession } from './usePeerlessSessionWorkbook';

const methods = ['version', 'readCells', 'findText', 'validateEdits', 'applyEdits', 'save'] as const;
function harness() {
  let release!: () => void;
  const pending = new Promise<void>((resolve) => { release = resolve; });
  const events: string[] = [];
  const state = { refused: false, readOnly: false };
  const bytes = new Uint8Array([1, 2, 3]);
  const request = { expectVersion: 'v0', steps: [] };
  const result = { ok: true, version: 'v1' };
  const call = Object.fromEntries(['version', 'readCells', 'findText', 'validateEdits'].map((method) =>
    [method, mock(async () => {
      events.push(method);
      return method === 'version' ? 'v1' : method === 'validateEdits'
        ? { ok: true, baseVersion: 'v1', wouldApply: true, previews: [] } : result;
    })]));
  const session = { current: true, retiring: false, ready: true, failure: null,
    session: { call, editState: {}, confirmedBasis: { version: 'v1' } } } as unknown as PeerlessWorkbookSession;
  const input = {
    async runAfterPendingInput<T>(operation: () => T) {
      await pending;
      if (state.refused) throw new XlsxCommandAdmissionError('input-failed');
      events.push('draft');
      return operation();
    },
    command: mock(() => { events.push('applyEdits'); return { result: Promise.resolve({ result }) }; }),
    async save() { return input.runAfterPendingInput(() => { events.push('save'); return bytes; }); },
  } as unknown as PeerlessWorkerInputCoordinator;
  const document = createPeerlessPluginDocument({ owner: () => session, coordinator: () => input,
    readOnly: () => state.readOnly, commands: () => null, selection: () => null!, navigator: {} as never });
  const saved = mock((_bytes: Uint8Array) => {});
  const bridge = { editorDocument: () => document, exportCoordinator: () => input, saved } as unknown as WorkerEditorApiBridge & XlsxAsyncExportBridge;
  const api = createAsyncEditorApi(session, createXlsxCommandController().store, () => bridge);
  const invoke = (method: typeof methods[number]) => Reflect.apply(Reflect.get(api, method), api, ['version', 'save'].includes(method) ? [] : [request]);
  return { api, session, state, release, events, invoke, saved, bytes, request, call, input };
}
for (const method of methods) {
  test(`host ${method} waits for pending input`, async () => {
    const h = harness();
    const work = h.invoke(method);
    expect(h.events).toEqual([]);
    h.release();
    await work;
    expect(h.events).toEqual(['draft', method]);
    if (method === 'save') {
      expect(await work).toBe(h.bytes);
      expect(h.saved).toHaveBeenCalledTimes(1);
      expect(h.saved).toHaveBeenCalledWith(h.bytes);
    }
  });
  for (const code of ['input-failed', 'document-replaced'] as const) {
    test(`host ${method} rejects ${code} during pending input`, async () => {
      const h = harness();
      const work = h.invoke(method);
      if (code === 'input-failed') h.state.refused = true;
      else Object.defineProperty(h.session, 'current', { value: false });
      h.release();
      await expect(work).rejects.toMatchObject({ code });
      expect(h.saved).not.toHaveBeenCalled();
    });
  }
  test(`host ${method} checks hydration and session failure`, async () => {
    const h = harness();
    Object.defineProperty(h.session, 'ready', { value: false, configurable: true });
    await expect(h.invoke(method)).rejects.toMatchObject({ code: 'editor-unavailable' });
    const failure = new Error('Session failed');
    Object.defineProperty(h.session, 'failure', { value: failure });
    await expect(h.invoke(method)).rejects.toBe(failure);
  });
}
for (const method of ['validateEdits', 'applyEdits'] as const) {
  test(`host ${method} refuses read-only`, async () => {
    const h = harness();
    h.state.readOnly = true;
    h.release();
    await expect(h.invoke(method)).resolves.toMatchObject({ ok: false, failure: { code: 'read-only' } });
  });
}
test('host requests are captured before input admission', async () => {
  const h = harness();
  const work = h.invoke('readCells');
  h.request.expectVersion = 'changed';
  h.release();
  await work;
  expect(h.call.readCells).toHaveBeenCalledWith({ expectVersion: 'v0', steps: [] });
});

for (const method of ['version', 'readCells', 'findText', 'validateEdits'] as const) {
  test(`host ${method} checks replacement after a read reply`, async () => {
    const h = harness();
    h.call[method]!.mockImplementation(async () => {
      await Promise.resolve();
      Object.defineProperty(h.session, 'current', { value: false });
      throw new Error('Read failed');
    });
    h.release();
    await expect(h.invoke(method)).rejects.toMatchObject({ code: 'document-replaced' });
  });
}
