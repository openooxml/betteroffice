import { expect, mock, test } from 'bun:test';
import { isClientMessage, type ClientMessage } from '../../../../shared/office-session/protocol';
import { createPptxWorkerSession } from './commandSession';
import type { PresentationCommandOpenOptions, PresentationCommandOpenResult } from './commandProtocol';
import { createPresentationCommandProjection } from './commandSharedAdapter';

function heldWorker() {
  const events = new EventTarget();
  const listeners = new Map<string, Set<EventListenerOrEventListenerObject>>();
  const calls: Extract<ClientMessage, { kind: 'call' }>[] = [];
  const arrivals = new Map<string, () => void>();
  const terminate = mock(() => {});
  const worker = {
    postMessage(message: unknown) {
      if (!isClientMessage(message) || message.kind !== 'call') return;
      calls.push(message);
      arrivals.get(message.method)?.();
      if (message.method !== 'openCommandEditor') return;
      const input = message.args[1] as PresentationCommandOpenOptions;
      const state: PresentationCommandOpenResult = {
        session: input.session, incarnation: input.incarnation, epoch: 1, throughSeq: 0, version: 'v0',
        selection: null, outcomes: [], dirtySlideIds: [], activeSlideId: 'slide',
        slides: [{ id: 'slide', index: 0, name: null, layoutPartPath: null }],
        size: { width: 960, height: 540 }, frame: null,
        editState: { readOnly: false, canUndo: false, canRedo: false, toolbar: { target: null,
          storedMarks: {}, formatting: null, alignment: null, hyperlink: null, table: null, canEditText: false } },
      };
      const projection = createPresentationCommandProjection();
      projection.initialize(state.version);
      queueMicrotask(() => events.dispatchEvent(new MessageEvent('message', { data: {
        protocol: 1, kind: 'reply', id: message.id, ok: true, value: { ...state, shared: projection.ack(state) },
      } })));
    },
    addEventListener(type: string, listener: EventListenerOrEventListenerObject) {
      const active = listeners.get(type) ?? new Set<EventListenerOrEventListenerObject>();
      listeners.set(type, active);
      active.add(listener);
      events.addEventListener(type, listener);
    },
    removeEventListener(type: string, listener: EventListenerOrEventListenerObject) {
      listeners.get(type)?.delete(listener);
      events.removeEventListener(type, listener);
    },
    terminate,
  } as unknown as Worker;
  return { worker, terminate,
    listenerCount: () => [...listeners.values()].reduce((count, active) => count + active.size, 0),
    waitFor(method: string) {
      if (calls.some((call) => call.method === method)) return Promise.resolve();
      return new Promise<void>((resolve) => { arrivals.set(method, resolve); });
    },
  };
}

for (const method of ['edit', 'readEditor', 'thumbnailFrame'] as const) {
  test(`disposing with a pending ${method} call does not leak a rejection`, async () => {
    const endpoint = heldWorker();
    const onError = mock(() => {});
    const failure = mock(() => {});
    const unhandled = mock((_reason: unknown) => {});
    process.on('unhandledRejection', unhandled);
    const session = createPptxWorkerSession(new Uint8Array(), { worker: () => endpoint.worker, onError });
    let pending: Promise<unknown> | undefined;
    try {
      const client = await session.ready;
      session.subscribeFailure(failure);
      pending = method === 'edit'
        ? client.dispatch({ kind: 'setSelection', payload: { selection: null, inputTime: 1, undoGroup: 'selection' } })
        : method === 'readEditor' ? client.read() : client.thumbnail('slide');
      await endpoint.waitFor(method);
      const disposal = session.dispose();
      expect(session.dispose()).toBe(disposal);
      await disposal;
      await new Promise<void>((resolve) => setTimeout(resolve, 0));
      expect(unhandled).not.toHaveBeenCalled();
      expect(await pending.then(() => undefined, (error: unknown) => error))
        .toMatchObject({ name: 'SessionFailure', code: 'disposed' });
      expect(session.failure).toBeUndefined();
      expect(failure).not.toHaveBeenCalled();
      expect(onError).not.toHaveBeenCalled();
      expect(endpoint.terminate).toHaveBeenCalledTimes(1);
      expect(endpoint.listenerCount()).toBe(0);
    } finally {
      void pending?.catch(() => {});
      await session.dispose();
      process.off('unhandledRejection', unhandled);
    }
  });
}

test('disposing before open settles ready and removes connection listeners', async () => {
  const endpoint = heldWorker();
  const unhandled = mock((_reason: unknown) => {});
  process.on('unhandledRejection', unhandled);
  const session = createPptxWorkerSession(new Uint8Array(), { worker: () => endpoint.worker });
  try {
    await session.dispose();
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
    expect(unhandled).not.toHaveBeenCalled();
    await expect(session.ready).rejects.toMatchObject({ name: 'PptxWorkerEditorDisposedError' });
    expect(endpoint.terminate).toHaveBeenCalledTimes(1);
    expect(endpoint.listenerCount()).toBe(0);
  } finally {
    await session.dispose();
    process.off('unhandledRejection', unhandled);
  }
});
