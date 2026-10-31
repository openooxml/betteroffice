import { afterEach, beforeEach, expect, mock, test } from 'bun:test';
import { createSessionWorker, sessionWorkerScope } from './nodeWorker';
import { deferSessionWorker } from './webWorker';

const original = Object.getOwnPropertyDescriptor(process, 'getBuiltinModule');
const threads: FakeThread[] = [];
const scope = { postMessage: mock(), addEventListener: mock(), removeEventListener: mock() };

class FakeThread {
  readonly listeners = new Map<string, (value: unknown) => void>();
  readonly postMessage = mock();
  readonly terminate = mock();
  constructor(readonly url: URL, readonly options: unknown) { threads.push(this); }
  on(type: string, listener: (value: unknown) => void): void { this.listeners.set(type, listener); }
  emit(type: string, value: unknown): void { this.listeners.get(type)?.(value); }
}

beforeEach(() => {
  threads.length = 0;
  Object.defineProperty(process, 'getBuiltinModule', { configurable: true, value: mock((name: string) => {
    expect(name).toBe('node:worker_threads');
    return { Worker: FakeThread, parentPort: scope };
  }) });
});

afterEach(() => {
  if (original) Object.defineProperty(process, 'getBuiltinModule', original);
  else Reflect.deleteProperty(process, 'getBuiltinModule');
});

const url = new URL('file:///session.mjs');
const open = () => createSessionWorker(url, { type: 'module', name: 'session' }, 'openDocument');

test('constructs a thread and passes messages, transfers and termination through', () => {
  const worker = open();
  const message = { bytes: new ArrayBuffer(8) };
  const transfer = [message.bytes];
  worker.postMessage(message, transfer);
  expect(threads[0].url).toBe(url);
  expect(threads[0].options).toEqual({ name: 'session' });
  expect(threads[0].postMessage).toHaveBeenCalledWith(message, transfer);
  expect(threads[0].postMessage.mock.calls[0][1]).toBe(transfer);
  worker.terminate();
  worker.terminate();
  expect(threads[0].terminate).toHaveBeenCalledTimes(1);
});

test('stops delivering events after termination', () => {
  const worker = open();
  const listener = mock();
  const handler = mock();
  worker.addEventListener('message', listener);
  worker.onmessage = handler;
  worker.terminate();
  threads[0].emit('message', 'late');
  expect(listener).not.toHaveBeenCalled();
  expect(handler).not.toHaveBeenCalled();
});

test('maps events and removes listeners without disturbing other listeners', () => {
  const worker = open();
  const message = mock();
  const second = { handleEvent: mock() };
  const error = mock();
  const messageerror = mock();
  worker.addEventListener('message', message);
  worker.addEventListener('message', message);
  worker.addEventListener('message', second);
  worker.addEventListener('error', error);
  worker.addEventListener('messageerror', messageerror);
  threads[0].emit('message', { value: 7 });
  expect(message).toHaveBeenCalledTimes(1);
  expect(message).toHaveBeenCalledWith({ data: { value: 7 } });
  worker.removeEventListener('message', message);
  threads[0].emit('message', 8);
  expect(message).toHaveBeenCalledTimes(1);
  expect(second.handleEvent).toHaveBeenLastCalledWith({ data: 8 });
  const failure = new Error('Thread failed');
  threads[0].emit('error', failure);
  threads[0].emit('messageerror', failure);
  expect(error).toHaveBeenCalledWith({ error: failure, message: 'Thread failed' });
  expect(messageerror).toHaveBeenCalledWith({ error: failure, message: 'Thread failed' });
});

test('replaces and clears event handlers used by the document client', () => {
  const worker = open();
  const previous = mock();
  const message = mock();
  const error = mock();
  const messageerror = mock();
  worker.onmessage = previous;
  worker.onmessage = message;
  worker.onerror = error;
  worker.onmessageerror = messageerror;
  threads[0].emit('message', 9);
  threads[0].emit('error', new Error('broken'));
  threads[0].emit('messageerror', new Error('unreadable'));
  expect(previous).not.toHaveBeenCalled();
  expect(message).toHaveBeenCalledWith({ data: 9 });
  expect(error.mock.calls[0][0].message).toBe('broken');
  expect(messageerror.mock.calls[0][0].message).toBe('unreadable');
  worker.onmessage = null;
  worker.onerror = null;
  worker.onmessageerror = null;
  for (const type of ['message', 'error', 'messageerror']) threads[0].emit(type, 10);
  for (const handler of [message, error, messageerror]) expect(handler).toHaveBeenCalledTimes(1);
});

test('uses parentPort as the worker scope', () => {
  expect(sessionWorkerScope()).toBe(scope);
});

test('reports the existing in-thread guidance without getBuiltinModule', () => {
  Object.defineProperty(process, 'getBuiltinModule', { configurable: true, value: undefined });
  expect(open).toThrow('use openDocument in-thread or pass the worker option');
  expect(() => deferSessionWorker(async () => open(), 'openDocument')).toThrow('use openDocument in-thread');
  expect(sessionWorkerScope).toThrow('Worker sessions need a Web Worker');
});

test('buffers posts and subscriptions while the adapter loads', async () => {
  let resolve!: (worker: Worker) => void;
  const pending = new Promise<Worker>(done => { resolve = done; });
  const worker = deferSessionWorker(() => pending, 'openDocument');
  const message = mock();
  const handler = mock();
  const bytes = new ArrayBuffer(4);
  worker.addEventListener('message', message);
  worker.onmessage = handler;
  worker.postMessage({ bytes }, [bytes]);
  resolve(open());
  await pending;
  expect(threads[0].postMessage).toHaveBeenCalledWith({ bytes }, [bytes]);
  threads[0].emit('message', 11);
  expect(message).toHaveBeenCalledWith({ data: 11 });
  expect(handler).toHaveBeenCalledWith({ data: 11 });
  worker.removeEventListener('message', message);
  threads[0].emit('message', 12);
  expect(message).toHaveBeenCalledTimes(1);
  worker.terminate();
  expect(threads[0].terminate).toHaveBeenCalledTimes(1);
});

test('terminates a worker disposed before its import completes', async () => {
  let resolve!: (worker: Worker) => void;
  const pending = new Promise<Worker>(done => { resolve = done; });
  const worker = deferSessionWorker(() => pending, 'openDocument');
  worker.postMessage('discard');
  worker.terminate();
  resolve(open());
  await pending;
  expect(threads[0].postMessage).not.toHaveBeenCalled();
  expect(threads[0].terminate).toHaveBeenCalledTimes(1);
});

test('reports asynchronous creation failures through listeners and setters', async () => {
  const failure = new Error('Cannot load thread');
  const worker = deferSessionWorker(() => Promise.reject(failure), 'openDocument');
  const error = mock();
  const handler = mock();
  worker.addEventListener('error', error);
  worker.onerror = handler;
  await Promise.resolve();
  await Promise.resolve();
  expect(error).toHaveBeenCalledWith({ error: failure, message: failure.message });
  expect(handler).toHaveBeenCalledWith({ error: failure, message: failure.message });
  worker.terminate();
});
