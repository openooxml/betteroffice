import { afterEach, expect, mock, test } from 'bun:test';
import { createDeferredSessionWorker } from './deferredWorker';

const original = Object.getOwnPropertyDescriptor(process, 'getBuiltinModule');

afterEach(() => {
  if (original) Object.defineProperty(process, 'getBuiltinModule', original);
  else Reflect.deleteProperty(process, 'getBuiltinModule');
});

class FakeWorker extends EventTarget {
  readonly postMessage = mock();
  readonly terminate = mock();
  emit(type: string, data?: unknown): void {
    this.dispatchEvent(Object.assign(new Event(type), { data }));
  }
}

function pending() {
  let resolve!: (worker: Worker) => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<Worker>((done, fail) => { resolve = done; reject = fail; });
  return { promise, resolve, reject };
}

const settle = () => new Promise((resolve) => setTimeout(resolve, 0));

test('queues messages with their transfers until the worker loads, then forwards them in order', async () => {
  const load = pending();
  const worker = createDeferredSessionWorker(() => load.promise, 'openDocument');
  const first = { bytes: new ArrayBuffer(4) };
  const transfer = [first.bytes];
  worker.postMessage(first, transfer);
  worker.postMessage('second');
  const target = new FakeWorker();
  load.resolve(target as unknown as Worker);
  await settle();
  expect(target.postMessage.mock.calls).toEqual([[first, transfer], ['second', undefined]]);
  expect(target.postMessage.mock.calls[0][1]).toBe(transfer);
  worker.postMessage('third');
  expect(target.postMessage.mock.calls[2]).toEqual(['third', undefined]);
});

test('delivers events to listeners and handler properties after loading', async () => {
  const target = new FakeWorker();
  const worker = createDeferredSessionWorker(async () => target as unknown as Worker, 'openDocument');
  const listener = mock();
  const handler = mock();
  worker.addEventListener('message', listener);
  worker.onmessage = handler;
  await settle();
  target.emit('message', 7);
  expect(listener.mock.calls[0][0].data).toBe(7);
  expect(handler.mock.calls[0][0].data).toBe(7);
  worker.removeEventListener('message', listener);
  target.emit('message', 8);
  expect(listener).toHaveBeenCalledTimes(1);
  expect(handler).toHaveBeenCalledTimes(2);
});

test('terminates a worker that loads after termination and drops queued messages', async () => {
  const load = pending();
  const worker = createDeferredSessionWorker(() => load.promise, 'openDocument');
  const listener = mock();
  worker.addEventListener('message', listener);
  worker.postMessage('dropped');
  worker.terminate();
  const target = new FakeWorker();
  load.resolve(target as unknown as Worker);
  await settle();
  expect(target.terminate).toHaveBeenCalledTimes(1);
  expect(target.postMessage).not.toHaveBeenCalled();
  target.emit('message', 1);
  expect(listener).not.toHaveBeenCalled();
});

test('terminates the loaded worker once', async () => {
  const target = new FakeWorker();
  const worker = createDeferredSessionWorker(async () => target as unknown as Worker, 'openDocument');
  await settle();
  worker.terminate();
  worker.terminate();
  expect(target.terminate).toHaveBeenCalledTimes(1);
});

test('reports a load failure as an error event', async () => {
  const load = pending();
  const worker = createDeferredSessionWorker(() => load.promise, 'openDocument');
  const error = mock();
  worker.addEventListener('error', error);
  load.reject(new Error('missing worker file'));
  await settle();
  expect(error.mock.calls[0][0].message).toBe('missing worker file');
  expect(error.mock.calls[0][0].error).toBeInstanceOf(Error);
});

test('throws the in-thread guidance where the runtime has neither Web Worker nor Node threads', () => {
  Object.defineProperty(process, 'getBuiltinModule', { configurable: true, value: undefined });
  const load = mock(async () => new FakeWorker() as unknown as Worker);
  expect(() => createDeferredSessionWorker(load, 'openDocument'))
    .toThrow('use openDocument in-thread or pass the worker option');
  expect(load).not.toHaveBeenCalled();
});
