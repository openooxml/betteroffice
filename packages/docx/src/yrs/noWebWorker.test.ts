import { afterEach, beforeEach, expect, test } from 'bun:test';
import { DocxSourceWorkerClient } from './residentSourceWorkerClient';

const BaseWorker = globalThis.Worker;
const builtin = Object.getOwnPropertyDescriptor(process, 'getBuiltinModule');

beforeEach(() => {
  Object.defineProperty(globalThis, 'Worker', { value: undefined, configurable: true, writable: true });
});

afterEach(() => {
  Object.defineProperty(globalThis, 'Worker', { value: BaseWorker, configurable: true, writable: true });
  if (builtin) Object.defineProperty(process, 'getBuiltinModule', builtin);
  else Reflect.deleteProperty(process, 'getBuiltinModule');
});

test('worker sessions without a Web Worker name the in-thread API', () => {
  Object.defineProperty(process, 'getBuiltinModule', { value: undefined, configurable: true });
  expect(() => new DocxSourceWorkerClient()).toThrow('use createYrsSession in-thread or pass the worker option');
});

test('a worker thread that fails to start fails the session with the thread error', async () => {
  const reached = new Error('Node adapter reached');
  let created = 0;
  Object.defineProperty(process, 'getBuiltinModule', { configurable: true, value: (name: string) => {
    expect(name).toBe('node:worker_threads');
    return { Worker: class { constructor() { created += 1; throw reached; } } };
  } });
  const client = new DocxSourceWorkerClient();
  try {
    await expect(client.openSource({ bytes: new Uint8Array([0x50, 0x4b]) })).rejects.toThrow(reached.message);
    expect(created).toBe(1);
  } finally {
    client.destroy();
  }
});
