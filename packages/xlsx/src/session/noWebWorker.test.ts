import { afterEach, beforeEach, expect, test } from 'bun:test';
import { openWorkbookSession } from './client';
import { workbookRecoveryBackend } from './clientRecovery';

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

test('worker sessions without a Web Worker name the in-thread API', async () => {
  Object.defineProperty(process, 'getBuiltinModule', { value: undefined, configurable: true });
  await expect(openWorkbookSession(new Uint8Array())).rejects.toThrow('use openWorkbook in-thread or pass the worker option');
  await expect(workbookRecoveryBackend.open(new Uint8Array(), {})).rejects.toThrow('use openWorkbook in-thread');
});

test('worker sessions without a Web Worker use worker_threads when available', async () => {
  const reached = new Error('Node adapter reached');
  let created = 0;
  Object.defineProperty(process, 'getBuiltinModule', { configurable: true, value: (name: string) => {
    expect(name).toBe('node:worker_threads');
    return { Worker: class { constructor() { created += 1; throw reached; } } };
  } });
  await expect(openWorkbookSession(new Uint8Array())).rejects.toBe(reached);
  await expect(workbookRecoveryBackend.open(new Uint8Array(), {})).rejects.toBe(reached);
  expect(created).toBe(2);
});
