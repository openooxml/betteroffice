import { afterEach, beforeEach, expect, test } from 'bun:test';
import { openPresentationSession } from './client';
import { createPptxWorkerEditorSession } from './editorSession';
import { preparePresentationFontDecoder } from './fontWorker';
import { createPptxWorkerSession } from './commandSession';

const BaseWorker = globalThis.Worker;
const builtin = Object.getOwnPropertyDescriptor(process, 'getBuiltinModule');
const message = 'use openPresentation in-thread or pass the worker option';

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
  await expect(openPresentationSession(new Uint8Array())).rejects.toThrow(message);
  expect(() => createPptxWorkerEditorSession(new Uint8Array())).toThrow(message);
  expect(() => preparePresentationFontDecoder(new Uint8Array())).toThrow(message);
  expect(() => createPptxWorkerSession(new Uint8Array())).toThrow(message);
});

test('a worker thread that fails to start fails the session with the thread error', async () => {
  const reached = new Error('Node adapter reached');
  let created = 0;
  Object.defineProperty(process, 'getBuiltinModule', { configurable: true, value: (name: string) => {
    expect(name).toBe('node:worker_threads');
    return { Worker: class { constructor() { created += 1; throw reached; } } };
  } });
  await expect(openPresentationSession(new Uint8Array())).rejects.toThrow(reached.message);
  await expect(createPptxWorkerEditorSession(new Uint8Array()).whenHydrated()).rejects.toThrow(reached.message);
  const decoder = preparePresentationFontDecoder(new Uint8Array());
  await expect(createPptxWorkerSession(new Uint8Array()).ready).rejects.toThrow(reached.message);
  decoder.release();
  expect(created).toBe(4);
});
