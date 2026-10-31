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

test('worker sessions without a Web Worker use worker_threads when available', async () => {
  const reached = new Error('Node adapter reached');
  let created = 0;
  Object.defineProperty(process, 'getBuiltinModule', { configurable: true, value: (name: string) => {
    expect(name).toBe('node:worker_threads');
    return { Worker: class { constructor() { created += 1; throw reached; } } };
  } });
  await expect(openPresentationSession(new Uint8Array())).rejects.toBe(reached);
  const editor = createPptxWorkerEditorSession(new Uint8Array());
  try { await expect(editor.whenHydrated()).rejects.toThrow(reached.message); }
  finally { await editor.dispose(); }
  const fonts = preparePresentationFontDecoder(new Uint8Array());
  try { await expect(fonts.decoder(new Uint8Array(), { codec: 'woff' })).rejects.toThrow(reached.message); }
  finally { fonts.release(); }
  const session = createPptxWorkerSession(new Uint8Array());
  try { await expect(session.ready).rejects.toThrow(reached.message); }
  finally { await session.dispose(); }
  expect(created).toBe(4);
});
