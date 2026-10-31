import { afterEach, beforeEach, expect, test } from 'bun:test';
import { openWorkbookSession } from './client';
import { workbookRecoveryBackend } from './clientRecovery';

const BaseWorker = globalThis.Worker;

beforeEach(() => {
  Object.defineProperty(globalThis, 'Worker', { value: undefined, configurable: true, writable: true });
});

afterEach(() => {
  Object.defineProperty(globalThis, 'Worker', { value: BaseWorker, configurable: true, writable: true });
});

test('worker sessions without a Web Worker name the in-thread API', async () => {
  await expect(openWorkbookSession(new Uint8Array())).rejects.toThrow('use openWorkbook in-thread or pass the worker option');
  await expect(workbookRecoveryBackend.open(new Uint8Array(), {})).rejects.toThrow('use openWorkbook in-thread');
});
