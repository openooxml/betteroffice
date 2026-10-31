import { afterEach, beforeEach, expect, test } from 'bun:test';
import { openPresentationSession } from './client';
import { createPptxWorkerEditorSession } from './editorSession';
import { preparePresentationFontDecoder } from './fontWorker';

const BaseWorker = globalThis.Worker;
const message = 'use openPresentation in-thread or pass the worker option';

beforeEach(() => {
  Object.defineProperty(globalThis, 'Worker', { value: undefined, configurable: true, writable: true });
});

afterEach(() => {
  Object.defineProperty(globalThis, 'Worker', { value: BaseWorker, configurable: true, writable: true });
});

test('worker sessions without a Web Worker name the in-thread API', async () => {
  await expect(openPresentationSession(new Uint8Array())).rejects.toThrow(message);
  expect(() => createPptxWorkerEditorSession(new Uint8Array())).toThrow(message);
  expect(() => preparePresentationFontDecoder(new Uint8Array())).toThrow(message);
});
