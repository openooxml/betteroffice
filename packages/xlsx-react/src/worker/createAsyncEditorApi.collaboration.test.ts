import { expect, test } from 'bun:test';
import type { WorkbookSession } from '@betteroffice/xlsx';
import { createAsyncCollaborationReplica } from '@betteroffice/xlsx/collaboration';
import { SessionFailure } from '../../../../shared/office-session';
import { createXlsxCommandController } from '../commands/createXlsxCommandStore';
import { testSessionCommands } from '../commands/testing';
import { createAsyncEditorApi } from './createAsyncEditorApi';
import { XlsxWorkerSessionFailedError } from './useEditableSessionWorkbook';
import { PeerlessWorkbookSession } from './usePeerlessSessionWorkbook';

function deferred<T>() {
  let reject!: (error: unknown) => void;
  const promise = new Promise<T>((_, failed) => { reject = failed; });
  return { promise, reject };
}

function pendingOperations() {
  const vector = deferred<Uint8Array>();
  const update = deferred<Uint8Array>();
  const apply = deferred<void>();
  let current = true;
  let onFailure: Parameters<WorkbookSession['onFailure']>[0] | undefined;
  const calls = {
    collaborationStateVector: () => vector.promise,
    collaborationStateAsUpdate: () => update.promise,
  };
  const session: WorkbookSession = {
    ...testSessionCommands(),
    state: { format: 'xlsx', stage: 'ready', version: 0, dirty: false, sheets: [], activeSheet: 0 },
    call: calls as WorkbookSession['call'],
    replica: createAsyncCollaborationReplica({ clientId: 101,
      encodeStateVector: calls.collaborationStateVector,
      encodeStateAsUpdate: calls.collaborationStateAsUpdate,
      applyUpdate: () => apply.promise, onUpdate: () => () => {},
    }),
    save: async () => new Uint8Array(), on: () => () => {}, failure: undefined,
    onFailure(listener) { onFailure = listener; return () => { onFailure = undefined; }; },
    dispose: async () => {},
  };
  const run = new PeerlessWorkbookSession(session, 1, {
    changed: () => {}, onError: () => {}, onReady: () => {}, isCurrent: () => current,
  });
  const api = createAsyncEditorApi(run, createXlsxCommandController().store,
    () => { throw new Error('Unexpected bridge access'); });
  return { run, api, replica: api.collaboration!, vector, update, apply,
    replace: () => { current = false; }, fail: (error: SessionFailure) => onFailure?.(error) };
}

test('pending replica operations reject with document replacement after disposal', async () => {
  const { run, replica, vector, update, apply } = pendingOperations();
  const readingVector = replica.encodeStateVector().catch((error: unknown) => error);
  const readingUpdate = replica.encodeStateAsUpdate(Uint8Array.of(1)).catch((error: unknown) => error);
  const applying = replica.applyUpdate(Uint8Array.of(2)).catch((error: unknown) => error);
  await run.dispose();
  const error = new SessionFailure('disposed', 'Session disposed');
  vector.reject(error);
  update.reject(error);
  apply.reject(error);
  expect(await readingVector).toMatchObject({ code: 'document-replaced' });
  expect(await readingUpdate).toMatchObject({ code: 'document-replaced' });
  expect(await applying).toMatchObject({ code: 'document-replaced' });
});

test('pending replica operations reject with the API failure after a worker crash', async () => {
  const { run, api, replica, vector, update, apply, fail } = pendingOperations();
  const readingVector = replica.encodeStateVector().catch((error: unknown) => error);
  const readingUpdate = replica.encodeStateAsUpdate().catch((error: unknown) => error);
  const applying = replica.applyUpdate(Uint8Array.of(2)).catch((error: unknown) => error);
  const error = new SessionFailure('crash', 'Worker stopped');
  fail(error);
  vector.reject(error);
  update.reject(error);
  apply.reject(error);
  try {
    expect(api.failure).toBeInstanceOf(XlsxWorkerSessionFailedError);
    expect(await readingVector).toBe(api.failure);
    expect(await readingUpdate).toBe(api.failure);
    expect(await applying).toBe(api.failure);
  } finally { await run.dispose(); }
});

test('pending replica rejections check currentness before disposal', async () => {
  const { run, replica, vector, update, apply, replace } = pendingOperations();
  const readingVector = replica.encodeStateVector().catch((error: unknown) => error);
  const readingUpdate = replica.encodeStateAsUpdate().catch((error: unknown) => error);
  const applying = replica.applyUpdate(Uint8Array.of(2)).catch((error: unknown) => error);
  replace();
  const error = new Error('Operation stopped');
  vector.reject(error);
  update.reject(error);
  apply.reject(error);
  try {
    expect(await readingVector).toMatchObject({ code: 'document-replaced' });
    expect(await readingUpdate).toMatchObject({ code: 'document-replaced' });
    expect(await applying).toMatchObject({ code: 'document-replaced' });
  } finally { await run.dispose(); }
});

test('replica rejections preserve operation errors while current', async () => {
  const { run, api, replica, apply } = pendingOperations();
  const applying = replica.applyUpdate(Uint8Array.of(2)).catch((error: unknown) => error);
  const error = new Error('invalid-update');
  apply.reject(error);
  try {
    expect(await applying).toBe(error);
    expect(api.failure).toBeNull();
  } finally { await run.dispose(); }
});
