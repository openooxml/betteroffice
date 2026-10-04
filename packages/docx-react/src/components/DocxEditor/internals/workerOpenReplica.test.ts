import { expect, mock, test } from 'bun:test';
import type { YrsSession } from '@betteroffice/docx/yrs';
import {
  adoptWorkerOpenHandoverVersion,
  adoptWorkerOpenMirrorVersion,
  awaitWorkerOpenReplica,
  deferWorkerOpenReplica,
  ensureWorkerOpenReplica,
  failWorkerOpenReplica,
  holdWorkerOpenDocument,
  releaseWorkerOpenDocument,
  requestWorkerOpenReplica,
  workerOpenDocumentHeld,
  WorkerOpenDocumentHeldError,
  workerOpenReplicaPending,
  workerOpenReplicaStarted,
  workerOpenSourceVersion,
} from './workerOpenReplica';

const fakeSession = () => ({ version: () => 'v1' }) as unknown as YrsSession;

test('awaiting a deferred editor replica leaves starting to its owner', async () => {
  const session = fakeSession();
  const hydrate = mock(async () => () => {});
  const replica = deferWorkerOpenReplica(session, hydrate, () => {}, () => {});
  expect(awaitWorkerOpenReplica(session)).toBe(replica.ready);
  await Promise.resolve();
  expect(hydrate).not.toHaveBeenCalled();
  expect(workerOpenReplicaPending(session)).toBe(true);
  replica.cancel();
});

test('a held document is pending and unstarted without loading its replica', () => {
  const session = fakeSession();
  const release = mock(() => deferWorkerOpenReplica(session, async () => () => {}, () => {}, () => {}));
  expect(workerOpenDocumentHeld(session)).toBe(false);
  holdWorkerOpenDocument(session, release);
  expect(workerOpenDocumentHeld(session)).toBe(true);
  expect(workerOpenReplicaPending(session)).toBe(true);
  expect(workerOpenReplicaStarted(session)).toBe(false);
  expect(release).not.toHaveBeenCalled();
});

test('requesting and awaiting a held document reject without releasing it', async () => {
  const session = fakeSession();
  const release = mock(() => deferWorkerOpenReplica(session, async () => () => {}, () => {}, () => {}));
  holdWorkerOpenDocument(session, release);
  await expect(requestWorkerOpenReplica(session)!).rejects.toBeInstanceOf(WorkerOpenDocumentHeldError);
  await expect(awaitWorkerOpenReplica(session)!).rejects.toBeInstanceOf(WorkerOpenDocumentHeldError);
  expect(workerOpenDocumentHeld(session)).toBe(true);
  expect(workerOpenReplicaPending(session)).toBe(true);
  expect(workerOpenReplicaStarted(session)).toBe(false);
  expect(release).not.toHaveBeenCalled();
});

test('ensuring a held document throws the access error without releasing it', () => {
  const session = fakeSession();
  const release = mock(() => deferWorkerOpenReplica(session, async () => () => {}, () => {}, () => {}));
  holdWorkerOpenDocument(session, release);
  expect(() => ensureWorkerOpenReplica(session)).toThrow(WorkerOpenDocumentHeldError);
  let error: unknown;
  try {
    ensureWorkerOpenReplica(session, 'getDocument');
  } catch (cause) {
    error = cause;
  }
  expect(error).toBeInstanceOf(Error);
  expect(error).toBeInstanceOf(WorkerOpenDocumentHeldError);
  expect((error as WorkerOpenDocumentHeldError).name).toBe('WorkerOpenDocumentHeldError');
  expect((error as WorkerOpenDocumentHeldError).access).toBe('getDocument');
  expect(workerOpenReplicaPending(session)).toBe(true);
  expect(release).not.toHaveBeenCalled();
});

test('a failed hold still never loads here and cannot be released', async () => {
  const session = fakeSession();
  const release = mock(() => deferWorkerOpenReplica(session, async () => () => {}, () => {}, () => {}));
  const failure = new Error('Worker failed');
  holdWorkerOpenDocument(session, release);
  failWorkerOpenReplica(session, failure);
  expect(workerOpenDocumentHeld(session)).toBe(true);
  expect(workerOpenReplicaPending(session)).toBe(false);
  expect(workerOpenReplicaStarted(session)).toBe(false);
  expect(releaseWorkerOpenDocument(session)).toBeUndefined();
  await expect(awaitWorkerOpenReplica(session)!).rejects.toBe(failure);
  await expect(requestWorkerOpenReplica(session)!).rejects.toBe(failure);
  expect(release).not.toHaveBeenCalled();
});

test('releasing a held document registers one unstarted editor replica with its versions', async () => {
  let version = 'initial';
  const session = { version: () => version } as unknown as YrsSession;
  const hydrate = mock(async () => () => { version = 'loaded'; });
  const onReady = mock(() => {});
  const release = mock(() => deferWorkerOpenReplica(session, hydrate, () => {}, onReady));
  holdWorkerOpenDocument(session, release);
  adoptWorkerOpenMirrorVersion(session, 'mirror');
  adoptWorkerOpenHandoverVersion(session, 'handover');
  expect(workerOpenSourceVersion(session, 'initial')).toBe('mirror');
  version = 'before-release';
  const replica = releaseWorkerOpenDocument(session)!;
  expect(release).toHaveBeenCalledTimes(1);
  expect(replica.initialVersion).toBe('initial');
  expect(replica.mirrorVersion).toBe('mirror');
  expect(replica.handoverVersion).toBe('handover');
  expect(replica.started).toBe(false);
  expect(workerOpenDocumentHeld(session)).toBe(false);
  expect(workerOpenReplicaPending(session)).toBe(true);
  expect(workerOpenReplicaStarted(session)).toBe(false);
  expect(workerOpenSourceVersion(session, 'initial')).toBe('mirror');
  expect(releaseWorkerOpenDocument(session)).toBeUndefined();
  expect(release).toHaveBeenCalledTimes(1);
  expect(hydrate).not.toHaveBeenCalled();
  await requestWorkerOpenReplica(session);
  expect(hydrate).toHaveBeenCalledTimes(1);
  expect(onReady).toHaveBeenCalledTimes(1);
  expect(workerOpenSourceVersion(session, 'handover')).toBe('loaded');
});

test('releasing a session without a hold returns nothing', () => {
  const session = fakeSession();
  expect(releaseWorkerOpenDocument(session)).toBeUndefined();
  const replica = deferWorkerOpenReplica(session, async () => () => {}, () => {}, () => {});
  expect(releaseWorkerOpenDocument(session)).toBeUndefined();
  replica.cancel();
});

test('a request starts the replica once and every caller gets the same readiness', async () => {
  const session = fakeSession();
  const load = mock(() => {});
  const hydrate = mock(async () => load);
  const onReady = mock(() => {});
  deferWorkerOpenReplica(session, hydrate, () => {}, onReady);
  expect(workerOpenReplicaPending(session)).toBe(true);
  const first = requestWorkerOpenReplica(session);
  const second = requestWorkerOpenReplica(session);
  expect(first).toBe(second);
  expect(first).toBe(awaitWorkerOpenReplica(session));
  await first;
  expect(hydrate).toHaveBeenCalledTimes(1);
  expect(load).toHaveBeenCalledTimes(1);
  expect(onReady).toHaveBeenCalledTimes(1);
  expect(workerOpenReplicaPending(session)).toBe(false);
  await requestWorkerOpenReplica(session);
  expect(hydrate).toHaveBeenCalledTimes(1);
});

test('a request for a session without a worker-open replica returns nothing', () => {
  expect(requestWorkerOpenReplica(fakeSession())).toBeUndefined();
});

test('a request after the replica was cancelled does not start it', async () => {
  const session = fakeSession();
  const hydrate = mock(async () => () => {});
  const replica = deferWorkerOpenReplica(session, hydrate, () => {}, () => {});
  replica.cancel();
  await expect(requestWorkerOpenReplica(session)!).rejects.toThrow();
  expect(hydrate).not.toHaveBeenCalled();
});

test.each(['rejection', 'load', 'ensure'] as const)('a replica %s fallback defaults to failure', async (cause) => {
  const session = fakeSession();
  const fallback = mock(() => {});
  deferWorkerOpenReplica(session, async () => {
    if (cause === 'rejection') throw new Error('Hydration rejected');
    return () => { throw new Error('Hydration failed'); };
  }, fallback, () => {});
  if (cause === 'ensure') ensureWorkerOpenReplica(session);
  else await requestWorkerOpenReplica(session);
  expect(fallback).toHaveBeenCalledTimes(1);
  expect(fallback).toHaveBeenCalledWith('failure');
  expect(workerOpenReplicaPending(session)).toBe(false);
});
