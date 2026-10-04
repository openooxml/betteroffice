import { expect, mock, test } from 'bun:test';
import type { YrsSession } from '@betteroffice/docx/yrs';
import {
  awaitWorkerOpenReplica,
  deferWorkerOpenReplica,
  ensureWorkerOpenReplica,
  requestOnDemandWorkerOpenReplica,
  requestWorkerOpenReplica,
  workerOpenReplicaOnDemand,
  workerOpenReplicaPending,
} from './workerOpenReplica';

const fakeSession = () => ({ version: () => 'v1' }) as unknown as YrsSession;

test('passive readiness waits for the owner without requesting hydration', async () => {
  const session = fakeSession();
  const request = mock(() => {});
  const hydrate = mock(async () => () => {});
  const replica = deferWorkerOpenReplica(session, hydrate, () => {}, () => {}, { active: () => true, request });
  const ready = awaitWorkerOpenReplica(session, { passive: true });
  expect(ready).toBe(replica.ready);
  expect(request).not.toHaveBeenCalled();
  expect(hydrate).not.toHaveBeenCalled();
  replica.start();
  await ready;
  expect(hydrate).toHaveBeenCalledTimes(1);
  expect(request).not.toHaveBeenCalled();
});

test('awaiting an on-demand replica asks its owner to start it and resolves when ready', async () => {
  const session = fakeSession();
  const load = mock(() => {});
  const hydrate = mock(async () => load);
  const onReady = mock(() => {});
  const request = mock(() => {});
  const replica = deferWorkerOpenReplica(session, hydrate, () => {}, onReady, {
    active: () => true,
    request,
  });
  const first = awaitWorkerOpenReplica(session);
  expect(first).toBe(replica.ready);
  expect(awaitWorkerOpenReplica(session)).toBe(first);
  expect(request).toHaveBeenCalledTimes(2);
  expect(hydrate).not.toHaveBeenCalled();
  replica.start();
  replica.start();
  await first;
  expect(request).toHaveBeenCalledTimes(2);
  expect(hydrate).toHaveBeenCalledTimes(1);
  expect(load).toHaveBeenCalledTimes(1);
  expect(onReady).toHaveBeenCalledTimes(1);
  expect(workerOpenReplicaPending(session)).toBe(false);
  await awaitWorkerOpenReplica(session);
  expect(hydrate).toHaveBeenCalledTimes(1);
  expect(request).toHaveBeenCalledTimes(2);
});

test.each([false, undefined])('awaiting a replica does not ask for it with onDemand=%s', async (onDemand) => {
  const session = fakeSession();
  const hydrate = mock(async () => () => {});
  const request = mock(() => {});
  const replica = deferWorkerOpenReplica(
    session,
    hydrate,
    () => {},
    () => {},
    onDemand === undefined ? undefined : { active: () => onDemand, request }
  );
  expect(awaitWorkerOpenReplica(session)).toBe(replica.ready);
  requestOnDemandWorkerOpenReplica(session);
  await Promise.resolve();
  expect(hydrate).not.toHaveBeenCalled();
  expect(request).not.toHaveBeenCalled();
  expect(workerOpenReplicaPending(session)).toBe(true);
  replica.cancel();
});

test('an explicit on-demand request asks the owner only while the replica is pending', async () => {
  const session = fakeSession();
  const request = mock(() => {});
  const replica = deferWorkerOpenReplica(session, async () => () => {}, () => {}, () => {}, {
    active: () => true,
    request,
  });
  requestOnDemandWorkerOpenReplica(session);
  expect(request).toHaveBeenCalledTimes(1);
  await requestWorkerOpenReplica(session);
  requestOnDemandWorkerOpenReplica(session);
  expect(request).toHaveBeenCalledTimes(1);
  expect(replica.pending).toBe(false);
});

test('on-demand status follows the current option only while the replica is pending', async () => {
  const session = fakeSession();
  expect(workerOpenReplicaOnDemand(session)).toBe(false);
  let onDemand = false;
  deferWorkerOpenReplica(session, async () => () => {}, () => {}, () => {}, {
    active: () => onDemand,
    request: () => {},
  });
  expect(workerOpenReplicaOnDemand(session)).toBe(false);
  onDemand = true;
  expect(workerOpenReplicaOnDemand(session)).toBe(true);
  onDemand = false;
  expect(workerOpenReplicaOnDemand(session)).toBe(false);
  onDemand = true;
  await requestWorkerOpenReplica(session);
  expect(workerOpenReplicaOnDemand(session)).toBe(false);
});

test('a pending replica without an on-demand option is not on demand', () => {
  const session = fakeSession();
  const replica = deferWorkerOpenReplica(session, async () => () => {}, () => {}, () => {});
  expect(workerOpenReplicaOnDemand(session)).toBe(false);
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
