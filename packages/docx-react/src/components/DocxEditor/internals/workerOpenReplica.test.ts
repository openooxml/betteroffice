import { expect, mock, test } from 'bun:test';
import type { YrsSession } from '@betteroffice/docx/yrs';
import {
  awaitWorkerOpenReplica,
  deferWorkerOpenReplica,
  requestWorkerOpenReplica,
  workerOpenReplicaPending,
} from './workerOpenReplica';

const fakeSession = () => ({ version: () => 'v1' }) as unknown as YrsSession;

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
