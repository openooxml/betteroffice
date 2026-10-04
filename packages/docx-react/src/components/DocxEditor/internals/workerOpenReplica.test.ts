import { afterEach, expect, mock, test } from 'bun:test';
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
const globalRestores = new Set<() => void>();

afterEach(() => {
  for (const restore of [...globalRestores]) restore();
});

function holdTasks() {
  const descriptor = Object.getOwnPropertyDescriptor(globalThis, 'scheduler');
  const tasks: Array<() => void> = [];
  Object.defineProperty(globalThis, 'scheduler', {
    configurable: true,
    value: { yield: () => new Promise<void>((resolve) => tasks.push(resolve)) },
  });
  const restore = () => {
    if (!globalRestores.delete(restore)) return;
    if (descriptor) Object.defineProperty(globalThis, 'scheduler', descriptor);
    else Reflect.deleteProperty(globalThis, 'scheduler');
  };
  globalRestores.add(restore);
  return {
    tasks,
    async run() {
      const task = tasks.shift();
      if (!task) throw new Error('No pending task');
      task();
      await Promise.resolve();
    },
    restore,
  };
}

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

test.each(['rejection', 'load', 'ensure'] as const)('a replica %s falls back once', async (cause) => {
  const session = fakeSession();
  const fallback = mock(() => {});
  deferWorkerOpenReplica(session, async () => {
    if (cause === 'rejection') throw new Error('Hydration rejected');
    return () => { throw new Error('Hydration failed'); };
  }, fallback, () => {});
  if (cause === 'ensure') ensureWorkerOpenReplica(session);
  else await requestWorkerOpenReplica(session);
  expect(fallback).toHaveBeenCalledTimes(1);
  expect(fallback).toHaveBeenCalledWith();
  expect(workerOpenReplicaPending(session)).toBe(false);
});

test('hydration opens, loads and publishes readiness in separate tasks', async () => {
  const tasks = holdTasks();
  try {
    const session = fakeSession();
    const openDocx = mock(() => {});
    const loadState = mock(() => {});
    const complete = mock(() => {});
    const onReady = mock(() => {});
    const replica = deferWorkerOpenReplica(session, async () => [openDocx, loadState, complete],
      () => {}, onReady);
    let ready = false;
    void replica.ready.then(() => { ready = true; });
    replica.start();
    await Promise.resolve();
    expect(openDocx).toHaveBeenCalledTimes(1);
    expect(loadState).not.toHaveBeenCalled();
    expect(replica.pending).toBe(true);
    expect(ready).toBe(false);
    expect(onReady).not.toHaveBeenCalled();
    await tasks.run();
    expect(loadState).toHaveBeenCalledTimes(1);
    expect(complete).not.toHaveBeenCalled();
    expect(replica.pending).toBe(true);
    expect(ready).toBe(false);
    expect(onReady).not.toHaveBeenCalled();
    await tasks.run();
    await replica.ready;
    expect(complete).toHaveBeenCalledTimes(1);
    expect(onReady).toHaveBeenCalledTimes(1);
    expect(replica.pending).toBe(false);
    expect(ready).toBe(true);
  } finally {
    tasks.restore();
  }
});

test.each([1, 2])('synchronous ensure completes hydration at yield %s exactly once', async (boundary) => {
  const tasks = holdTasks();
  try {
    const session = fakeSession();
    const openDocx = mock(() => {});
    const loadState = mock(() => {});
    const complete = mock(() => {});
    const fallback = mock(() => {});
    const onReady = mock(() => {});
    const replica = deferWorkerOpenReplica(session, async () => [openDocx, loadState, complete],
      fallback, onReady);
    replica.start();
    await Promise.resolve();
    if (boundary === 2) await tasks.run();
    ensureWorkerOpenReplica(session);
    expect(replica.pending).toBe(false);
    expect(openDocx).toHaveBeenCalledTimes(1);
    expect(loadState).toHaveBeenCalledTimes(1);
    expect(complete).toHaveBeenCalledTimes(1);
    expect(onReady).toHaveBeenCalledTimes(1);
    expect(fallback).not.toHaveBeenCalled();
    await tasks.run();
    await replica.ready;
    ensureWorkerOpenReplica(session);
    expect(openDocx).toHaveBeenCalledTimes(1);
    expect(loadState).toHaveBeenCalledTimes(1);
    expect(complete).toHaveBeenCalledTimes(1);
    expect(onReady).toHaveBeenCalledTimes(1);
    expect(tasks.tasks).toHaveLength(0);
  } finally {
    tasks.restore();
  }
});

test.each([1, 2])('a stale hydration at yield %s cancels without readiness or fallback', async (boundary) => {
  const tasks = holdTasks();
  try {
    const session = fakeSession();
    const openDocx = mock(() => {});
    const loadState = mock(() => {});
    const complete = mock(() => {});
    const fallback = mock(() => {});
    const onReady = mock(() => {});
    const cancel = mock(() => {});
    let current = true;
    const replica = deferWorkerOpenReplica(session, async () => [openDocx, loadState, complete],
      fallback, onReady, undefined, { current: () => current, cancel });
    replica.start();
    await Promise.resolve();
    if (boundary === 2) await tasks.run();
    current = false;
    await tasks.run();
    await expect(replica.ready).rejects.toThrow('The document changed');
    expect(openDocx).toHaveBeenCalledTimes(1);
    expect(loadState).toHaveBeenCalledTimes(boundary - 1);
    expect(complete).not.toHaveBeenCalled();
    expect(onReady).not.toHaveBeenCalled();
    expect(fallback).not.toHaveBeenCalled();
    expect(cancel).toHaveBeenCalledTimes(1);
  } finally {
    tasks.restore();
  }
});

test.each([false, true])('loadState failure after a yield preserves fallback with failure=%s', async (fails) => {
  const tasks = holdTasks();
  try {
    const session = fakeSession();
    const loadError = new Error('Load failed');
    const fallbackError = new Error('Fallback failed');
    const openDocx = mock(() => {});
    const loadState = mock(() => { throw loadError; });
    const complete = mock(() => {});
    const fallback = mock(() => { if (fails) throw fallbackError; });
    const onReady = mock(() => {});
    const replica = deferWorkerOpenReplica(session, async () => [openDocx, loadState, complete],
      fallback, onReady);
    replica.start();
    await Promise.resolve();
    expect(fallback).not.toHaveBeenCalled();
    await tasks.run();
    if (fails) await expect(replica.ready).rejects.toBe(fallbackError);
    else await replica.ready;
    expect(openDocx).toHaveBeenCalledTimes(1);
    expect(loadState).toHaveBeenCalledTimes(1);
    expect(fallback).toHaveBeenCalledTimes(1);
    expect(complete).not.toHaveBeenCalled();
    expect(onReady).toHaveBeenCalledTimes(fails ? 0 : 1);
    expect(replica.pending).toBe(false);
  } finally {
    tasks.restore();
  }
});
