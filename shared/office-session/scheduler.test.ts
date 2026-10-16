import { expect, test } from 'bun:test';
import {
  createResidentScheduler,
  IDLE_SLICE_MS,
  INPUT_ACTIVE_MS,
  INPUT_SLICE_MS,
  type SchedulerTask,
  type TaskExecutor,
} from './scheduler';

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<T>((done, fail) => {
    resolve = done;
    reject = fail;
  });
  return { promise, resolve, reject };
}

async function microtasks() {
  for (let index = 0; index < 20; index += 1) await Promise.resolve();
}

function fakeHost(executor?: TaskExecutor) {
  let now = 0;
  let turnCount = 0;
  let nextTimer = 0;
  const turns: Array<() => void> = [];
  const timers = new Map<number, { at: number; callback: () => void }>();
  const failures: unknown[] = [];
  const scheduler = createResidentScheduler({
    now: () => now,
    turn: (callback) => turns.push(callback),
    timer: (callback, ms) => {
      const id = nextTimer++;
      timers.set(id, { at: now + ms, callback });
      return () => { timers.delete(id); };
    },
    failed: (error) => failures.push(error),
    executor,
  });
  async function turn() {
    const callback = turns.shift();
    if (callback) {
      turnCount += 1;
      callback();
    }
    await microtasks();
  }
  return {
    scheduler,
    failures,
    turn,
    get turnCount() { return turnCount; },
    get timerCount() { return timers.size; },
    async flush() {
      await microtasks();
      let count = 0;
      while (turns.length > 0) {
        if (count++ >= 1000) throw new Error('scheduler did not settle');
        await turn();
      }
    },
    advance(ms: number) {
      now += ms;
      for (const [id, timer] of [...timers].sort((a, b) => a[1].at - b[1].at)) {
        if (timer.at > now || !timers.has(id)) continue;
        timers.delete(id);
        timer.callback();
      }
    },
  };
}

function fakeExecutor() {
  const jobs: Array<{ kind: string; input: unknown; transfer?: Transferable[] }> = [];
  const result = deferred<unknown>();
  const executor: TaskExecutor = {
    async run<Input, Result>(job: { kind: string; input: Input; transfer?: Transferable[] }) {
      jobs.push(job);
      return await result.promise as Result;
    },
  };
  return { executor, jobs, result };
}

test('foreground lanes preserve arrival order without reorderable units', async () => {
  const { scheduler, flush } = fakeHost();
  const order: string[] = [];
  scheduler.submit({ lane: 'interactive', run: () => { order.push('read'); } });
  scheduler.submit({ lane: 'input', run: () => { order.push('input'); } });
  scheduler.submit({ lane: 'collab', run: () => { order.push('collab'); } });
  await flush();
  expect(order).toEqual(['read', 'input', 'collab']);
});

test('input overtakes reorderable reads after any earlier non-reorderable unit', async () => {
  for (const barrier of [false, true]) {
    const { scheduler, flush } = fakeHost();
    const order: string[] = [];
    if (barrier) {
      scheduler.submit({ lane: 'interactive', run: () => { order.push('barrier'); } });
    }
    scheduler.submit({
      lane: 'interactive', reorderable: true, run: () => { order.push('read'); },
    });
    scheduler.submit({ lane: 'input', run: () => { order.push('input'); } });
    await flush();
    expect(order).toEqual(barrier ? ['barrier', 'input', 'read'] : ['input', 'read']);
  }
});

test('background slices yield to foreground units queued before and during a slice', async () => {
  const { scheduler, flush } = fakeHost();
  const order: string[] = [];
  let slices = 0;
  scheduler.schedule({
    kind: 'pages', version: 0, generation: 0,
    run: async () => {
      expect(scheduler.pending().foreground).toBe(0);
      order.push(`slice:${++slices}`);
      if (slices === 1) {
        scheduler.submit({ lane: 'interactive', run: () => { order.push('during'); } });
        await Promise.resolve();
        return 'yield';
      }
      return 'done';
    },
  });
  scheduler.submit({ lane: 'input', run: () => { order.push('before'); } });
  await flush();
  expect(order).toEqual(['before', 'slice:1', 'during', 'slice:2']);
  expect(scheduler.pending()).toEqual({ foreground: 0, background: 0 });
});

test('adjacent collab units run as one unit and stop at an earlier unit of another lane', async () => {
  const host = fakeHost();
  const order: string[] = [];
  const unit = (name: string) => () => { order.push(name); };
  host.scheduler.submit({ lane: 'collab', run: unit('a') });
  host.scheduler.submit({ lane: 'collab', run: unit('b') });
  host.scheduler.submit({ lane: 'interactive', reorderable: true, run: unit('read') });
  host.scheduler.submit({ lane: 'interactive', run: unit('barrier') });
  host.scheduler.submit({ lane: 'collab', run: unit('c') });
  await microtasks();
  expect(order).toEqual(['a', 'b']);
  expect(host.turnCount).toBe(0);
  await host.turn();
  expect(order).toEqual(['a', 'b', 'read', 'barrier', 'c']);
});

test('units run without a turn unless a later message could overtake the next one', async () => {
  const host = fakeHost();
  const order: string[] = [];
  const unit = (name: string) => () => { order.push(name); };
  host.scheduler.submit({ lane: 'interactive', run: unit('first') });
  host.scheduler.submit({ lane: 'input', run: unit('second') });
  await microtasks();
  expect(order).toEqual(['first', 'second']);
  host.scheduler.submit({ lane: 'interactive', run: unit('third') });
  host.scheduler.submit({ lane: 'interactive', reorderable: true, run: unit('read') });
  await microtasks();
  expect(order).toEqual(['first', 'second', 'third']);
  host.scheduler.submit({ lane: 'input', run: unit('input') });
  await host.turn();
  expect(order).toEqual(['first', 'second', 'third', 'input']);
  await host.flush();
  expect(order).toEqual(['first', 'second', 'third', 'input', 'read']);
  expect(host.turnCount).toBe(2);
});

test('a later key supersedes a replaceable queued unit', async () => {
  const { scheduler, flush } = fakeHost();
  const order: string[] = [];
  scheduler.submit({
    lane: 'interactive', replaceableBy: 'pages',
    supersede: () => { order.push('superseded'); },
    run: () => { order.push('old'); },
  });
  scheduler.submit({ lane: 'interactive', key: 'pages', run: () => { order.push('new'); } });
  expect(order).toEqual(['superseded']);
  await flush();
  expect(order).toEqual(['superseded', 'new']);
});

test('a later key does not supersede a running unit', async () => {
  const host = fakeHost();
  const held = deferred<void>();
  const order: string[] = [];
  host.scheduler.submit({
    lane: 'interactive', replaceableBy: 'pages',
    supersede: () => { order.push('superseded'); },
    run: async () => {
      order.push('start');
      await held.promise;
      order.push('finish');
    },
  });
  await host.turn();
  host.scheduler.submit({ lane: 'interactive', key: 'pages', run: () => { order.push('new'); } });
  await host.flush();
  expect(order).toEqual(['start']);
  held.resolve();
  await host.flush();
  expect(order).toEqual(['start', 'finish', 'new']);
});

test('background tasks receive the input budget during activity and the idle budget afterwards', async () => {
  const host = fakeHost();
  const budgets: number[] = [];
  const schedule = () => host.scheduler.schedule({
    kind: 'measure', version: 0, generation: 0,
    run: (budget) => { budgets.push(budget); return 'done'; },
  });
  expect(host.scheduler.budget()).toBe(IDLE_SLICE_MS);
  host.scheduler.submit({ lane: 'input', userInput: true, run: () => {} });
  expect(host.scheduler.budget()).toBe(INPUT_SLICE_MS);
  schedule();
  await host.flush();
  host.advance(INPUT_ACTIVE_MS - 1);
  expect(host.scheduler.budget()).toBe(INPUT_SLICE_MS);
  schedule();
  await host.flush();
  host.advance(1);
  expect(host.scheduler.budget()).toBe(IDLE_SLICE_MS);
  schedule();
  await host.flush();
  expect(budgets).toEqual([INPUT_SLICE_MS, INPUT_SLICE_MS, IDLE_SLICE_MS]);
});

test('idle holds wake by timer and new input postpones the next slice', async () => {
  const host = fakeHost();
  const order: string[] = [];
  let slices = 0;
  const input = () => host.scheduler.submit({
    lane: 'input', userInput: true, holdsIdleTasks: true, run: () => { order.push('input'); },
  });
  input();
  host.scheduler.schedule({
    kind: 'completion', version: 0, generation: 0, idleAfterInputMs: 300,
    run: () => {
      order.push(`slice:${++slices}`);
      if (slices === 1) {
        input();
        return 'yield';
      }
      return 'done';
    },
  });
  await host.flush();
  expect(order).toEqual(['input']);
  expect(host.timerCount).toBe(1);
  host.advance(200);
  input();
  await host.flush();
  host.advance(299);
  await host.flush();
  expect(order).toEqual(['input', 'input']);
  host.advance(1);
  await host.flush();
  expect(order).toEqual(['input', 'input', 'slice:1', 'input']);
  host.advance(299);
  await host.flush();
  expect(slices).toBe(1);
  host.advance(1);
  await host.flush();
  expect(order).toEqual(['input', 'input', 'slice:1', 'input', 'slice:2']);
  expect(host.timerCount).toBe(0);
});

test('input that does not hold idle tasks selects the input budget without postponing them', async () => {
  const host = fakeHost();
  const budgets: number[] = [];
  host.scheduler.submit({ lane: 'input', userInput: true, holdsIdleTasks: true, run: () => {} });
  host.scheduler.schedule({
    kind: 'completion', version: 0, generation: 0, idleAfterInputMs: 300,
    run: (budget) => {
      budgets.push(budget);
      return budgets.length < 2 ? 'yield' : 'done';
    },
  });
  await host.flush();
  for (let step = 0; step < 3; step += 1) {
    host.advance(100);
    host.scheduler.submit({ lane: 'collab', userInput: true, run: () => {} });
    await host.flush();
  }
  expect(budgets).toEqual([INPUT_SLICE_MS, INPUT_SLICE_MS]);
  expect(host.timerCount).toBe(0);
});

test('mutations bump both stamps before running and reframes bump only generation', async () => {
  const { scheduler, flush } = fakeHost();
  const stamps: number[][] = [];
  const run = () => { stamps.push([scheduler.version, scheduler.generation]); };
  scheduler.submit({ lane: 'input', mutates: true, run });
  scheduler.submit({ lane: 'interactive', reframes: true, run });
  scheduler.submit({ lane: 'interactive', run });
  await flush();
  expect(stamps).toEqual([[1, 1], [1, 2], [1, 2]]);
});

test('stale tasks cancel by default or continue with refreshed stamps', async () => {
  const { scheduler, flush } = fakeHost();
  const order: string[] = [];
  scheduler.schedule({
    kind: 'cancel', version: 0, generation: 0,
    cancel: () => { order.push('cancel'); },
    run: () => { order.push('stale'); return 'done'; },
  });
  const continued: SchedulerTask = {
    kind: 'continue', version: 0, generation: 0,
    onStale: () => 'continue',
    run: () => {
      expect([continued.version, continued.generation]).toEqual([1, 2]);
      order.push('continue');
      return 'done';
    },
  };
  scheduler.schedule(continued);
  scheduler.submit({ lane: 'input', mutates: true, run: () => {} });
  scheduler.submit({ lane: 'interactive', reframes: true, run: () => {} });
  await flush();
  expect(order).toEqual(['cancel', 'continue']);
});

test('executor results install as background work after queued foreground units', async () => {
  const fake = fakeExecutor();
  const { scheduler, flush } = fakeHost(fake.executor);
  const order: string[] = [];
  const buffer = new ArrayBuffer(1);
  scheduler.dispatch({
    kind: 'pure', version: 0, generation: 0, input: 3, transfer: [buffer],
    compute: () => { throw new Error('local compute must not run'); },
    install: (result: number) => { order.push(`install:${result}`); },
  });
  expect(fake.jobs).toEqual([{ kind: 'pure', input: 3, transfer: [buffer] }]);
  fake.result.resolve(6);
  await microtasks();
  expect(order).toEqual([]);
  scheduler.submit({ lane: 'interactive', run: () => { order.push('foreground'); } });
  await flush();
  expect(order).toEqual(['foreground', 'install:6']);
});

test('an executor result whose stamps become stale cancels without installing', async () => {
  const fake = fakeExecutor();
  const { scheduler, flush } = fakeHost(fake.executor);
  const order: string[] = [];
  scheduler.dispatch({
    kind: 'pure', version: 0, generation: 0, input: 3,
    compute: () => { throw new Error('local compute must not run'); },
    install: () => { order.push('install'); },
    cancel: () => { order.push('cancel'); },
  });
  fake.result.resolve(6);
  await microtasks();
  scheduler.submit({ lane: 'input', mutates: true, run: () => { order.push('input'); } });
  await flush();
  expect(order).toEqual(['input', 'cancel']);
});

test('executor rejection falls back to local compute in a background unit', async () => {
  const fake = fakeExecutor();
  const { scheduler, flush } = fakeHost(fake.executor);
  const order: string[] = [];
  scheduler.dispatch({
    kind: 'pure', version: 0, generation: 0, input: 3,
    compute: (input) => { order.push(`compute:${input}`); return input * 2; },
    install: (result) => { order.push(`install:${result}`); },
  });
  fake.result.reject(new Error('executor failed'));
  await microtasks();
  expect(order).toEqual([]);
  scheduler.submit({ lane: 'interactive', run: () => { order.push('foreground'); } });
  await flush();
  expect(order).toEqual(['foreground', 'compute:3', 'install:6']);
});

test('an executor that fails after taking a transfer cancels instead of computing locally', async () => {
  const fake = fakeExecutor();
  const { scheduler, flush } = fakeHost(fake.executor);
  const order: string[] = [];
  const failure = new Error('executor failed');
  scheduler.dispatch({
    kind: 'pure', version: 0, generation: 0, input: 3, transfer: [new ArrayBuffer(1)],
    compute: () => { order.push('compute'); return 0; },
    install: () => { order.push('install'); },
    cancel: (reason) => { order.push(reason === failure ? 'cancel' : 'other'); },
  });
  fake.result.reject(failure);
  await flush();
  expect(order).toEqual(['cancel']);
});

test('without an executor pure work computes locally after foreground units', async () => {
  const { scheduler, flush } = fakeHost();
  const order: string[] = [];
  scheduler.dispatch({
    kind: 'pure', version: 0, generation: 0, input: 3,
    compute: (input) => { order.push(`compute:${input}`); return input * 2; },
    install: (result) => { order.push(`install:${result}`); },
  });
  expect(order).toEqual([]);
  scheduler.submit({ lane: 'interactive', run: () => { order.push('foreground'); } });
  await flush();
  expect(order).toEqual(['foreground', 'compute:3', 'install:6']);
});

test('foreground failures report through the host and task failures use their fail hook', async () => {
  const { scheduler, flush, failures } = fakeHost();
  const foregroundError = new Error('foreground failed');
  const taskError = new Error('task failed');
  const taskFailures: unknown[] = [];
  const order: string[] = [];
  scheduler.submit({ lane: 'input', run: () => { throw foregroundError; } });
  scheduler.submit({ lane: 'interactive', run: () => { order.push('later'); } });
  scheduler.schedule({
    kind: 'failed', version: 0, generation: 0,
    run: () => { throw taskError; },
    fail: (error) => taskFailures.push(error),
  });
  scheduler.schedule({
    kind: 'later', version: 0, generation: 0,
    run: () => { order.push('task'); return 'done'; },
  });
  await flush();
  expect(failures).toEqual([foregroundError]);
  expect(taskFailures).toEqual([taskError]);
  expect(order).toEqual(['later', 'task']);
});

test('a throwing supersede or stale cancel is reported and the queue keeps running', async () => {
  const { scheduler, flush, failures } = fakeHost();
  const supersedeError = new Error('supersede failed');
  const cancelError = new Error('cancel failed');
  const taskFailures: unknown[] = [];
  const order: string[] = [];
  scheduler.submit({ lane: 'input', mutates: true, run: () => { order.push('input'); } });
  scheduler.submit({
    lane: 'interactive', key: 'pages', replaceableBy: 'pages',
    run: () => { order.push('replaced'); },
    supersede: () => { throw supersedeError; },
  });
  scheduler.submit({ lane: 'interactive', key: 'pages', run: () => { order.push('latest'); } });
  scheduler.schedule({
    kind: 'stale', version: 0, generation: 0,
    cancel: () => { throw cancelError; },
    fail: (error) => taskFailures.push(error),
    run: () => { order.push('stale'); return 'done'; },
  });
  scheduler.schedule({
    kind: 'later', version: 0, generation: 0, onStale: () => 'continue',
    run: () => { order.push('later'); return 'done'; },
  });
  await flush();
  expect(failures).toEqual([supersedeError]);
  expect(taskFailures).toEqual([cancelError]);
  expect(order).toEqual(['input', 'latest', 'later']);
});
