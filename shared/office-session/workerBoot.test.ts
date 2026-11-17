import { afterEach, expect, spyOn, test } from 'bun:test';
import { FakeWorkerScope } from './testing/workerScope';

type Start = (scope: FakeWorkerScope) => void;
type Boot = (scope: FakeWorkerScope, load: () => Promise<Start>, failure: (message: string) => unknown, guard: string) => void;

const docxFailure = (error: string) => ({ id: 0, ok: false, terminal: true, failureCode: 'crash', error });
const sessionFailure = (message: string) => ({ protocol: 1, kind: 'failure', code: 'crash', message });
const timers: (() => void)[] = [];
let timer: ReturnType<typeof spyOn> | undefined;

afterEach(() => {
  timer?.mockRestore();
  timer = undefined;
  timers.length = 0;
});

async function helper(): Promise<Boot | undefined> {
  return import('./workerBoot').then((module) => module.bootWorker as unknown as Boot, () => undefined);
}

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<T>((done, fail) => { resolve = done; reject = fail; });
  return { promise, resolve, reject };
}

async function settle(): Promise<void> {
  for (let turn = 0; turn < 20; turn += 1) await Promise.resolve();
}

function captureTimers(): void {
  timer = spyOn(globalThis, 'setTimeout').mockImplementation(((callback: () => void) => {
    timers.push(callback);
    return 0;
  }) as unknown as typeof setTimeout);
}

function thrownBy(callback: () => void): unknown {
  try { callback(); } catch (error) { return error; }
  return undefined;
}

test('messages sent while the worker loads reach it once, in order, with the same data and ports', async () => {
  const bootWorker = await helper();
  expect(bootWorker).toBeFunction();
  const scope = new FakeWorkerScope();
  const load = deferred<Start>();
  bootWorker!(scope, () => load.promise, sessionFailure, 'betteroffice.test.worker');
  const channel = new MessageChannel();
  const first = { kind: 'call', id: 1 };
  const second = { kind: 'call', id: 2 };
  const early: unknown[] = [];
  scope.dispatchEvent(new MessageEvent('message', { data: first, ports: [channel.port1] }));
  scope.addEventListener('message', (event) => early.push((event as MessageEvent).data));
  scope.dispatchEvent(new MessageEvent('message', { data: second }));
  scope.dispatchEvent(new MessageEvent('messageerror'));
  expect(early).toEqual([]);
  const listened: MessageEvent[] = [];
  const handled: MessageEvent[] = [];
  const unreadable: string[] = [];
  load.resolve((target) => {
    target.addEventListener('message', (event) => listened.push(event as MessageEvent));
    target.onmessage = (event) => { handled.push(event); };
    target.onmessageerror = (event) => { unreadable.push(event.type); };
  });
  await settle();
  for (const events of [listened, handled]) {
    expect(events.map((event) => event.data)).toEqual([first, second]);
    expect(events[0]!.data).toBe(first);
    expect(events[1]!.data).toBe(second);
    expect(events[0]!.ports).toHaveLength(1);
    expect(events[0]!.ports[0]).toBe(channel.port1);
  }
  expect(early).toEqual([first, second]);
  expect(unreadable).toEqual(['messageerror']);
  const third = { kind: 'call', id: 3 };
  scope.dispatchEvent(new MessageEvent('message', { data: third }));
  expect(listened.map((event) => event.data)).toEqual([first, second, third]);
  expect(handled.map((event) => event.data)).toEqual([first, second, third]);
  expect(early).toEqual([first, second, third]);
  expect(scope.posted).toEqual([]);
  channel.port1.close();
  channel.port2.close();
});

test('an import that fails posts the document fatal message, rethrows on a timer and holds nothing', async () => {
  const bootWorker = await helper();
  expect(bootWorker).toBeFunction();
  captureTimers();
  const scope = new FakeWorkerScope();
  const failure = new Error('chunk failed to load');
  const load = deferred<Start>();
  bootWorker!(scope, () => load.promise, docxFailure, 'betteroffice.test.worker');
  scope.dispatchEvent(new MessageEvent('message', { data: { id: 1, type: 'warm' } }));
  load.reject(failure);
  await settle();
  expect(scope.posted).toEqual([{ id: 0, ok: false, terminal: true, failureCode: 'crash', error: 'chunk failed to load' }]);
  expect(timers).toHaveLength(1);
  expect(thrownBy(timers[0]!)).toBe(failure);
  const late: unknown[] = [];
  scope.addEventListener('message', (event) => late.push((event as MessageEvent).data));
  scope.dispatchEvent(new MessageEvent('message', { data: { id: 2, type: 'warm' } }));
  await settle();
  expect(late).toEqual([]);
  expect(scope.posted).toHaveLength(1);
  expect(timers).toHaveLength(1);
});

test('a start that throws posts the session fatal message and rethrows on a timer', async () => {
  const bootWorker = await helper();
  expect(bootWorker).toBeFunction();
  captureTimers();
  const scope = new FakeWorkerScope();
  const failure = new Error('engine start failed');
  bootWorker!(scope, () => Promise.resolve(() => { throw failure; }), sessionFailure, 'betteroffice.test.worker');
  scope.dispatchEvent(new MessageEvent('message', { data: { protocol: 1, kind: 'call', id: 1 } }));
  await settle();
  expect(scope.posted).toEqual([{ protocol: 1, kind: 'failure', code: 'crash', message: 'engine start failed' }]);
  expect(timers).toHaveLength(1);
  expect(thrownBy(timers[0]!)).toBe(failure);
});

test('a second evaluation on the same scope does nothing', async () => {
  const bootWorker = await helper();
  expect(bootWorker).toBeFunction();
  const scope = new FakeWorkerScope();
  const received: unknown[] = [];
  let loads = 0;
  const load = () => {
    loads += 1;
    return Promise.resolve<Start>((target) => {
      target.addEventListener('message', (event) => received.push((event as MessageEvent).data));
    });
  };
  bootWorker!(scope, load, sessionFailure, 'betteroffice.test.worker');
  bootWorker!(scope, load, sessionFailure, 'betteroffice.test.worker');
  const message = { kind: 'call', id: 1 };
  scope.dispatchEvent(new MessageEvent('message', { data: message }));
  await settle();
  expect(loads).toBe(1);
  expect(received).toEqual([message]);
  expect(scope.listened.filter((type) => type === 'message')).toHaveLength(2);
});
