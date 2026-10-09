import { describe, expect, it } from 'bun:test';
import { createSessionClient } from './client';
import { createSessionHost } from './host';
import { deferReply, isClientMessage, isHostMessage, transferable, type HostMessage } from './protocol';
import type { SchedulerTask, TaskExecutor } from './scheduler';
import { createInProcessPair } from './testing/inProcessTransport';
import { SESSION_SUPERSEDED, SessionFailure, type MethodPolicies, type SessionFailureCode, type SessionScheduler } from './types';

type Methods = {
  echo(value: unknown): unknown;
  add(a: number, b: number): number;
  slow(): string;
  mutate(value: string): string;
  throw(): never;
  trap(): never;
  exhaust(): never;
  transfer(buffer: ArrayBuffer): ArrayBuffer;
};
type Events = { tick: number };
const methods = {
  echo: true, add: true, slow: true, mutate: true, throw: true, trap: true,
  exhaust: true, transfer: true,
} as const;
const policies: MethodPolicies<Methods> = {
  echo: { lane: 'interactive' }, add: { lane: 'interactive' },
  slow: { lane: 'interactive' }, mutate: { lane: 'input', mutates: true },
  throw: { lane: 'interactive' }, trap: { lane: 'interactive' }, transfer: { lane: 'interactive' },
  exhaust: { lane: 'interactive' },
};

function deferred<T>() {
  let resolve!: (value: T | PromiseLike<T>) => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<T>((done, fail) => { resolve = done; reject = fail; });
  return { promise, resolve, reject };
}

function session(executor?: TaskExecutor) {
  const pair = createInProcessPair();
  const held = deferred<string>();
  const started = deferred<void>();
  const disposed = deferred<void>();
  const order: string[] = [];
  let disposeCount = 0;
  let hostBuffer: ArrayBuffer | undefined;
  const host = createSessionHost<Methods, Events, string[]>(pair.host, {
    policies, context: order, executor,
    handlers: {
      echo: (_, value) => value,
      add: (_, a, b) => a + b,
      slow: async (context) => { context.push('slow'); started.resolve(); return await held.promise; },
      mutate: (context, value) => { context.push(value); return value; },
      throw: () => { throw Object.assign(new Error('refused'), { refusal: { reason: 'busy' } }); },
      trap: () => { throw new WebAssembly.RuntimeError('unreachable'); },
      exhaust: () => { throw new SessionFailure('out-of-memory', 'Memory exhausted'); },
      transfer: (_, buffer) => { hostBuffer = buffer; return transferable(buffer, [buffer]); },
    },
    onDispose: () => { disposeCount += 1; disposed.resolve(); },
  });
  const client = createSessionClient<Methods, Events>(pair.client, { methods });
  return { pair, host, client, held, started, disposed, order,
    get disposeCount() { return disposeCount; },
    get hostBuffer() { return hostBuffer; },
  };
}

describe('session host and cloned transport', () => {
  it('round trips calls, serializes handlers and delivers events in order', async () => {
    const s = session();
    const ticks: number[] = [];
    s.client.on('tick', (tick) => { ticks.push(tick); });
    s.host.emit('tick', 1);
    s.host.emit('tick', 2);
    expect(await s.client.call.add(2, 3)).toBe(5);
    const value = { text: 'echo', nested: [1, 2] };
    const echoed = await s.client.call.echo(value);
    expect(echoed).toEqual(value);
    expect(echoed).not.toBe(value);
    expect(ticks).toEqual([1, 2]);
    const completions: string[] = [];
    const slow = s.client.call.slow().then((value) => { completions.push(value); });
    await s.started.promise;
    const first = s.client.call.mutate('first').then((value) => { completions.push(value); });
    const second = s.client.call.mutate('second').then((value) => { completions.push(value); });
    s.held.resolve('done');
    await Promise.all([slow, first, second]);
    expect(s.order).toEqual(['slow', 'first', 'second']);
    expect(completions).toEqual(['done', 'first', 'second']);
    await s.client.dispose();
  });

  it('rejects non-cloneable arguments without failing the session', async () => {
    const s = session();
    const error = (await s.client.call.echo(() => {}).catch((error: unknown) => error)) as Error;
    expect(error).toBeInstanceOf(Error);
    expect(error.name).toBe('DataCloneError');
    expect(error.message.length).toBeGreaterThan(0);
    expect(s.client.failure).toBeUndefined();
    expect(await s.client.call.add(1, 2)).toBe(3);
    await s.client.dispose();
  });

  it('detaches buffers in both directions and transfers event payloads', async () => {
    const s = session();
    const buffer = new Uint8Array([1, 2, 3]).buffer;
    const reply = s.client.callWithTransfer('transfer', [buffer], [buffer]);
    expect(buffer.byteLength).toBe(0);
    const received = await reply;
    expect([...new Uint8Array(received)]).toEqual([1, 2, 3]);
    expect(s.hostBuffer?.byteLength).toBe(0);
    await s.client.dispose();
    const pair = createInProcessPair();
    const events = deferred<ArrayBuffer>();
    const host = createSessionHost<{}, { buffer: ArrayBuffer }, null>(pair.host, {
      handlers: {}, policies: {}, context: null,
    });
    const client = createSessionClient<{}, { buffer: ArrayBuffer }>(pair.client, { methods: {} });
    client.on('buffer', (value) => events.resolve(value));
    const payload = new Uint8Array([4]).buffer;
    host.emit('buffer', payload, [payload]);
    expect(payload.byteLength).toBe(0);
    expect([...new Uint8Array(await events.promise)]).toEqual([4]);
    await client.dispose();
  });

  it('keeps ordinary errors and unknown methods non-terminal', async () => {
    const s = session();
    const error = await s.client.call.throw().catch((error) => error);
    expect(error.message).toBe('refused');
    expect(error.refusal).toEqual({ reason: 'busy' });
    expect(await s.client.call.add(3, 4)).toBe(7);
    const reply = deferred<unknown>();
    const off = s.pair.client.listen((message) => reply.resolve(message));
    s.pair.client.post({ protocol: 1, kind: 'call', id: 999, method: 'toString', args: [] });
    expect(await reply.promise).toEqual({ protocol: 1, kind: 'reply', id: 999, ok: false,
      error: { name: 'Error', message: 'Unknown session method: toString' } });
    off();
    expect(await s.client.call.add(4, 5)).toBe(9);
    await s.client.dispose();
  });

  it('fails once on a trap, rejects queued and later calls, and skips queued mutations', async () => {
    const s = session();
    const failures: SessionFailure[] = [];
    s.client.onFailure((error) => { failures.push(error); });
    const trap = s.client.call.trap().catch((error) => error);
    const queued = s.client.call.mutate('skipped').catch((error) => error);
    expect(await trap).toBeInstanceOf(SessionFailure);
    expect(s.client.failure?.code).toBe('trap');
    expect(await queued).toBe(s.client.failure);
    expect(await s.client.call.add(1, 2).catch((error) => error)).toBe(s.client.failure);
    expect(failures).toEqual([s.client.failure!]);
    expect(s.order).toEqual([]);
  });

  it('makes handler session failures terminal and skips queued handlers', async () => {
    const s = session();
    const failures: SessionFailure[] = [];
    s.client.onFailure((error) => { failures.push(error); });
    const running = s.client.call.slow();
    await s.started.promise;
    const received = deferred<void>();
    const off = s.pair.host.listen((message) => {
      if (isClientMessage(message) && message.kind === 'call' && message.method === 'mutate') {
        received.resolve();
      }
    });
    const failed = s.client.call.exhaust().catch((error) => error);
    const queued = s.client.call.mutate('skipped').catch((error) => error);
    await received.promise;
    off();
    expect(s.host.scheduler.pending().foreground).toBe(2);
    s.held.resolve('done');
    expect(await running).toBe('done');
    expect(await failed).toBe(s.client.failure);
    expect(s.client.failure?.code).toBe('out-of-memory');
    expect(await queued).toBe(s.client.failure);
    expect(failures).toEqual([s.client.failure!]);
    expect(s.order).toEqual(['slow']);
    await s.client.dispose();
  });

  it('cleans up a failed host before the client closes and disposes only once', async () => {
    const pair = createInProcessPair();
    let disposals = 0;
    let calls = 0;
    let unlistens = 0;
    let unerrors = 0;
    createSessionHost<{ trap(): never; echo(): string }, {}, null>({
      ...pair.host,
      listen(listener) {
        const off = pair.host.listen(listener);
        return () => { unlistens += 1; off(); };
      },
      onError(listener) {
        const off = pair.host.onError(listener);
        return () => { unerrors += 1; off(); };
      },
    }, {
      context: null, policies: { trap: { lane: 'interactive' }, echo: { lane: 'interactive' } },
      handlers: {
        trap: () => { throw new WebAssembly.RuntimeError('unreachable'); },
        echo: () => { calls += 1; return 'value'; },
      },
      onDispose: () => { disposals += 1; throw new Error('cleanup'); },
    });
    const client = createSessionClient<{ trap(): never; echo(): string }, {}>(
      { ...pair.client, close: () => {} }, { methods: { trap: true, echo: true } }
    );
    expect((await client.call.trap().catch((error) => error)).code).toBe('trap');
    expect(disposals).toBe(1);
    expect(unlistens).toBe(1);
    expect(unerrors).toBe(1);
    const delivered = deferred<void>();
    pair.host.listen(() => { delivered.resolve(); });
    pair.client.post({ protocol: 1, kind: 'call', id: 99, method: 'echo', args: [] });
    await delivered.promise;
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
    expect(calls).toBe(0);
    pair.client.post({ protocol: 1, kind: 'dispose' });
    await client.dispose();
    expect(disposals).toBe(1);
    pair.client.close();
  });

  it('does not dispose again when a running handler fails after disposal', async () => {
    const s = session();
    const running = s.client.call.slow().catch((error) => error);
    await s.started.promise;
    await s.client.dispose();
    await s.disposed.promise;
    s.held.reject(new SessionFailure('out-of-memory', 'Memory exhausted'));
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
    expect((await running).code).toBe('disposed');
    expect(s.disposeCount).toBe(1);
  });

  it('replies once to superseded calls and runs only the latest replacement', async () => {
    type Methods = { hold(): string; viewport(value: string): string };
    const pair = createInProcessPair();
    const held = deferred<string>();
    const started = deferred<void>();
    const received = deferred<void>();
    const calls: string[] = [];
    const replies: Array<Extract<HostMessage, { kind: 'reply' }>> = [];
    createSessionHost<Methods, {}, null>(pair.host, {
      context: null,
      policies: {
        hold: { lane: 'interactive' },
        viewport: { lane: 'interactive', key: 'viewport', replaceableBy: 'viewport' },
      },
      handlers: {
        hold: () => { started.resolve(); return held.promise; },
        viewport: (_, value) => { calls.push(value); return value; },
      },
    });
    pair.client.listen((message) => {
      if (isHostMessage(message) && message.kind === 'reply') replies.push(message);
    });
    pair.host.listen((message) => {
      if (isClientMessage(message) && message.kind === 'call' && message.args[0] === 'B') {
        received.resolve();
      }
    });
    const client = createSessionClient<Methods, {}>(pair.client, {
      methods: { hold: true, viewport: true },
    });
    const running = client.call.hold();
    await started.promise;
    const first = client.call.viewport('A').catch((error) => error);
    const second = client.call.viewport('B');
    await received.promise;
    held.resolve('done');
    const error = await first;
    expect(error.name).toBe(SESSION_SUPERSEDED);
    expect(error.message).toBe('Superseded by a newer request');
    expect(await running).toBe('done');
    expect(await second).toBe('B');
    expect(calls).toEqual(['B']);
    expect(replies.map((message) => message.id)
      .sort((a, b) => a - b)).toEqual([1, 2, 3]);
    expect(client.failure).toBeUndefined();
    await client.dispose();
  });

  it('runs foreground calls before a deferred background reply completes', async () => {
    type Methods = { background(): string; foreground(): string };
    const pair = createInProcessPair();
    const done = deferred<string>();
    const context: { scheduler?: SessionScheduler } = {};
    const completions: string[] = [];
    let slices = 0;
    const host = createSessionHost<Methods, {}, typeof context>(pair.host, {
      context, policies: { background: { lane: 'interactive' }, foreground: { lane: 'input' } },
      handlers: {
        background(context) {
          const scheduler = context.scheduler!;
          scheduler.schedule({
            kind: 'slices', version: scheduler.version, generation: scheduler.generation,
            run() {
              slices += 1;
              if (slices < 3) return 'yield';
              done.resolve('background');
              return 'done';
            },
          });
          return deferReply(done.promise);
        },
        foreground: () => 'foreground',
      },
    });
    context.scheduler = host.scheduler;
    const client = createSessionClient<Methods, {}>(pair.client, {
      methods: { background: true, foreground: true },
    });
    const background = client.call.background().then((value) => { completions.push(value); return value; });
    const foreground = client.call.foreground().then((value) => { completions.push(value); return value; });
    expect(await foreground).toBe('foreground');
    expect(completions).toEqual(['foreground']);
    expect(await background).toBe('background');
    expect(slices).toBe(3);
    expect(completions).toEqual(['foreground', 'background']);
    await client.dispose();
  });

  it('resolves per-call policies from arguments to control ordering', async () => {
    type Methods = { hold(): void; read(value: string, reorderable: boolean): string; input(): void };
    for (const reorderable of [false, true]) {
      const pair = createInProcessPair();
      const held = deferred<void>();
      const started = deferred<void>();
      const received = deferred<void>();
      const order: string[] = [];
      const args: Array<[string, boolean]> = [];
      const host = createSessionHost<Methods, {}, string[]>(pair.host, {
        context: order,
        policies: {
          hold: { lane: 'interactive' },
          read: (value, reorderable) => {
            args.push([value, reorderable]);
            return { lane: 'interactive', reorderable };
          },
          input: { lane: 'input' },
        },
        handlers: {
          hold: () => { started.resolve(); return held.promise; },
          read: (context, value) => { context.push(value); return value; },
          input: (context) => { context.push('input'); },
        },
      });
      const client = createSessionClient<Methods, {}>(pair.client, {
        methods: { hold: true, read: true, input: true },
      });
      const off = pair.host.listen((message) => {
        if (isClientMessage(message) && message.kind === 'call' && message.method === 'input') {
          received.resolve();
        }
      });
      const hold = client.call.hold();
      await started.promise;
      const read = client.call.read('read', reorderable);
      const input = client.call.input();
      await received.promise;
      off();
      expect(host.scheduler.pending().foreground).toBe(2);
      expect(args).toEqual([['read', reorderable]]);
      held.resolve();
      await Promise.all([hold, read, input]);
      expect(order).toEqual(reorderable ? ['input', 'read'] : ['read', 'input']);
      await client.dispose();
    }
  });

  it('replies to throwing policy functions and fails the session on terminal ones', async () => {
    type Methods = { read(value: string): string; echo(): string };
    const cases: Array<[Error, SessionFailureCode | undefined]> = [
      [new Error('policy'), undefined],
      [new WebAssembly.RuntimeError('policy'), 'trap'],
      [new SessionFailure('out-of-memory', 'policy'), 'out-of-memory'],
    ];
    for (const [error, code] of cases) {
      const pair = createInProcessPair();
      const args: string[] = [];
      let calls = 0;
      createSessionHost<Methods, {}, null>(pair.host, {
        context: null,
        policies: {
          read: (value) => { args.push(value); throw error; },
          echo: { lane: 'interactive' },
        },
        handlers: {
          read: (_, value) => { calls += 1; return value; },
          echo: () => 'value',
        },
      });
      const client = createSessionClient<Methods, {}>(pair.client, {
        methods: { read: true, echo: true },
      });
      const reply = await client.call.read('read').catch((reason) => reason);
      expect(args).toEqual(['read']);
      expect(calls).toBe(0);
      if (code === undefined) {
        expect(reply.name).toBe(error.name);
        expect(reply.message).toBe(error.message);
        expect(client.failure).toBeUndefined();
        expect(await client.call.echo()).toBe('value');
      } else {
        expect(reply).toBeInstanceOf(SessionFailure);
        expect(client.failure?.code).toBe(code);
        await expect(client.call.echo()).rejects.toBeInstanceOf(SessionFailure);
      }
      await client.dispose();
    }
  });

  it('makes deferred session failures terminal', async () => {
    const pair = createInProcessPair();
    const done = deferred<string>();
    const started = deferred<void>();
    let disposals = 0;
    createSessionHost<{ background(): string }, {}, null>(pair.host, {
      context: null, policies: { background: { lane: 'interactive' } },
      handlers: { background: () => { started.resolve(); return deferReply(done.promise); } },
      onDispose: () => { disposals += 1; },
    });
    const client = createSessionClient<{ background(): string }, {}>(pair.client, {
      methods: { background: true },
    });
    const failures: SessionFailure[] = [];
    client.onFailure((error) => { failures.push(error); });
    const reply = client.call.background().catch((error) => error);
    await started.promise;
    done.reject(new SessionFailure('out-of-memory', 'Memory exhausted'));
    expect(await reply).toBe(client.failure);
    expect(client.failure?.code).toBe('out-of-memory');
    expect(failures).toEqual([client.failure!]);
    expect(disposals).toBe(1);
    await client.dispose();
    expect(disposals).toBe(1);
  });

  it('makes pure compute and install traps terminal without cancel hooks', async () => {
    for (const phase of ['compute', 'install'] as const) {
      const s = session();
      const failed = deferred<SessionFailure>();
      s.client.onFailure((error) => { failed.resolve(error); });
      s.host.scheduler.dispatch({
        kind: 'pure', version: 0, generation: 0, input: 1,
        compute(input) {
          if (phase === 'compute') throw new WebAssembly.RuntimeError('boom');
          return input;
        },
        install() {
          if (phase === 'install') throw new WebAssembly.RuntimeError('boom');
        },
      });
      expect((await failed.promise).code).toBe('trap');
      expect(s.client.failure?.message).toBe('boom');
      expect(s.disposeCount).toBe(1);
      expect(await s.client.call.add(1, 2).catch((error) => error)).toBe(s.client.failure);
      await s.client.dispose();
      expect(s.disposeCount).toBe(1);
    }
  });

  it('installs dispatched results through the executor', async () => {
    const result = deferred<unknown>();
    const installed = deferred<number>();
    const jobs: Array<{ kind: string; input: unknown; transfer?: Transferable[] }> = [];
    const executor: TaskExecutor = {
      run<Input, Result>(job: { kind: string; input: Input; transfer?: Transferable[] }): Promise<Result> {
        jobs.push(job);
        return result.promise as Promise<Result>;
      },
    };
    const s = session(executor);
    const transfer: Transferable[] = [];
    let computes = 0;
    let cancels = 0;
    s.host.scheduler.dispatch({
      kind: 'pure', version: 0, generation: 0, input: 4, transfer,
      compute(input) { computes += 1; return input; },
      install(value) { installed.resolve(value); },
      cancel() { cancels += 1; },
    });
    expect(jobs).toEqual([{ kind: 'pure', input: 4, transfer }]);
    result.resolve(8);
    expect(await installed.promise).toBe(8);
    expect(computes).toBe(0);
    expect(await s.client.call.add(1, 2)).toBe(3);
    expect(s.client.failure).toBeUndefined();
    await s.client.dispose();
    await s.disposed.promise;
    expect(cancels).toBe(0);
  });

  it('guards install rejections through the executor', async () => {
    const result = deferred<unknown>();
    const s = session({
      run<Input, Result>(_job: { kind: string; input: Input; transfer?: Transferable[] }) {
        return result.promise as Promise<Result>;
      },
    });
    const failed = deferred<SessionFailure>();
    let computes = 0;
    let installs = 0;
    s.client.onFailure((error) => { failed.resolve(error); });
    s.host.scheduler.dispatch({
      kind: 'pure', version: 0, generation: 0, input: 1,
      compute(input) { computes += 1; return input; },
      async install() { installs += 1; throw new WebAssembly.RuntimeError('trap'); },
    });
    result.resolve(2);
    expect((await failed.promise).code).toBe('trap');
    expect(computes).toBe(0);
    expect(installs).toBe(1);
    expect(s.disposeCount).toBe(1);
    await s.client.dispose();
  });

  it('fails the session when a cancel throws after an executor rejects transferred input', async () => {
    const result = deferred<unknown>();
    const s = session({
      run<Input, Result>(_job: { kind: string; input: Input; transfer?: Transferable[] }) {
        return result.promise as Promise<Result>;
      },
    });
    const failed = deferred<SessionFailure>();
    let cancels = 0;
    s.client.onFailure((error) => { failed.resolve(error); });
    s.host.scheduler.dispatch({
      kind: 'pure', version: 0, generation: 0, input: new ArrayBuffer(4), transfer: [new ArrayBuffer(4)],
      compute(input) { return input; },
      install() {},
      cancel() { cancels += 1; throw new Error('cancel'); },
    });
    result.reject(new Error('executor'));
    expect((await failed.promise).code).toBe('crash');
    expect(cancels).toBe(1);
    expect(s.disposeCount).toBe(1);
    await s.client.dispose();
  });

  it('drops executor results that settle after disposal', async () => {
    const result = deferred<unknown>();
    const s = session({
      run<Input, Result>(_job: { kind: string; input: Input; transfer?: Transferable[] }) {
        return result.promise as Promise<Result>;
      },
    });
    let installs = 0;
    let cancels = 0;
    s.host.scheduler.dispatch({
      kind: 'pure', version: 0, generation: 0, input: 1,
      compute(input) { return input; },
      install() { installs += 1; },
      cancel() { cancels += 1; },
    });
    await s.client.dispose();
    await s.disposed.promise;
    result.resolve(2);
    await new Promise((done) => setTimeout(done, 10));
    expect(installs).toBe(0);
    expect(cancels).toBe(1);
    expect(s.host.scheduler.pending()).toEqual({ foreground: 0, background: 0 });
  });

  it('makes background session failures terminal before task fail hooks', async () => {
    const s = session();
    const failed = deferred<SessionFailure>();
    let taskFailures = 0;
    s.client.onFailure((error) => { failed.resolve(error); });
    s.host.scheduler.schedule({
      kind: 'slice', version: 0, generation: 0,
      run() { throw new SessionFailure('out-of-memory', 'oom', 'diag'); },
      fail() { taskFailures += 1; },
    });
    const error = await failed.promise;
    expect(error.code).toBe('out-of-memory');
    expect(error.message).toBe('oom');
    expect(error.diagnostics).toBe('diag');
    expect(taskFailures).toBe(0);
    expect(s.disposeCount).toBe(1);
    await s.client.dispose();
  });

  it('classifies unhandled background errors as traps or crashes', async () => {
    for (const [error, code] of [
      [new WebAssembly.RuntimeError('boom'), 'trap'],
      [new Error('boom'), 'crash'],
    ] as const) {
      const s = session();
      const failed = deferred<SessionFailure>();
      s.client.onFailure((error) => { failed.resolve(error); });
      s.host.scheduler.schedule({
        kind: 'slice', version: 0, generation: 0,
        run() { throw error; },
      });
      expect((await failed.promise).code).toBe(code);
      expect(s.disposeCount).toBe(1);
      await s.client.dispose();
    }
  });

  it('keeps ordinary background errors routed to task fail hooks', async () => {
    const s = session();
    const failed = deferred<unknown>();
    const error = new Error('refused');
    let cancels = 0;
    s.host.scheduler.schedule({
      kind: 'slice', version: 0, generation: 0,
      run() { throw error; },
      fail(error) { failed.resolve(error); },
      cancel() { cancels += 1; },
    });
    expect(await failed.promise).toBe(error);
    expect(await s.client.call.add(1, 2)).toBe(3);
    expect(s.client.failure).toBeUndefined();
    await s.client.dispose();
    await s.disposed.promise;
    expect(cancels).toBe(0);
  });

  it('refreshes the original task stamps when stale work continues', async () => {
    const s = session();
    const stamps = deferred<number[]>();
    const task: SchedulerTask = {
      kind: 'slice', version: 0, generation: 0,
      onStale: () => 'continue',
      run() {
        stamps.resolve([task.version, task.generation]);
        return 'done';
      },
    };
    s.host.scheduler.schedule(task);
    s.host.scheduler.bump({ version: true });
    s.host.scheduler.bump({ generation: true });
    expect(await stamps.promise).toEqual([1, 2]);
    await s.client.dispose();
    await s.disposed.promise;
  });

  it('delivers a stale cancel error to the task fail hook once', async () => {
    const s = session();
    const failed = deferred<void>();
    const error = new Error('x');
    const failures: unknown[] = [];
    let cancels = 0;
    let runs = 0;
    s.host.scheduler.schedule({
      kind: 'slice', version: 0, generation: 0,
      run() { runs += 1; return 'done'; },
      cancel() { cancels += 1; throw error; },
      fail(error) { failures.push(error); failed.resolve(); },
    });
    s.host.scheduler.bump({ version: true });
    await failed.promise;
    expect(failures).toEqual([error]);
    expect(await s.client.call.add(1, 2)).toBe(3);
    expect(s.client.failure).toBeUndefined();
    await s.client.dispose();
    await s.disposed.promise;
    expect(failures).toEqual([error]);
    expect(cancels).toBe(1);
    expect(runs).toBe(0);
  });

  it('makes stale hook traps terminal before task fail hooks', async () => {
    const s = session();
    const failed = deferred<SessionFailure>();
    let taskFailures = 0;
    let runs = 0;
    s.client.onFailure((error) => { failed.resolve(error); });
    s.host.scheduler.schedule({
      kind: 'slice', version: 0, generation: 0,
      onStale() { throw new WebAssembly.RuntimeError('trap'); },
      run() { runs += 1; return 'done'; },
      fail() { taskFailures += 1; },
    });
    s.host.scheduler.bump({ generation: true });
    expect((await failed.promise).code).toBe('trap');
    expect(s.client.failure?.message).toBe('trap');
    expect(taskFailures).toBe(0);
    expect(runs).toBe(0);
    expect(s.disposeCount).toBe(1);
    await s.client.dispose();
  });

  it('makes stale cancel traps terminal before task fail hooks', async () => {
    const s = session();
    const failed = deferred<SessionFailure>();
    let taskFailures = 0;
    let cancels = 0;
    s.client.onFailure((error) => { failed.resolve(error); });
    s.host.scheduler.schedule({
      kind: 'slice', version: 0, generation: 0,
      run: () => 'done',
      cancel() { cancels += 1; throw new WebAssembly.RuntimeError('trap'); },
      fail() { taskFailures += 1; },
    });
    s.host.scheduler.bump({ version: true });
    expect((await failed.promise).code).toBe('trap');
    expect(taskFailures).toBe(0);
    expect(cancels).toBe(1);
    expect(s.disposeCount).toBe(1);
    await s.client.dispose();
    expect(cancels).toBe(1);
  });

  it('makes pure cancel traps terminal', async () => {
    const s = session();
    const failed = deferred<SessionFailure>();
    let cancels = 0;
    let computes = 0;
    let installs = 0;
    s.client.onFailure((error) => { failed.resolve(error); });
    s.host.scheduler.dispatch({
      kind: 'pure', version: 0, generation: 0, input: 1,
      compute(input) { computes += 1; return input; },
      install() { installs += 1; },
      cancel() { cancels += 1; throw new WebAssembly.RuntimeError('trap'); },
    });
    s.host.scheduler.bump({ version: true });
    expect((await failed.promise).code).toBe('trap');
    expect(cancels).toBe(1);
    expect(computes).toBe(0);
    expect(installs).toBe(0);
    expect(s.disposeCount).toBe(1);
    await s.client.dispose();
    expect(cancels).toBe(1);
  });

  it('fails the session when a stale pure cancel throws twice', async () => {
    const s = session();
    const failed = deferred<SessionFailure>();
    const error = new Error('x');
    const reasons: unknown[] = [];
    let computes = 0;
    let installs = 0;
    s.client.onFailure((error) => { failed.resolve(error); });
    s.host.scheduler.dispatch({
      kind: 'pure', version: 0, generation: 0, input: 1,
      compute(input) { computes += 1; return input; },
      install() { installs += 1; },
      cancel(reason) { reasons.push(reason); throw error; },
    });
    s.host.scheduler.bump({ version: true });
    expect((await failed.promise).code).toBe('crash');
    expect(s.client.failure?.message).toBe('x');
    expect(reasons).toEqual([undefined, error]);
    expect(computes).toBe(0);
    expect(installs).toBe(0);
    expect(s.disposeCount).toBe(1);
    await s.client.dispose();
    expect(reasons).toEqual([undefined, error]);
  });

  it('does not cancel completed or stale-cancelled tasks during shutdown', async () => {
    for (const state of ['completed', 'cancelled'] as const) {
      const s = session();
      const drained = deferred<void>();
      let runs = 0;
      let cancels = 0;
      s.host.scheduler.schedule({
        kind: 'slice', version: 0, generation: 0,
        run() { runs += 1; return 'done'; },
        cancel() { cancels += 1; },
      });
      if (state === 'cancelled') s.host.scheduler.bump({ version: true });
      s.host.scheduler.schedule({
        kind: 'drain', version: s.host.scheduler.version, generation: s.host.scheduler.generation,
        run() { drained.resolve(); return 'done'; },
      });
      await drained.promise;
      expect(runs).toBe(state === 'completed' ? 1 : 0);
      expect(cancels).toBe(state === 'completed' ? 0 : 1);
      await s.client.dispose();
      await s.disposed.promise;
      expect(cancels).toBe(state === 'completed' ? 0 : 1);
    }
  });

  it('cancels queued background tasks once during an idle hold', async () => {
    const pair = createInProcessPair();
    const disposed = deferred<void>();
    let runs = 0;
    let cancels = 0;
    const host = createSessionHost<{ input(): void }, {}, null>(pair.host, {
      context: null, policies: { input: { lane: 'input', holdsIdleTasks: true } },
      handlers: { input() {} },
      onDispose() { disposed.resolve(); },
    });
    const client = createSessionClient<{ input(): void }, {}>(pair.client, {
      methods: { input: true },
    });
    await client.call.input();
    host.scheduler.schedule({
      kind: 'idle', version: 0, generation: 0, idleAfterInputMs: 10_000,
      run() { runs += 1; return 'done'; },
      cancel() { cancels += 1; },
    });
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
    expect(host.scheduler.pending().background).toBe(1);
    await client.dispose();
    await disposed.promise;
    expect(cancels).toBe(1);
    expect(runs).toBe(0);
    await client.dispose();
    expect(cancels).toBe(1);
  });

  it('cancels a running slice once and suppresses late yields and errors after disposal', async () => {
    for (const settle of ['yield', 'reject'] as const) {
      const s = session();
      const started = deferred<void>();
      const held = deferred<'yield'>();
      const finished = deferred<void>();
      let runs = 0;
      let cancels = 0;
      let failures = 0;
      s.host.scheduler.schedule({
        kind: 'slice', version: 0, generation: 0,
        async run() {
          runs += 1;
          started.resolve();
          try { return await held.promise; }
          finally { finished.resolve(); }
        },
        cancel() { cancels += 1; },
        fail() { failures += 1; },
      });
      await started.promise;
      await s.client.dispose();
      await s.disposed.promise;
      expect(cancels).toBe(1);
      if (settle === 'yield') held.resolve('yield');
      else held.reject(new Error('x'));
      await finished.promise;
      await new Promise<void>((resolve) => setTimeout(resolve, 0));
      expect(runs).toBe(1);
      expect(cancels).toBe(1);
      expect(failures).toBe(0);
      expect(s.host.scheduler.pending().background).toBe(0);
    }
  });

  it('transfers deferred results and keeps ordinary deferred errors non-terminal', async () => {
    type Methods = { transfer(): ArrayBuffer; refuse(): string; echo(): string };
    const pair = createInProcessPair();
    const buffer = new Uint8Array([1, 2, 3]).buffer;
    createSessionHost<Methods, {}, null>(pair.host, {
      context: null,
      policies: {
        transfer: { lane: 'interactive' }, refuse: { lane: 'interactive' }, echo: { lane: 'interactive' },
      },
      handlers: {
        transfer: () => deferReply<ArrayBuffer>(Promise.resolve(transferable(buffer, [buffer]))),
        refuse: () => deferReply<string>(Promise.reject(
          Object.assign(new Error('refused'), { refusal: { reason: 'busy' } })
        )),
        echo: () => 'value',
      },
    });
    const client = createSessionClient<Methods, {}>(pair.client, {
      methods: { transfer: true, refuse: true, echo: true },
    });
    expect([...new Uint8Array(await client.call.transfer())]).toEqual([1, 2, 3]);
    expect(buffer.byteLength).toBe(0);
    const error = await client.call.refuse().catch((error) => error);
    expect(error.name).toBe('Error');
    expect(error.message).toBe('refused');
    expect(error.refusal).toEqual({ reason: 'busy' });
    expect(client.failure).toBeUndefined();
    expect(await client.call.echo()).toBe('value');
    await client.dispose();
  });

  it('makes deferred traps terminal', async () => {
    const pair = createInProcessPair();
    createSessionHost<{ trap(): string }, {}, null>(pair.host, {
      context: null, policies: { trap: { lane: 'interactive' } },
      handlers: {
        trap: () => deferReply<string>(Promise.reject(new WebAssembly.RuntimeError('unreachable'))),
      },
    });
    const client = createSessionClient<{ trap(): string }, {}>(pair.client, {
      methods: { trap: true },
    });
    expect(await client.call.trap().catch((error) => error)).toBe(client.failure);
    expect(client.failure?.code).toBe('trap');
    await client.dispose();
  });

  it('sends nothing when deferred replies settle after failure or disposal', async () => {
    for (const end of ['failure', 'dispose'] as const) {
      type Methods = { background(value: string): string; trap(): never };
      const pair = createInProcessPair();
      const resolved = deferred<string>();
      const rejected = deferred<string>();
      const started = deferred<void>();
      const posts: unknown[] = [];
      let calls = 0;
      let disposals = 0;
      createSessionHost<Methods, {}, null>({
        ...pair.host,
        post(message, transfer) { posts.push(message); pair.host.post(message, transfer); },
      }, {
        context: null,
        policies: { background: { lane: 'interactive' }, trap: { lane: 'interactive' } },
        handlers: {
          background(_, value) {
            calls += 1;
            if (calls === 2) started.resolve();
            return deferReply(value === 'resolve' ? resolved.promise : rejected.promise);
          },
          trap: () => { throw new WebAssembly.RuntimeError('unreachable'); },
        },
        onDispose: () => { disposals += 1; },
      });
      const client = createSessionClient<Methods, {}>(pair.client, {
        methods: { background: true, trap: true },
      });
      const first = client.call.background('resolve').catch((error) => error);
      const second = client.call.background('reject').catch((error) => error);
      await started.promise;
      if (end === 'failure') await client.call.trap().catch((error) => error);
      else await client.dispose();
      const count = posts.length;
      resolved.resolve('ignored');
      rejected.reject(new SessionFailure('out-of-memory', 'Memory exhausted'));
      await new Promise<void>((resolve) => setTimeout(resolve, 0));
      expect(posts).toHaveLength(count);
      expect(await first).toBe(client.failure);
      expect(await second).toBe(client.failure);
      expect(disposals).toBe(1);
      await client.dispose();
    }
  });

  it('disposes once, rejects pending and later calls, and suppresses failure notifications', async () => {
    const s = session();
    const failures: SessionFailure[] = [];
    s.client.onFailure((error) => { failures.push(error); });
    const pending = s.client.call.slow().catch((error) => error);
    await s.started.promise;
    const disposal = s.client.dispose();
    expect(s.client.dispose()).toBe(disposal);
    expect((await pending).code).toBe('disposed');
    expect(await s.client.call.add(1, 2).catch((error) => error)).toBe(s.client.failure);
    await disposal;
    await s.disposed.promise;
    s.held.resolve('ignored');
    expect(s.disposeCount).toBe(1);
    expect(failures).toEqual([]);
  });

  it('ignores wasm compile messages without replying or disposing', async () => {
    const s = session();
    const messages: unknown[] = [];
    const off = s.pair.client.listen((message) => { messages.push(message); });
    s.pair.client.post({ protocol: 1, kind: 'wasm-compile' });
    s.pair.client.post({ protocol: 1, kind: 'wasm-compile' });
    expect(await s.client.call.add(1, 2)).toBe(3);
    expect(messages).toEqual([{ protocol: 1, kind: 'reply', id: 1, ok: true, value: 3 }]);
    expect(s.client.failure).toBeUndefined();
    expect(s.disposeCount).toBe(0);
    off();
    await s.client.dispose();
  });

  it('makes malformed client messages terminal', async () => {
    const s = session();
    s.pair.client.post({ protocol: 2, kind: 'call', id: 1, method: 'echo', args: [] });
    const pending = s.client.call.add(1, 2).catch((error) => error);
    expect((await pending).code).toBe('message');
    expect(s.client.failure?.code).toBe('message');
  });

  it('closing either endpoint stops already queued delivery', async () => {
    for (const side of ['client', 'host'] as const) {
      const pair = createInProcessPair();
      const messages: unknown[] = [];
      pair.host.listen((message) => { messages.push(message); });
      pair.client.post({ value: 1 });
      pair[side].close();
      await new Promise<void>((resolve) => setTimeout(resolve, 0));
      expect(messages).toEqual([]);
    }
  });
});
