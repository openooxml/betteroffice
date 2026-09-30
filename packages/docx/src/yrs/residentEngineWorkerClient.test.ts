import { afterEach, beforeEach, describe, expect, test } from 'bun:test';
import type { YrsResidentWorkerSnapshot, YrsSelection } from './index';
import {
  RESIDENT_WORKER_SILENCE_MS,
  ResidentEngineWorkerClient,
  ResidentWorkerFailureError,
  preloadResidentEngineWorker,
  retainPreloadedResidentEngineWorker,
  takePreloadedResidentEngineWorker,
  ResidentWorkerOutOfMemoryError,
  type ResidentEngineWorkerPort,
} from './residentEngineWorkerClient';
import type {
  ResidentEngineWorkerRequest,
  ResidentEngineWorkerResponse,
} from './residentEngineWorkerProtocol';

class FakeWorker implements ResidentEngineWorkerPort {
  static instances: FakeWorker[] = [];
  onmessage: ResidentEngineWorkerPort['onmessage'] = null;
  onerror: ResidentEngineWorkerPort['onerror'] = null;
  onmessageerror: ResidentEngineWorkerPort['onmessageerror'] = null;
  readonly posted: ResidentEngineWorkerRequest[] = [];
  terminated = false;

  constructor(readonly url?: string | URL, readonly options?: WorkerOptions) {
    FakeWorker.instances.push(this);
  }

  postMessage(message: ResidentEngineWorkerRequest): void {
    this.posted.push(message);
  }

  terminate(): void {
    this.terminated = true;
  }

  reply(response: ResidentEngineWorkerResponse): void {
    this.onmessage?.({ data: response } as MessageEvent<ResidentEngineWorkerResponse>);
  }

  lastId(): number {
    return this.posted[this.posted.length - 1].id;
  }
}

type Timer = { callback: () => void; ms: number };
const timers = new Map<number, Timer>();
let nextTimer = 1;
const realSetTimeout = globalThis.setTimeout;
const realClearTimeout = globalThis.clearTimeout;
const originalWorker = globalThis.Worker;

beforeEach(() => {
  timers.clear();
  FakeWorker.instances = [];
  globalThis.setTimeout = ((callback: () => void, ms: number) => {
    const id = nextTimer++;
    timers.set(id, { callback, ms });
    return id;
  }) as unknown as typeof setTimeout;
  globalThis.clearTimeout = ((id: number) => {
    timers.delete(id);
  }) as unknown as typeof clearTimeout;
});

afterEach(() => {
  takePreloadedResidentEngineWorker()?.destroy();
  globalThis.Worker = originalWorker;
  globalThis.setTimeout = realSetTimeout;
  globalThis.clearTimeout = realClearTimeout;
});

function expireTimers(): void {
  const armed = [...timers.values()];
  timers.clear();
  for (const { callback } of armed) callback();
}

function armedBudgets(): number[] {
  return [...timers.values()].map((timer) => timer.ms);
}

const snapshot: YrsResidentWorkerSnapshot = {
  clientId: 1,
  state: new Uint8Array([1]),
  selection: null,
  fonts: [],
  fontsRevision: 0,
  renderInputs: [],
  measureInputs: [],
  layoutInput: '',
  layoutWithRegions: false,
  layoutRevision: 1,
};

const selection: YrsSelection = {
  anchor: { story: 'body', paraId: 'p1', offset: 0 },
  head: { story: 'body', paraId: 'p1', offset: 0 },
};

function frameReply(id: number): Extract<ResidentEngineWorkerResponse, { ok: true }> {
  return {
    id,
    ok: true,
    frame: new ArrayBuffer(0),
    caret: { frameEpoch: 0, caretRect: null },
    selection: null,
    layoutRevision: 1,
  };
}

function errorMessage(error: Error): string {
  return error.message;
}

function setup() {
  const worker = new FakeWorker();
  const client = new ResidentEngineWorkerClient(worker);
  return { worker, client };
}

describe('warmup', () => {
  test('waits for warm without marking a session ready or bootstrapped', async () => {
    const { worker, client } = setup();
    const warm = client.warm();
    expect(worker.posted).toEqual([{ id: 1, type: 'warm' }]);
    expect(client.isReady()).toBe(false);
    expect(client.bootstrapSent()).toBe(false);
    worker.reply({ id: worker.lastId(), ok: true });
    await warm;
    expect(client.isReady()).toBe(false);
    expect(client.remoteStateVector()).toBeNull();
    const bootstrap = client.bootstrap(snapshot, '');
    worker.reply(frameReply(worker.lastId()));
    await bootstrap;
    expect(client.isReady()).toBe(true);
  });

  test('a rejected warm leaves bootstrap usable', async () => {
    const { worker, client } = setup();
    const warm = client.warm();
    worker.reply({ id: worker.lastId(), ok: false, error: 'init failed' });
    await expect(warm).rejects.toThrow('init failed');
    const bootstrap = client.bootstrap(snapshot, '');
    worker.reply(frameReply(worker.lastId()));
    await bootstrap;
    expect(worker.terminated).toBe(false);
    expect(client.isReady()).toBe(true);
  });
});

describe('preloaded worker', () => {
  function installWorker(): void {
    globalThis.Worker = FakeWorker as unknown as typeof Worker;
  }

  async function preloaded(): Promise<FakeWorker> {
    installWorker();
    const warm = preloadResidentEngineWorker();
    const worker = FakeWorker.instances.at(-1)!;
    worker.reply({ id: worker.lastId(), ok: true });
    await warm;
    return worker;
  }

  test('keeps one spare and consumes it exactly once', async () => {
    installWorker();
    const first = preloadResidentEngineWorker();
    const second = preloadResidentEngineWorker();
    expect(first).toBe(second);
    expect(FakeWorker.instances).toHaveLength(1);
    const worker = FakeWorker.instances[0];
    expect(String(worker.url).endsWith('/yrs/residentEngineWorker.mjs')).toBe(true);
    expect(worker.options).toEqual({ type: 'module', name: 'openooxml-resident-engine' });
    worker.reply({ id: worker.lastId(), ok: true });
    await first;
    const client = takePreloadedResidentEngineWorker();
    expect(client).not.toBeNull();
    expect(takePreloadedResidentEngineWorker()).toBeNull();
    expect(timers.size).toBe(0);
    client?.destroy();
  });

  test('drops a failed spare so the next preload creates a fresh worker', async () => {
    installWorker();
    const warm = preloadResidentEngineWorker();
    const failed = FakeWorker.instances[0];
    failed.reply({ id: failed.lastId(), ok: false, error: 'init failed' });
    await expect(warm).rejects.toThrow('init failed');
    expect(failed.terminated).toBe(true);
    expect(takePreloadedResidentEngineWorker()).toBeNull();
    const next = await preloaded();
    expect(next).not.toBe(failed);
    expect(FakeWorker.instances).toHaveLength(2);
  });

  test('an adopted spare can retry bootstrap after its pending warm fails', async () => {
    installWorker();
    const warm = preloadResidentEngineWorker();
    const worker = FakeWorker.instances[0];
    const client = takePreloadedResidentEngineWorker()!;
    const bootstrap = client.bootstrap(snapshot, '');
    worker.reply({ id: worker.posted[0].id, ok: false, error: 'init failed' });
    await expect(warm).rejects.toThrow('init failed');
    worker.reply(frameReply(worker.lastId()));
    await bootstrap;
    expect(worker.terminated).toBe(false);
    expect(client.isReady()).toBe(true);
    client.destroy();
  });

  test('does not adopt a spare made by another Worker constructor', async () => {
    const worker = await preloaded();
    class OtherWorker extends FakeWorker {}
    globalThis.Worker = OtherWorker as unknown as typeof Worker;
    expect(takePreloadedResidentEngineWorker()).toBeNull();
    expect(worker.terminated).toBe(true);
    const fresh = new ResidentEngineWorkerClient();
    expect(FakeWorker.instances.at(-1)).toBeInstanceOf(OtherWorker);
    fresh.destroy();
  });

  test('does not adopt a spare that crashed after warming', async () => {
    const worker = await preloaded();
    worker.onerror?.({ message: 'out of memory' } as ErrorEvent);
    expect(takePreloadedResidentEngineWorker()).toBeNull();
    expect(worker.terminated).toBe(true);
    const next = await preloaded();
    expect(next).not.toBe(worker);
  });

  test('keeps the spare through StrictMode cleanup and frees it after the last owner leaves', async () => {
    const worker = await preloaded();
    const release = retainPreloadedResidentEngineWorker();
    release();
    release();
    await preloadResidentEngineWorker();
    const remounted = retainPreloadedResidentEngineWorker();
    const secondOwner = retainPreloadedResidentEngineWorker();
    expireTimers();
    expect(worker.terminated).toBe(false);
    remounted();
    expireTimers();
    expect(worker.terminated).toBe(false);
    secondOwner();
    expireTimers();
    expect(worker.terminated).toBe(true);
    expect(takePreloadedResidentEngineWorker()).toBeNull();
  });

  test('releasing an owner cannot terminate an adopted worker or a later spare', async () => {
    const adopted = await preloaded();
    const release = retainPreloadedResidentEngineWorker();
    const client = takePreloadedResidentEngineWorker()!;
    const spare = await preloaded();
    const releaseSpare = retainPreloadedResidentEngineWorker();
    release();
    expireTimers();
    expect(adopted.terminated).toBe(false);
    expect(spare.terminated).toBe(false);
    releaseSpare();
    expireTimers();
    expect(spare.terminated).toBe(true);
    client.destroy();
  });

  test('expires an unused public preload', async () => {
    const worker = await preloaded();
    expireTimers();
    expect(worker.terminated).toBe(true);
    expect(takePreloadedResidentEngineWorker()).toBeNull();
  });

  test('skips worker startup where Worker is unavailable', async () => {
    globalThis.Worker = undefined as unknown as typeof Worker;
    await preloadResidentEngineWorker();
    expect(FakeWorker.instances).toHaveLength(0);
    expect(takePreloadedResidentEngineWorker()).toBeNull();
  });
});

describe('watchdog', () => {
  test('queued requests share one silence budget that each reply restarts', async () => {
    const { worker, client } = setup();
    const bootstrap = client.bootstrap(snapshot, '');
    const frame = client.buildFrame('', 0);
    expect(armedBudgets()).toEqual([RESIDENT_WORKER_SILENCE_MS]);
    const armed = [...timers.keys()];
    worker.reply(frameReply(worker.posted[0].id));
    await bootstrap;
    expect(armedBudgets()).toEqual([RESIDENT_WORKER_SILENCE_MS]);
    expect([...timers.keys()]).not.toEqual(armed);
    worker.reply(frameReply(worker.posted[1].id));
    await frame;
    expect(timers.size).toBe(0);
    expect(worker.terminated).toBe(false);
  });

  test('a reply slower than a keystroke budget keeps the worker', async () => {
    const { worker, client } = setup();
    const bootstrap = client.bootstrap(snapshot, '');
    worker.reply(frameReply(worker.lastId()));
    await bootstrap;
    const input = client.applyInput('a', selection, 0);
    expect(armedBudgets().every((ms) => ms >= 60_000)).toBe(true);
    worker.reply(frameReply(worker.lastId()));
    expect(await input).toMatchObject({ applied: true });
    expect(worker.terminated).toBe(false);
    expect(client.isReady()).toBe(true);
  });

  test('rejects an unanswered request, terminates the worker, refuses later ones', async () => {
    const { worker, client } = setup();
    const frame = client.buildFrame('', 0);
    expireTimers();
    const failure = await frame.then(
      () => { throw new Error('unanswered request resolved'); },
      (error: Error) => error
    );
    expect(failure.message).toContain('did not answer buildFrame');
    expect(failure).toBeInstanceOf(ResidentWorkerFailureError);
    expect(worker.terminated).toBe(true);
    await expect(client.buildFrame('', 0)).rejects.toThrow('did not answer buildFrame');
    expect(worker.posted).toHaveLength(1);
  });

  test('disarms the budget once the worker answers', async () => {
    const { worker, client } = setup();
    const frame = client.buildFrame('', 0);
    worker.reply(frameReply(worker.lastId()));
    await frame;
    expect(timers.size).toBe(0);
    expireTimers();
    void client.buildFrame('', 0);
    expect(worker.posted).toHaveLength(2);
    expect(worker.terminated).toBe(false);
  });
});

describe('worker failure', () => {
  test('onerror rejects every pending request and refuses later ones', async () => {
    const { worker, client } = setup();
    const first = client.buildFrame('', 0).catch(errorMessage);
    const second = client
      .attachCanvases([], [], 1, 1, { color: '#000', width: 2 })
      .catch(errorMessage);
    worker.onerror?.({ message: 'boom' } as ErrorEvent);
    expect(await first).toBe('Resident engine worker failed: boom');
    expect(await second).toBe('Resident engine worker failed: boom');
    expect(worker.terminated).toBe(true);
    expect(client.isReady()).toBe(false);
    await expect(client.buildFrame('', 0)).rejects.toThrow('worker failed: boom');
    expect(worker.posted).toHaveLength(2);
  });

  test('a worker crash rejects input with a failure error, not an op error', async () => {
    const { worker, client } = setup();
    const bootstrap = client.bootstrap(snapshot, '');
    worker.reply(frameReply(worker.lastId()));
    await bootstrap;
    const input = client.applyInput('a', selection, 0);
    worker.onerror?.({ message: 'boom' } as ErrorEvent);
    const failure = await input.then(
      () => { throw new Error('crashed input resolved'); },
      (error: Error) => error
    );
    expect(failure).toBeInstanceOf(ResidentWorkerFailureError);
  });

  test('an engine-level input rejection stays a plain error', async () => {
    const { worker, client } = setup();
    const bootstrap = client.bootstrap(snapshot, '');
    worker.reply(frameReply(worker.lastId()));
    await bootstrap;
    const input = client.applyInput('a', selection, 0);
    worker.reply({
      id: worker.lastId(),
      ok: false,
      error: 'apply_input requires a collapsed selection',
    });
    const failure = await input.then(
      () => { throw new Error('rejected input resolved'); },
      (error: Error) => error
    );
    expect(failure.message).toContain('collapsed selection');
    expect(failure).not.toBeInstanceOf(ResidentWorkerFailureError);
    expect(worker.terminated).toBe(false);
  });

  test('a bootstrap carries the heap limit and the frame epoch to follow', () => {
    const { worker, client } = setup();
    void client.bootstrap(snapshot, '', { heapLimitBytes: 2048, frameEpoch: 7 });
    expect(worker.posted[0]).toMatchObject({
      type: 'bootstrap',
      heapLimitBytes: 2048,
      expectedFrameEpoch: 7,
    });
    const other = setup();
    void other.client.bootstrap(snapshot, '');
    expect('heapLimitBytes' in other.worker.posted[0]).toBe(false);
  });

  test('memory() is the memory of the latest reply that carried one', async () => {
    const { worker, client } = setup();
    expect(client.memory()).toBeNull();
    const bootstrap = client.bootstrap(snapshot, '');
    worker.reply({ ...frameReply(worker.lastId()), memory: [{ label: 'docx-edit', bufferBytes: 2 }] });
    await bootstrap;
    expect(client.memory()).toEqual([{ label: 'docx-edit', bufferBytes: 2 }]);
    const frame = client.buildFrame('', 0);
    worker.reply(frameReply(worker.lastId()));
    await frame;
    expect(client.memory()).toEqual([{ label: 'docx-edit', bufferBytes: 2 }]);
  });

  test('onmessageerror is terminal too', async () => {
    const { worker, client } = setup();
    const frame = client.buildFrame('', 0);
    worker.onmessageerror?.({} as MessageEvent);
    await expect(frame).rejects.toThrow('unreadable message');
    expect(worker.terminated).toBe(true);
    await expect(client.buildFrame('', 0)).rejects.toThrow('unreadable message');
  });
});

describe('wasm trap', () => {
  test('a terminal reply rejects the request and refuses later ones', async () => {
    const { worker, client } = setup();
    const frame = client.buildFrame('', 0);
    worker.reply({ id: worker.lastId(), ok: false, error: 'trapped: unreachable', terminal: true });
    await expect(frame).rejects.toThrow('trapped: unreachable');
    expect(worker.terminated).toBe(true);
    await expect(client.sync(snapshot, '', 0)).rejects.toThrow('trapped: unreachable');
    expect(worker.posted).toHaveLength(1);
  });

  test('an out-of-memory trap rejects with the memory the worker reported', async () => {
    const { worker, client } = setup();
    const frame = client.buildFrame('', 0);
    const memory = [
      { label: 'docx-edit', bufferBytes: 4294901760, liveBytes: 4172000000, peakBytes: 4172000000, failedAllocationBytes: 65536 },
    ];
    worker.reply({
      id: worker.lastId(),
      ok: false,
      error: 'Resident engine worker ran out of memory allocating 65536 bytes: unreachable',
      terminal: true,
      outOfMemory: true,
      memory,
    });
    const failure = await frame.then(
      () => { throw new Error('trapped frame resolved'); },
      (error: Error) => error
    );
    expect(failure).toBeInstanceOf(ResidentWorkerOutOfMemoryError);
    expect(failure).not.toBeInstanceOf(ResidentWorkerFailureError);
    expect((failure as ResidentWorkerOutOfMemoryError).memory).toEqual(memory);
    expect(client.memory()).toEqual(memory);
    expect(worker.terminated).toBe(true);
  });

  test('input that runs the worker out of memory rejects instead of reporting not applied', async () => {
    const { worker, client } = setup();
    const bootstrap = client.bootstrap(snapshot, '');
    worker.reply(frameReply(worker.lastId()));
    await bootstrap;
    const input = client.applyInput('x', selection, 1);
    worker.reply({
      id: worker.lastId(),
      ok: false,
      error: 'Resident engine worker ran out of memory allocating 64 bytes: unreachable',
      terminal: true,
      outOfMemory: true,
    });
    await expect(input).rejects.toBeInstanceOf(ResidentWorkerOutOfMemoryError);
    expect(await client.applyDelete('backward', selection, 1)).toEqual({ applied: false });
  });

  test('input is reported as not applied after a trap', async () => {
    const { worker, client } = setup();
    const bootstrap = client.bootstrap(snapshot, '');
    worker.reply(frameReply(worker.lastId()));
    await bootstrap;
    expect(client.isReady()).toBe(true);
    const input = client.applyInput('a', selection, 0);
    worker.reply({ id: worker.lastId(), ok: false, error: 'trapped: unreachable', terminal: true });
    expect(await input).toEqual({ applied: false });
    expect(await client.applyInput('b', selection, 0)).toEqual({ applied: false });
    expect(worker.posted).toHaveLength(2);
  });

  test('a transient unavailable reply keeps the worker alive', async () => {
    const { worker, client } = setup();
    const bootstrap = client.bootstrap(snapshot, '');
    worker.reply(frameReply(worker.lastId()));
    await bootstrap;
    const input = client.applyInput('a', selection, 0);
    worker.reply({
      id: worker.lastId(),
      ok: false,
      error: 'resident input state is not ready',
      residentUnavailable: true,
    });
    expect(await input).toEqual({ applied: false });
    expect(worker.terminated).toBe(false);
    void client.buildFrame('', 0);
    expect(worker.posted).toHaveLength(3);
  });
});

describe('queued snapshots', () => {
  test('a sync sent before the bootstrap answers diffs against the bootstrap state', async () => {
    const { worker, client } = setup();
    const sent = new Uint8Array([7, 7]);
    const bootstrap = client.bootstrap({ ...snapshot, fontsRevision: 3 }, '', {
      stateVector: sent,
      layoutExtras: '{}',
    });
    expect(client.bootstrapSent()).toBe(true);
    expect(client.remoteStateVector()).toEqual(sent);
    expect(client.syncedFontsRevision()).toBe(3);
    expect(worker.posted[0]).toMatchObject({ type: 'bootstrap', layoutExtras: '{}' });

    const sync = client.sync(snapshot, '', 0, false, { stateVector: new Uint8Array([8]) });
    expect(client.remoteStateVector()).toEqual(new Uint8Array([8]));
    expect(worker.posted[1]).not.toHaveProperty('layoutExtras');

    const layoutReply = frameReply(worker.posted[0].id);
    if (layoutReply.ok) layoutReply.layoutJson = '{"layout":{}}';
    worker.reply(layoutReply);
    expect((await bootstrap).layoutJson).toBe('{"layout":{}}');
    worker.reply(frameReply(worker.posted[1].id));
    expect((await sync).layoutJson).toBeUndefined();
  });
});

describe('sent snapshot state', () => {
  test('a reply to an earlier request does not replace a later snapshot hint', async () => {
    const { worker, client } = setup();
    const bootstrap = client.bootstrap({ ...snapshot, fontsRevision: 1 }, '', {
      stateVector: new Uint8Array([1]),
    });
    const sync = client.sync({ ...snapshot, fontsRevision: 2 }, '', 0, false, {
      stateVector: new Uint8Array([2]),
    });
    const early = frameReply(worker.posted[0].id);
    if (early.ok) early.stateVector = new Uint8Array([9]).buffer;
    worker.reply(early);
    await bootstrap;
    expect(client.remoteStateVector()).toEqual(new Uint8Array([2]));
    expect(client.syncedFontsRevision()).toBe(2);
    const late = frameReply(worker.posted[1].id);
    if (late.ok) late.stateVector = new Uint8Array([3]).buffer;
    worker.reply(late);
    await sync;
    expect(client.remoteStateVector()).toEqual(new Uint8Array([3]));
  });
});

describe('provisional layout', () => {
  test('a bootstrap asks for the first pages and completeLayout brings the rest', async () => {
    const { worker, client } = setup();
    const bootstrap = client.bootstrap(snapshot, '', { layoutExtras: '{}', provisionalPages: 3 });
    expect(worker.posted[0]).toMatchObject({ type: 'bootstrap', provisionalPages: 3 });
    worker.reply({ ...frameReply(worker.lastId()), layoutJson: '{}', layoutProvisional: true });
    expect((await bootstrap).layoutProvisional).toBe(true);

    const complete = client.completeLayout(4);
    expect(worker.posted[1]).toMatchObject({ type: 'completeLayout', expectedFrameEpoch: 4 });
    worker.reply({ ...frameReply(worker.lastId()), layoutJson: '{"full":1}' });
    expect(await complete).toMatchObject({ layoutJson: '{"full":1}' });

    const superseded = client.completeLayout(5);
    worker.reply({ id: worker.lastId(), ok: true });
    expect(await superseded).toBeNull();
  });
});
