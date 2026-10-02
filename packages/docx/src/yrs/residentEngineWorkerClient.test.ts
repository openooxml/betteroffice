import { afterEach, beforeEach, describe, expect, spyOn, test } from 'bun:test';
import * as wasm from './wasm/index';
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
  ResidentEngineWorkerHostModule,
  ResidentEngineWorkerRequest,
  ResidentEngineWorkerResponse,
} from './residentEngineWorkerProtocol';

class FakeWorker implements ResidentEngineWorkerPort {
  static instances: FakeWorker[] = [];
  onmessage: ResidentEngineWorkerPort['onmessage'] = null;
  onerror: ResidentEngineWorkerPort['onerror'] = null;
  onmessageerror: ResidentEngineWorkerPort['onmessageerror'] = null;
  readonly posted: (ResidentEngineWorkerRequest | ResidentEngineWorkerHostModule)[] = [];
  readonly transfers: Transferable[][] = [];
  terminated = false;

  constructor(readonly url?: string | URL, readonly options?: WorkerOptions) {
    FakeWorker.instances.push(this);
  }

  postMessage(
    message: ResidentEngineWorkerRequest | ResidentEngineWorkerHostModule,
    transfer: Transferable[] = []
  ): void {
    this.posted.push(message);
    this.transfers.push(transfer);
  }

  terminate(): void {
    this.terminated = true;
  }

  reply(response: ResidentEngineWorkerResponse): void {
    this.onmessage?.({ data: response } as MessageEvent<ResidentEngineWorkerResponse>);
  }

  requestAt(index: number): ResidentEngineWorkerRequest {
    const message = this.posted.at(index)!;
    if (!('id' in message)) throw new Error('Expected a worker request');
    return message;
  }

  lastId(): number {
    const requests = this.posted.filter((message): message is ResidentEngineWorkerRequest =>
      'id' in message
    );
    return requests.at(-1)!.id;
  }
}

type Timer = { callback: () => void; ms: number };
const timers = new Map<number, Timer>();
let nextTimer = 1;
const realSetTimeout = globalThis.setTimeout;
const realClearTimeout = globalThis.clearTimeout;
const originalWorker = globalThis.Worker;
const editModule = new WebAssembly.Module(
  new Uint8Array([0x00, 0x61, 0x73, 0x6d, 0x01, 0x00, 0x00, 0x00])
);
let compileModule: ReturnType<typeof spyOn<typeof wasm, 'editWasmModule'>>;

beforeEach(() => {
  compileModule = spyOn(wasm, 'editWasmModule').mockResolvedValue(editModule);
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
  compileModule.mockRestore();
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

test('pending proposal and navigation reads keep background page builds waiting', async () => {
  const { worker, client } = setup();
  const bootstrap = client.bootstrap(snapshot, '');
  worker.reply(frameReply(worker.lastId()));
  await bootstrap;
  const proposal = client.proposal({ kind: 'snapshot' });
  expect(client.frameRequestPending()).toBe(true);
  worker.reply({ id: worker.lastId(), ok: false, error: 'proposal failed' });
  await expect(proposal).rejects.toThrow('proposal failed');
  expect(client.frameRequestPending()).toBe(false);
  const read = client.documentRead({ kind: 'navigationTarget', story: 'body', paraId: 'p1' });
  expect(client.frameRequestPending()).toBe(true);
  worker.reply({ id: worker.lastId(), ok: true, read: { version: 'v1', value: 'missing-target' } });
  expect(await read).toEqual({ version: 'v1', value: 'missing-target' });
  expect(client.frameRequestPending()).toBe(false);
  client.destroy();
});

test('a version-checked read reports a superseded answer apart from a read value', async () => {
  const { worker, client } = setup();
  const read = () => client.documentReadAt({ kind: 'navigationTarget', story: 'body', paraId: 'p1' }, 'v1');
  const current = read();
  expect(worker.posted.at(-1)).toMatchObject({ type: 'documentRead', expectVersion: 'v1' });
  worker.reply({ id: worker.lastId(), ok: true, read: { version: 'v1', value: 'missing-target' } });
  expect(await current).toEqual({ status: 'ok', version: 'v1', value: 'missing-target' });
  const stale = read();
  worker.reply({ id: worker.lastId(), ok: true, superseded: true });
  expect(await stale).toEqual({ status: 'superseded' });
  expect(client.hasFailed()).toBe(false);
  client.destroy();
});

test('a superseded background build completes without failing the worker', async () => {
  const { worker, client } = setup();
  const pending = client.buildPages([5, 6, 7, 8, 9], 1, false, true);
  expect(worker.posted.at(-1)).toMatchObject({ type: 'buildPages', background: true });
  worker.reply({ id: worker.lastId(), ok: true, pageBuildSuperseded: true });
  expect(await pending).toBeNull();
  expect(client.hasFailed()).toBe(false);
  expect(worker.terminated).toBe(false);
  client.destroy();
});

test('sliced page replies preserve the ordered frames for idle adoption', async () => {
  const { worker, client } = setup();
  const pending = client.buildPages([5, 6, 7, 8, 9], 1, false, true);
  const first = Uint8Array.of(1).buffer;
  const last = Uint8Array.of(2).buffer;
  const reply = frameReply(worker.lastId());
  worker.reply({ ...reply, ok: true, frame: last, pageFrames: [first, last] });
  expect((await pending)?.pageFrames).toEqual([Uint8Array.of(1), Uint8Array.of(2)]);
  client.destroy();
});

test('identical font requirement reads share one request until another request is posted', async () => {
  const { worker, client } = setup();
  const first = client.fontRequirements('{"a":1}');
  const second = client.fontRequirements('{"a":1}');
  const other = client.fontRequirements('{"b":2}');
  const afterOther = client.fontRequirements('{"a":1}');
  client.eraseCaret();
  const afterPost = client.fontRequirements('{"a":1}');
  const reads = worker.posted.flatMap((request) =>
    request.type === 'fontRequirements' ? [request] : []
  );
  expect(reads.map(({ layoutInput }) => layoutInput)).toEqual(['{"a":1}', '{"b":2}', '{"a":1}', '{"a":1}']);
  for (const [index, { id }] of reads.entries()) {
    worker.reply({ id, ok: true, requirementsJson: `[${index}]` });
  }
  expect(await Promise.all([first, second, other, afterOther, afterPost])).toEqual([
    '[0]', '[0]', '[1]', '[2]', '[3]',
  ]);
  const settled = client.fontRequirements('{"a":1}');
  expect(worker.posted.at(-1)).toMatchObject({ type: 'fontRequirements', layoutInput: '{"a":1}' });
  worker.reply({ id: worker.lastId(), ok: true, requirementsJson: '[4]' });
  expect(await settled).toBe('[4]');
});

test('a shared font requirement read rejects every caller when the worker fails', async () => {
  const { worker, client } = setup();
  const first = client.fontRequirements('{}');
  const second = client.fontRequirements('{}');
  expect(worker.posted).toHaveLength(1);
  worker.reply({ id: worker.lastId(), ok: false, error: 'boom' });
  await expect(first).rejects.toThrow('boom');
  await expect(second).rejects.toThrow('boom');
});

test('release requests carry page identities and return a frame', async () => {
  const { worker, client } = setup();
  const pages = [{ index: 8, pageId: '9007199254740993' }];
  const released = client.releasePages(pages, 42, true);
  expect(worker.posted.at(-1)).toEqual({
    id: worker.lastId(),
    type: 'releasePages',
    pages,
    expectedFrameEpoch: 42,
    paintCaret: true,
  });
  expect(client.frameRequestPending()).toBe(true);
  worker.reply(frameReply(worker.lastId()));
  expect(await released).toMatchObject({ frame: new Uint8Array(), selection: null });
  expect(client.frameRequestPending()).toBe(false);
});

test('superseded release replies need no frame and leave the client usable', async () => {
  const { worker, client } = setup();
  const released = client.releasePages([{ index: 2, pageId: '3' }], 7);
  expect(worker.posted.at(-1)).toMatchObject({ type: 'releasePages', paintCaret: false });
  worker.reply({ id: worker.lastId(), ok: true, superseded: true });
  expect(await released).toEqual({ superseded: true });
  expect(client.frameRequestPending()).toBe(false);
  const built = client.buildPages([2], 7);
  worker.reply(frameReply(worker.lastId()));
  expect(await built).toHaveProperty('frame');
});

test('frame and edit requests carry the current display window and retention flag', async () => {
  const { worker, client } = setup();
  client.setRetainBuiltPages(true);
  const bootstrap = client.bootstrap(snapshot, '', { displayWindow: [0, 5] });
  expect(worker.posted.at(-1)).toMatchObject({
    type: 'bootstrap',
    displayWindow: [0, 5],
    retainBuiltPages: true,
  });
  worker.reply(frameReply(worker.lastId()));
  await bootstrap;

  const sync = client.sync(snapshot, '', 0, false, { displayWindow: [1, 6] });
  expect(worker.posted.at(-1)).toMatchObject({
    type: 'sync',
    displayWindow: [1, 6],
    retainBuiltPages: true,
  });
  worker.reply(frameReply(worker.lastId()));
  await sync;

  const frame = client.buildFrame('', 0, false, [8, 11]);
  expect(worker.posted.at(-1)).toMatchObject({
    type: 'buildFrame',
    displayWindow: [8, 11],
    retainBuiltPages: true,
  });
  worker.reply(frameReply(worker.lastId()));
  await frame;

  const input = client.applyInput('a', selection, 0, false, false, [9, 12]);
  expect(worker.posted.at(-1)).toMatchObject({
    type: 'applyInput',
    displayWindow: [9, 12],
    retainBuiltPages: true,
  });
  worker.reply(frameReply(worker.lastId()));
  expect(await input).toMatchObject({ applied: true });

  const deletion = client.applyDelete('backward', selection, 0, false, false, 1, [10, 13]);
  expect(worker.posted.at(-1)).toMatchObject({
    type: 'applyDelete',
    displayWindow: [10, 13],
    retainBuiltPages: true,
  });
  worker.reply(frameReply(worker.lastId()));
  expect(await deletion).toMatchObject({ applied: true });

  client.setRetainBuiltPages(false);
  const released = client.buildFrame('', 0, false, [10, 13]);
  expect(worker.posted.at(-1)).not.toHaveProperty('retainBuiltPages');
  worker.reply(frameReply(worker.lastId()));
  await released;

  client.setRetainBuiltPages(true);
  const unwindowed = client.buildFrame('', 0);
  expect(worker.posted.at(-1)).not.toHaveProperty('retainBuiltPages');
  worker.reply(frameReply(worker.lastId()));
  await unwindowed;
});

test('default windowed requests omit the retention flag', async () => {
  const { worker, client } = setup();
  const requests = [
    () => client.bootstrap(snapshot, '', { displayWindow: [0, 5] }),
    () => client.sync(snapshot, '', 0, false, { displayWindow: [0, 5] }),
    () => client.buildFrame('', 0, false, [0, 5]),
    () => client.applyInput('a', selection, 0, false, false, [0, 5]),
    () => client.applyDelete('backward', selection, 0, false, false, 1, [0, 5]),
  ];
  for (const request of requests) {
    const pending = request();
    expect(worker.posted.at(-1)).toHaveProperty('displayWindow', [0, 5]);
    expect(worker.posted.at(-1)).not.toHaveProperty('retainBuiltPages');
    worker.reply(frameReply(worker.lastId()));
    await pending;
  }
});

describe('warmup', () => {
  test('waits for warm without marking a session ready or bootstrapped', async () => {
    const { worker, client } = setup();
    const warm = client.warm();
    expect(worker.posted).toEqual([{ id: 1, type: 'warm', hostModule: true }]);
    expect(client.isReady()).toBe(false);
    expect(client.bootstrapSent()).toBe(false);
    worker.reply({ id: worker.lastId(), ok: true });
    await warm;
    expect(worker.posted).toEqual([
      { id: 1, type: 'warm', hostModule: true },
      { type: 'editModule', module: editModule },
    ]);
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

  test('posts the compiled module after warm without adding a request or watchdog', async () => {
    let resolve!: (module: WebAssembly.Module) => void;
    compileModule.mockReturnValue(
      new Promise((settle) => {
        resolve = settle;
      })
    );
    const { worker, client } = setup();
    const warm = client.warm();
    expect(compileModule).toHaveBeenCalledTimes(1);
    expect(worker.posted).toEqual([{ id: 1, type: 'warm', hostModule: true }]);
    const watchdogs = [...timers.keys()];
    resolve(editModule);
    await Promise.resolve();
    expect(worker.posted).toEqual([
      { id: 1, type: 'warm', hostModule: true },
      { type: 'editModule', module: editModule },
    ]);
    expect(worker.transfers).toEqual([[], []]);
    expect([...timers.keys()]).toEqual(watchdogs);
    worker.reply({ id: 1, ok: true });
    await warm;
    expect(timers.size).toBe(0);
    const bootstrap = client.bootstrap(snapshot, '');
    expect(worker.posted.at(-1)).toMatchObject({ id: 2, type: 'bootstrap' });
    worker.reply(frameReply(2));
    await bootstrap;
  });

  test('posts null once when the host has no shared module', async () => {
    compileModule.mockResolvedValue(null);
    const { worker, client } = setup();
    const warm = client.warm();
    await Promise.resolve();
    expect(compileModule).toHaveBeenCalledTimes(1);
    expect(worker.posted).toEqual([
      { id: 1, type: 'warm', hostModule: true },
      { type: 'editModule' as const, module: null },
    ]);
    expect(worker.transfers).toEqual([[], []]);
    worker.reply({ id: 1, ok: true });
    await warm;
    expect(worker.posted.filter((message) => message.type === 'editModule')).toHaveLength(1);
  });

  test('posts null when compilation fails', async () => {
    compileModule.mockRejectedValue(new Error('compile failed'));
    const { worker, client } = setup();
    const warm = client.warm();
    await Promise.resolve();
    expect(worker.posted).toEqual([
      { id: 1, type: 'warm', hostModule: true },
      { type: 'editModule' as const, module: null },
    ]);
    worker.reply({ id: 1, ok: true });
    await warm;
  });

  test.each(['destroy', 'failure'])('does not post a compiled module after %s', async (ending) => {
    let resolve!: (module: WebAssembly.Module) => void;
    compileModule.mockReturnValue(
      new Promise((settle) => {
        resolve = settle;
      })
    );
    const { worker, client } = setup();
    const warm = client.warm();
    if (ending === 'destroy') client.destroy();
    else worker.onerror?.({ message: 'crashed' } as ErrorEvent);
    await expect(warm).rejects.toThrow();
    const posted = [...worker.posted];
    resolve(editModule);
    await Promise.resolve();
    expect(worker.posted).toEqual(posted);
    expect(worker.posted.some((message) => message.type === 'editModule')).toBe(false);
  });

  test.each([false, true])('falls back to null on DataCloneError, even if null also throws=%s', async (rejectNull) => {
    const { worker, client } = setup();
    const post = worker.postMessage.bind(worker);
    const posting = spyOn(worker, 'postMessage').mockImplementation((message, transfer) => {
      if (message.type === 'editModule' && (message.module !== null || rejectNull)) {
        throw new DOMException('Module cannot be cloned', 'DataCloneError');
      }
      post(message, transfer);
    });
    try {
      const warm = client.warm();
      await Promise.resolve();
      expect(posting.mock.calls.map(([message]) => message)).toEqual([
        { id: 1, type: 'warm', hostModule: true },
        { type: 'editModule', module: editModule },
        { type: 'editModule' as const, module: null },
      ]);
      expect(worker.posted).toEqual([
        { id: 1, type: 'warm', hostModule: true },
        ...(!rejectNull ? [{ type: 'editModule' as const, module: null }] : []),
      ]);
      worker.reply({ id: 1, ok: true });
      await warm;
    } finally {
      posting.mockRestore();
    }
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
    expect(worker.posted).toEqual([
      { id: 1, type: 'warm', hostModule: true },
      { type: 'editModule', module: editModule },
    ]);
    expect(compileModule).toHaveBeenCalledTimes(1);
    const client = takePreloadedResidentEngineWorker();
    expect(client).not.toBeNull();
    expect(takePreloadedResidentEngineWorker()).toBeNull();
    expect(timers.size).toBe(0);
    client?.destroy();
  });

  test('keeps a spare usable after a non-terminal warm failure', async () => {
    installWorker();
    const warm = preloadResidentEngineWorker();
    const worker = FakeWorker.instances[0];
    worker.reply({ id: worker.lastId(), ok: false, error: 'init failed' });
    await expect(warm).rejects.toThrow('init failed');
    expect(worker.terminated).toBe(false);
    const client = takePreloadedResidentEngineWorker()!;
    expect(client).not.toBeNull();
    expect(client.hasFailed()).toBe(false);
    const bootstrap = client.bootstrap(snapshot, '');
    expect(worker.posted.at(-1)).toMatchObject({ type: 'bootstrap' });
    worker.reply(frameReply(worker.lastId()));
    await bootstrap;
    client.destroy();
  });

  test('retries a non-terminal warm failure on the same spare', async () => {
    installWorker();
    const warm = preloadResidentEngineWorker();
    const worker = FakeWorker.instances[0];
    worker.reply({ id: worker.lastId(), ok: false, error: 'init failed' });
    await expect(warm).rejects.toThrow('init failed');
    const retry = preloadResidentEngineWorker();
    const sharedRetry = preloadResidentEngineWorker();
    expect(FakeWorker.instances).toHaveLength(1);
    expect(retry).not.toBe(warm);
    expect(sharedRetry).toBe(retry);
    expect(worker.posted.filter((message) => 'id' in message)).toEqual([
      { id: 1, type: 'warm', hostModule: true },
      { id: 2, type: 'warm', hostModule: true },
    ]);
    worker.reply({ id: worker.lastId(), ok: true });
    await retry;
    expect(worker.posted.filter((message) => message.type === 'editModule')).toHaveLength(2);
    expect(worker.terminated).toBe(false);
  });

  test('expires an unused spare after a non-terminal warm failure', async () => {
    installWorker();
    const warm = preloadResidentEngineWorker();
    const worker = FakeWorker.instances[0];
    worker.reply({ id: worker.lastId(), ok: false, error: 'init failed' });
    await expect(warm).rejects.toThrow('init failed');
    expect(worker.terminated).toBe(false);
    expect(armedBudgets()).toEqual([RESIDENT_WORKER_SILENCE_MS]);
    expireTimers();
    expect(worker.terminated).toBe(true);
    expect(takePreloadedResidentEngineWorker()).toBeNull();
  });

  test('drops a terminally failed spare so the next preload creates a fresh worker', async () => {
    installWorker();
    const warm = preloadResidentEngineWorker();
    const failed = FakeWorker.instances[0];
    failed.reply({ id: failed.lastId(), ok: false, error: 'init failed', terminal: true });
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
    worker.reply({ id: worker.requestAt(0).id, ok: false, error: 'init failed' });
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
    worker.reply(frameReply(worker.requestAt(0).id));
    await bootstrap;
    expect(armedBudgets()).toEqual([RESIDENT_WORKER_SILENCE_MS]);
    expect([...timers.keys()]).not.toEqual(armed);
    worker.reply(frameReply(worker.requestAt(1).id));
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

  test('the failure listener hears an idle crash once and never a destroy', () => {
    const { worker, client } = setup();
    const failures: string[] = [];
    client.onFailure((error) => failures.push(error.message));
    worker.onerror?.({ message: 'boom' } as ErrorEvent);
    worker.onerror?.({ message: 'again' } as ErrorEvent);
    expect(failures).toEqual(['Resident engine worker failed: boom']);
    const destroyed = setup();
    destroyed.client.onFailure((error) => failures.push(error.message));
    destroyed.client.destroy();
    expect(failures).toHaveLength(1);
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

describe('resident worker opening', () => {
  test('an opened bootstrap under another heap limit fails before touching its bookkeeping', async () => {
    const { worker, client } = setup();
    void client.open(new Uint8Array([1]), { heapLimitBytes: 1024 });
    await expect(
      client.bootstrap(snapshot, '', { opened: true, heapLimitBytes: 2048 })
    ).rejects.toThrow('another heap limit');
    expect(client.bootstrapSent()).toBe(false);
    expect(worker.posted).toHaveLength(1);
  });

  test('opens one document per worker and sends only the bytes of the view it gets', async () => {
    const { worker, client } = setup();
    // A Buffer view into a larger store: its `slice` shares that store rather than copying.
    const view = Buffer.from(new Uint8Array([0, 1, 2, 3, 4]).buffer, 1, 3);
    expect(view.buffer.byteLength).toBeGreaterThan(3);
    void client.open(view);
    const request = worker.posted[0];
    if (request.type !== 'open') throw new Error('open request missing');
    expect(request.bytes).not.toBe(view.buffer);
    expect(request.bytes.byteLength).toBe(3);
    expect(new Uint8Array(request.bytes)).toEqual(new Uint8Array([1, 2, 3]));
    expect(worker.transfers[0]).toEqual([request.bytes]);
    await expect(client.open(new Uint8Array([4]))).rejects.toThrow('already holds a document');
    expect(worker.posted).toHaveLength(1);
  });

  test('open, font requirements, and state handover report memory before an opened bootstrap', async () => {
    const { worker, client } = setup();
    const bytes = new Uint8Array([1, 2, 3]);
    const opened = client.open(bytes, { digest: 'abc', generation: 'opening', heapLimitBytes: 2048 });
    const request = worker.posted[0];
    expect(request).toMatchObject({
      type: 'open',
      digest: 'abc',
      generation: 'opening',
      heapLimitBytes: 2048,
    });
    if (request.type !== 'open') throw new Error('open request missing');
    expect(request.bytes).not.toBe(bytes.buffer);
    expect(new Uint8Array(request.bytes)).toEqual(bytes);
    const stateVector = new Uint8Array([5]);
    worker.reply({
      id: request.id,
      ok: true,
      hostJson: '{"host":1}',
      stateVector: stateVector.buffer,
      memory: [{ label: 'docx-edit', bufferBytes: 1 }],
    });
    expect(await opened).toEqual({ hostJson: '{"host":1}', stateVector });
    expect(client.remoteStateVector()).toEqual(stateVector);
    expect(client.memory()).toEqual([{ label: 'docx-edit', bufferBytes: 1 }]);

    const requirements = client.fontRequirements('{}');
    expect(worker.posted[1]).toMatchObject({ type: 'fontRequirements', layoutInput: '{}' });
    worker.reply({
      id: worker.lastId(),
      ok: true,
      requirementsJson: '[]',
      memory: [{ label: 'docx-edit', bufferBytes: 2 }],
    });
    expect(await requirements).toBe('[]');
    expect(client.memory()).toEqual([{ label: 'docx-edit', bufferBytes: 2 }]);

    const state = client.encodeState();
    expect(worker.posted[2]).toMatchObject({ type: 'encodeState' });
    worker.reply({
      id: worker.lastId(),
      ok: true,
      state: new Uint8Array([7, 8]).buffer,
      memory: [{ label: 'docx-edit', bufferBytes: 3 }],
    });
    expect(await state).toEqual(new Uint8Array([7, 8]));
    expect(client.memory()).toEqual([{ label: 'docx-edit', bufferBytes: 3 }]);

    const bootstrap = client.bootstrap(snapshot, '', { opened: true, frameEpoch: 7 });
    expect(worker.posted[3]).toMatchObject({ type: 'bootstrap', opened: true, expectedFrameEpoch: 7 });
    worker.reply(frameReply(worker.lastId()));
    await bootstrap;
    expect(client.isReady()).toBe(true);
  });

  for (const count of [0, 2]) {
    test(`reads a revision count of ${count} and reports memory`, async () => {
      const { worker, client } = setup();
      const pending = client.revisionCount();
      expect(worker.posted[0]).toMatchObject({ type: 'revisionCount' });
      const memory = [{ label: 'docx-edit', bufferBytes: 65536 }];
      worker.reply({ id: worker.lastId(), ok: true, revisionCount: count, memory });
      expect(await pending).toBe(count);
      expect(client.memory()).toEqual(memory);
    });
  }

  for (const count of [undefined, -1, 0.5, NaN, Infinity, '1', null]) {
    test(`rejects a malformed revision count of ${String(count)}`, async () => {
      const { worker, client } = setup();
      const pending = client.revisionCount();
      worker.reply({
        id: worker.lastId(),
        ok: true,
        revisionCount: count,
      } as ResidentEngineWorkerResponse);
      await expect(pending).rejects.toBeInstanceOf(ResidentWorkerFailureError);
    });
  }

  test('a package that fails to open leaves the client able to open another', async () => {
    const { worker, client } = setup();
    const failed = client.open(new Uint8Array([1]));
    worker.reply({ id: worker.lastId(), ok: false, error: 'not a package' });
    await expect(failed).rejects.toThrow('not a package');
    const opened = client.open(new Uint8Array([2]));
    expect(worker.posted).toHaveLength(2);
    worker.reply({
      id: worker.lastId(),
      ok: true,
      hostJson: '{}',
      stateVector: new Uint8Array([5]).buffer,
    });
    expect((await opened).hostJson).toBe('{}');
    await expect(client.open(new Uint8Array([3]))).rejects.toThrow('already holds a document');
  });

  test('an opened bootstrap queued behind a failed open leaves the client able to open again', async () => {
    const { worker, client } = setup();
    const failed = client.open(new Uint8Array([1]));
    const bootstrap = client.bootstrap(snapshot, '', { opened: true });
    expect(worker.posted).toHaveLength(2);
    worker.reply({ id: worker.requestAt(0).id, ok: false, error: 'not a package' });
    await expect(failed).rejects.toThrow('not a package');
    worker.reply({
      id: worker.requestAt(1).id,
      ok: false,
      error: 'Resident engine worker has no opened document',
    });
    await expect(bootstrap).rejects.toThrow('no opened document');
    expect(client.bootstrapSent()).toBe(false);
    expect(client.remoteStateVector()).toBeNull();
    const opened = client.open(new Uint8Array([2]));
    expect(worker.posted).toHaveLength(3);
    worker.reply({
      id: worker.lastId(),
      ok: true,
      hostJson: '{}',
      stateVector: new Uint8Array([5]).buffer,
    });
    expect((await opened).hostJson).toBe('{}');
  });

  test('a failed opened bootstrap leaves a later bootstrap in place', async () => {
    const { worker, client } = setup();
    const failed = client.open(new Uint8Array([1]));
    const opened = client.bootstrap(snapshot, '', { opened: true });
    worker.reply({ id: worker.requestAt(0).id, ok: false, error: 'not a package' });
    await expect(failed).rejects.toThrow('not a package');
    const recovery = client.bootstrap(snapshot, '');
    worker.reply({
      id: worker.requestAt(1).id,
      ok: false,
      error: 'Resident engine worker has no opened document',
    });
    await expect(opened).rejects.toThrow('no opened document');
    worker.reply(frameReply(worker.requestAt(2).id));
    await recovery;
    expect(client.bootstrapSent()).toBe(true);
    await expect(client.open(new Uint8Array([2]))).rejects.toThrow('already holds a document');
  });

  test('an opened bootstrap without an open fails before touching its bookkeeping', async () => {
    const { worker, client } = setup();
    await expect(client.bootstrap(snapshot, '', { opened: true })).rejects.toThrow(
      'no opened document'
    );
    expect(client.bootstrapSent()).toBe(false);
    expect(worker.posted).toHaveLength(0);
  });

  test('bytes that cannot be copied leave the client able to open another', async () => {
    const { worker, client } = setup();
    const unreadable = {
      get length(): number {
        throw new TypeError('detached');
      },
    } as unknown as Uint8Array;
    await expect(client.open(unreadable)).rejects.toThrow('detached');
    expect(worker.posted).toHaveLength(0);
    void client.open(new Uint8Array([2]));
    expect(worker.posted).toHaveLength(1);
  });

  for (const type of ['open', 'fontRequirements', 'encodeState', 'revisionCount'] as const) {
    test(`${type} propagates the worker's OOM error and memory`, async () => {
      const { worker, client } = setup();
      const pending =
        type === 'open'
          ? client.open(new Uint8Array([1]))
          : type === 'fontRequirements'
            ? client.fontRequirements('{}')
            : type === 'encodeState'
              ? client.encodeState()
              : client.revisionCount();
      const memory = [{ label: 'docx-edit', bufferBytes: 65536, failedAllocationBytes: 64 }];
      worker.reply({
        id: worker.lastId(),
        ok: false,
        error: 'Resident engine worker ran out of memory allocating 64 bytes: unreachable',
        terminal: true,
        outOfMemory: true,
        memory,
      });
      const failure = await pending.then(
        () => { throw new Error('trapped request resolved'); },
        (error: Error) => error
      );
      expect(failure).toBeInstanceOf(ResidentWorkerOutOfMemoryError);
      expect((failure as ResidentWorkerOutOfMemoryError).memory).toEqual(memory);
      expect(client.memory()).toEqual(memory);
      expect(worker.terminated).toBe(true);
      await expect(client.encodeState()).rejects.toBe(failure);
      expect(worker.posted).toHaveLength(1);
    });
  }
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

    const layoutReply = frameReply(worker.requestAt(0).id);
    if (layoutReply.ok) layoutReply.layoutJson = '{"layout":{}}';
    worker.reply(layoutReply);
    expect((await bootstrap).layoutJson).toBe('{"layout":{}}');
    worker.reply(frameReply(worker.requestAt(1).id));
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
    const early = frameReply(worker.requestAt(0).id);
    if (early.ok) early.stateVector = new Uint8Array([9]).buffer;
    worker.reply(early);
    await bootstrap;
    expect(client.remoteStateVector()).toEqual(new Uint8Array([2]));
    expect(client.syncedFontsRevision()).toBe(2);
    const late = frameReply(worker.requestAt(1).id);
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
    expect(client.frameRequestPending()).toBe(false);
    worker.reply({ ...frameReply(worker.lastId()), layoutJson: '{"full":1}' });
    expect(await complete).toMatchObject({ layoutJson: '{"full":1}' });

    const superseded = client.completeLayout(5);
    worker.reply({ id: worker.lastId(), ok: true });
    expect(await superseded).toBeNull();
  });
});
