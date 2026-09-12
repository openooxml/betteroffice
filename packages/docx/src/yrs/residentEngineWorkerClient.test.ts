import { afterEach, beforeEach, describe, expect, test } from 'bun:test';
import type { YrsResidentWorkerSnapshot, YrsSelection } from './index';
import {
  ResidentEngineWorkerClient,
  type ResidentEngineWorkerPort,
} from './residentEngineWorkerClient';
import type {
  ResidentEngineWorkerRequest,
  ResidentEngineWorkerResponse,
} from './residentEngineWorkerProtocol';

class FakeWorker implements ResidentEngineWorkerPort {
  onmessage: ResidentEngineWorkerPort['onmessage'] = null;
  onerror: ResidentEngineWorkerPort['onerror'] = null;
  onmessageerror: ResidentEngineWorkerPort['onmessageerror'] = null;
  readonly posted: ResidentEngineWorkerRequest[] = [];
  terminated = false;

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

beforeEach(() => {
  timers.clear();
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

function frameReply(id: number): ResidentEngineWorkerResponse {
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

describe('watchdog', () => {
  test('gives bootstrap a larger budget than buildFrame', () => {
    const { client } = setup();
    void client.bootstrap(snapshot, '').catch(() => {});
    void client.buildFrame('', 0).catch(() => {});
    const [bootstrapMs, buildFrameMs] = armedBudgets();
    expect(bootstrapMs).toBeGreaterThan(buildFrameMs);
  });

  test('rejects an unanswered request, terminates the worker, refuses later ones', async () => {
    const { worker, client } = setup();
    const frame = client.buildFrame('', 0);
    expireTimers();
    await expect(frame).rejects.toThrow('did not answer buildFrame');
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
