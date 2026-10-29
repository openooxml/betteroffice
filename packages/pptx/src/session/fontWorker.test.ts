import { beforeAll, expect, spyOn, test } from 'bun:test';
import { resolve } from 'node:path';
import { sessionWasmFactory } from '../../../../shared/office-session/testing/wasmWorker';

type Client = typeof import('./fontWorker');
let factory: Awaited<ReturnType<typeof sessionWasmFactory<Client>>>;

beforeAll(async () => {
  factory = await sessionWasmFactory<Client>(resolve(import.meta.dir, 'fontWorker.ts'),
    resolve(import.meta.dir, 'commandSessionWasmTestWorker.ts'));
});

test('preparation decodes before open and hands the same worker to its editor', async () => {
  const h = factory();
  const source = new Uint8Array([1, 2, 3]);
  const prepared = h.client.preparePresentationFontDecoder(source);
  expect(h.client.preparePresentationFontDecoder(source)).toBe(prepared);
  try {
    expect(await prepared.families).toEqual(['Arial']);
    const backing = new Uint8Array([99, 0, 1, 0, 0, 99]);
    const result = await prepared.decoder(backing.subarray(1, 5), { codec: 'zstd' });
    expect(new Uint8Array(result.font)).toEqual(new Uint8Array([0, 1, 0, 0]));
    expect(backing.byteLength).toBe(6);
    await expect(prepared.decoder(backing, { codec: 'unsupported' }))
      .rejects.toThrow('unsupported font face codec: unsupported');
    expect(h.workers).toHaveLength(1);
    expect(h.workers[0]!.requests.some(message => message.kind === 'call' && message.method === 'openCommandEditor')).toBe(false);
    expect(h.client.takePresentationFontWorker(source) as unknown).toBe(h.workers[0]);
    expect(h.client.takePresentationFontWorker(source)).toBeUndefined();
    await new Promise(resolve => setTimeout(resolve, 0));
    expect(h.workers[0]!.terminated).toBe(false);
    await expect(prepared.decoder(backing, { codec: 'zstd' })).rejects.toThrow('released');
  } finally { prepared.release(); h.workers[0]?.terminate(); }
});

test('abandoned preparation releases its worker', async () => {
  const h = factory();
  const source = new Uint8Array([1]);
  const prepared = h.client.preparePresentationFontDecoder(source);
  await prepared.families;
  prepared.release();
  expect(h.workers[0]!.terminated).toBe(true);
  expect(h.client.takePresentationFontWorker(source)).toBeUndefined();
});

test('font preparation can load past 60 seconds and still decode on the retained worker', async () => {
  let now = 0;
  let nextTimer = 0;
  const timers = new Map<number, { at: number; run(): void }>();
  const clock = spyOn(Date, 'now').mockImplementation(() => now);
  const schedule = (...args: Parameters<typeof setTimeout>): ReturnType<typeof setTimeout> => {
    const [callback, delay, ...values] = args;
    if (typeof callback !== 'function') throw new Error('Expected a timer callback');
    const id = ++nextTimer;
    timers.set(id, { at: now + (delay ?? 0), run: () => callback(...values) });
    return id as unknown as ReturnType<typeof setTimeout>;
  };
  const timeout = spyOn(globalThis, 'setTimeout').mockImplementation(schedule as typeof setTimeout);
  const clear = spyOn(globalThis, 'clearTimeout').mockImplementation(id => {
    timers.delete(id as unknown as number);
  });
  async function advance(ms: number): Promise<void> {
    now += ms;
    for (const [id, timer] of [...timers]) {
      if (timer.at > now || !timers.delete(id)) continue;
      timer.run();
    }
    for (let index = 0; index < 32; index += 1) await Promise.resolve();
  }
  const h = factory();
  const source = new Uint8Array([1]);
  const prepared = h.client.preparePresentationFontDecoder(source);
  try {
    await prepared.families;
    const loading = new Promise<void>(resolve => setTimeout(resolve, 61_000)).then(() =>
      prepared.decoder(new Uint8Array([0, 1, 0, 0]), { codec: 'zstd' }));
    await advance(60_001);
    expect(h.workers[0]!.terminated).toBe(false);
    await advance(999);
    expect(new Uint8Array((await loading).font)).toEqual(new Uint8Array([0, 1, 0, 0]));
    expect(h.workers).toHaveLength(1);
    expect(h.client.takePresentationFontWorker(source) as unknown).toBe(h.workers[0]);
    await advance(0);
    expect(h.workers[0]!.terminated).toBe(false);
  } finally {
    prepared.release();
    h.workers[0]?.terminate();
    clock.mockRestore();
    timeout.mockRestore();
    clear.mockRestore();
  }
});
