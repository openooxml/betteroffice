import { beforeAll, expect, spyOn, test } from 'bun:test';
import JSZip from 'jszip';
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
  const again = h.client.preparePresentationFontDecoder(source);
  expect(again).not.toBe(prepared);
  again.release();
  try {
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
  prepared.release();
  expect(h.workers[0]!.terminated).toBe(true);
  expect(h.client.takePresentationFontWorker(source)).toBeUndefined();
});

test('releasing during a pending decode rejects it and terminates the worker', async () => {
  const h = factory();
  const releaseCompile = h.holdCompile();
  const source = new Uint8Array([1]);
  const prepared = h.client.preparePresentationFontDecoder(source);
  const decoding = prepared.decoder(new Uint8Array([0, 1, 0, 0]), { codec: 'zstd' });
  const rejected = decoding.catch((error: unknown) => error);
  try {
    for (let tick = 0; tick < 50 && h.compiles === 0; tick++) await Promise.resolve();
    expect(h.compiles).toBe(1);
    prepared.release();
    expect(await rejected).toMatchObject({ name: 'SessionFailure', code: 'disposed' });
    expect(h.workers[0]!.terminated).toBe(true);
    expect(h.client.takePresentationFontWorker(source)).toBeUndefined();
    await expect(prepared.decoder(new Uint8Array([0]), { codec: 'zstd' })).rejects.toThrow('released');
  } finally { prepared.release(); releaseCompile(); }
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

async function kanaDeck(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('ppt/presentation.xml', '<p:presentation><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst></p:presentation>');
  zip.file('ppt/_rels/presentation.xml.rels', '<Relationships><Relationship Id="rId1" Target="slides/slide1.xml"/>' +
    '<Relationship Id="rId2" Target="slides/slide2.xml"/></Relationships>');
  zip.file('ppt/slides/slide1.xml', '<p:sld><a:p><a:r><a:rPr/><a:t>Latin</a:t></a:r></a:p></p:sld>');
  zip.file('ppt/slides/slide2.xml', '<p:sld><a:p><a:r><a:rPr/><a:t>かな</a:t></a:r></a:p></p:sld>');
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

const prefetches = (h: ReturnType<typeof factory>) => h.workers[0]!.requests
  .filter(message => message.kind === 'call' && message.method === 'prefetchCoverageFonts')
  .map(message => message as { id: number; args: unknown[] });

test('preparation starts the opening slide coverage faces on the prepared worker', async () => {
  const h = factory();
  const source = await kanaDeck();
  const prepared = h.client.preparePresentationFontDecoder(source, { activeSlideIndex: 1 });
  try {
    for (let tick = 0; tick < 200 && !prefetches(h).length; tick++) await new Promise(resolve => setTimeout(resolve, 5));
    expect(prefetches(h).map(message => message.args)).toEqual([[['cjk-jp']]]);
  } finally { prepared.release(); }
});

test('preparation skips the coverage prefetch without default fonts or coverage text', async () => {
  for (const [options, label] of [[{ activeSlideIndex: 1, defaultFonts: false }, 'off'], [{}, 'latin']] as const) {
    const h = factory();
    const prepared = h.client.preparePresentationFontDecoder(await kanaDeck(), options);
    try {
      await new Promise(resolve => setTimeout(resolve, 100));
      expect({ label, prefetches: prefetches(h) }).toEqual({ label, prefetches: [] });
    } finally { prepared.release(); }
  }
});

test('a coverage prefetch in flight cannot answer the adopting client', async () => {
  const h = factory();
  const source = await kanaDeck();
  const prepared = h.client.preparePresentationFontDecoder(source, { activeSlideIndex: 1 });
  try {
    for (let tick = 0; tick < 200 && !prefetches(h).length; tick++) await new Promise(resolve => setTimeout(resolve, 5));
    const { id } = prefetches(h)[0]!;
    expect(Number.isSafeInteger(id)).toBe(true);
    expect(id).toBeGreaterThan(2 ** 32);
    const decoded = await prepared.decoder(new Uint8Array([0, 1, 0, 0]), { codec: 'zstd' });
    expect(new Uint8Array(decoded.font)).toEqual(new Uint8Array([0, 1, 0, 0]));
    const decode = h.workers[0]!.requests.find(message => message.kind === 'call' && message.method === 'decodeFontFace') as { id: number };
    expect(decode.id).toBe(1);
    expect(h.client.takePresentationFontWorker(source) as unknown).toBe(h.workers[0]);
  } finally { prepared.release(); h.workers[0]?.terminate(); }
});

test('two editors preparing one file share a worker until both release it', async () => {
  const h = factory();
  const source = new Uint8Array([4, 5, 6]);
  const first = h.client.preparePresentationFontDecoder(source);
  const second = h.client.preparePresentationFontDecoder(source);
  try {
    expect(h.workers).toHaveLength(1);
    const pending = second.decoder(new Uint8Array([0, 1, 0, 0]), { codec: 'zstd' });
    first.release();
    expect(h.workers[0]!.terminated).toBe(false);
    expect(new Uint8Array((await pending).font)).toEqual(new Uint8Array([0, 1, 0, 0]));
    await expect(first.decoder(new Uint8Array([0, 1, 0, 0]), { codec: 'zstd' })).rejects.toThrow('released');
    second.release();
    expect(h.workers[0]!.terminated).toBe(true);
  } finally { first.release(); second.release(); h.workers[0]?.terminate(); }
});

test('a session adopts a shared worker only once the other editor has released it', async () => {
  const h = factory();
  const source = new Uint8Array([7, 8, 9]);
  const first = h.client.preparePresentationFontDecoder(source);
  const second = h.client.preparePresentationFontDecoder(source);
  try {
    expect(h.client.takePresentationFontWorker(source)).toBeUndefined();
    expect(new Uint8Array((await second.decoder(new Uint8Array([0, 1, 0, 0]), { codec: 'zstd' })).font))
      .toEqual(new Uint8Array([0, 1, 0, 0]));
    first.release();
    expect(h.client.takePresentationFontWorker(source) as unknown).toBe(h.workers[0]);
    expect(h.workers[0]!.terminated).toBe(false);
  } finally { first.release(); second.release(); h.workers[0]?.terminate(); }
});
