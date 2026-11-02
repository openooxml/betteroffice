import { expect, mock, test } from 'bun:test';
import { createMetafileSvg } from './metafileSvg';

const bytes = new Uint8Array([1, 2, 3]);

test('start shares a load and conversion is unavailable until ready', async () => {
  let finish!: (module: { metafileToSvg: (source: Uint8Array) => Uint8Array }) => void;
  const load = mock(() => new Promise<{ metafileToSvg: (source: Uint8Array) => Uint8Array }>((resolve) => { finish = resolve; }));
  const converter = createMetafileSvg(load);
  expect(converter.convert('picture', bytes)).toBeUndefined();
  const first = converter.start();
  expect(converter.start()).toBe(first);
  await Promise.resolve();
  expect(load).toHaveBeenCalledTimes(1);
  expect(converter.ready()).toBe(false);
  finish({ metafileToSvg: (source) => source });
  await first;
  expect(converter.ready()).toBe(true);
  await converter.start();
  expect(load).toHaveBeenCalledTimes(1);
});

test('a failed start is retryable', async () => {
  let attempts = 0;
  const load = mock(async () => {
    if (++attempts === 1) throw new Error('unavailable');
    return { metafileToSvg: () => bytes };
  });
  const converter = createMetafileSvg(load);
  await expect(converter.start()).rejects.toThrow('unavailable');
  expect(converter.ready()).toBe(false);
  await converter.start();
  expect(converter.ready()).toBe(true);
  expect(load).toHaveBeenCalledTimes(2);
});

test('cache identity includes source length and content, even for reused buffers', async () => {
  const convert = mock((source: Uint8Array) => source.slice());
  const converter = createMetafileSvg(async () => ({ metafileToSvg: convert }));
  await converter.start();
  const source = bytes.slice();
  const first = converter.convert('picture', source);
  expect(converter.convert('picture', source.slice())).toBe(first);
  source[0] = 4;
  expect(converter.convert('picture', source)).toEqual(new Uint8Array([4, 2, 3]));
  expect(converter.convert('picture', source.subarray(0, 2))).toEqual(new Uint8Array([4, 2]));
  expect(convert).toHaveBeenCalledTimes(3);
  expect(first).toEqual(bytes);
});

test('refusals and exceptions are cached', async () => {
  const convert = mock((): Uint8Array | undefined => { throw new Error('refused'); });
  const converter = createMetafileSvg(async () => ({ metafileToSvg: convert }));
  await converter.start();
  expect(converter.convert('picture', bytes)).toBeUndefined();
  expect(converter.convert('picture', bytes)).toBeUndefined();
  expect(convert).toHaveBeenCalledTimes(1);
});

test('per-picture and aggregate source limits include refused conversions', async () => {
  const convert = mock(() => undefined);
  const converter = createMetafileSvg(async () => ({ metafileToSvg: convert }));
  await converter.start();
  expect(converter.convert('large', new Uint8Array(16 * 1024 * 1024 + 1))).toBeUndefined();
  expect(convert).not.toHaveBeenCalled();
  const source = new Uint8Array(16 * 1024 * 1024);
  for (let index = 0; index < 4; index++) converter.convert(String(index), source);
  expect(convert).toHaveBeenCalledTimes(4);
  expect(converter.convert('extra', bytes)).toBeUndefined();
  expect(convert).toHaveBeenCalledTimes(4);
}, 60_000);

test('aggregate SVG limit preserves cached results and refuses excess output', async () => {
  const output = new Uint8Array(32 * 1024 * 1024);
  const convert = mock(() => output);
  const converter = createMetafileSvg(async () => ({ metafileToSvg: convert }));
  await converter.start();
  const first = converter.convert('first', bytes);
  converter.convert('second', bytes);
  expect(converter.convert('extra', bytes)).toBeUndefined();
  expect(converter.convert('first', bytes)).toBe(first);
  expect(convert).toHaveBeenCalledTimes(2);
});

test('output larger than the remaining SVG budget is refused', async () => {
  const convert = mock(() => new Uint8Array(33 * 1024 * 1024));
  const converter = createMetafileSvg(async () => ({ metafileToSvg: convert }));
  await converter.start();
  expect(converter.convert('first', bytes)).toHaveLength(33 * 1024 * 1024);
  expect(converter.convert('second', bytes)).toBeUndefined();
  expect(converter.convert('second', bytes)).toBeUndefined();
  expect(convert).toHaveBeenCalledTimes(2);
});
