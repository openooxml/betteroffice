import { expect, mock, test } from 'bun:test';
import { createMetafileSvg, type MetafileSvgModule } from './metafileSvg';

const bytes = new Uint8Array([1, 2, 3]);

function moduleFor(toSvg: (source: Uint8Array) => Uint8Array | undefined): MetafileSvgModule {
  return { MetafileConverter: class { toSvg = toSvg; } };
}

test('start shares a load and conversion is unavailable until ready', async () => {
  let finish!: (module: MetafileSvgModule) => void;
  const load = mock(() => new Promise<MetafileSvgModule>((resolve) => { finish = resolve; }));
  const converter = createMetafileSvg(load);
  expect(converter.convert('picture', bytes)).toBeUndefined();
  const first = converter.start();
  expect(converter.start()).toBe(first);
  await Promise.resolve();
  expect(load).toHaveBeenCalledTimes(1);
  expect(converter.ready()).toBe(false);
  finish(moduleFor((source) => source));
  await first;
  expect(converter.ready()).toBe(true);
  await converter.start();
  expect(load).toHaveBeenCalledTimes(1);
});

test('a failed start is retryable', async () => {
  let attempts = 0;
  const load = mock(async () => {
    if (++attempts === 1) throw new Error('unavailable');
    return moduleFor(() => bytes);
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
  const converter = createMetafileSvg(async () => moduleFor(convert));
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
  const converter = createMetafileSvg(async () => moduleFor(convert));
  await converter.start();
  expect(converter.convert('picture', bytes)).toBeUndefined();
  expect(converter.convert('picture', bytes)).toBeUndefined();
  expect(convert).toHaveBeenCalledTimes(1);
});

test('oversized sources do not reach the converter', async () => {
  const convert = mock(() => undefined);
  const converter = createMetafileSvg(async () => moduleFor(convert));
  await converter.start();
  expect(converter.convert('large', new Uint8Array(16 * 1024 * 1024 + 1))).toBeUndefined();
  expect(convert).not.toHaveBeenCalled();
});

test('one converter is reused across pictures and isolated from other sessions', async () => {
  const constructed = mock(() => {});
  const load = async () => ({ MetafileConverter: class {
    private conversions = 0;
    constructor() { constructed(); }
    toSvg() { return new Uint8Array([++this.conversions]); }
  } });
  const first = createMetafileSvg(load);
  const second = createMetafileSvg(load);
  await first.start();
  await first.start();
  expect(first.convert('first', bytes)).toEqual(new Uint8Array([1]));
  expect(first.convert('second', bytes)).toEqual(new Uint8Array([2]));
  expect(first.convert('first', bytes)).toEqual(new Uint8Array([1]));
  expect(constructed).toHaveBeenCalledTimes(1);
  await second.start();
  expect(second.convert('first', bytes)).toEqual(new Uint8Array([1]));
  expect(constructed).toHaveBeenCalledTimes(2);
});

test('cached output owns its bytes', async () => {
  const output = bytes.slice();
  const converter = createMetafileSvg(async () => moduleFor(() => output));
  await converter.start();
  const first = converter.convert('first', bytes);
  output.fill(0);
  expect(first).toEqual(bytes);
  expect(converter.convert('first', bytes)).toBe(first);
});
