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

test('failed starts back off and stop after four attempts', async () => {
  let now = 0;
  const load = mock(async () => { throw new Error('unavailable'); });
  const converter = createMetafileSvg(load, () => now);
  await expect(converter.start()).rejects.toThrow('unavailable');
  for (const [index, delay] of [5_000, 30_000, 120_000].entries()) {
    now += delay - 1;
    await expect(converter.start()).rejects.toThrow('unavailable');
    expect(load).toHaveBeenCalledTimes(index + 1);
    now += 1;
    await expect(converter.start()).rejects.toThrow('unavailable');
    expect(load).toHaveBeenCalledTimes(index + 2);
  }
  now += 120_001;
  await expect(converter.start()).rejects.toThrow('unavailable');
  expect(converter.ready()).toBe(false);
  expect(load).toHaveBeenCalledTimes(4);
});

test('retryIn reports when the next load may start', async () => {
  let now = 0;
  const converter = createMetafileSvg(async () => { throw new Error('unavailable'); }, () => now);
  expect(converter.retryIn()).toBe(0);
  await expect(converter.start()).rejects.toThrow('unavailable');
  for (const delay of [5_000, 30_000, 120_000]) {
    expect(converter.retryIn()).toBe(delay);
    now += delay;
    expect(converter.retryIn()).toBe(0);
    await expect(converter.start()).rejects.toThrow('unavailable');
  }
  expect(converter.retryIn()).toBeUndefined();

  const ready = createMetafileSvg(async () => moduleFor((source) => source));
  await ready.start();
  expect(ready.retryIn()).toBeUndefined();
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
