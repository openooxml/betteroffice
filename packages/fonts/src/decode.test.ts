import { describe, expect, test } from 'bun:test';
import { Buffer } from 'node:buffer';
import { existsSync } from 'node:fs';
import * as nodeZlib from 'node:zlib';
import { decodeFontBytes } from './decode';

interface Zlib {
  zstdCompressSync?: (
    bytes: Uint8Array,
    options: {
      dictionary?: Uint8Array;
      pledgedSrcSize: number;
      params: Record<number, number>;
    },
  ) => Uint8Array;
  zstdDecompressSync?: unknown;
  constants: Record<string, number>;
}

const zlib = nodeZlib as unknown as Zlib;
const proc = globalThis as {
  process?: { getBuiltinModule?: (id: string) => unknown };
};
const native = proc.process?.getBuiltinModule?.('node:zlib') as Zlib | undefined;
const wasmUrl = new URL('../dist/decoder/fontpack.wasm', import.meta.url);
const hasWasm = existsSync(wasmUrl);
const canDecode = typeof native?.zstdDecompressSync === 'function' || hasWasm;
const canCompress = typeof zlib.zstdCompressSync === 'function';
const options = typeof native?.zstdDecompressSync === 'function' ? undefined : { wasmUrl };
const hello = new TextEncoder().encode('hello');
const helloFrame = new Uint8Array([
  0x28, 0xb5, 0x2f, 0xfd, 0x24, 0x05, 0x29, 0x00, 0x00,
  0x68, 0x65, 0x6c, 0x6c, 0x6f, 0xa3, 0x6d, 0x9f, 0x88,
]);

function synthetic(): Uint8Array<ArrayBuffer> {
  const bytes = new Uint8Array(65536);
  let state = 42;
  for (let index = 0; index < 16384; index++) {
    state ^= state << 13;
    state ^= state >>> 17;
    state ^= state << 5;
    bytes[index] = state & 255;
  }
  for (let index = 16384; index < bytes.length; index++) bytes[index] = bytes[index % 16384];
  return bytes;
}

function compress(bytes: Uint8Array, dictionary?: Uint8Array): Uint8Array {
  return zlib.zstdCompressSync!(bytes, {
    dictionary: dictionary === undefined ? undefined : Buffer.from(dictionary),
    pledgedSrcSize: bytes.length,
    params: {
      [zlib.constants.ZSTD_c_compressionLevel]: 19,
      [zlib.constants.ZSTD_c_checksumFlag]: 1,
      [zlib.constants.ZSTD_c_contentSizeFlag]: 1,
      [zlib.constants.ZSTD_c_windowLog]: 20,
    },
  });
}

describe('internal font decoder', () => {
  test('unsupported manifest codecs fail before decoding or passthrough', async () => {
    for (const bytes of [hello, helloFrame]) {
      await expect(decodeFontBytes(bytes, undefined, { codec: 'woff2t+zstd' }))
        .rejects.toThrow('Unsupported font face codec: woff2t+zstd');
      await expect(decodeFontBytes(bytes, undefined, { codec: '' }))
        .rejects.toThrow('Unsupported font face codec');
    }
    expect(new Uint8Array(await decodeFontBytes(hello, undefined, { codec: 'zstd' }))).toEqual(hello);
  });

  for (const [name, magic] of [
    ['TTF', [0, 1, 0, 0]],
    ['OTF', [0x4f, 0x54, 0x54, 0x4f]],
    ['WOFF', [0x77, 0x4f, 0x46, 0x46]],
    ['WOFF2', [0x77, 0x4f, 0x46, 0x32]],
  ] as const) {
    test(`${name} passes through in fresh owned buffers`, async () => {
      const input = new Uint8Array([...magic, 5, 6, 7]);
      for (const bytes of [input, input.buffer]) {
        const first = await decodeFontBytes(bytes);
        const second = await decodeFontBytes(bytes);
        expect(new Uint8Array(first)).toEqual(input);
        expect(first).not.toBe(input.buffer);
        expect(first).not.toBe(second);
        new Uint8Array(first)[4] = 0;
        expect(new Uint8Array(second)).toEqual(input);
      }
    });
  }

  test('passthrough copies only the input view', async () => {
    const input = new Uint8Array([9, 0, 1, 0, 0, 5, 9]);
    const result = await decodeFontBytes(input.subarray(1, 6), undefined, { maxOutput: 0 });
    expect(new Uint8Array(result)).toEqual(new Uint8Array([0, 1, 0, 0, 5]));
    expect(result).not.toBe(input.buffer);
    expect((await decodeFontBytes(new Uint8Array())).byteLength).toBe(0);
  });

  test.skipIf(!canCompress || !canDecode)('synthetic face round-trips without a dictionary', async () => {
    const bytes = synthetic();
    const frame = compress(bytes);
    const first = await decodeFontBytes(frame, undefined, options);
    const second = await decodeFontBytes(frame, undefined, options);
    expect(new Uint8Array(first)).toEqual(bytes);
    expect(first).not.toBe(second);
    structuredClone(first, { transfer: [first] });
    expect(new Uint8Array(second)).toEqual(bytes);
  });

  test.skipIf(!canCompress || !canDecode)('synthetic face round-trips with its raw dictionary', async () => {
    const dict = synthetic();
    const bytes = dict.slice();
    bytes[100] ^= 1;
    const frame = compress(bytes, dict);
    expect(frame.byteLength).toBeLessThan(compress(bytes).byteLength);
    expect(new Uint8Array(await decodeFontBytes(frame, dict, options))).toEqual(bytes);
    const wrong = dict.slice();
    wrong.fill(0);
    await expect(decodeFontBytes(frame, wrong, options)).rejects.toThrow(/corrupt|dictionary|checksum/);
  });

  test('content size is bounded before the decoder is loaded', async () => {
    const hugeHeader = new Uint8Array([
      0x28, 0xb5, 0x2f, 0xfd, 0xe4,
      0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0x7f,
    ]);
    await expect(decodeFontBytes(hugeHeader)).rejects.toThrow('too large');
    await expect(decodeFontBytes(helloFrame, undefined, { maxOutput: 4 })).rejects.toThrow('too large');
    await expect(decodeFontBytes(helloFrame, undefined, { maxOutput: -1 })).rejects.toThrow('maxOutput');
  });

  test('truncated and concatenated frames are rejected', async () => {
    await expect(decodeFontBytes(helloFrame.subarray(0, -1))).rejects.toThrow('corrupt');
    await expect(decodeFontBytes(new Uint8Array([...helloFrame, ...helloFrame]))).rejects.toThrow('one zstd frame');
  });

  test.skipIf(!canDecode)('corrupt checksum is rejected and decoded bytes pass through once', async () => {
    const corrupt = helloFrame.slice();
    corrupt[corrupt.length - 1] ^= 1;
    await expect(decodeFontBytes(corrupt, undefined, options)).rejects.toThrow(/corrupt|checksum/);
    const decoded = await decodeFontBytes(helloFrame, undefined, options);
    const again = await decodeFontBytes(decoded);
    expect(new Uint8Array(again)).toEqual(hello);
    expect(again).not.toBe(decoded);
  });

  test.skipIf(!hasWasm)('wasm path decodes without native zstd and releases each output', async () => {
    const results = await Promise.all([
      decodeFontBytes(helloFrame, undefined, { wasmUrl }),
      decodeFontBytes(helloFrame, undefined, { wasmUrl }),
    ]);
    for (const result of results) expect(new Uint8Array(result)).toEqual(hello);
    expect(results[0]).not.toBe(results[1]);
    const corrupt = helloFrame.slice();
    corrupt[corrupt.length - 1] ^= 1;
    await expect(decodeFontBytes(corrupt, undefined, { wasmUrl })).rejects.toThrow('checksum');
    expect(new Uint8Array(await decodeFontBytes(helloFrame, undefined, { wasmUrl }))).toEqual(hello);
  });

  test.skipIf(!hasWasm)('wasm path refuses a 2 GiB content size without allocating it', async () => {
    const bomb = new Uint8Array([
      0x28, 0xb5, 0x2f, 0xfd, 0xa4, 0xff, 0xff, 0xff, 0x7f,
      0x09, 0x00, 0x00, 0x41, 0x00, 0x00, 0x00, 0x00,
    ]);
    await expect(decodeFontBytes(bomb, undefined, { wasmUrl, maxOutput: 0x7fffffff })).rejects.toThrow(
      'Cannot allocate memory for font face',
    );
    expect(new Uint8Array(await decodeFontBytes(helloFrame, undefined, { wasmUrl }))).toEqual(hello);
  });

  test.skipIf(!hasWasm || !canCompress)('wasm path grows memory and decodes a raw dictionary', async () => {
    const dict = synthetic();
    const bytes = new Uint8Array(3 * 1024 * 1024);
    for (let offset = 0; offset < bytes.length; offset += dict.length) bytes.set(dict, offset);
    bytes[100] ^= 1;
    const frame = compress(bytes, dict);
    const first = await decodeFontBytes(frame, dict, { wasmUrl });
    const second = await decodeFontBytes(frame, dict, { wasmUrl });
    expect(new Uint8Array(first)).toEqual(bytes);
    expect(new Uint8Array(second)).toEqual(bytes);
    const wrong = new Uint8Array(dict.length);
    await expect(decodeFontBytes(frame, wrong, { wasmUrl })).rejects.toThrow(/corrupt|dictionary|checksum/);
  });
});
