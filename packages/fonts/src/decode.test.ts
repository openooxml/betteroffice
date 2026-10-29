import { describe, expect, spyOn, test } from 'bun:test';
import { Buffer } from 'node:buffer';
import { existsSync, readFileSync } from 'node:fs';
import * as nodeZlib from 'node:zlib';
import { decodeFontBytes, decodeFontFace } from './decode';

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
  test.skipIf(process.env.CI === undefined)('decoder wasm is built in CI', () => {
    expect(hasWasm).toBe(true);
  });

  test('unsupported manifest codecs fail before decoding or passthrough', async () => {
    for (const bytes of [hello, helloFrame]) {
      await expect(decodeFontBytes(bytes, undefined, { codec: 'woff2t' }))
        .rejects.toThrow('Unsupported font face codec: woff2t');
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

  test.skipIf(!hasWasm)('a trapping wasm instance is discarded and the next decode reloads it', async () => {
    const trapUrl = new URL('https://font-decoder.invalid/trap-recovery.wasm');
    const trapModule = trappingDecoder();
    const originalFetch = globalThis.fetch;
    let loads = 0;
    const fetchSpy = spyOn(globalThis, 'fetch').mockImplementation(((input: RequestInfo | URL, init?: RequestInit) => {
      const url = input instanceof Request ? input.url : String(input);
      if (url !== trapUrl.href) return originalFetch(input, init);
      loads++;
      const bytes = loads === 1 ? trapModule : new Uint8Array(readFileSync(wasmUrl));
      return Promise.resolve(new Response(bytes, { headers: { 'Content-Type': 'application/wasm' } }));
    }) as typeof fetch);
    try {
      await expect(decodeFontBytes(helloFrame, undefined, { wasmUrl: trapUrl }))
        .rejects.toThrow('Font face could not be decoded');
      expect(loads).toBe(1);
      expect(new Uint8Array(await decodeFontBytes(helloFrame, undefined, { wasmUrl: trapUrl })))
        .toEqual(hello);
      expect(loads).toBe(2);
      const corrupt = helloFrame.slice();
      corrupt[corrupt.length - 1] ^= 1;
      await expect(decodeFontBytes(corrupt, undefined, { wasmUrl: trapUrl })).rejects.toThrow('checksum');
      expect(new Uint8Array(await decodeFontBytes(helloFrame, undefined, { wasmUrl: trapUrl })))
        .toEqual(hello);
      expect(loads).toBe(2);
    } finally {
      fetchSpy.mockRestore();
    }
  });

  test.skipIf(!hasWasm)('without native zstd or wasmUrl the default decoder url loads the built wasm', async () => {
    const host = globalThis.process as unknown as { getBuiltinModule?: (id: string) => unknown };
    const original = host.getBuiltinModule;
    host.getBuiltinModule = (id: string) => (id === 'node:zlib' ? {} : original?.call(host, id));
    try {
      expect(new Uint8Array(await decodeFontBytes(helloFrame))).toEqual(hello);
      const corrupt = helloFrame.slice();
      corrupt[corrupt.length - 1] ^= 1;
      await expect(decodeFontBytes(corrupt)).rejects.toThrow('Font face checksum failed');
    } finally {
      host.getBuiltinModule = original;
    }
  });

  test.skipIf(!hasWasm)('wasm linear memory cannot grow past 1 GiB', async () => {
    const { instance } = await WebAssembly.instantiate(readFileSync(wasmUrl), {});
    const memory = instance.exports.memory as WebAssembly.Memory;
    const pages = memory.buffer.byteLength / 65536;
    expect(() => memory.grow(16384 - pages + 1)).toThrow();
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

function fixture(name: string): Uint8Array<ArrayBuffer> {
  return new Uint8Array(readFileSync(new URL(`../../../crates/fontpack/tests/fixtures/${name}`, import.meta.url)));
}

function trappingDecoder(): Uint8Array<ArrayBuffer> {
  function leb(value: number): number[] {
    const bytes: number[] = [];
    do {
      const byte = value & 127;
      value >>>= 7;
      bytes.push(byte | (value === 0 ? 0 : 128));
    } while (value !== 0);
    return bytes;
  }
  function section(id: number, bytes: number[]): number[] {
    return [id, ...leb(bytes.length), ...bytes];
  }
  const names = [
    'fp_alloc', 'fp_free', 'fp_decode', 'fp_decode_face',
    'fp_out_ptr', 'fp_out_len', 'fp_out_free',
    'fp_payload_ptr', 'fp_payload_len', 'fp_payload_free',
  ];
  const exports = [1 + names.length, 6, ...new TextEncoder().encode('memory'), 2, 0];
  for (const name of names) {
    const bytes = new TextEncoder().encode(name);
    exports.push(...leb(bytes.length), ...bytes, 0, 0);
  }
  return new Uint8Array([
    0, 97, 115, 109, 1, 0, 0, 0,
    ...section(1, [1, 96, 1, 127, 1, 127]),
    ...section(3, [1, 0]),
    ...section(5, [1, 0, 1]),
    ...section(7, exports),
    ...section(10, [1, 3, 0, 0, 11]),
  ]);
}

function glyphCount(font: ArrayBuffer): number {
  const bytes = new Uint8Array(font);
  expect(bytes.subarray(0, 4)).toEqual(new Uint8Array([0, 1, 0, 0]));
  const data = new DataView(font);
  for (let index = 0; index < data.getUint16(4); index++) {
    const record = 12 + index * 16;
    if (data.getUint32(record) === 0x6d617870) {
      return data.getUint16(data.getUint32(record + 8) + 4);
    }
  }
  throw new Error('Missing maxp table');
}

describe('internal woff2t font decoder', () => {
  const woff2t = { wasmUrl, codec: 'woff2t+zstd' };

  test('woff2t inputs above 64 MiB are refused before the decoder sees them', async () => {
    const oversized = new Uint8Array(64 * 1024 * 1024 + 1);
    await expect(decodeFontFace(helloFrame, oversized, { ...woff2t, maxOutput: 0x7fffffff }))
      .rejects.toThrow('too large');
  });

  test.skipIf(!hasWasm)('reconstructs synthetic faces and returns the raw dictionary payload', async () => {
    const regularFrame = fixture('Synthetic-Regular.ttf.zst');
    const boldFrame = fixture('Synthetic-Bold.ttf.zst');
    const regular = await decodeFontFace(regularFrame, undefined, woff2t);
    expect(new Uint8Array(regular.payload!)).toEqual(fixture('Synthetic-Regular.uw2'));
    expect(glyphCount(regular.font)).toBe(5);
    const bold = await decodeFontFace(boldFrame, regular.payload, woff2t);
    const independent = await decodeFontFace(fixture('Synthetic-Bold.nodict.zst'), undefined, woff2t);
    expect(new Uint8Array(bold.font)).toEqual(new Uint8Array(independent.font));
    expect(new Uint8Array(bold.payload!)).toEqual(fixture('Synthetic-Bold.uw2'));
    expect(glyphCount(bold.font)).toBe(glyphCount(new Uint8Array(fixture('Synthetic-Bold.ttf')).buffer));
    expect(new Uint8Array(await decodeFontBytes(boldFrame, regular.payload, woff2t)))
      .toEqual(new Uint8Array(bold.font));
    expect(bold.font).not.toBe(bold.payload);
    expect(new Uint8Array(regular.font)).not.toEqual(new Uint8Array(regular.payload!));
  });

  test.skipIf(!hasWasm)('wrong dictionaries, checksum errors and reconstruction failures release outputs', async () => {
    const regularFrame = fixture('Synthetic-Regular.ttf.zst');
    const boldFrame = fixture('Synthetic-Bold.ttf.zst');
    const wrong = new Uint8Array(fixture('Synthetic-Regular.uw2').length);
    await expect(decodeFontFace(boldFrame, wrong, woff2t)).rejects.toThrow(/corrupt|dictionary|checksum/);
    await expect(decodeFontFace(boldFrame, undefined, woff2t)).rejects.toThrow(/corrupt|dictionary|checksum/);
    const regularFont = await decodeFontBytes(regularFrame, undefined, woff2t);
    await expect(decodeFontFace(boldFrame, regularFont, woff2t)).rejects.toThrow(/corrupt|dictionary|checksum/);
    const corrupt = regularFrame.slice();
    corrupt[corrupt.length - 1] ^= 1;
    await expect(decodeFontFace(corrupt, undefined, woff2t)).rejects.toThrow('checksum');
    await expect(decodeFontFace(helloFrame, undefined, woff2t)).rejects.toThrow('WOFF2 reconstruction');
    expect(glyphCount((await decodeFontFace(regularFrame, undefined, woff2t)).font)).toBe(5);
  });

  test.skipIf(!hasWasm)('bounds both the transformed payload and reconstructed font', async () => {
    const frame = fixture('Synthetic-Regular.ttf.zst');
    const payload = fixture('Synthetic-Regular.uw2');
    await expect(decodeFontFace(frame, undefined, { ...woff2t, maxOutput: payload.length - 1 }))
      .rejects.toThrow('too large');
    const decoded = await decodeFontFace(frame, undefined, woff2t);
    expect(decoded.font.byteLength).toBeGreaterThan(payload.length);
    await expect(decodeFontFace(frame, undefined, { ...woff2t, maxOutput: decoded.font.byteLength - 1 }))
      .rejects.toThrow('WOFF2 reconstruction');
  });

  test.skipIf(!hasWasm)('font and payload buffers remain owned across calls and transfers', async () => {
    const frame = fixture('Synthetic-Regular.ttf.zst');
    const [first, second] = await Promise.all([
      decodeFontFace(frame, undefined, woff2t),
      decodeFontFace(frame, undefined, woff2t),
    ]);
    expect(first.font).not.toBe(second.font);
    expect(first.payload).not.toBe(second.payload);
    structuredClone(first.font, { transfer: [first.font] });
    structuredClone(first.payload!, { transfer: [first.payload!] });
    expect(glyphCount(second.font)).toBe(5);
    const bold = await decodeFontFace(fixture('Synthetic-Bold.ttf.zst'), second.payload, woff2t);
    expect(glyphCount(bold.font)).toBe(5);
    const plain = await decodeFontFace(helloFrame, undefined, { wasmUrl, codec: 'zstd' });
    expect(new Uint8Array(plain.font)).toEqual(hello);
    expect(plain.payload).toBeUndefined();
  });

  test.skipIf(!hasWasm)('raw ABI clears both output cells on errors and legacy decodes', async () => {
    const { instance } = await WebAssembly.instantiate(readFileSync(wasmUrl), {});
    const wasm = instance.exports as unknown as {
      memory: WebAssembly.Memory;
      fp_alloc(len: number): number;
      fp_free(ptr: number, len: number): void;
      fp_decode(...args: number[]): number;
      fp_decode_face(...args: number[]): number;
      fp_out_ptr(): number;
      fp_out_len(): number;
      fp_out_free(): void;
      fp_payload_ptr(): number;
      fp_payload_len(): number;
      fp_payload_free(): void;
    };
    const frame = fixture('Synthetic-Regular.ttf.zst');
    const ptr = wasm.fp_alloc(frame.length) >>> 0;
    expect(ptr).not.toBe(0);
    try {
      new Uint8Array(wasm.memory.buffer, ptr, frame.length).set(frame);
      expect(wasm.fp_decode_face(ptr, frame.length, 0, 0, 65536, 1)).toBe(0);
      expect(wasm.fp_out_len()).toBeGreaterThan(0);
      expect(wasm.fp_payload_len()).toBe(fixture('Synthetic-Regular.uw2').length);
      expect(wasm.fp_decode_face(ptr, frame.length, 0, 0, 65536, 2)).toBe(-6);
      expect(wasm.fp_out_ptr()).toBe(0);
      expect(wasm.fp_payload_ptr()).toBe(0);
      expect(wasm.fp_decode_face(ptr, frame.length, 0, 0, 65536, 1)).toBe(0);
      expect(wasm.fp_decode(ptr, frame.length, 0, 0, 65536)).toBe(0);
      expect(wasm.fp_out_len()).toBe(fixture('Synthetic-Regular.uw2').length);
      expect(wasm.fp_payload_len()).toBe(0);
      expect(wasm.fp_decode_face(ptr, 1, 0, 0, 65536, 1)).toBe(-1);
      expect(wasm.fp_out_len()).toBe(0);
      expect(wasm.fp_payload_len()).toBe(0);
    } finally {
      wasm.fp_out_free();
      wasm.fp_payload_free();
      wasm.fp_free(ptr, frame.length);
    }
  });

  test('recognized woff2t codec preserves passthrough ownership', async () => {
    const source = new Uint8Array([0, 1, 0, 0, 5]);
    const result = await decodeFontFace(source, undefined, { codec: 'woff2t+zstd' });
    expect(new Uint8Array(result.font)).toEqual(source);
    expect(result.font).not.toBe(source.buffer);
    expect(result.payload).toBeUndefined();
  });
});
