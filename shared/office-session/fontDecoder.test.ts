import { expect, test } from 'bun:test';
import { copyDecodedFace, rpcFontDecoder } from './fontDecoder';

test('engine results copy exact font and dictionary payload ranges', () => {
  const backing = new Uint8Array([99, 0, 1, 0, 0, 99, 10, 11, 99]);
  const result = copyDecodedFace({ font: backing.subarray(1, 5), payload: backing.subarray(6, 8) });
  expect(new Uint8Array(result.font)).toEqual(new Uint8Array([0, 1, 0, 0]));
  expect(new Uint8Array(result.payload!)).toEqual(new Uint8Array([10, 11]));
  backing.fill(0);
  expect(result.font.byteLength).toBe(4);
  expect(new Uint8Array(result.payload!)).toEqual(new Uint8Array([10, 11]));
});

test('RPC copies input views and dictionaries before transferring them', async () => {
  const backing = new Uint8Array([99, 1, 2, 99]);
  const decoder = rpcFontDecoder(async (bytes, options, transfer) => {
    expect(new Uint8Array(bytes)).toEqual(new Uint8Array([1, 2]));
    expect(new Uint8Array(options.dict as ArrayBuffer)).toEqual(new Uint8Array([2]));
    expect(options.codec).toBe('woff2t+zstd');
    expect(options.maxOutput).toBe(100);
    structuredClone({ bytes, dict: options.dict }, { transfer });
    return { font: new Uint8Array([0, 1, 0, 0]).buffer };
  });
  expect((await decoder(backing.subarray(1, 3), {
    codec: 'woff2t+zstd', dict: backing.subarray(2, 3), maxOutput: 100,
  })).font.byteLength).toBe(4);
  expect(backing).toEqual(new Uint8Array([99, 1, 2, 99]));
});
