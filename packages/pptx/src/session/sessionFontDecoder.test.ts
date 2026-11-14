import { expect, test } from 'bun:test';
import type { FontDecoder } from '../../../../shared/office-session/fontDecoder';
import { decodeLargeFacesOutsideEngine } from './sessionFontDecoder';

const decoded = (mark: number) => ({ font: new Uint8Array([mark]).buffer });

test('large pack files decode outside the engine and small ones in it', async () => {
  const calls: string[] = [];
  const decoder = decodeLargeFacesOutsideEngine(async () => { calls.push('engine'); return decoded(1); },
    async () => async () => { calls.push('outside'); return decoded(2); });
  expect(new Uint8Array((await decoder(new Uint8Array(512 * 1024 - 1), { codec: 'zstd' })).font)[0]).toBe(1);
  expect(new Uint8Array((await decoder(new Uint8Array(512 * 1024), { codec: 'zstd' })).font)[0]).toBe(2);
  expect(calls).toEqual(['engine', 'outside']);
});

test('the engine decodes a large pack file when the outside decoder is unavailable or fails', async () => {
  const engine: FontDecoder = async () => decoded(1);
  const large = new Uint8Array(512 * 1024);
  const missing = decodeLargeFacesOutsideEngine(engine, async () => { throw new Error('no decoder'); });
  expect(new Uint8Array((await missing(large, { codec: 'zstd' })).font)[0]).toBe(1);
  const failing = decodeLargeFacesOutsideEngine(engine, async () => async () => { throw new Error('decode'); });
  expect(new Uint8Array((await failing(large, { codec: 'zstd' })).font)[0]).toBe(1);
});
