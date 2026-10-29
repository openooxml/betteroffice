import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { expectFontFace, fontCases, fontErrorCases } from '../../../../bindings/font-fixtures';
import { initWasm } from './loader';
import { decodeFontFace } from './generated/vsdx_wasm.js';

beforeAll(async () => {
  await initWasm(new Uint8Array(readFileSync(new URL('./generated/vsdx_wasm_bg.wasm', import.meta.url))));
});

test('the engine exports its linked font decoder with known answers', () => {
  expect(typeof decodeFontFace).toBe('function');
  for (const entry of fontCases()) {
    const decoded = decodeFontFace(entry.bytes, entry.codec, entry.dict, entry.maxOutput);
    expectFontFace(decoded, entry);
    expect(decoded.font).toBeInstanceOf(Uint8Array);
    expect(decoded.font).not.toBe(entry.bytes);
    expect(decoded.font.buffer).not.toBe(entry.bytes.buffer);
    expect('payload' in decoded).toBe(entry.payload !== undefined);
    if (decoded.payload) {
      expect(decoded.payload).toBeInstanceOf(Uint8Array);
      expect(decoded.payload.buffer).not.toBe(decoded.font.buffer);
    }
  }
  for (const entry of fontErrorCases()) {
    expect(() => decodeFontFace(entry.bytes, entry.codec, entry.dict, entry.maxOutput))
      .toThrow(new Error(entry.message));
  }
});
