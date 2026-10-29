import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { expectFontFace, fontCases, fontErrorCases } from '../../../../bindings/font-fixtures';
import { decodeFontFace as decodePublicFontFace } from '../core';
import { preloadEditWasm } from './edit';
import { decodeFontFace } from './generated/edit/docx_edit.js';

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(new URL('./generated/edit/docx_edit_bg.wasm', import.meta.url))));
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

test('the root font decoder returns ArrayBuffers and rejects with engine errors', async () => {
  for (const entry of fontCases()) {
    const backing = new Uint8Array(entry.bytes.length + 2);
    backing.set(entry.bytes, 1);
    const result = decodePublicFontFace(backing.subarray(1, -1), entry);
    expect(result).toBeInstanceOf(Promise);
    const decoded = await result;
    expect(decoded.font).toBeInstanceOf(ArrayBuffer);
    if (decoded.payload !== undefined) expect(decoded.payload).toBeInstanceOf(ArrayBuffer);
    expectFontFace({ font: new Uint8Array(decoded.font),
      ...(decoded.payload === undefined ? {} : { payload: new Uint8Array(decoded.payload) }) }, entry);
    expect(backing.subarray(1, -1)).toEqual(entry.bytes);
  }
  for (const entry of fontErrorCases()) {
    await expect(decodePublicFontFace(entry.bytes.slice(), entry)).rejects.toThrow(new Error(entry.message));
  }
});
