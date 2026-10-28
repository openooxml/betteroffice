import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { preloadEditWasm, setEditWasmHeapLimit } from './edit';
import { wasmModuleMemories } from './loadWasmAsset';
import { rezipContainer, unzipContainer } from './opc';

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(
  resolve(import.meta.dir, 'generated/edit/docx_edit_bg.wasm')
))));

test('the edit heap limit also caps OPC calls routed through the resident engine', () => {
  let seed = 1;
  const payload = Uint8Array.from({ length: 2 * 1024 * 1024 }, () => {
    seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
    return seed >>> 24;
  });
  const archive = rezipContainer({ 'payload.bin': payload });
  const before = wasmModuleMemories().find(({ label }) => label === 'docx-edit')!;
  expect(before.liveBytes).toBeDefined();
  expect(archive.byteLength).toBeGreaterThan(128 * 1024);
  try {
    setEditWasmHeapLimit(before.liveBytes! + 128 * 1024);
    expect(() => unzipContainer(archive)).toThrow();
    const capped = wasmModuleMemories().find(({ label }) => label === 'docx-edit')!;
    expect(capped.failedAllocationBytes).toBeGreaterThan(0);
  } finally {
    setEditWasmHeapLimit(undefined);
  }
  expect(unzipContainer(archive)['payload.bin']).toEqual(payload);
  expect(wasmModuleMemories().some(({ label }) => label === 'ooxml-opc')).toBe(false);
});
