import { beforeAll, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, wasmModuleMemories, WASM32_MEMORY_LIMIT_BYTES } from './index';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const edit = () => wasmModuleMemories().find((module) => module.label === 'docx-edit')!;

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

it('counts the editing core heap as sessions grow and go', async () => {
  const before = edit();
  expect(before.bufferBytes).toBeGreaterThan(0);
  expect(before.bufferBytes).toBeLessThanOrEqual(WASM32_MEMORY_LIMIT_BYTES);
  expect(before.liveBytes).toBeLessThanOrEqual(before.bufferBytes);
  expect(before.failedAllocationBytes).toBe(0);

  const session = await createYrsSession({ clientId: 76101 });
  session.createStory('body', 'x'.repeat(4 * 1024 * 1024));
  const grown = edit();
  expect(grown.liveBytes!).toBeGreaterThan(before.liveBytes! + 4 * 1024 * 1024);
  expect(grown.peakBytes!).toBeGreaterThanOrEqual(grown.liveBytes!);
  expect(grown.bufferBytes).toBeGreaterThanOrEqual(grown.liveBytes!);

  session.destroy();
  const after = edit();
  expect(after.liveBytes!).toBeLessThan(grown.liveBytes! - 4 * 1024 * 1024);
  expect(after.peakBytes!).toBeGreaterThanOrEqual(grown.liveBytes!);
  expect(after.bufferBytes).toBeGreaterThanOrEqual(grown.bufferBytes);
});
