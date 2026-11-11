import { afterAll, beforeAll, expect, test } from 'bun:test';
import { createRequire } from 'node:module';
import { engineKind, XlsxDocument } from '../wasm/engine';
import { createTestWorkbookSession, initRecordedWasm, limitEngineMemoryHeadroom, numbersWorkbook, ranInOwnEngineProcess, restoreEngineMemory } from './testHelpers';

beforeAll(initRecordedWasm);
afterAll(restoreEngineMemory);

function liveBytes(): number {
  return (createRequire(import.meta.url)('../../../../bindings/xlsx-native/index.js') as { engineHeapLiveBytes(): number }).engineHeapLiveBytes();
}

async function collectLeftovers(): Promise<void> {
  const leftoverBytes = await numbersWorkbook(50_000);
  const leftovers: unknown[] = Array.from({ length: 6 }, () => XlsxDocument.open(leftoverBytes));
  const before = liveBytes();
  leftovers.length = 0;
  for (let round = 0; round < 5; round += 1) {
    Bun.gc(true);
    await new Promise((resolve) => setTimeout(resolve, 10));
  }
  expect(before - liveBytes()).toBeGreaterThan(64 * 2 ** 10);
}

const name = 'leftover documents collected by the test process leave a native save over budget out-of-memory';
test.skipIf(engineKind !== 'native')(name, async () => {
  if (await ranInOwnEngineProcess(import.meta.path, name, collectLeftovers)) return;
  const session = await createTestWorkbookSession(await numbersWorkbook(20_000));
  limitEngineMemoryHeadroom(64 * 2 ** 10);
  const outcome = await session.save().then(() => 'resolved', (error: unknown) => error);
  restoreEngineMemory();
  expect(outcome).toMatchObject({ code: 'out-of-memory' });
  expect(session.failure).toMatchObject({ code: 'out-of-memory' });
}, 60_000);
