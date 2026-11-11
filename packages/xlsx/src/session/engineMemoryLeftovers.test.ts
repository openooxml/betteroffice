import { afterAll, beforeAll, expect, test } from 'bun:test';
import { createRequire } from 'node:module';
import type { SessionTransport } from '../../../../shared/office-session';
import { isClientMessage } from '../../../../shared/office-session/protocol';
import { engineKind, XlsxDocument } from '../wasm/engine';
import { createTestWorkbookSession, initRecordedWasm, limitEngineMemoryHeadroom, numbersWorkbook, restoreEngineMemory } from './testHelpers';

beforeAll(initRecordedWasm);
afterAll(restoreEngineMemory);

function liveBytes(): number {
  return (createRequire(import.meta.url)('../../../../bindings/xlsx-native/index.js') as { engineHeapLiveBytes(): number }).engineHeapLiveBytes();
}

async function collect(): Promise<void> {
  for (let round = 0; round < 5; round += 1) {
    Bun.gc(true);
    await new Promise((resolve) => setTimeout(resolve, 10));
  }
}

test.skipIf(engineKind !== 'native')('a native save stays out-of-memory when documents left by earlier work are collected before it runs', async () => {
  const leftoverBytes = await numbersWorkbook(50_000);
  const leftovers: unknown[] = Array.from({ length: 6 }, () => XlsxDocument.open(leftoverBytes));
  let freed = 0;
  let collected = false;
  const collectBeforeSave = (transport: SessionTransport): SessionTransport => ({ ...transport, listen(listener) {
    return transport.listen((message) => {
      if (collected || !isClientMessage(message) || message.kind !== 'call' || !message.method.startsWith('save')) {
        listener(message);
        return;
      }
      collected = true;
      leftovers.length = 0;
      const before = liveBytes();
      void collect().then(() => {
        freed = before - liveBytes();
        listener(message);
      });
    });
  } });
  const session = await createTestWorkbookSession(await numbersWorkbook(20_000), collectBeforeSave);
  limitEngineMemoryHeadroom(64 * 2 ** 10);
  const outcome = await session.save().then(() => 'resolved', (error: unknown) => error);
  restoreEngineMemory();
  expect(collected).toBe(true);
  expect(freed).toBeGreaterThan(64 * 2 ** 10);
  expect(outcome).toMatchObject({ code: 'out-of-memory' });
  expect(session.failure).toMatchObject({ code: 'out-of-memory' });
});
