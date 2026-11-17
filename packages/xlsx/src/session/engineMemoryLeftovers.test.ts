import { afterAll, beforeAll, expect, test } from 'bun:test';
import type { SessionTransport } from '../../../../shared/office-session';
import { isClientMessage } from '../../../../shared/office-session/protocol';
import { engineKind, XlsxDocument } from '../wasm/engine';
import { createTestWorkbookSession, initRecordedWasm, isOwnEngineProcess, limitEngineMemoryHeadroom, numbersWorkbook, ranInOwnEngineProcess, restoreEngineMemory } from './testHelpers';

beforeAll(initRecordedWasm);
afterAll(restoreEngineMemory);

async function collect(): Promise<void> {
  for (let round = 0; round < 5; round += 1) {
    Bun.gc(true);
    await new Promise((resolve) => setTimeout(resolve, 10));
  }
}

const leftovers: unknown[] = [];

const name = 'leftover documents of the test process collected during a native save over budget leave it out-of-memory';
test.skipIf(engineKind !== 'native')(name, async () => {
  if (!isOwnEngineProcess(name)) {
    const leftoverBytes = await numbersWorkbook(50_000);
    for (let index = 0; index < 6; index += 1) leftovers.push(XlsxDocument.open(leftoverBytes));
  }
  if (await ranInOwnEngineProcess(import.meta.path, name)) {
    leftovers.length = 0;
    return;
  }
  let collected = false;
  const collectBeforeSave = (transport: SessionTransport): SessionTransport => ({ ...transport, listen(listener) {
    return transport.listen((message) => {
      if (collected || !isClientMessage(message) || message.kind !== 'call' || !message.method.startsWith('save')) {
        listener(message);
        return;
      }
      collected = true;
      leftovers.length = 0;
      void collect().then(() => listener(message));
    });
  } });
  const session = await createTestWorkbookSession(await numbersWorkbook(20_000), collectBeforeSave);
  limitEngineMemoryHeadroom(64 * 2 ** 10);
  const outcome = await session.save().then(() => 'resolved', (error: unknown) => error);
  restoreEngineMemory();
  expect(collected).toBe(true);
  expect(outcome).toMatchObject({ code: 'out-of-memory' });
  expect(session.failure).toMatchObject({ code: 'out-of-memory' });
}, 60_000);
