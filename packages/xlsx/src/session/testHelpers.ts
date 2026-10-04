import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import type { SessionTransport } from '../../../../shared/office-session';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import { initWasm } from '../wasm/loader';
import { createWorkbookSession, type OpenWorkbookSessionOptions, type WorkbookSession } from './client';
import { createWorkbookSessionHost } from './host';

export async function loadWorkbookSessionFixtures(): Promise<{
  fixture: Uint8Array;
  chartFixture: Uint8Array;
  wasmBytes: Uint8Array<ArrayBuffer>;
}> {
  const [wasm, xlsx, charts] = await Promise.all([
    readFile(resolve(import.meta.dir, '../wasm/generated/xlsx_wasm_bg.wasm')),
    readFile(resolve(import.meta.dir, '../../test-fixtures/sample.xlsx')),
    readFile(resolve(import.meta.dir, '../../test-fixtures/charts.xlsx')),
  ]);
  const wasmBytes = new Uint8Array(wasm);
  await initWasm(wasmBytes);
  return { wasmBytes, fixture: new Uint8Array(xlsx), chartFixture: new Uint8Array(charts) };
}

export async function createTestWorkbookSession(
  bytes: Uint8Array,
  wrapHost?: (transport: SessionTransport) => SessionTransport,
  options: OpenWorkbookSessionOptions = {}
): Promise<WorkbookSession> {
  const pair = createInProcessPair();
  createWorkbookSessionHost(wrapHost ? wrapHost(pair.host) : pair.host);
  return createWorkbookSession(bytes, options, pair.client);
}

export function batchRequests(transport: SessionTransport): SessionTransport {
  return { ...transport, listen(listener) {
    let messages: unknown[] = [];
    let timer: ReturnType<typeof setTimeout> | undefined;
    const off = transport.listen((message) => {
      messages.push(message);
      if (timer !== undefined) return;
      timer = setTimeout(() => {
        timer = undefined;
        const batch = messages;
        messages = [];
        for (const queued of batch) listener(queued);
      }, 0);
    });
    return () => {
      off();
      if (timer !== undefined) clearTimeout(timer);
      messages = [];
    };
  } };
}
