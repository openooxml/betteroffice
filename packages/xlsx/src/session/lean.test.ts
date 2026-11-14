import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { isHostMessage } from '../../../../shared/office-session/protocol';
import { engineKind } from '../wasm/engine';
import { openWorkbookSession, type OpenWorkbookSessionOptions, type WorkbookSession } from './client';
import { workbookRecoveryBackend } from './clientRecovery';

const fixture = new Uint8Array(readFileSync(new URL('../../test-fixtures/sample.xlsx', import.meta.url)));
const calculation = { nowSerial: 45000, randSeed: 1 };

function worker(requests: string[]) {
  const created = new Worker(new URL('./testing/pngWorker.ts', import.meta.url), { type: 'module' });
  created.addEventListener('message', (event: MessageEvent) => {
    const message: unknown = event.data;
    if (isHostMessage(message) && message.kind === 'event' && message.name === 'png-test-fetch' &&
      typeof message.payload === 'string') requests.push(message.payload);
  });
  return created;
}

async function open(options: OpenWorkbookSessionOptions, requests: string[] = []) {
  return openWorkbookSession(fixture.slice(), { ...options, worker: () => worker(requests),
    ...(options.collaborative ? {} : { calculation }) });
}

async function edit(session: WorkbookSession, input: string) {
  const ticket = session.editSession!.submit('editCell', {
    payload: { args: [0, 2, 1, input] },
    frame: { sheet: 0, viewport: { x: 0, y: 0, width: 800, height: 600 }, zoom: 1, dpr: 1, generation: 1 },
  });
  expect((await ticket.ack).status).toBe('applied');
}

async function texts(saved: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(saved);
  const parts = Object.keys(zip.files).filter((name) => /^xl\/(sharedStrings|worksheets\/sheet\d+)\.xml$/.test(name));
  return (await Promise.all(parts.map((name) => zip.file(name)!.async('string')))).join('\n');
}

describe.skipIf(engineKind === 'native')('lean worker core', () => {
  test('a lean collaborative session joins, applies a remote update, encodes state and saves', async () => {
    const full = await open({ png: true, collaborative: true, clientId: 801 });
    const requests: string[] = [];
    let lean: WorkbookSession | undefined;
    try {
      await edit(full, 'seed');
      const initialUpdate = await full.replica!.async.encodeStateAsUpdate();
      lean = await open({ png: false, collaborative: true, clientId: 802, initialUpdate }, requests);
      await edit(full, 'remote');
      await lean.replica!.async.applyUpdate(await full.replica!.async.encodeStateAsUpdate());
      const state = await lean.replica!.async.encodeStateAsUpdate();
      expect(state.byteLength).toBeGreaterThan(0);
      expect(await texts(await lean.save())).toContain('remote');
      expect(requests.some((url) => url.endsWith('/xlsx_wasm_lean_bg.wasm'))).toBe(true);
      expect(requests.some((url) => url.endsWith('/xlsx_wasm_bg.wasm'))).toBe(false);
      expect(lean.failure).toBeUndefined();
    } finally {
      await lean?.dispose();
      await full.dispose();
    }
  }, 30_000);

  test('a lean session exports structured content and Markdown', async () => {
    const requests: string[] = [];
    const session = await open({ png: false }, requests);
    const full = await open({ png: true });
    try {
      const scope = { scope: [{ sheet: 0, range: 'A1:C5' }] };
      const content = <T>(exported: { ok: boolean; content?: T }) => {
        expect(exported.ok).toBe(true);
        return exported.content;
      };
      expect(content(await session.call.exportStructured(scope)))
        .toEqual(content(await full.call.exportStructured(scope)));
      expect(content(await session.call.exportMarkdown(scope)))
        .toEqual(content(await full.call.exportMarkdown(scope)));
      expect(requests.filter((url) => url.endsWith('xlsx_wasm_bg.wasm'))).toEqual([]);
    } finally {
      await session.dispose();
      await full.dispose();
    }
  });

  test('a lean recovery reopen from the source adopts the session version and round-trips', async () => {
    const requests: string[] = [];
    const session = await open({ png: false, includeEditState: true }, requests);
    try {
      const version = await session.call.version();
      const recovery = await workbookRecoveryBackend.open(fixture.slice(),
        { png: false, calculation, worker: () => worker(requests) }, undefined, version);
      try {
        expect(await recovery.save()).toEqual(await session.save());
      } finally { await recovery.dispose(); }
      expect(requests.some((url) => url.endsWith('/xlsx_wasm_bg.wasm'))).toBe(false);
    } finally { await session.dispose(); }
  }, 30_000);
});
