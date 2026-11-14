import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { PngExportError } from '../../../../shared/office-session/pngExportError';
import { isHostMessage } from '../../../../shared/office-session/protocol';
import { engineKind } from '../wasm/engine';
import { openWorkbookSession, type WorkbookSession } from './client';

const fixture = new Uint8Array(readFileSync(new URL('../../test-fixtures/sample.xlsx', import.meta.url)));
const viewport = { x: 13.37, y: 21.125, width: 300.3, height: 180.7 };

async function opened(png: boolean, bytes = fixture) {
  const worker = new Worker(new URL('./testing/pngWorker.ts', import.meta.url), { type: 'module' });
  const requests: string[] = [];
  worker.addEventListener('message', (event: MessageEvent) => {
    const message: unknown = event.data;
    if (isHostMessage(message) && message.kind === 'event' && message.name === 'png-test-fetch' &&
      typeof message.payload === 'string') requests.push(message.payload);
  });
  try {
    const session = await openWorkbookSession(bytes, { worker: () => worker, png,
      calculation: { nowSerial: 45000, randSeed: 1 } });
    return { session, requests };
  } catch (error) {
    worker.terminate();
    throw error;
  }
}

async function failure(call: Promise<unknown>): Promise<string> {
  try { await call; }
  catch (error) { return error instanceof Error ? error.message : String(error); }
  throw new Error('Expected a rejected export');
}

describe.skipIf(engineKind === 'native')('worker PNG export without a canvas', () => {
  test('full and lean sessions reject exports and preserve edits, frames and saves', async () => {
    const lean = await opened(false);
    const full = await opened(true);
    try {
      expect(lean.session.pngExport).toBe(false);
      expect(full.session.pngExport).toBe(false);
      expect(lean.requests.some((url) => url.endsWith('/xlsx_wasm_lean_bg.wasm'))).toBe(true);
      expect(lean.requests.some((url) => url.endsWith('/xlsx_wasm_bg.wasm'))).toBe(false);
      expect(full.requests.some((url) => url.endsWith('/xlsx_wasm_bg.wasm'))).toBe(true);
      for (const [target, options] of [
        [{ sheet: 0, viewport: { ...viewport, width: -1 } }, {}],
        [{ sheet: 0, viewport: { ...viewport, width: 17000 } }, {}],
        [{ sheet: 0, range: 'nope' }, {}],
        [{ sheet: 0, range: 'A1:B2' }, { scale: 0 }],
        [{ sheet: 9, range: 'A1:B2' }, {}],
      ] as const) {
        const call = (session: WorkbookSession) => session.call.renderPng(target, options);
        expect(await failure(call(lean.session))).toBe(await failure(call(full.session)));
      }
      for (const session of [lean.session, full.session]) {
        for (const target of [{ sheet: 1, viewport }, { sheet: 0, range: 'A1:B3' }]) {
          const png = session.call.renderPng(target);
          await expect(png).rejects.toBeInstanceOf(PngExportError);
          await expect(png).rejects.toMatchObject({
            name: 'PngExportError', code: 'canvas-unavailable', message: expect.stringContaining('@betteroffice/xlsx-native'),
          });
        }
        await session.submit({ ...session.commandIdentity, basis: session.confirmedBasis, seq: 1,
          kind: 'editCell', payload: { args: [1, 0, 0, '1234'] } });
        expect(session.state.activeSheet).toBe(0);
        expect((await session.call.frame(viewport)).sheet).toBe(0);
        expect(session.failure).toBeUndefined();
      }
      expect((await lean.session.call.frame(viewport)).displayList)
        .toEqual((await full.session.call.frame(viewport)).displayList);
      expect(await lean.session.save()).toEqual(await full.session.save());
    } finally {
      await lean.session.dispose();
      await full.session.dispose();
    }
  }, 30_000);

  test('a lean session decodes font faces without loading the full core', async () => {
    const lean = await opened(false);
    try {
      const face = await lean.session.call.decodeFontFace(new Uint8Array([0, 1, 0, 0, 10]).buffer, { codec: 'zstd' });
      expect(new Uint8Array(face.font)).toEqual(new Uint8Array([0, 1, 0, 0, 10]));
      expect(lean.requests.some((url) => url.endsWith('/xlsx_wasm_lean_bg.wasm'))).toBe(true);
      expect(lean.requests.some((url) => url.endsWith('/xlsx_wasm_bg.wasm'))).toBe(false);
      expect(lean.session.failure).toBeUndefined();
    } finally { await lean.session.dispose(); }
  }, 30_000);
});
