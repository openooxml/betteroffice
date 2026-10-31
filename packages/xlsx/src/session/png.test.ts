import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { isHostMessage } from '../../../../shared/office-session/protocol';
import { engineKind } from '../wasm/engine';
import { openWorkbookSession, type WorkbookSession } from './client';

const fixture = new Uint8Array(readFileSync(new URL('../../test-fixtures/sample.xlsx', import.meta.url)));
const charts = new Uint8Array(readFileSync(new URL('../../test-fixtures/charts.xlsx', import.meta.url)));
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

function rasterRequests(requests: string[]) {
  return requests.filter((url) => url.endsWith('/xlsx_raster_bg.wasm'));
}

async function failure(call: Promise<unknown>): Promise<string> {
  try { await call; }
  catch (error) { return error instanceof Error ? error.message : String(error); }
  throw new Error('Expected a rejected export');
}

async function matchingPng(lean: WorkbookSession, full: WorkbookSession, sheet: number) {
  expect(await lean.call.renderPng(viewport, { sheet })).toEqual(await full.call.renderPng(viewport, { sheet }));
  for (const options of [undefined, { range: 'A1:B3' }, { range: 'B2:D5', scale: 0.75 },
    { range: 'A1:B3', scale: 1.3 }, { range: 'A1:B3', scale: 2 }]) {
    expect(await lean.call.renderRangePng(options, { sheet }))
      .toEqual(await full.call.renderRangePng(options, { sheet }));
  }
}

describe.skipIf(engineKind === 'native')('lazy worker PNG export', () => {
  test('loads raster only on the first export and matches full-core bytes after edits', async () => {
    const lean = await opened(false);
    const full = await opened(true);
    try {
      expect(lean.session.pngExport).toBe(true);
      expect(full.session.pngExport).toBe(true);
      expect(lean.requests.some((url) => url.endsWith('/xlsx_wasm_lean_bg.wasm'))).toBe(true);
      expect(lean.requests.some((url) => url.endsWith('/xlsx_wasm_bg.wasm'))).toBe(false);
      expect(rasterRequests(lean.requests)).toEqual([]);
      for (const session of [lean.session, full.session]) {
        await session.call.frame(viewport);
        await session.save();
      }
      expect(rasterRequests(lean.requests)).toEqual([]);
      for (const [invalid, range] of [
        [{ ...viewport, width: -1 }, undefined],
        [{ ...viewport, width: 17000 }, undefined],
        [undefined, { range: 'nope' }],
        [undefined, { range: 'A1:B2', scale: 0 }],
      ] as const) {
        const call = (session: WorkbookSession) => invalid
          ? session.call.renderPng(invalid) : session.call.renderRangePng(range);
        expect(await failure(call(lean.session))).toBe(await failure(call(full.session)));
      }
      expect(rasterRequests(lean.requests)).toEqual([]);
      const [png, range] = await Promise.all([
        lean.session.call.renderPng(viewport, { sheet: 1 }),
        lean.session.call.renderRangePng({ range: 'A1:B3', scale: 1.3 }, { sheet: 0 }),
      ]);
      expect(png).toEqual(await full.session.call.renderPng(viewport, { sheet: 1 }));
      expect(range).toEqual(await full.session.call.renderRangePng({ range: 'A1:B3', scale: 1.3 }, { sheet: 0 }));
      expect([...png.subarray(0, 8)]).toEqual([137, 80, 78, 71, 13, 10, 26, 10]);
      expect(rasterRequests(lean.requests)).toHaveLength(1);
      for (const session of [lean.session, full.session]) {
        await session.submit({ ...session.commandIdentity, basis: session.confirmedBasis, seq: 1,
          kind: 'editCell', payload: { args: [1, 0, 0, '1234'] } });
      }
      await matchingPng(lean.session, full.session, 1);
      await matchingPng(lean.session, full.session, 2);
      expect(lean.session.state.activeSheet).toBe(0);
      expect((await lean.session.call.frame(viewport)).sheet).toBe(0);
      expect(await lean.session.save()).toEqual(await full.session.save());
      expect(rasterRequests(lean.requests)).toHaveLength(1);
      expect(rasterRequests(full.requests)).toEqual([]);
      expect(lean.session.failure).toBeUndefined();
    } finally {
      await lean.session.dispose();
      await full.session.dispose();
    }
  }, 30_000);

  test('preserves chart PNG bytes across the module boundary', async () => {
    const lean = await opened(false, charts);
    const full = await opened(true, charts);
    try {
      await matchingPng(lean.session, full.session, 0);
      expect(rasterRequests(lean.requests)).toHaveLength(1);
      expect(rasterRequests(full.requests)).toEqual([]);
    } finally {
      await lean.session.dispose();
      await full.session.dispose();
    }
  }, 30_000);
});
