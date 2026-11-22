import { afterEach, expect, mock, spyOn, test } from 'bun:test';
import { fontPackDiagnostics as docxDiagnostics } from '../../packages/docx/src/fontPackDiagnostics';
import { createYrsSession } from '../../packages/docx/src/yrs/index';
import { ResidentProgressiveFonts } from '../../packages/docx/src/yrs/residentProgressiveFonts';
import { preloadEngines, restoreBrowser, scriptDocument, stubBrowser } from '../../packages/docx/src/yrs/renderPngTestSupport';
import { setWorkerFontAssets } from '../../packages/fonts/src/workerAssets';
import { importFonts, importProgressiveFonts } from '../../packages/fonts/src/workerFontModules';
import { fontPackDiagnostics as pptxDiagnostics, initWasm as initPptx, openPresentation } from '../../packages/pptx/src/index';
import { engineKind as pptxEngine } from '../../packages/pptx/src/wasm/engine';
import { fontPackDiagnostics as xlsxDiagnostics } from '../../packages/xlsx/src/fontPackDiagnostics';
import { engineKind as xlsxEngine, renderPng as renderXlsx } from '../../packages/xlsx/src/wasm/enginePng';
import * as fontModules from './fontModules';
import { createWorkerFontPacks, type FontPackRequest } from './optionalFontPacks';

afterEach(() => {
  restoreBrowser();
  setWorkerFontAssets(async () => undefined);
});

async function expectMissingFonts(run: () => Promise<unknown>, answer = false): Promise<void> {
  const post = mock((request: FontPackRequest): void => {
    if (answer) channel.receive({ protocol: 1, kind: 'font-pack-reply', id: request.id });
  });
  const channel = createWorkerFontPacks(post, 5);
  const load = spyOn(fontModules, 'importFonts').mockImplementation(importFonts);
  setWorkerFontAssets(channel.source);
  try {
    const error = await run().catch((reason: unknown) => reason);
    expect(error).toBeInstanceOf(TypeError);
    expect(error).toMatchObject({ message: "Failed to resolve module specifier '@betteroffice/fonts'" });
    expect(post.mock.calls.map(([request]) => request.pack)).toEqual(['base']);
  } finally {
    load.mockRestore();
    channel.dispose();
  }
}

test.skipIf(xlsxEngine === 'native')('XLSX renderPng rejects after one wait for a silent page', async () => {
  const canvas = stubBrowser();
  const displayList = JSON.stringify({ width: 100, height: 40, commands: [
    { op: 'text', x: 0, y: 0, text: '日本語', fontSize: 11, fontFamily: 'Calibri', color: '#000000' },
    { op: 'text', x: 0, y: 20, text: 'مرحبا', fontSize: 11, fontFamily: 'Arial', color: '#000000' },
  ] });
  const engine = {
    onSheet: <T>(_sheet: number, read: () => T): T => read(),
    doc: { pngDisplayListJson: () => displayList, rangePngDisplayListJson: () => displayList },
  };
  await expectMissingFonts(() => renderXlsx(engine, { sheet: 0, range: 'used' }, {}));
  expect(canvas.encoded).toBe(0);
});

async function exportPptx(answer: boolean): Promise<void> {
  const wasm = Bun.file(new URL('../../packages/pptx/src/wasm/generated/pptx_wasm_bg.wasm', import.meta.url));
  await initPptx(new Uint8Array(await wasm.arrayBuffer()));
  const slides = Bun.file(new URL('../../apps/demo/public/betteroffice-demo.pptx', import.meta.url));
  const handle = openPresentation(new Uint8Array(await slides.arrayBuffer()), { clientId: 7303 });
  const canvas = stubBrowser();
  try {
    await expectMissingFonts(() => handle.renderPng({ slide: 0 }), answer);
    expect(canvas.encoded).toBe(0);
  } finally {
    handle.dispose();
  }
}

test.skipIf(pptxEngine === 'native')('PPTX renderPng rejects after one wait for a silent page', () => exportPptx(false));
test.skipIf(pptxEngine === 'native')('PPTX renderPng preserves the missing-module TypeError when the page answers none', () => exportPptx(true));

test('DOCX renderPng rejects after one wait for a silent page', async () => {
  await preloadEngines();
  const session = await createYrsSession();
  const canvas = stubBrowser();
  try {
    session.openDocx(scriptDocument(), true);
    await expectMissingFonts(() => session.renderPng({ page: 0 }));
    expect(canvas.encoded).toBe(0);
  } finally {
    session.destroy();
  }
});

for (const [name, check] of [['XLSX', xlsxDiagnostics], ['PPTX', pptxDiagnostics], ['DOCX', docxDiagnostics]] as const) {
  test(`${name} checkFontPacks rejects after one wait for a silent page`, () => expectMissingFonts(check));
}

test('DOCX reports unavailable progressive fonts before the frame can continue', async () => {
  const post = mock((_request: FontPackRequest) => {});
  const channel = createWorkerFontPacks(post, 15);
  setWorkerFontAssets(channel.source);
  const logs: string[] = [];
  const spec = { slot: 0, source: { ref: { format: 1, index: 'slices.json', indexSha256: '' }, urls: {} } };
  const updateFontOutlines = mock(() => {});
  const resident = new ResidentProgressiveFonts([spec], [10], [], { updateFontOutlines }, async (bytes) => bytes,
    (message) => logs.push(message), importProgressiveFonts);
  const refused = mock((_owner: number) => {});
  resident.onRefused(refused);
  const glyphs = new Map([[10, new Set([1])]]);
  try {
    expect(await resident.ensure(glyphs, 0)).toEqual(new Set());
    expect(refused.mock.calls).toEqual([[10]]);
    expect(logs).toEqual(["[docx] the progressive font module is unavailable here: Failed to resolve module specifier '@betteroffice/fonts/progressive'"]);
    expect(await resident.ensure(glyphs, 0)).toEqual(new Set());
    expect(post.mock.calls.map(([request]) => request.pack)).toEqual(['base']);
    expect(updateFontOutlines).not.toHaveBeenCalled();
  } finally {
    resident.dispose();
    channel.dispose();
  }
});
