import { expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { fontPreflight, preloadPreflightWasm, initWasm, openPresentation } from '../index';
import { fontPreflightFixture } from './__fixtures__/fontPreflight';
import { preflightWasmState } from './preflight';
import { engineKind, PptxDocument, PptxRenderer } from './engine';

test.skipIf(engineKind === 'native')('opening an engine leaves preflight unloaded until requested', async () => {
  const fetch = spyOn(globalThis, 'fetch');
  try {
    const bytes = await fontPreflightFixture();
    expect(preflightWasmState.status()).toBe('idle');
    await initWasm(readFileSync(new URL('./generated/pptx_wasm_bg.wasm', import.meta.url)));
    const handle = openPresentation(bytes, { clientId: 94601 });
    try {
      expect(handle.snapshot().slides).toHaveLength(1);
      expect(preflightWasmState.status()).toBe('idle');
      expect(preflightWasmState.memory()).toBeUndefined();
      expect(fetch.mock.calls.filter(([input]) => String(input).includes('pptx_preflight'))).toHaveLength(0);
      const result = await fontPreflight(bytes);
      expect(result).toEqual({ faces: [{ family: 'Body', bold: true, italic: false, scripts: ['cjk-jp'] }], complete: true });
      const doc = PptxDocument.openCollaborative(bytes, 94602);
      const renderer = new PptxRenderer();
      try {
        const engine = JSON.parse(renderer.requestedFontFacesJson(doc, (JSON.parse(doc.snapshotJson()) as { slides: { id: string }[] }).slides[0]!.id)) as { faces: typeof result.faces };
        expect(engine.faces.map((face) => ({ ...face, scripts: face.scripts ?? [] }))).toEqual(result.faces);
      } finally {
        renderer.free();
        doc.free();
      }
      expect(preflightWasmState.status()).toBe('ready');
      expect(preflightWasmState.memory()).toBeInstanceOf(WebAssembly.Memory);
      await preloadPreflightWasm();
      expect(await fontPreflight(bytes)).toEqual(result);
    } finally { handle.dispose(); }
  } finally { fetch.mockRestore(); }
});
