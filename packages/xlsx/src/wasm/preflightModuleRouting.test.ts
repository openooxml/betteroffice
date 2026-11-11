import { expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { fontPreflight, preloadPreflightWasm, initWasm, openWorkbook } from '../index';
import { fontPreflightFixture } from './__fixtures__/fontPreflight';
import { preflightWasmState } from './preflight';
import { engineKind } from './engine';

test.skipIf(engineKind === 'native')('opening an engine leaves preflight unloaded until requested', async () => {
  const fetch = spyOn(globalThis, 'fetch');
  try {
    const bytes = await fontPreflightFixture();
    expect(preflightWasmState.status()).toBe('idle');
    await initWasm(readFileSync(new URL('./generated/xlsx_wasm_bg.wasm', import.meta.url)));
    const handle = openWorkbook(bytes);
    try {
      expect(handle.sheetInfo().activeSheet).toBe(0);
      expect(preflightWasmState.status()).toBe('idle');
      expect(preflightWasmState.memory()).toBeUndefined();
      expect(fetch.mock.calls.filter(([input]) => String(input).includes('xlsx_preflight'))).toHaveLength(0);
      const result = await fontPreflight(bytes);
      expect(result).toEqual({ faces: [{ family: 'Base', bold: false, italic: false, scripts: ['cjk-jp'] }, { family: 'Styled', bold: true, italic: true, scripts: [] }], complete: true });
      const corner = handle.cellRect(0, 49, 19);
      const display = handle.displayList({ x: 0, y: 0, width: corner.x + corner.w - 0.01, height: corner.y + corner.h - 0.01 });
      const engineKeys = new Set(display.commands.flatMap((command) => command.op === 'text' && command.fontFamily && !command.chart
        ? [JSON.stringify([command.fontFamily, command.bold ?? false, command.italic ?? false])] : []));
      expect([...engineKeys].sort()).toEqual(result.faces.map((face) => JSON.stringify([face.family, face.bold, face.italic])).sort());
      expect(preflightWasmState.status()).toBe('ready');
      expect(preflightWasmState.memory()).toBeInstanceOf(WebAssembly.Memory);
      await preloadPreflightWasm();
      expect(await fontPreflight(bytes)).toEqual(result);
    } finally { handle.dispose(); }
  } finally { fetch.mockRestore(); }
});
