import { expect, spyOn, test } from 'bun:test';
import { fontPreflight, preloadPreflightWasm } from '../core';
import { createYrsSession } from '../yrs';
import { routingDocxEntries } from './__fixtures__/routingDocx';
import { wasmModuleMemories } from './loadWasmAsset';
import { preloadOpcWasm, rezipContainer } from './opc';
import { preflightWasmState } from './preflight';

test('opening an engine leaves preflight unloaded until requested', async () => {
  const fetch = spyOn(globalThis, 'fetch');
  try {
    expect(preflightWasmState.status()).toBe('idle');
    await preloadOpcWasm();
    const bytes = rezipContainer(routingDocxEntries);
    const session = await createYrsSession({ clientId: 94601 });
    try {
      session.openDocx(bytes, true);
      expect(session.storyIds()).toContain('body');
      expect(preflightWasmState.status()).toBe('idle');
      expect(wasmModuleMemories().map(({ label }) => label)).not.toContain('docx-preflight');
      expect(fetch.mock.calls.filter(([input]) => String(input).includes('docx_preflight')))
        .toHaveLength(0);

      const result = await fontPreflight(bytes);
      expect(result).toEqual({
        faces: [{ family: 'Calibri', bold: false, italic: false, scripts: [] }],
        complete: false,
      });
      expect(preflightWasmState.status()).toBe('ready');
      expect(wasmModuleMemories().map(({ label }) => label)).toContain('docx-preflight');
      await preloadPreflightWasm();
      expect(await fontPreflight(bytes)).toEqual(result);
    } finally {
      session.destroy();
    }
  } finally {
    fetch.mockRestore();
  }
});
