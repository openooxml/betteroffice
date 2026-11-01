import { readFileSync, readdirSync } from 'node:fs';
import { describe, expect, spyOn, test } from 'bun:test';

const manifest = JSON.parse(readFileSync(new URL('./package.json', import.meta.url), 'utf8'));
const wasmManifest = JSON.parse(readFileSync(new URL('../../packages/xlsx/package.json', import.meta.url), 'utf8'));
const fixture = readFileSync(new URL('../../crates/ooxml-opc/tests/fixtures/sample.xlsx', import.meta.url));
const dist = new URL('./dist/', import.meta.url);
const target = { sheetId: 'sheet:0', range: { kind: 'a1' as const, a1: 'B2' } };

async function entry(name: string) {
  return import(new URL(manifest.exports[name].import, import.meta.url).href);
}

describe('@betteroffice/xlsx-native', () => {
  test('exposes the entries of @betteroffice/xlsx and refuses browsers', () => {
    expect(Object.keys(manifest.exports)).toEqual(Object.keys(wasmManifest.exports));
    for (const [name, conditions] of Object.entries(wasmManifest.exports) as [string, Record<string, string>][]) {
      expect(manifest.exports[name]).toEqual({
        types: conditions.types,
        browser: './dist/browser.js',
        import: conditions.import,
      });
    }
    expect(manifest.dependencies).toBeUndefined();
    expect(JSON.parse(readFileSync(new URL('package.json', dist), 'utf8'))).toEqual({ type: 'module' });
  });

  test('the browser entry throws', async () => {
    await expect(import(new URL('browser.js', dist).href)).rejects.toThrow(
      'use @betteroffice/xlsx in browsers'
    );
  });

  test('ships no wasm engine', () => {
    for (const file of readdirSync(dist)) {
      expect(file.endsWith('.wasm')).toBe(false);
      const text = readFileSync(new URL(file, dist), 'utf8');
      expect(text).not.toContain('__wbindgen');
      expect(text).not.toContain('_bg.wasm');
    }
  });

  test('opens, recalculates, saves and reopens on the native engine', async () => {
    const instantiate = spyOn(WebAssembly, 'instantiate');
    const compile = spyOn(WebAssembly, 'compile');
    try {
      const xlsx = await entry('.');
      await xlsx.initWasm();
      expect(xlsx.isPngExportAvailable()).toBe(true);
      const workbook = xlsx.openWorkbook(fixture);
      workbook.editCell(0, 1, 1, '=1+2');
      const read = workbook.readCells({ ranges: [target] });
      expect(read.ok).toBe(true);
      expect(read.ranges[0].cells[0][0].value).toEqual({ kind: 'number', value: 3 });
      const reopened = xlsx.openWorkbook(workbook.save());
      expect(reopened.cell(0, 1, 1).input).toBe('=1+2');
      expect(reopened.readCells({ ranges: [target] }).ranges[0].cells[0][0].displayText).toBe('3');
      workbook.dispose();
      reopened.dispose();
      expect(instantiate).not.toHaveBeenCalled();
      expect(compile).not.toHaveBeenCalled();
    } finally {
      instantiate.mockRestore();
      compile.mockRestore();
    }
  });

  test('runs a worker session on the native engine', async () => {
    const { openWorkbookSession } = await entry('./worker');
    const session = await openWorkbookSession(fixture);
    try {
      const read = await session.call.readCells({ ranges: [{ ...target, sheetId: session.state.sheets[0].id }] });
      expect(read.ok).toBe(true);
      expect(new Uint8Array(await session.call.save()).subarray(0, 2)).toEqual(new Uint8Array([0x50, 0x4b]));
    } finally {
      await session.dispose();
    }
  });
});
