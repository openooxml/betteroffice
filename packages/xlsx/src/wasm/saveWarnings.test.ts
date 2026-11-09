import { beforeAll, describe, expect, it, spyOn } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { initWasm, openWorkbook } from './loader';

const WASM = resolve(import.meta.dir, './generated/xlsx_wasm_bg.wasm');
const source = new Uint8Array(readFileSync(resolve(import.meta.dir, '../../test-fixtures/sample.xlsx')));

describe('save warnings', () => {
  beforeAll(() => initWasm(new Uint8Array(readFileSync(WASM))));

  it('an ordinary save carries none, emits nothing and writes the bytes plain save writes', () => {
    const emitted = spyOn(process, 'emitWarning').mockImplementation(() => {});
    const handle = openWorkbook(source);
    try {
      handle.editCell(0, 0, 0, '42');
      const { bytes, warnings } = handle.saveWithWarnings();
      expect(warnings).toEqual([]);
      expect(handle.save()).toEqual(bytes);
      expect(emitted.mock.calls).toEqual([]);
    } finally {
      handle.dispose();
      emitted.mockRestore();
    }
  });
});
