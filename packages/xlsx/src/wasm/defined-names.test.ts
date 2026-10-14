/**
 * Structural edits beside a dynamic-range defined name, across the wasm
 * boundary: the whole-column reference inside `COUNTA` and the range operator
 * applied to `INDEX` are both beyond the formula parser, and used to refuse
 * every row and column edit on the sheet they name. The spill `#`, the
 * implicit intersection `@` and a range written around whitespace are beyond
 * it too, and used to be saved on their pre-edit addresses.
 */

import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { inflateRawSync } from 'node:zlib';

import { initWasm, openWorkbook } from './loader';

const FIXTURE = resolve(import.meta.dir, '../../test-fixtures/defined-names.xlsx');
const WASM = resolve(import.meta.dir, './generated/xlsx_wasm_bg.wasm');

function fixtureBytes(): Uint8Array {
  return new Uint8Array(readFileSync(FIXTURE));
}

/** the text of one zip member, read from its local file header. */
function zipEntry(bytes: Uint8Array, wanted: string): string {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  for (let at = 0; at + 30 <= bytes.length; at += 1) {
    if (view.getUint32(at, true) !== 0x04034b50) continue;
    const nameLength = view.getUint16(at + 26, true);
    const start = at + 30 + nameLength + view.getUint16(at + 28, true);
    const name = new TextDecoder().decode(bytes.subarray(at + 30, at + 30 + nameLength));
    if (name !== wanted) continue;
    const data = bytes.subarray(start, start + view.getUint32(at + 18, true));
    return view.getUint16(at + 8, true) === 8
      ? inflateRawSync(data).toString()
      : new TextDecoder().decode(data);
  }
  throw new Error(`${wanted} is not in the archive`);
}

describe('wasm defined names', () => {
  beforeAll(() => initWasm(new Uint8Array(readFileSync(WASM))));

  it('inserts rows beside a dynamic-range name and shifts it', () => {
    const handle = openWorkbook(fixtureBytes());
    try {
      expect(() =>
        handle.applyOps([{ type: 'insertRows', sheet: 0, at: 9998, count: 1 }])
      ).not.toThrow();
      expect(zipEntry(handle.save(), 'xl/workbook.xml')).toContain(
        'Data!$A$1:INDEX(Data!$A:$A,COUNTA(Data!$A:$A))'
      );

      expect(handle.applyOps([{ type: 'insertRows', sheet: 0, at: 0, count: 1 }]).applied).toBe(
        true
      );
      const workbookXml = zipEntry(handle.save(), 'xl/workbook.xml');
      expect(workbookXml).toContain('Data!$A$2:INDEX(Data!$A:$A,COUNTA(Data!$A:$A))');
      expect(workbookXml).toContain('SUM(Data!$2:$2)');

      handle.undo();
      expect(zipEntry(handle.save(), 'xl/workbook.xml')).toContain('SUM(Data!$1:$1)');
    } finally {
      handle.dispose();
    }
  });

  it('moves the reference operators the formula parser cannot read', () => {
    const handle = openWorkbook(fixtureBytes());
    try {
      expect(handle.applyOps([{ type: 'insertRows', sheet: 0, at: 0, count: 1 }]).applied).toBe(
        true
      );
      const workbookXml = zipEntry(handle.save(), 'xl/workbook.xml');
      expect(workbookXml).toContain('SUM(Data!A2#)');
      expect(workbookXml).toContain('SUM(@Data!A2)');
      expect(workbookXml).toContain('Data!A2: Data!B3');
    } finally {
      handle.dispose();
    }
  });

  it('clips a range written around whitespace as one span', () => {
    const handle = openWorkbook(fixtureBytes());
    try {
      expect(handle.applyOps([{ type: 'deleteRows', sheet: 0, at: 0, count: 1 }]).applied).toBe(
        true
      );
      expect(zipEntry(handle.save(), 'xl/workbook.xml')).toContain('Data!A1: Data!B1');
    } finally {
      handle.dispose();
    }
  });
});
