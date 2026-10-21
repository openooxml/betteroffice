import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { inflateRawSync } from 'node:zlib';

import { initWasm, openWorkbook } from './loader';

const FIXTURE = resolve(import.meta.dir, '../../test-fixtures/defined-names.xlsx');
const WASM = resolve(import.meta.dir, './generated/xlsx_wasm_bg.wasm');

function zipEntry(bytes: Uint8Array, wanted: string): string {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  for (let at = 0; at + 30 <= bytes.length; at += 1) {
    if (view.getUint32(at, true) !== 0x04034b50) continue;
    const nameLength = view.getUint16(at + 26, true);
    const start = at + 30 + nameLength + view.getUint16(at + 28, true);
    const name = new TextDecoder().decode(bytes.subarray(at + 30, at + 30 + nameLength));
    if (name !== wanted) continue;
    const data = bytes.subarray(start, start + view.getUint32(at + 18, true));
    return view.getUint16(at + 8, true) === 8 ? inflateRawSync(data).toString() : new TextDecoder().decode(data);
  }
  if (wanted.includes('worksheets/')) {
    for (let at = 0; at + 30 <= bytes.length; at += 1) {
      if (view.getUint32(at, true) !== 0x04034b50) continue;
      const nameLength = view.getUint16(at + 26, true);
      const name = new TextDecoder().decode(bytes.subarray(at + 30, at + 30 + nameLength));
      if (/worksheets\/[^/]+\.xml$/.test(name)) return zipEntry(bytes, name);
    }
  }
  throw new Error(`${wanted} is not in the archive`);
}

describe('wasm conditional formats', () => {
  beforeAll(() => initWasm(new Uint8Array(readFileSync(WASM))));

  it('adds, lists, deletes and reopens cellIs and expression rules', () => {
    const handle = openWorkbook(new Uint8Array(readFileSync(FIXTURE)));
    try {
      expect(handle.conditionalFormats(0)).toEqual([]);
      handle.applyOps([{
        type: 'addConditionalFormat', sheet: 0,
        range: { start: { row: 0, col: 0 }, end: { row: 4, col: 0 } },
        kind: 'cellIs', operator: 'greaterThan', formulas: ['5'],
        format: { fontColor: '#FF0000', bold: true, fill: '#00FF00' },
      }]);
      handle.applyOps([{
        type: 'addConditionalFormat', sheet: 0,
        range: { start: { row: 0, col: 1 }, end: { row: 4, col: 1 } },
        kind: 'expression', operator: null, formulas: ['=A1>5'],
        format: { bold: false },
      }]);
      const listed = handle.conditionalFormats(0);
      expect(listed).toHaveLength(2);
      expect(listed[0]).toMatchObject({ range: 'A1:A5', type: 'cellIs', operator: 'greaterThan', formulas: ['5'], priority: expect.any(Number), editable: true, format: { fontColor: '#FF0000', bold: true, fill: '#00FF00' } });
      expect(listed[1]).toMatchObject({ range: 'B1:B5', type: 'expression', formulas: ['A1>5'], format: { bold: false } });
      const saved = handle.save();
      expect(zipEntry(saved, 'xl/worksheets/sheet1.xml')).toContain('conditionalFormatting');
      expect(zipEntry(saved, 'xl/styles.xml')).toContain('<dxfs');
      const reopened = openWorkbook(saved);
      try {
        const again = reopened.conditionalFormats(0);
        expect(again).toHaveLength(2);
        reopened.applyOps([{ type: 'removeConditionalFormat', sheet: 0, id: again[0].id }]);
        expect(reopened.conditionalFormats(0)).toHaveLength(1);
      } finally { reopened.dispose(); }
    } finally { handle.dispose(); }
  });

  it('refuses bad arity through applyOps', () => {
    const handle = openWorkbook(new Uint8Array(readFileSync(FIXTURE)));
    try {
      expect(() => handle.applyOps([{
        type: 'addConditionalFormat', sheet: 0,
        range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
        kind: 'cellIs', operator: 'between', formulas: ['1'],
        format: {},
      }])).toThrow();
      expect(handle.conditionalFormats(0)).toEqual([]);
    } finally { handle.dispose(); }
  });
});
