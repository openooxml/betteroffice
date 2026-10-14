import { expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import * as vsdx from './index';
import { shapeDataRows } from './shapeData';

const root = resolve(import.meta.dir, '../../..');

async function openFixture() {
  const [wasm, fixture] = await Promise.all([
    readFile(resolve(import.meta.dir, 'wasm/generated/vsdx_wasm_bg.wasm')),
    readFile(resolve(root, 'crates/vsdx-parse/tests/fixtures/guard-format.vsdx')),
  ]);
  await vsdx.initWasm(wasm);
  return vsdx.openDiagram(fixture);
}

function valueOf(handle: vsdx.DiagramHandle, shapeIndex: number, rowName: string): string | null {
  const shape = handle.snapshot().pages[0].shapes[shapeIndex];
  return shapeDataRows(shape).find((row) => row.rowName === rowName)?.formula ?? null;
}

test('setShapeData writes every row of a batch and reports one receipt each', async () => {
  const handle = await openFixture();
  try {
    const page = handle.snapshot().pages[0];
    const shape = page.shapes[0];
    const named = shapeDataRows(shape).map((row) => row.rowName).filter((name): name is string => name !== null);
    expect(named.length).toBeGreaterThan(1);
    const receipts = handle.setShapeData(page.id, shape.id, named.map((rowName) => ({ rowName, formula: `"${rowName}-set"` })));
    expect(receipts).toHaveLength(named.length);
    expect(receipts.every((receipt) => receipt.refusal === null)).toBe(true);
    for (const rowName of named) expect(valueOf(handle, 0, rowName)).toBe(`"${rowName}-set"`);
  } finally { handle.dispose(); }
});

test('one rejected row leaves every other row in the batch unwritten', async () => {
  const handle = await openFixture();
  try {
    const page = handle.snapshot().pages[0];
    const shape = page.shapes[0];
    const named = shapeDataRows(shape).map((row) => row.rowName).filter((name): name is string => name !== null);
    const before = named.map((rowName) => valueOf(handle, 0, rowName));
    const receipts = handle.setShapeData(page.id, shape.id, [
      { rowName: named[0], formula: '"written"' },
      { rowName: named[1], formula: '  ' },
    ]);
    expect(receipts[1].refusal).toContain('empty formula');
    expect(receipts[0].after).toBeNull();
    expect(named.map((rowName) => valueOf(handle, 0, rowName))).toEqual(before);
  } finally { handle.dispose(); }
});

test('missing and duplicate rows return refusals without changing the batch', async () => {
  const handle = await openFixture();
  try {
    const page = handle.snapshot().pages[0];
    const shape = page.shapes[0];
    const rowName = shapeDataRows(shape).find((row) => row.rowName !== null)!.rowName!;
    const before = handle.snapshot();
    for (const second of ['missing-row', rowName]) {
      const receipts = handle.setShapeData(page.id, shape.id, [
        { rowName, formula: '"first"' },
        { rowName: second, formula: '"second"' },
      ]);
      expect(receipts).toHaveLength(2);
      expect(receipts[1].refusal).not.toBeNull();
      expect(receipts.every((receipt) => receipt.after === null)).toBe(true);
      expect(handle.snapshot()).toEqual(before);
    }
  } finally { handle.dispose(); }
});
