import { expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import * as vsdx from '@betteroffice/vsdx';
import { GATED_CELLS, isCellWriteBlocked, isDeleteBlocked, isHandleResizeBlocked, isRotateBlocked, probeKey, selectionWriteProbes } from './components/ribbon/commands';

const root = resolve(import.meta.dir, '../../..');

async function probes() {
  const [wasm, fixture] = await Promise.all([
    readFile(resolve(root, 'packages/vsdx/src/wasm/generated/vsdx_wasm_bg.wasm')),
    readFile(resolve(root, 'crates/vsdx-parse/tests/fixtures/guard-format.vsdx')),
  ]);
  await vsdx.initWasm(wasm);
  const handle = vsdx.openDiagram(fixture);
  const page = handle.snapshot().pages[0];
  const selection = page.shapes.map((shape) => ({ pageId: page.id, shapeId: shape.id, hit: { kind: 'shape' as const, shapeId: shape.id } }));
  const map = selectionWriteProbes(handle, selection, GATED_CELLS);
  const bySourceId = new Map(page.shapes.map((shape) => [shape.sourceId, map.get(probeKey(page.id, shape.id)) ?? null]));
  handle.dispose();
  return bySourceId;
}

test('a reference named like GUARD leaves the resize handles enabled', async () => {
  const byShape = await probes();
  expect(isHandleResizeBlocked(byShape.get(4))).toBe(false);
});

test('a real GUARD on Width still disables the resize handles', async () => {
  const byShape = await probes();
  expect(isHandleResizeBlocked(byShape.get(5))).toBe(true);
});

test('the engine answers each gated control for a guarded shape', async () => {
  const byShape = await probes();
  const guarded = byShape.get(1);
  expect(isCellWriteBlocked(guarded, 'FillForegnd')).toBe(true);
  expect(isRotateBlocked(guarded)).toBe(true);
  expect(isDeleteBlocked(guarded)).toBe(true);
  expect(guarded?.get('FillForegnd')?.reason).toContain('GUARD');
});

test('a SETATREF hop onto a guarded cell is refused, and an unguarded shape is not', async () => {
  const byShape = await probes();
  expect(isCellWriteBlocked(byShape.get(2), 'LineColor')).toBe(true);
  const plain = byShape.get(3);
  expect(isCellWriteBlocked(plain, 'LineColor')).toBe(false);
  expect(isCellWriteBlocked(plain, 'FillForegnd')).toBe(false);
  expect(isRotateBlocked(plain)).toBe(false);
  expect(isDeleteBlocked(plain)).toBe(false);
});

test('the engine decides each Control row, including the ones the editor used to guess at', async () => {
  const [wasm, fixture] = await Promise.all([
    readFile(resolve(root, 'packages/vsdx/src/wasm/generated/vsdx_wasm_bg.wasm')),
    readFile(resolve(root, 'crates/vsdx-parse/tests/fixtures/guard-format.vsdx')),
  ]);
  await vsdx.initWasm(wasm);
  const handle = vsdx.openDiagram(fixture);
  try {
    const page = handle.snapshot().pages[0];
    const shape = page.shapes.find((entry) => entry.sourceId === 6)!;
    const rows = ['Row_1', 'Row_2', 'Row_3'];
    const queries = rows.flatMap((rowName) => (['X', 'Y'] as const).map((cellName) => ({ section: 'Control', rowName, cellName })));
    const answers = handle.probeCellWrites(page.id, shape.id, queries);
    const allowed = Object.fromEntries(queries.map((query, index) => [`${query.rowName}.${query.cellName}`, answers[index].allowed]));

    expect(allowed['Row_1.X']).toBe(false);
    expect(answers[0].reason).toContain('GUARD');
    expect(allowed['Row_1.Y']).toBe(true);
    expect(allowed['Row_2.Y']).toBe(false);
    // A reference merely named like GUARD is not a GUARD call.
    expect(allowed['Row_3.X']).toBe(true);

    // The editor used to offer this one: a plain SETATREF redirect looked harmless to the
    // TypeScript walk, but the engine cannot resolve a sectioned reference and would refuse
    // the commit. Disabling it is the honest half of that disagreement.
    expect(allowed['Row_2.X']).toBe(false);
    expect(answers[2].reason).toContain('SETATREF target does not exist');
  } finally { handle.dispose(); }
});
