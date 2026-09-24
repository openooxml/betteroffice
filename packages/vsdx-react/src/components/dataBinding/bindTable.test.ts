import { expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import * as vsdx from '@betteroffice/vsdx';
import { shapeDataRows } from '@betteroffice/vsdx';
import { bindRowToShape, matchColumnsToRows, type ImportedTable } from './bindTable';

const root = resolve(import.meta.dir, '../../../../..');
const wasm = await readFile(resolve(root, 'packages/vsdx/src/wasm/generated/vsdx_wasm_bg.wasm'));
const fixture = await readFile(resolve(root, 'crates/vsdx-parse/tests/fixtures/data-binding.vsdx'));
await vsdx.initWasm(wasm);

function open() {
  const handle = vsdx.openDiagram(fixture);
  const page = handle.snapshot().pages[0];
  return { handle, pageId: page.id, shapeId: page.shapes[0].id };
}

function shapeOf(handle: vsdx.DiagramHandle) {
  return handle.snapshot().pages[0].shapes[0];
}

function valueOf(handle: vsdx.DiagramHandle, rowName: string): string | null {
  const row = shapeDataRows(shapeOf(handle)).find((entry) => entry.rowName === rowName);
  return row ? row.formula ?? row.value : null;
}

const TABLE: ImportedTable = {
  headers: ['Device', 'Owner', 'Count', 'Due date'],
  rows: [['Mixer', 'Grace', '7', '4 April']],
};

test('columns match shape-data rows by label or row name, case and spacing aside', () => {
  const { handle } = open();
  try {
    const bindings = matchColumnsToRows({ headers: ['  device ', 'OWNER', 'Nothing'], rows: [] }, shapeOf(handle));
    expect(bindings).toEqual([
      { column: 0, rowName: 'Device' },
      { column: 1, rowName: 'Owner' },
    ]);
  } finally { handle.dispose(); }
});

test('a header two rows answer to binds nothing, rather than picking one of them', () => {
  const handle = vsdx.openDiagram(fixture);
  try {
    const ambiguous = handle.snapshot().pages[0].shapes[1];
    // 'Owner' and 'owner ' are two distinct rows with the same normalized label.
    expect(matchColumnsToRows({ headers: ['Owner'], rows: [] }, ambiguous)).toEqual([]);
    // An unambiguous header on the same shape still binds.
    expect(matchColumnsToRows({ headers: ['Site'], rows: [] }, ambiguous)).toEqual([{ column: 0, rowName: 'Site' }]);
  } finally { handle.dispose(); }
});

test('a row the engine would refuse is never offered as a binding', () => {
  const { handle } = open();
  try {
    // 'Due' is a Type=5 date row: the panel must not offer a bind the commit would reject.
    expect(matchColumnsToRows({ headers: ['Due date'], rows: [] }, shapeOf(handle))).toEqual([]);
  } finally { handle.dispose(); }
});

test('binding the editable columns writes them all as one undo entry', () => {
  const { handle, pageId, shapeId } = open();
  try {
    const bindings = matchColumnsToRows(TABLE, shapeOf(handle));
    expect(bindings.some((binding) => binding.rowName === 'Due')).toBe(false);
    const outcome = bindRowToShape(handle, pageId, shapeId, shapeOf(handle), TABLE, 0, bindings);
    expect(outcome.applied).toBe(true);
    expect(outcome.refusals).toHaveLength(0);
    expect(valueOf(handle, 'Device')).toBe('"Mixer"');
    expect(valueOf(handle, 'Owner')).toBe('"Grace"');
    expect(valueOf(handle, 'Count')).toBe('7');

    expect(handle.undo().applied).toBe(true);
    expect(valueOf(handle, 'Device')).toBe('"Amp"');
    expect(valueOf(handle, 'Owner')).toBe('"Ada"');
    expect(valueOf(handle, 'Count')).toBe('4');
  } finally { handle.dispose(); }
});

test('a date column bound past the filter refuses and leaves every other column unwritten', () => {
  const { handle, pageId, shapeId } = open();
  try {
    // Bypasses matchColumnsToRows deliberately: the engine stays the backstop.
    const forced = [{ column: 0, rowName: 'Device' }, { column: 2, rowName: 'Count' }, { column: 3, rowName: 'Due' }];
    const outcome = bindRowToShape(handle, pageId, shapeId, shapeOf(handle), TABLE, 0, forced);
    expect(outcome.applied).toBe(false);
    expect(outcome.refusals.map((receipt) => receipt.rowName)).toEqual(['Due']);
    expect(valueOf(handle, 'Device')).toBe('"Amp"');
    expect(valueOf(handle, 'Count')).toBe('4');
    expect(valueOf(handle, 'Due')).toBe('DATETIME(45000)');
  } finally { handle.dispose(); }
});

test('a number column takes the value unquoted, and an empty cell is skipped', () => {
  const { handle, pageId, shapeId } = open();
  try {
    const table: ImportedTable = { headers: ['Count', 'Device'], rows: [['9', '   ']] };
    const outcome = bindRowToShape(handle, pageId, shapeId, shapeOf(handle), table, 0, matchColumnsToRows(table, shapeOf(handle)));
    expect(outcome.applied).toBe(true);
    expect(outcome.receipts.map((receipt) => receipt.rowName)).toEqual(['Count']);
    expect(valueOf(handle, 'Count')).toBe('9');
    expect(valueOf(handle, 'Device')).toBe('"Amp"');
  } finally { handle.dispose(); }
});

test('a row with nothing bindable applies nothing rather than an empty write', () => {
  const { handle, pageId, shapeId } = open();
  try {
    const table: ImportedTable = { headers: ['Unrelated'], rows: [['x']] };
    const outcome = bindRowToShape(handle, pageId, shapeId, shapeOf(handle), table, 0, matchColumnsToRows(table, shapeOf(handle)));
    expect(outcome).toEqual({ receipts: [], refusals: [], applied: false });
    expect(handle.canUndo()).toBe(false);
  } finally { handle.dispose(); }
});
