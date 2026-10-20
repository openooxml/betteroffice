import { expect, spyOn, test } from 'bun:test';
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

function propertyCells(rowName: string, values: Record<string, string>, sectionIndex = 0): vsdx.CellSnapshot[] {
  return Object.entries(values).map(([name, value]) => ({
    locator: { sheet: { page: 1 }, shapeId: 1, section: 'Property', sectionIndex, row: { name: rowName }, cellName: name },
    name, formula: name === 'Label' || name === 'Format' ? vsdx.quoteShapeDataValue(value) : value, value,
  }));
}

function shapeWithCells(cells: vsdx.CellSnapshot[]): vsdx.ShapeSnapshot {
  return { id: 'page:1:shape:1', sourceId: 1, name: null, children: [], cells };
}

function openWithProperties(cells: vsdx.CellSnapshot[]) {
  const original = vsdx.openDiagram(fixture);
  try {
    const pageId = original.snapshot().pages[0].id;
    original.addShape(pageId, { name: 'BindingTest', cells: cells.map((cell) => ({
      locator: {
        section: 'Property', sectionIndex: cell.locator.sectionIndex ?? undefined,
        rowName: cell.locator.row && 'name' in cell.locator.row ? cell.locator.row.name : undefined,
        rowIndex: cell.locator.row && 'index' in cell.locator.row ? cell.locator.row.index : undefined,
        cellName: cell.name,
      },
      formula: cell.formula ?? '',
    })) });
    const handle = vsdx.openDiagram(original.save());
    const shapeId = handle.snapshot().pages[0].shapes.find((shape) => shape.name === 'BindingTest')!.id;
    return { handle, pageId, shapeId };
  } finally { original.dispose(); }
}

function boundShape(handle: vsdx.DiagramHandle, shapeId: string) {
  return handle.snapshot().pages[0].shapes.find((shape) => shape.id === shapeId)!;
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
      { column: 0, rowName: 'Device', sectionIndex: undefined },
      { column: 1, rowName: 'Owner', sectionIndex: undefined },
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
    expect(matchColumnsToRows({ headers: ['Site'], rows: [] }, ambiguous)).toEqual([{ column: 0, rowName: 'Site', sectionIndex: undefined }]);
  } finally { handle.dispose(); }
});

test('a date row is excluded from matching', () => {
  const { handle } = open();
  try {
    // 'Due' is a Type=5 date row: the panel must not offer a bind the commit would reject.
    expect(matchColumnsToRows({ headers: ['Due date'], rows: [] }, shapeOf(handle))).toEqual([]);
  } finally { handle.dispose(); }
});

test('same-named rows in different sections stay distinct and their shared name is ambiguous', () => {
  const shape = shapeWithCells([
    ...propertyCells('Shared', { Label: 'Primary owner', Value: '"Ada"' }, 0),
    ...propertyCells('Shared', { Label: 'Backup owner', Value: '"Grace"' }, 1),
  ]);
  const table: ImportedTable = { headers: ['Primary owner', 'Backup owner', 'Shared', 'Primary owner'], rows: [] };
  expect(matchColumnsToRows(table, shape)).toEqual([
    { column: 0, rowName: 'Shared', sectionIndex: 0 },
    { column: 1, rowName: 'Shared', sectionIndex: 1 },
  ]);
});

test('hidden properties are excluded without making visible labels ambiguous', () => {
  const shape = shapeWithCells([
    ...propertyCells('Visible', { Label: 'Owner', Value: '"Ada"' }),
    ...propertyCells('Hidden', { Label: 'Owner', Value: '"Grace"', Invisible: '1' }),
  ]);
  expect(matchColumnsToRows({ headers: ['Owner', 'Hidden'], rows: [] }, shape)).toEqual([
    { column: 0, rowName: 'Visible', sectionIndex: 0 },
  ]);
});

test('binding carries section indices to the engine and uses each section row type', () => {
  const { handle, pageId, shapeId } = openWithProperties([
    ...propertyCells('Shared', { Label: 'Name', Type: '0', Value: '"Ada"' }, 0),
    ...propertyCells('Shared', { Label: 'Count', Type: '2', Value: '4' }, 1),
  ]);
  const write = spyOn(handle, 'setShapeData');
  try {
    const table: ImportedTable = { headers: ['Name', 'Count'], rows: [['Grace', '7']] };
    const outcome = bindRowToShape(handle, pageId, shapeId, table, 0, matchColumnsToRows(table, boundShape(handle, shapeId)));
    expect(write).toHaveBeenCalledWith(pageId, shapeId, [
      { rowName: 'Shared', sectionIndex: 0, formula: '"Grace"' },
      { rowName: 'Shared', sectionIndex: 1, formula: '7' },
    ]);
    expect(outcome.applied).toBe(true);
    expect(outcome.receipts.map((receipt) => receipt.sectionIndex)).toEqual([0, 1]);
    expect(shapeDataRows(boundShape(handle, shapeId)).map((row) => row.formula)).toEqual(['"Grace"', '7']);
    expect(handle.undo().applied).toBe(true);
    expect(shapeDataRows(boundShape(handle, shapeId)).map((row) => row.formula)).toEqual(['"Ada"', '4']);
    expect(handle.canUndo()).toBe(false);
  } finally { write.mockRestore(); handle.dispose(); }
});

test('binding the editable columns writes them all as one undo entry', () => {
  const { handle, pageId, shapeId } = open();
  try {
    const bindings = matchColumnsToRows(TABLE, shapeOf(handle));
    expect(bindings.some((binding) => binding.rowName === 'Due')).toBe(false);
    const outcome = bindRowToShape(handle, pageId, shapeId, TABLE, 0, bindings);
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
    const outcome = bindRowToShape(handle, pageId, shapeId, TABLE, 0, forced);
    expect(outcome.applied).toBe(false);
    expect(outcome.refusals.map((receipt) => receipt.rowName)).toEqual(['Due']);
    expect(valueOf(handle, 'Device')).toBe('"Amp"');
    expect(valueOf(handle, 'Count')).toBe('4');
    expect(valueOf(handle, 'Due')).toBe('DATETIME(45000)');
    expect(handle.canUndo()).toBe(false);
  } finally { handle.dispose(); }
});

test('a number column takes the value unquoted, and an empty cell is skipped', () => {
  const { handle, pageId, shapeId } = open();
  try {
    const table: ImportedTable = { headers: ['Count', 'Device'], rows: [['9', '   ']] };
    const outcome = bindRowToShape(handle, pageId, shapeId, table, 0, matchColumnsToRows(table, shapeOf(handle)));
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
    const outcome = bindRowToShape(handle, pageId, shapeId, table, 0, matchColumnsToRows(table, shapeOf(handle)));
    expect(outcome).toEqual({ receipts: [], refusals: [], applied: false });
    expect(handle.canUndo()).toBe(false);
  } finally { handle.dispose(); }
});

test('binding reads current row types after the matching snapshot becomes stale', () => {
  const { handle, pageId, shapeId } = openWithProperties(propertyCells('Flag', { Type: '0', Value: '"old"' }));
  try {
    const table: ImportedTable = { headers: ['Flag'], rows: [['No']] };
    const stale = boundShape(handle, shapeId);
    const bindings = matchColumnsToRows(table, stale);
    handle.setCellFormula(pageId, shapeId, { section: 'Property', sectionIndex: 0, rowName: 'Flag', cellName: 'Type' }, '3');
    expect(shapeDataRows(stale)[0].type).toBe('string');
    const outcome = bindRowToShape(handle, pageId, shapeId, table, 0, bindings);
    expect(outcome.applied).toBe(true);
    expect(shapeDataRows(boundShape(handle, shapeId))[0].formula).toBe('0');
  } finally { handle.dispose(); }
});

test('an off-list value refuses the entire bind and names the field without adding undo history', () => {
  const { handle, pageId, shapeId } = openWithProperties([
    ...propertyCells('Device', { Value: '"Amp"' }),
    ...propertyCells('Use', { Type: '1', Format: 'Office; Conference', Value: '"Office"' }),
  ]);
  try {
    const before = handle.snapshot();
    const table: ImportedTable = { headers: ['Device', 'Use'], rows: [['Mixer', 'Warehouse']] };
    const outcome = bindRowToShape(handle, pageId, shapeId, table, 0, matchColumnsToRows(table, boundShape(handle, shapeId)));
    expect(outcome.applied).toBe(false);
    expect(outcome.refusals.map((receipt) => receipt.rowName)).toEqual(['Use']);
    expect(outcome.receipts.every((receipt) => receipt.after === null)).toBe(true);
    expect(handle.snapshot()).toEqual(before);
    expect(handle.canUndo()).toBe(false);
  } finally { handle.dispose(); }
});

test('a fixed-list value from the format options is accepted', () => {
  const { handle, pageId, shapeId } = openWithProperties(propertyCells('Use', { Type: '1', Format: 'Office; Conference', Value: '"Office"' }));
  try {
    const table: ImportedTable = { headers: ['Use'], rows: [['Conference']] };
    const outcome = bindRowToShape(handle, pageId, shapeId, table, 0, matchColumnsToRows(table, boundShape(handle, shapeId)));
    expect(outcome.applied).toBe(true);
    expect(shapeDataRows(boundShape(handle, shapeId))[0].formula).toBe('"Conference"');
  } finally { handle.dispose(); }
});

test.each(['maybe', '2', '-1', 'FALSE()'])('an unclear boolean %s refuses every field without adding undo history', (input) => {
  const { handle, pageId, shapeId } = openWithProperties([
    ...propertyCells('Device', { Value: '"Amp"' }),
    ...propertyCells('Flag', { Type: '3', Value: '1' }),
  ]);
  try {
    const before = handle.snapshot();
    const table: ImportedTable = { headers: ['Device', 'Flag'], rows: [['Mixer', input]] };
    const outcome = bindRowToShape(handle, pageId, shapeId, table, 0, matchColumnsToRows(table, boundShape(handle, shapeId)));
    expect(outcome.applied).toBe(false);
    expect(outcome.refusals.map((receipt) => receipt.rowName)).toEqual(['Flag']);
    expect(outcome.receipts.every((receipt) => receipt.after === null)).toBe(true);
    expect(handle.snapshot()).toEqual(before);
    expect(handle.canUndo()).toBe(false);
  } finally { handle.dispose(); }
});

test.each([
  ['TRUE', '1'], ['false', '0'], ['1', '1'], ['0', '0'], [' yes ', '1'], ['No', '0'],
])('boolean %s is written as %s', (input, formula) => {
  const { handle, pageId, shapeId } = openWithProperties(propertyCells('Flag', { Type: '3', Value: formula === '0' ? '1' : '0' }));
  try {
    const table: ImportedTable = { headers: ['Flag'], rows: [[input]] };
    const outcome = bindRowToShape(handle, pageId, shapeId, table, 0, matchColumnsToRows(table, boundShape(handle, shapeId)));
    expect(outcome.applied).toBe(true);
    expect(shapeDataRows(boundShape(handle, shapeId))[0].formula).toBe(formula);
  } finally { handle.dispose(); }
});
