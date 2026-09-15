import { expect, test } from 'bun:test';
import { formatOptions, quoteShapeDataValue, shapeDataRows, shapeDataValueFormula, unquoteFormula, visibleShapeDataRows } from './shapeData';
import type { CellSnapshot, ShapeSnapshot } from './types';

function cell(section: string, row: string, name: string, formula: string | null, value: string | null): CellSnapshot {
  return {
    locator: { sheet: { page: 1 }, shapeId: 1, section, row: { name: row }, cellName: name },
    name, formula, value,
  };
}

function shape(cells: CellSnapshot[]): ShapeSnapshot {
  return { id: 'page:1:shape:1', sourceId: 1, name: null, cells, children: [] };
}

test('groups property cells into rows with labels and display values', () => {
  const rows = shapeDataRows(shape([
    cell('Property', 'Device', 'Label', null, 'Device name'),
    cell('Property', 'Device', 'Value', '"Amp"', 'Amp'),
    cell('Property', 'Device', 'Type', null, '0'),
    cell('Property', 'Device', 'Invisible', null, '0'),
    cell('Width', '', 'Width', '2', '2'),
  ]));
  expect(rows.length).toBe(1);
  expect(rows[0].rowName).toBe('Device');
  expect(rows[0].rowIndex).toBe(null);
  expect(rows[0].label).toBe('Device name');
  expect(rows[0].type).toBe('string');
  expect(rows[0].displayValue).toBe('Amp');
  expect(rows[0].invisible).toBe(false);
});

test('falls back to row names and quoted formulas', () => {
  const rows = shapeDataRows(shape([
    cell('Property', 'Cost', 'Value', '"42"', null),
    cell('Property', 'Cost', 'Type', null, '7'),
  ]));
  expect(rows[0].label).toBe('Cost');
  expect(rows[0].type).toBe('currency');
  expect(rows[0].displayValue).toBe('42');
});

test('hides invisible rows from the visible view only', () => {
  const snapshot = shape([
    cell('Property', 'Shown', 'Value', null, 'a'),
    cell('Property', 'Hidden', 'Value', null, 'b'),
    cell('Property', 'Hidden', 'Invisible', null, '1'),
    cell('Property', 'FormulaHidden', 'Value', null, 'c'),
    cell('Property', 'FormulaHidden', 'Invisible', 'TRUE', null),
  ]);
  expect(shapeDataRows(snapshot).length).toBe(3);
  const visible = visibleShapeDataRows(snapshot);
  expect(visible.length).toBe(1);
  expect(visible[0].rowName).toBe('Shown');
});

test('supports index-addressed rows', () => {
  const indexed: CellSnapshot = {
    locator: { sheet: { page: 1 }, shapeId: 1, section: 'Property', row: { index: 2 }, cellName: 'Value' },
    name: 'Value', formula: '3', value: '3',
  };
  const rows = shapeDataRows(shape([indexed]));
  expect(rows.length).toBe(1);
  expect(rows[0].rowName).toBe(null);
  expect(rows[0].rowIndex).toBe(2);
  expect(rows[0].label).toBe('Row 2');
});

test('sorts by sort key and keeps snapshot order otherwise', () => {
  const snapshot = shape([
    cell('Property', 'Zebra', 'Value', null, '1'),
    cell('Property', 'Zebra', 'SortKey', null, 'b'),
    cell('Property', 'Alpha', 'Value', null, '2'),
    cell('Property', 'Alpha', 'SortKey', null, 'a'),
    cell('Property', 'Middle', 'Value', null, '3'),
    cell('Property', 'Middle', 'SortKey', null, 'b'),
  ]);
  expect(shapeDataRows(snapshot).map((row) => row.rowName)).toEqual(['Alpha', 'Zebra', 'Middle']);
  const unsorted = shape([
    cell('Property', 'Zebra', 'Value', null, '1'),
    cell('Property', 'Alpha', 'Value', null, '2'),
  ]);
  expect(shapeDataRows(unsorted).map((row) => row.rowName)).toEqual(['Zebra', 'Alpha']);
});

test('quotes values and builds typed formulas', () => {
  expect(quoteShapeDataValue('a"b')).toBe('"a""b"');
  expect(unquoteFormula('"a""b"')).toBe('a"b');
  expect(unquoteFormula('INDEX(0,Prop.X.Format)')).toBe(null);
  expect(shapeDataValueFormula('boolean', 'yes')).toBe('1');
  expect(shapeDataValueFormula('boolean', 'false')).toBe('0');
  expect(shapeDataValueFormula('number', ' 2.5 ')).toBe('2.5');
  expect(shapeDataValueFormula('string', 'Amp')).toBe('"Amp"');
  expect(shapeDataValueFormula('fixed-list', 'Conference')).toBe('"Conference"');
});

test('splits list formats into options', () => {
  expect(formatOptions('Sequence Flow;Message Flow;Association')).toEqual(['Sequence Flow', 'Message Flow', 'Association']);
  expect(formatOptions(null)).toEqual([]);
});
