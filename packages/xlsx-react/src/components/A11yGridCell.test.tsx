import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import type { A11yGrid } from '@betteroffice/xlsx';
import { A11yGridTable } from './A11yGridCell';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, render } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const cell = (row: number, col: number, address: string) =>
  ({ row, col, address, text: '', selected: false, label: `${address}, blank` });
const grid: A11yGrid = {
  label: 'Sheet1',
  sheetName: 'Sheet1',
  charts: [],
  columnHeaders: [
    { col: 2, text: 'C', label: 'Column C' },
    { col: 4, text: 'E', label: 'Column E' },
  ],
  rows: [
    { row: 9, header: 'Row 10', cells: [cell(9, 2, 'C10'), cell(9, 4, 'E10')] },
    { row: 11, header: 'Row 12', cells: [cell(11, 2, 'C12'), cell(11, 4, 'E12')] },
  ],
};
const attributes = (root: Element, name: string) =>
  Array.from(root.querySelectorAll(`[${name}]`), (node) => node.getAttribute(name));

test('a scrolled viewport with hidden tracks says where each row and cell sits in the sheet', () => {
  const { container } = render(<A11yGridTable grid={grid} style={{}} />);
  expect(attributes(container, 'aria-rowindex')).toEqual(['1', '11', '13']);
  const rows = Array.from(container.querySelectorAll('[role="row"]'));
  expect(rows.map((row) => attributes(row, 'aria-colindex'))).toEqual([
    ['1', '4', '6'], ['1', '4', '6'], ['1', '4', '6'],
  ]);
});

test('the sheet size is declared unknown, since the grid holds only the viewport', () => {
  const { container } = render(<A11yGridTable grid={grid} style={{}} />);
  const root = container.querySelector('[role="grid"]');
  expect([root?.getAttribute('aria-rowcount'), root?.getAttribute('aria-colcount')]).toEqual(['-1', '-1']);
});

test('the corner is a plain cell, so every column header has a name', () => {
  const { container } = render(<A11yGridTable grid={grid} style={{}} />);
  const headers = Array.from(container.querySelectorAll('[role="columnheader"]'), (node) => node.textContent);
  expect(headers).toEqual(['Column C', 'Column E']);
  const corner = container.querySelector('[role="row"]')?.firstElementChild;
  expect([corner?.getAttribute('role'), corner?.textContent, corner?.hasAttribute('aria-selected')]).toEqual(['gridcell', '', false]);
});
