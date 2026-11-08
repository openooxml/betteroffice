import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, test } from 'bun:test';
import type { DisplayPage, DisplayPrimitive } from './displayList';
import { buildMirrorPage, buildMirrorPageLinks, buildMirrorPageText } from './mirrorDom';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const table = { tableId: 't' };
const ref = (row: number, col: number, extra: object = {}) =>
  ({ row, col, rowSpan: 1, colSpan: 1, cellId: `t-r${row}-c${col}`, ...extra });
const run = (cell: object, x: number, y: number, value: string, extra: object = {}): DisplayPrimitive =>
  ({
    kind: 'text', x, y, baselineY: y + 10, width: 30, text: value, font: '11px sans-serif', color: '#000',
    blockKey: 't', table, cell, ...extra,
  }) as unknown as DisplayPrimitive;
const caret = (cell: object, x: number, y: number): DisplayPrimitive =>
  ({
    kind: 'text', x, y, baselineY: y + 10, width: 0, text: '\u200b', font: '11px sans-serif', color: '#000',
    blockKey: 't', table, cell,
  }) as unknown as DisplayPrimitive;

/** Header row "Region | Q1", then a row headed "North" and a row whose
 *  first-column cell is empty. */
function page(): DisplayPage {
  return {
    pageIndex: 0,
    width: 200,
    height: 200,
    primitives: [
      run(ref(0, 0, { isHeader: true }), 0, 0, 'Region'),
      run(ref(0, 1, { isHeader: true }), 50, 0, 'Q1'),
      run(ref(1, 0, { isRowHeader: true }), 0, 30, 'North'),
      run(ref(1, 1, { headerIds: ['t-r0-c1', 't-r1-c0'] }), 50, 30, '12', { href: '#north' }),
      caret(ref(2, 0, { isRowHeader: true }), 0, 60),
      run(ref(2, 1, { headerIds: ['t-r0-c1', 't-r2-c0'] }), 50, 60, '8'),
    ],
  };
}

const roles = (root: HTMLElement) =>
  Array.from(root.querySelectorAll('[role="row"]'), (row) =>
    Array.from(row.querySelectorAll('.layout-table-cell'), (cell) => [cell.getAttribute('role'), (cell.textContent ?? '').replace(/\u200b/g, '')]));

test('first-column cells marked as row headers read as row headers', () => {
  expect(roles(buildMirrorPage(page()))).toEqual([
    [['columnheader', 'Region'], ['columnheader', 'Q1']],
    [['rowheader', 'North'], ['cell', '12']],
    [['cell', ''], ['cell', '8']],
  ]);
});

test('data cells name their row header beside their column header', () => {
  const mirror = buildMirrorPage(page());
  const labels = mirror.querySelector('[id="t-r1-c1"]')!.getAttribute('aria-labelledby')!.split(' ');
  expect(labels.map((id) => mirror.querySelector(`[id="${id}"]`)?.textContent)).toEqual(['Q1', 'North']);
});

test('a links-only mirror keeps the row header a linked cell names', () => {
  const links = buildMirrorPageLinks(page());
  expect(links.querySelector('[id="t-r1-c0"]')?.getAttribute('role')).toBe('rowheader');
  expect(links.querySelector('[id="t-r1-c0"]')?.textContent).toBe('North');
});

test('reduced mirrors keep a row header with its id and text even when no cell links', () => {
  const plain: DisplayPage = { ...page(), primitives: page().primitives.map((primitive) => ({ ...primitive, href: undefined }) as DisplayPrimitive) };
  for (const mirror of [buildMirrorPageLinks(plain), buildMirrorPageText(plain)]) {
    expect(mirror.querySelector('[id="t-r1-c0"]')?.getAttribute('role')).toBe('rowheader');
    expect(mirror.querySelector('[id="t-r1-c0"]')?.textContent).toBe('North');
  }
});
