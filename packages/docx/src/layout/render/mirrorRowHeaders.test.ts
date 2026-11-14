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

test('a header-row cell with nothing to read is a plain cell, and one holding a named picture is a header', () => {
  const picture = (cell: object, x: number, y: number) => ({
    kind: 'image', x, y, w: 20, h: 20, src: 'x', altText: 'Logo', blockKey: 't', table, cell,
  }) as unknown as DisplayPrimitive;
  const mirror = buildMirrorPage({
    pageIndex: 0, width: 200, height: 200,
    primitives: [
      caret(ref(0, 0, { isHeader: true }), 0, 0),
      run(ref(0, 1, { isHeader: true }), 50, 0, 'Q1'),
      picture(ref(0, 2, { isHeader: true }), 100, 0),
      picture(ref(1, 0, { isRowHeader: true }), 0, 30),
      run(ref(1, 1, { headerIds: ['t-r0-c1'] }), 50, 30, '12'),
    ],
  });
  expect(Array.from(mirror.querySelectorAll('.layout-table-cell'), (cell) => cell.getAttribute('role'))).toEqual([
    'cell', 'columnheader', 'columnheader', 'rowheader', 'cell',
  ]);
  const links = buildMirrorPageLinks({ pageIndex: 0, width: 200, height: 200, primitives: [caret(ref(0, 0, { isHeader: true }), 0, 0)] });
  expect(links.querySelector('.layout-table-cell')).toBeNull();
});

test('a header holding only what assistive technology cannot read is a plain cell', () => {
  const chart = { label: 'Sales', frame: { x: 0, y: 0, w: 40, h: 20 } };
  const mirror = buildMirrorPage({
    pageIndex: 0, width: 200, height: 200,
    primitives: [
      run(ref(0, 0, { isHeader: true }), 0, 0, 'Q1', { chart, decorative: true, blockKey: 'c' }),
      run(ref(0, 1, { isHeader: true }), 50, 0, 'Q2'),
    ],
  });
  expect(Array.from(mirror.querySelectorAll('.layout-table-cell'), (cell) => cell.getAttribute('role'))).toEqual(['cell', 'columnheader']);
});

test('a header holding only a nested table reads as a header', () => {
  const outer = ref(0, 0, { isHeader: true });
  const nested = { tableId: 'n', parentTableId: 't' };
  const mirror = buildMirrorPage({
    pageIndex: 0, width: 200, height: 200,
    primitives: [
      { kind: 'rect', x: 0, y: 0, w: 100, h: 50, fill: '#fff', blockKey: 't', table, cell: outer },
      run({ row: 0, col: 0, rowSpan: 1, colSpan: 1, cellId: 'n-r0-c0' }, 10, 10, 'Inner', { table: nested, blockKey: 'n' }),
    ] as unknown as DisplayPrimitive[],
  });
  const cell = mirror.querySelector('[id="t-r0-c0"]')!;
  expect(cell.querySelector('[id="n-r0-c0"]')?.textContent).toContain('Inner');
  expect(cell.getAttribute('role')).toBe('columnheader');
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
