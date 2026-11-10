import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, test } from 'bun:test';
import type { DisplayPage, DisplayPrimitive } from './displayList';
import { buildMirrorPage, buildMirrorPageLinks, buildMirrorPageText, updateMirrorPage } from './mirrorDom';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const run = (value: string, x: number, extra: object = {}): DisplayPrimitive =>
  ({
    kind: 'text', x, y: 10, baselineY: 20, width: 10, text: value, font: '11px sans-serif', color: '#000', blockKey: 'p', ...extra,
  }) as unknown as DisplayPrimitive;
const page = (primitives: DisplayPrimitive[]): DisplayPage => ({ pageIndex: 0, width: 400, height: 200, primitives });
const item = run('Item', 100, { docStart: 5, docEnd: 9 });
const numbered = { listMarker: true, listMarkerNumbered: true };
const markers = (root: HTMLElement) => Array.from(root.querySelectorAll('.layout-list-marker'));

test('a numbered marker reads as one number, and a bullet stays hidden', () => {
  const mirror = buildMirrorPage(page([run('3', 70, numbered), run('.', 80, numbered), run('2', 90, numbered), item]));
  expect(markers(mirror).map((marker) => [marker.textContent, marker.getAttribute('aria-hidden')])).toEqual([['3.2', null]]);
  expect((markers(mirror)[0] as HTMLElement).style.width).toBe('30px');
  const bullet = buildMirrorPage(page([run('\u2022', 90, { listMarker: true }), item]));
  expect(markers(bullet).map((marker) => marker.getAttribute('aria-hidden'))).toEqual(['true']);
});

test('only touching marker runs on one line join, in either direction', () => {
  const numbers = (primitives: DisplayPrimitive[]) => markers(buildMirrorPage(page(primitives))).map((marker) => marker.textContent);
  expect(numbers([run('1.', 70, numbered), run('2.', 120, numbered), item])).toEqual(['1.', '2.']);
  expect(numbers([run('1.', 70, numbered), run('2.', 80, { ...numbered, y: 40, baselineY: 50 }), item])).toEqual(['1.', '2.']);
  const rtl = { ...numbered, rtl: true };
  const joined = buildMirrorPage(page([run('3', 100, rtl), run('.', 90, rtl), item]));
  expect(markers(joined).map((marker) => marker.textContent)).toEqual(['3.']);
  expect([(markers(joined)[0] as HTMLElement).style.left, (markers(joined)[0] as HTMLElement).style.width]).toEqual(['90px', '20px']);
});

test('a numbered marker takes no document position, so positions in the text stay as they were', () => {
  const positions = (root: HTMLElement) =>
    Array.from(root.querySelectorAll('[data-doc-start]'), (el) => [el.getAttribute('data-doc-start'), el.getAttribute('data-doc-end')]);
  const exposed = buildMirrorPage(page([run('1.', 90, numbered), item]));
  const hidden = buildMirrorPage(page([run('1.', 90, { listMarker: true }), item]));
  expect(markers(exposed)[0]!.hasAttribute('data-doc-start')).toBe(false);
  expect(positions(exposed)).toEqual(positions(hidden));
});

test('the text-only mirror keeps a number apart from its item and drops a bullet', () => {
  expect(buildMirrorPageText(page([run('1.', 90, numbered), item])).textContent).toBe('1. Item');
  expect(buildMirrorPageText(page([run('\u2022', 90, { listMarker: true }), item])).textContent).toBe('Item');
  expect(buildMirrorPageLinks(page([run('1.', 90, numbered), item])).textContent).not.toContain('1.');
});

test('a first-column cell holding only a number reads as a row header, and one holding only a bullet does not', () => {
  const table = { tableId: 't' };
  const cell = (row: number) => ({ row, col: 0, rowSpan: 1, colSpan: 1, cellId: `c${row}`, isRowHeader: true });
  const mirror = buildMirrorPage(page([
    run('1.', 0, { ...numbered, table, cell: cell(0), blockKey: 'a' }),
    run('\u2022', 0, { listMarker: true, table, cell: cell(1), blockKey: 'b' }),
  ]));
  expect(Array.from(mirror.querySelectorAll('.layout-table-cell'), (el) => el.getAttribute('role'))).toEqual(['rowheader', 'cell']);
});

const bullet = { listMarker: true, listMarkerBullet: true };
const labels = { bullet: 'bullet' };

test('a bullet reads as the word, never its stored letter, and takes no document position', () => {
  const mirror = buildMirrorPage(page([run('q', 90, { ...bullet, font: '11px Wingdings' }), item]), { labels });
  expect(markers(mirror).map((marker) => [marker.textContent, marker.getAttribute('aria-hidden')])).toEqual([['bullet', null]]);
  expect(markers(mirror)[0]!.hasAttribute('data-doc-start')).toBe(false);
  expect(mirror.textContent).not.toContain('q');
});

test('a bullet stays hidden without the word, and an unmarked marker stays hidden with it', () => {
  const hidden = (primitive: DisplayPrimitive, options?: { labels: typeof labels }) =>
    markers(buildMirrorPage(page([primitive, item]), options)).map((marker) => [marker.textContent, marker.getAttribute('aria-hidden')]);
  expect(hidden(run('\u2022', 90, bullet))).toEqual([['\u2022', 'true']]);
  expect(hidden(run('3.', 90, { listMarker: true }), { labels })).toEqual([['3.', 'true']]);
});

test('a bullet shaped run by run reads as one word, and does not join a number', () => {
  const words = (primitives: DisplayPrimitive[]) => markers(buildMirrorPage(page(primitives), { labels })).map((marker) => marker.textContent);
  expect(words([run('\u2022', 80, bullet), run('\u2022', 90, bullet), item])).toEqual(['bullet']);
  expect(words([run('1', 80, numbered), run('\u2022', 90, bullet), item])).toEqual(['1', 'bullet']);
});

test('the text-only mirror reads the word apart from its item, and the links mirror does not', () => {
  expect(buildMirrorPageText(page([run('\u2022', 90, bullet), item]), { labels }).textContent).toBe('bullet Item');
  expect(buildMirrorPageLinks(page([run('\u2022', 90, bullet), item]), { labels }).textContent).not.toContain('bullet');
});

test('a header cell holding only a bullet is a plain cell and names nothing, and one with text reads both', () => {
  const table = { tableId: 't' };
  const cell = (id: string, row: number) => ({ row, col: 0, rowSpan: 1, colSpan: 1, cellId: id, isRowHeader: true });
  const mirror = buildMirrorPage(page([
    run('\u2022', 0, { ...bullet, table, cell: cell('a', 0), blockKey: 'a' }),
    run('\u2022', 0, { ...bullet, table, cell: cell('b', 1), blockKey: 'b' }),
    run('Item', 20, { table, cell: cell('b', 1), blockKey: 'b', docStart: 5, docEnd: 9 }),
  ]), { labels });
  expect(Array.from(mirror.querySelectorAll('.layout-table-cell'), (el) => el.getAttribute('role'))).toEqual(['cell', 'rowheader']);
  expect(markers(mirror).map((marker) => marker.getAttribute('aria-hidden'))).toEqual(['true', null]);
});

test('an update in another language reads the new word', () => {
  const content = page([run('\u2022', 90, bullet), item]);
  const mirror = buildMirrorPage(content, { labels });
  expect(updateMirrorPage(mirror, content, { labels: { bullet: 'puce' } })).toBe(true);
  expect(markers(mirror)[0]?.textContent).toBe('puce');
});
