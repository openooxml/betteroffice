import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, test } from 'bun:test';
import type { DisplayPage, DisplayPrimitive } from './displayList';
import { buildInteractiveOverlayPage, interactiveOverlayHasTabStops } from './interactiveOverlay';
import {
  buildMirrorPage,
  buildMirrorPageLinks,
  displayPageHoldsMirrorId,
  mirrorPageHasHeaderCells,
  mirrorPageHasTabStops,
} from './mirrorDom';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const text = (extra: Partial<DisplayPrimitive>): DisplayPrimitive =>
  ({ kind: 'text', x: 10, y: 10, text: '1', font: '11px sans-serif', color: '#000', ...extra }) as DisplayPrimitive;

const pages: DisplayPage[] = [
  { pageIndex: 0, width: 100, height: 100, primitives: [text({})] },
  {
    pageIndex: 1,
    width: 100,
    height: 100,
    primitives: [text({ noteRef: { id: 7, kind: 'endnote' } })],
    noteAreas: [
      {
        kind: 'footnote',
        y: 50,
        height: 20,
        noteIds: [3],
        notes: [{ id: 3, label: '3' }],
        primitives: [text({ groupId: 'footnote-3' })],
      },
    ],
  },
  { pageIndex: 2, width: 100, height: 100, primitives: [text({ href: 'https://example.com' })] },
  {
    pageIndex: 3,
    width: 100,
    height: 100,
    primitives: [text({ inlineSdtWidget: { kind: 'checkbox', groupId: 'g', pos: 1 } })],
  },
  {
    pageIndex: 4,
    width: 100,
    height: 100,
    primitives: [
      text({ inlineSdtWidget: { kind: 'checkbox', groupId: 'l', pos: 1, locked: true } }),
    ],
    noteAreas: [
      {
        kind: 'endnote',
        y: 50,
        height: 20,
        noteIds: [9],
        primitives: [text({ groupId: 'endnote-9' })],
      },
    ],
  },
];

test('a page reports the note ids and tab stops its chrome builds', () => {
  for (const page of pages) {
    const mirror = buildMirrorPage(page);
    const overlay = buildInteractiveOverlayPage(page);
    const ids = Array.from(mirror.querySelectorAll('[id]'), (element) => element.id);
    for (const id of ids) expect(displayPageHoldsMirrorId(page, id)).toBe(true);
    for (const id of ['oox-footnote-3', 'oox-endnote-7', 'oox-noteref-endnote-7']) {
      if (!ids.includes(id)) expect(displayPageHoldsMirrorId(page, id)).toBe(false);
    }
    const stops = (root: HTMLElement) =>
      Array.from(root.querySelectorAll<HTMLButtonElement>('a[href], button')).filter(
        (element) => !element.disabled
      ).length;
    expect(mirrorPageHasTabStops(page)).toBe(stops(mirror) > 0);
    expect(interactiveOverlayHasTabStops(page)).toBe(stops(overlay) > 0);
  }
});

test("a page's links-only mirror holds the full mirror's links and ids, in order", () => {
  const links = (root: HTMLElement) =>
    Array.from(root.querySelectorAll('a'), (a) => [a.getAttribute('href'), a.id, a.textContent]);
  const ids = (root: HTMLElement) => Array.from(root.querySelectorAll('[id]'), (el) => el.id);
  for (const page of pages) {
    const full = buildMirrorPage(page);
    const linksOnly = buildMirrorPageLinks(page);
    expect(links(linksOnly)).toEqual(links(full));
    expect(ids(linksOnly)).toEqual(ids(full));
  }
});

test("a links-only mirror keeps the full mirror's nesting order and header labels", () => {
  const outer = { tableId: 't1' };
  const inner = { tableId: 't2', parentTableId: 't1' };
  const cell = (row: number, extra: object = {}) => ({
    row,
    col: 0,
    rowSpan: 1,
    colSpan: 1,
    ...extra,
  });
  const header = cell(0, { cellId: 'h', isHeader: true });
  const data = (row: number) => cell(row, { headerIds: ['h'] });
  const shade = (ref: object, y: number, h: number) =>
    ({
      kind: 'rect',
      x: 0,
      y,
      w: 200,
      h,
      fill: '#eee',
      blockKey: 'outer',
      table: outer,
      cell: ref,
    }) as unknown as DisplayPrimitive;
  const run = (key: string, table: object, ref: object, x: number, y: number, extra: object) => {
    const fields = { blockKey: key, table, cell: ref, x, baselineY: y, width: 10, ...extra };
    return text(fields as Partial<DisplayPrimitive>);
  };
  const page: DisplayPage = {
    pageIndex: 0,
    width: 200,
    height: 200,
    primitives: [
      run('inner', inner, cell(0), 20, 40, { text: 'C', href: '#c' }),
      shade(header, 0, 20),
      run('outer', outer, header, 5, 15, { text: 'Name' }),
      shade(data(1), 20, 80),
      run('outer', outer, data(1), 5, 90, { text: 'Account ', href: undefined }),
      run('outer', outer, data(1), 50, 90, { text: 'A', href: '#a' }),
      shade(data(2), 100, 40),
      run('outer', outer, data(2), 5, 120, { text: 'B', href: '#b' }),
      text({ blockKey: 'after', x: 5, baselineY: 180, width: 10, text: 'words' }),
    ],
  };
  const full = buildMirrorPage(page);
  const linksOnly = buildMirrorPageLinks(page);
  const hrefs = (root: HTMLElement) =>
    Array.from(root.querySelectorAll('a'), (a) => a.getAttribute('href'));
  expect(full.querySelector('[data-table-id="t1"] [data-table-id="t2"]')).not.toBeNull();
  expect(hrefs(linksOnly)).toEqual(hrefs(full));
  expect(linksOnly.textContent).not.toContain('words');
  // A kept cell keeps the text that names it.
  const cellText = (root: HTMLElement) =>
    Array.from(root.querySelectorAll('[role="cell"]'), (cell) => cell.textContent);
  expect(cellText(linksOnly)).toEqual(cellText(full));
  const labelled = Array.from(linksOnly.querySelectorAll('[aria-labelledby]'));
  expect(labelled.length).toBeGreaterThan(0);
  for (const element of labelled) {
    for (const id of element.getAttribute('aria-labelledby')!.split(' ')) {
      expect(linksOnly.querySelector(`[id="${id}"]`)?.textContent).toBe('Name');
    }
  }
});

test("a page's header cells stay in its links-only mirror, with their whole text", () => {
  const ref = (extra: object) => ({ row: 0, col: 0, rowSpan: 1, colSpan: 1, ...extra });
  const header = ref({ cellId: 'account', isHeader: true });
  const run = (extra: object) => {
    const fields = { blockKey: 't', table: { tableId: 't' }, cell: header, baselineY: 15, ...extra };
    return text(fields as Partial<DisplayPrimitive>);
  };
  const withLink: DisplayPage = {
    pageIndex: 0,
    width: 200,
    height: 100,
    primitives: [
      run({ x: 5, width: 40, text: 'Account ' }),
      run({ x: 50, width: 20, text: 'help', href: '#help' }),
    ],
  };
  const withoutLink: DisplayPage = { ...withLink, primitives: [withLink.primitives[0]!] };
  for (const page of [withLink, withoutLink]) {
    expect(mirrorPageHasHeaderCells(page)).toBe(true);
    const cell = (root: HTMLElement) => root.querySelector('[id="account"]')?.textContent;
    expect(cell(buildMirrorPageLinks(page))).toBe(cell(buildMirrorPage(page)));
  }
  expect(mirrorPageHasHeaderCells(pages[2]!)).toBe(false);
});
