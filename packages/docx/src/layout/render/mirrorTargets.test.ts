import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, test } from 'bun:test';
import type { DisplayPage, DisplayPrimitive } from './displayList';
import { buildInteractiveOverlayPage, interactiveOverlayHasTabStops } from './interactiveOverlay';
import {
  buildMirrorPage,
  buildMirrorPageLinks,
  buildMirrorPageText,
  displayPageHoldsMirrorId,
  mirrorPageHasHeaderCells,
  mirrorPageHasTabStops,
  reduceMirrorToLinks,
  reduceMirrorToText,
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

test('overlapping header/footer widgets stack beneath body widgets of either kind', () => {
  const widget = (groupId: string, inline: boolean, depth: number): DisplayPrimitive => ({
    kind: 'rect',
    x: 10,
    y: 10,
    w: 20,
    h: 20,
    fill: '#fff',
    ...(inline
      ? { inlineSdtWidget: { kind: 'checkbox' as const, groupId, pos: 1 } }
      : { sdt: { groupId, sdtType: 'checkbox', depth } }),
  });
  for (const bodyInline of [false, true]) {
    const page: DisplayPage = {
      pageIndex: 0,
      width: 100,
      height: 100,
      primitives: [widget('body', bodyInline, 0)],
      header: {
        rId: 'rIdHeader',
        kind: 'header',
        y: 0,
        height: 40,
        primitives: [widget('header', !bodyInline, 2)],
      },
      footer: {
        rId: 'rIdFooter',
        kind: 'footer',
        y: 60,
        height: 40,
        primitives: [widget('footer', !bodyInline, 1)],
      },
    };
    const overlay = buildInteractiveOverlayPage(page);
    const widgets = Array.from(overlay.querySelectorAll<HTMLElement>('.layout-sdt-widget'));
    expect(widgets.map((element) => element.dataset.sdtGroupId)).toEqual(['header', 'footer', 'body']);
    expect(
      widgets.map((element) =>
        (element.closest<HTMLElement>('.layout-block-sdt-box') ?? element).style.zIndex
      )
    ).toEqual(['1', '2', '3']);
  }
});

test('a covered header checkbox mirror is inert and hidden until the cover goes', () => {
  const page: DisplayPage = {
    pageIndex: 0, width: 100, height: 100,
    primitives: [{ kind: 'rect', x: 25, y: 10, w: 10, h: 20, fill: '#fff' }],
    header: {
      rId: 'rIdHeader', kind: 'header', y: 0, height: 40,
      primitives: [10, 30].map((x): DisplayPrimitive => ({
        kind: 'rect', x, y: 10, w: 20, h: 20, fill: '#fff',
        inlineSdtWidget: { kind: 'checkbox', groupId: 'header', pos: 1, checked: false },
      })),
    },
  };
  const mirrors = buildMirrorPage(page).querySelectorAll<HTMLElement>('.layout-inline-sdt-widget');
  expect(mirrors.length).toBe(2);
  const overlay = buildInteractiveOverlayPage(page)
    .querySelector<HTMLButtonElement>('.layout-inline-sdt-widget')!;
  for (const element of [...mirrors, overlay]) {
    expect(element.style.visibility).toBe('hidden');
    expect(element.style.pointerEvents).toBe('none');
    expect(element.tabIndex).toBe(-1);
    expect(element.hasAttribute('inert')).toBe(true);
    expect(element.getAttribute('aria-hidden')).toBe('true');
    expect(element.getAttribute('aria-disabled')).toBe('true');
  }
  expect(overlay.disabled).toBe(true);

  page.primitives = [];
  for (const restored of buildMirrorPage(page).querySelectorAll<HTMLElement>('.layout-inline-sdt-widget')) {
    expect(restored.style.visibility).not.toBe('hidden');
    expect(restored.style.pointerEvents).not.toBe('none');
    expect(restored.hasAttribute('tabindex')).toBe(false);
    expect(restored.hasAttribute('inert')).toBe(false);
    expect(restored.hasAttribute('aria-hidden')).toBe(false);
    expect(restored.hasAttribute('aria-disabled')).toBe(false);
    expect(restored.getAttribute('role')).toBe('checkbox');
    expect(restored.getAttribute('aria-checked')).toBe('false');
  }
});

test("a page's links-only mirror holds the full mirror's links and ids, in order", () => {
  const links = (root: HTMLElement) =>
    Array.from(root.querySelectorAll('a'), (a) => [a.getAttribute('href'), a.id, a.textContent]);
  const ids = (root: HTMLElement) => Array.from(root.querySelectorAll('[id]'), (el) => el.id);
  for (const page of pages) {
    const full = buildMirrorPage(page);
    const linksOnly = buildMirrorPageLinks(page);
    const textOnly = buildMirrorPageText(page);
    expect(links(linksOnly)).toEqual(links(full));
    expect(ids(linksOnly)).toEqual(ids(full));
    const kept = (root: HTMLElement) =>
      Array.from(root.querySelectorAll('a, .layout-note'), (element) => element.outerHTML);
    expect(kept(textOnly)).toEqual(kept(linksOnly));
    expect(ids(textOnly)).toEqual(ids(linksOnly));
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
  const textOnly = buildMirrorPageText(page);
  const hrefs = (root: HTMLElement) =>
    Array.from(root.querySelectorAll('a'), (a) => a.getAttribute('href'));
  expect(full.querySelector('[data-table-id="t1"] [data-table-id="t2"]')).not.toBeNull();
  expect(hrefs(linksOnly)).toEqual(hrefs(full));
  expect(linksOnly.textContent).not.toContain('words');
  expect(textOnly.textContent).toContain('words');
  const keptCells = (root: HTMLElement) =>
    Array.from(root.querySelectorAll('.layout-table-cell'), (cell) => cell.outerHTML);
  expect(keptCells(textOnly)).toEqual(keptCells(linksOnly));
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
    expect(buildMirrorPageText(page).querySelector('[id="account"]')?.outerHTML).toBe(
      buildMirrorPageLinks(page).querySelector('[id="account"]')?.outerHTML
    );
  }
  expect(mirrorPageHasHeaderCells(pages[2]!)).toBe(false);
});

test('a text-only mirror reads every paragraph in order without run elements or painter data', () => {
  const page: DisplayPage = {
    pageIndex: 0,
    width: 200,
    height: 200,
    primitives: [
      text({ blockKey: 'first', paraId: 'first', text: 'First ', docStart: 1, docEnd: 7 }),
      text({ blockKey: 'first', paraId: 'first', text: 'paragraph', docStart: 7, docEnd: 16 }),
      text({ blockKey: 'second', text: '1. ', listMarker: true }),
      text({ blockKey: 'second', text: 'Second paragraph', docStart: 20, docEnd: 36 }),
      text({
        blockKey: 'third', text: 'Third paragraph', docStart: 40, docEnd: 55,
        sdt: { groupId: 'control', sdtType: 'text', alias: 'Content control' },
      }),
    ],
    header: {
      rId: 'header',
      kind: 'header',
      y: 0,
      height: 20,
      primitives: [text({ blockKey: 'header', text: 'Header' })],
    },
    footer: {
      rId: 'footer',
      kind: 'footer',
      y: 180,
      height: 20,
      primitives: [text({ blockKey: 'footer', text: 'Footer' })],
    },
  };
  const mirror = buildMirrorPage(page, {
    labels: { page: 'Page', header: 'Header', footer: 'Footer' },
  });
  for (const element of [mirror, ...mirror.querySelectorAll('.layout-paragraph')]) {
    element.id = 'unused';
  }
  expect(reduceMirrorToText(mirror)).toBe(mirror);
  expect(Array.from(mirror.querySelectorAll('[role="paragraph"]'), (block) => block.textContent)).toEqual([
    'First paragraph',
    'Second paragraph',
    'Third paragraph',
    'Header',
    'Footer',
  ]);
  const paragraphs = mirror.querySelectorAll('[role="paragraph"]');
  for (const paragraph of paragraphs) {
    expect(paragraph.children).toHaveLength(0);
    expect(paragraph.childNodes).toHaveLength(1);
    expect(paragraph.firstChild?.nodeType).toBe(Node.TEXT_NODE);
  }
  expect(mirror.querySelector('.layout-run')).toBeNull();
  expect(
    [mirror, ...mirror.querySelectorAll('*')].flatMap((element) =>
      Array.from(element.attributes, (attribute) => attribute.name).filter((name) =>
        name === 'id' || name.startsWith('data-')
      )
    )
  ).toEqual([]);
  expect(mirror.getAttribute('role')).toBe('document');
  expect(mirror.getAttribute('aria-label')).toBe('Page');
  expect(mirror.querySelectorAll('[role="region"]')).toHaveLength(2);
  expect(mirror.querySelector('.layout-block-sdt')?.getAttribute('aria-label')).toBe('Content control');
  expect(mirror.style.contentVisibility).toBe('auto');
});

test('a text-only mirror keeps language changes and drops runs the full mirror hides', () => {
  const page: DisplayPage = {
    pageIndex: 0,
    width: 200,
    height: 200,
    primitives: [
      text({ blockKey: 'mixed', text: 'Hello ', lang: 'en-US' }),
      text({ blockKey: 'mixed', text: 'שלום', lang: 'he-IL', bidiLevel: 1 }),
      text({ blockKey: 'mixed', text: ' again', lang: 'en-US' }),
      text({ blockKey: 'covered', text: 'Visible' }),
      text({ blockKey: 'covered', text: ' covered' }),
    ],
  };
  const mirror = buildMirrorPage(page);
  const covered = Array.from(mirror.querySelectorAll('.layout-run')).find(
    (run) => run.textContent === ' covered'
  )!;
  covered.setAttribute('aria-hidden', 'true');
  reduceMirrorToText(mirror);
  const [mixed, visible] = Array.from(mirror.querySelectorAll('[role="paragraph"]'));
  expect(mixed!.textContent).toBe('Hello שלום again');
  expect(
    Array.from(mixed!.children, (child) => [child.getAttribute('lang'), child.textContent])
  ).toEqual([['he-IL', 'שלום']]);
  expect(visible!.textContent).toBe('Visible');
  expect(mirror.querySelector('.layout-run')).toBeNull();
});

test('a text-only mirror keeps image names and table semantics and removes visual leaves', () => {
  const page: DisplayPage = {
    pageIndex: 0,
    width: 100,
    height: 100,
    primitives: [
      text({
        blockKey: 'table',
        text: 'Cell',
        docStart: 1,
        docEnd: 5,
        table: { tableId: 'table' },
        cell: { row: 0, col: 0, rowSpan: 1, colSpan: 1 },
      }),
      {
        kind: 'image', relId: 'picture', x: 0, y: 0, w: 10, h: 10,
        altText: 'A picture', docStart: 6, docEnd: 7,
      },
      { kind: 'image', relId: 'decoration', x: 0, y: 0, w: 10, h: 10, decorative: true },
      { kind: 'rect', x: 0, y: 0, w: 10, h: 10, fill: '#fff' },
      { kind: 'line', x1: 0, y1: 0, x2: 10, y2: 10, strokeWidth: 1, color: '#000' },
      { kind: 'line', x1: 0, y1: 0, x2: 10, y2: 10, strokeWidth: 1, color: '#000', role: 'table-cut' },
      { kind: 'decoration', deco: 'underline', x: 0, y: 0, w: 10, h: 10, color: '#000' },
      text({
        text: '¶',
        structuralRevision: { scope: 'pmark', kind: 'ins', author: 'Author', revisionId: 'revision' },
      }),
      text({ text: '' }),
    ],
    pageBorders: [{ kind: 'pageBorder', x: 0, y: 0, w: 100, h: 100 }],
  };
  const mirror = buildMirrorPageText(page);
  expect(mirror.textContent).toBe('Cell');
  const image = mirror.querySelector('[role="img"]')!;
  expect(image.getAttribute('aria-label')).toBe('A picture');
  expect(image.hasAttribute('data-doc-start')).toBe(false);
  expect(mirror.querySelectorAll('.layout-run-image')).toHaveLength(1);
  expect(mirror.querySelector('[role="table"] [role="row"] [role="cell"]')?.textContent).toBe('Cell');
  expect(mirror.querySelector('[data-doc-start], [data-table-id]')).toBeNull();
  expect(mirror.querySelector([
    '.layout-decoration', '.layout-mirror-rect', '.layout-mirror-line',
    '.layout-table-cut-border', '.layout-page-border', '.layout-revision-pmark-glyph',
  ].join(', '))).toBeNull();
});

test('a text-only mirror keeps labels and attributes of ancestors of retained content', () => {
  const full = buildMirrorPage({
    pageIndex: 0,
    width: 100,
    height: 100,
    primitives: [
      text({ blockKey: 'label', text: 'Link description' }),
      text({ blockKey: 'link', href: '#target', text: 'Link', docStart: 10, docEnd: 14 }),
      text({ blockKey: 'other', text: 'Other text', docStart: 20, docEnd: 30 }),
    ],
  });
  const label = full.querySelector('.layout-paragraph')!;
  label.id = 'label';
  full.querySelector('a')!.parentElement!.setAttribute('aria-describedby', 'label');
  const linksOnly = reduceMirrorToLinks(full.cloneNode(true) as HTMLElement);
  const textOnly = reduceMirrorToText(full);
  for (const selector of ['#label', 'a', '[aria-describedby]']) {
    expect(textOnly.querySelector(selector)?.outerHTML).toBe(linksOnly.querySelector(selector)?.outerHTML);
  }
  expect(textOnly.dataset.pageIndex).toBe(linksOnly.dataset.pageIndex);
  expect(textOnly.querySelectorAll('[data-doc-start]')).toHaveLength(
    linksOnly.querySelectorAll('[data-doc-start]').length
  );
  expect(textOnly.textContent).toBe('Link descriptionLinkOther text');
});
