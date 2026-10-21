import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, test } from 'bun:test';
import type { DisplayPage, DisplayPrimitive } from './displayList';
import { buildMirrorPage, reduceMirrorToText, updateMirrorPage } from './mirrorDom';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const options = { labels: { page: 'Page 1', header: 'Header', footer: 'Footer', noteBacklink: 'Back' } };

function run(blockKey: string, start: number, text: string, y: number, extra: Partial<DisplayPrimitive> = {}): DisplayPrimitive {
  return {
    kind: 'text', text, x: 10, baselineY: y, width: 6 * text.length, font: '12px Arial', color: '#000',
    blockKey, docStart: start, docEnd: start + text.length,
    fragmentDocStart: start - 1, fragmentDocEnd: start + text.length + 1, ...extra,
  } as DisplayPrimitive;
}

function cell(tableId: string, parentTableId?: string): Partial<DisplayPrimitive> {
  return {
    table: { tableId, rowStart: 0, ...(parentTableId ? { parentTableId } : {}) },
    cell: { row: 0, col: 0, rowSpan: 1, colSpan: 1, cellId: `${tableId}-0-0` },
  };
}

/** `count` paragraphs of five runs each; `edit` inserts "x" at the end of the first run of paragraph `edit`. */
function prose(count: number, edit = -1, shift = 0): DisplayPage {
  const primitives: DisplayPrimitive[] = [];
  let pos = 100 + shift;
  for (let block = 0; block < count; block++) {
    for (let r = 0; r < 5; r++) {
      const text = block === edit && r === 0 ? 'wordsx' : 'words';
      primitives.push(run(`b${block}`, pos, text, 20 + block * 14));
      pos += text.length;
    }
    pos += 2;
  }
  return { pageIndex: 0, width: 600, height: 800, primitives };
}

function canonical(node: Node): string {
  if (node.nodeType !== 1) return JSON.stringify(node.textContent);
  const el = node as Element;
  const attributes = Array.from(el.attributes, (a) => `${a.name}=${JSON.stringify(a.value)}`).sort();
  return `<${el.tagName} ${attributes.join(' ')}>${Array.from(el.childNodes, canonical).join('')}</>`;
}

/** Updates `mirror` to `page`, checks it matches a fresh build, and returns the elements it created. */
function update(mirror: HTMLElement, page: DisplayPage): Element[] {
  const before = new Set(mirror.querySelectorAll('*'));
  expect(updateMirrorPage(mirror, page, options)).toBe(true);
  expect(canonical(mirror)).toBe(canonical(buildMirrorPage(page, options)));
  return Array.from(mirror.querySelectorAll('*')).filter((el) => !before.has(el));
}

test('a one-character edit rebuilds only the edited paragraph and shifts the rest', () => {
  const mirror = buildMirrorPage(prose(40), options);
  const created = update(mirror, prose(40, 20));
  expect(created.length).toBeLessThanOrEqual(7);
  expect(created[0]?.getAttribute('data-block-id')).toBe('b20');
  const after = mirror.querySelector('[data-block-id="b21"] > span:not(.layout-mirror-structural-anchor)')!;
  expect(after.getAttribute('data-doc-start')).toBe(String(100 + 21 * 27 + 1));
});

test('a page whose positions only shifted keeps every element', () => {
  const mirror = buildMirrorPage(prose(40), options);
  expect(update(mirror, prose(40, -1, 7))).toHaveLength(0);
  expect(update(mirror, prose(40, -1, 3))).toHaveLength(0);
});

test('blocks that come, go or reorder leave the others in place', () => {
  const page = prose(10);
  const mirror = buildMirrorPage(page, options);
  const kept = mirror.querySelector('[data-block-id="b5"]');
  const without = { ...page, primitives: page.primitives.filter((p) => p.blockKey !== 'b3') };
  expect(update(mirror, without)).toHaveLength(0);
  expect(mirror.querySelector('[data-block-id="b5"]')).toBe(kept);
  const reordered = { ...page, primitives: [...page.primitives.slice(25), ...page.primitives.slice(0, 25)] };
  update(mirror, reordered);
  expect(update(mirror, page).length).toBeLessThanOrEqual(30);
});

test('an uneven shift inside a block rebuilds that block', () => {
  const page = prose(5);
  const mirror = buildMirrorPage(page, options);
  const moved = (p: DisplayPrimitive): DisplayPrimitive => ({
    ...p, docStart: p.docStart! + 1, docEnd: p.docEnd! + 1,
    fragmentDocStart: p.fragmentDocStart! + 1, fragmentDocEnd: p.fragmentDocEnd! + 1,
  });
  const inner = page.primitives.map((p, i) =>
    i === 7 ? { ...p, docEnd: p.docEnd! + 1 } : i > 7 ? moved(p) : p
  ) as DisplayPrimitive[];
  expect(update(mirror, { ...page, primitives: inner }).length).toBe(7);
});

test('a note reference id stays on the first mark of the page', () => {
  const noteRef = { id: 1, kind: 'footnote' } as DisplayPrimitive['noteRef'];
  const page = (first: number): DisplayPage => {
    const base = prose(4);
    return {
      ...base,
      primitives: base.primitives.map((p, i) => (i === first || i === 17 ? { ...p, noteRef } : p)),
      noteAreas: [{
        kind: 'footnote', y: 700, height: 40, noteIds: [1],
        notes: [{ id: 1, label: '1', anchorDocStart: 120, anchorDocEnd: 121 }],
        primitives: [run('n1', 400, 'note', 710, { groupId: 'footnote-1' })],
      }],
    } as DisplayPage;
  };
  const mirror = buildMirrorPage(page(17), options);
  update(mirror, page(2));
  expect(mirror.querySelectorAll('#oox-noteref-footnote-1')).toHaveLength(1);
  update(mirror, page(17));
  update(mirror, { ...page(17), noteAreas: [] });
});

test('nested tables, bands and page borders update like a fresh build', () => {
  const page = (text: string, shift = 0): DisplayPage => ({
    pageIndex: 0, width: 600, height: 800,
    primitives: [
      { kind: 'rect', x: 0, y: 0, w: 300, h: 300, fill: '#fff', blockKey: 't-p0', ...cell('t') } as DisplayPrimitive,
      run('t-p1', 100 + shift, 'outer', 30, cell('t')),
      run('n-p1', 120 + shift, text, 60, { x: 20, ...cell('n', 't') }),
      run('b9', 300 + shift, 'after', 400),
      { kind: 'line', x1: 0, y1: 500, x2: 600, y2: 500, strokeWidth: 1, color: '#000' },
    ],
    header: { kind: 'header', rId: 'rId1', y: 0, height: 20, primitives: [run('h0', 1, 'head', 10)] },
    footer: { kind: 'footer', rId: 'rId2', y: 780, height: 20, primitives: [run('f0', 1, 'foot', 790)] },
    pageBorders: [{ kind: 'pageBorder', x: 5, y: 5, w: 590, h: 790, zOrder: 'back', top: { width: 1, color: '#000', style: 'solid' } }],
  });
  const mirror = buildMirrorPage(page('inner'), options);
  expect(mirror.querySelector('[role="cell"] [role="table"]')).not.toBeNull();
  expect(update(mirror, page('inner', 4))).toHaveLength(0);
  expect(update(mirror, page('inner!', 4)).length).toBeGreaterThan(0);
  expect(update(mirror, { ...page('inner!', 4), header: undefined, pageIndex: 3 })).toHaveLength(0);
});

test('nested tables that go or reorder inside a kept table update like a fresh build', () => {
  const at = { n: 0, m: 1, k: 2 } as const;
  const page = (...nested: (keyof typeof at)[]): DisplayPage => ({
    pageIndex: 0, width: 600, height: 800,
    primitives: [
      { kind: 'rect', x: 0, y: 0, w: 300, h: 300, fill: '#fff', blockKey: 't-p0', ...cell('t') } as DisplayPrimitive,
      run('t-p1', 100, 'outer', 30, cell('t')),
      ...nested.map((id) => run(`${id}-p1`, 120 + at[id] * 20, id, 60 + at[id] * 30, { x: 20, ...cell(id, 't') })),
    ],
  });
  const mirror = buildMirrorPage(page('n', 'm', 'k'), options);
  expect(mirror.querySelectorAll('[role="cell"] [role="table"]')).toHaveLength(3);
  update(mirror, page('m', 'n', 'k'));
  update(mirror, page('m', 'k'));
  update(mirror, page('k'));
  update(mirror, page());
  update(mirror, page('n', 'm'));
});

test('a primitive that gains a direction key set to undefined updates like a fresh build', () => {
  const page = prose(3);
  const keyed = { ...page, primitives: page.primitives.map((p, i) => (i === 5 ? ({ ...p, rtl: undefined } as DisplayPrimitive) : p)) };
  const mirror = buildMirrorPage(page, options);
  update(mirror, keyed);
  expect(mirror.querySelector('[data-block-id="b1"]')!.getAttribute('dir')).toBe('ltr');
  update(mirror, page);
  expect(mirror.querySelector('[data-block-id="b1"]')!.hasAttribute('dir')).toBe(false);
});

test('a drawing payload the mirror does not show keeps the shape, and its accessibility does not', () => {
  const page = (src: string, decorative: boolean): DisplayPage => ({
    ...prose(2),
    primitives: [
      ...prose(2).primitives,
      {
        kind: 'shape', x: 50, y: 300, w: 100, h: 80, blockKey: 's0', docStart: 200, docEnd: 201, decorative,
        geometryPath: [{ type: 'move', x: 50, y: 300 }, { type: 'line', x: src.length, y: 380 }, { type: 'close' }],
        fillPaint: { kind: 'picture', pictureSrc: src },
        drawingScene: { version: 1, nodes: [{ src }] },
      } as unknown as DisplayPrimitive,
    ],
  });
  const mirror = buildMirrorPage(page('data:image/png;base64,AAAA', false), options);
  expect(update(mirror, page(`data:image/png;base64,${'B'.repeat(4096)}`, false))).toHaveLength(0);
  expect(update(mirror, page('data:image/png;base64,AAAA', true)).length).toBeGreaterThan(0);
});

test('a mirror reduced to text is not updated in place', () => {
  const mirror = reduceMirrorToText(buildMirrorPage(prose(3), options));
  expect(updateMirrorPage(mirror, prose(3, 1), options)).toBe(false);
});
