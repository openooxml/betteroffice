/**
 * PNG export on the native and wasm engines hands each page the same display
 * list. Face ids are canonicalised before comparing because they follow font
 * load completion order.
 */
import { describe, expect, test } from 'bun:test';
import { canonicalFaceIds, dumpDisplayLists, firstDifferences, type DisplayListDump } from './displayListParity';

const dumps = new Map<string, DisplayListDump>();

function dump(engine: 'wasm' | 'native', run: 1 | 2): DisplayListDump {
  const key = `${engine} ${run}`;
  let found = dumps.get(key);
  if (!found) {
    found = dumpDisplayLists(engine, 'pinned');
    dumps.set(key, found);
  }
  return found;
}

test('with the same fonts, PNG export draws the same display list on the native and wasm engines', () => {
  const wasm = dump('wasm', 1);
  const native = dump('native', 1);
  expect(Object.keys(wasm.pages).sort()).toEqual(['demo', 'leaders', 'picture-corrupt', 'picture-drawn', 'picture-missing', 'picture-none', 'typography']);
  expect(Object.entries(wasm.pages).filter(([, pages]) => pages.length === 0).map(([name]) => name)).toEqual([]);
  expect(firstDifferences(wasm, native)).toEqual([]);
}, 240_000);

for (const engine of ['wasm', 'native'] as const) {
  test(`the ${engine} engine draws the same display list, font chains included, on every run`, () => {
    expect(firstDifferences(dump(engine, 1), dump(engine, 2))).toEqual([]);
  }, 240_000);
}

describe('face id canonicalisation', () => {
  const recorded = (fontChains: Record<string, number[]>, ids: [number, number, number, number]) => JSON.stringify({
    fontChains,
    page: {
      pageIndex: 0,
      primitives: [
        { kind: 'glyphRun', fontId: ids[0], text: 'a', glyphs: [{ id: 4, x: 0, y: 0, cluster: 0 }] },
        { kind: 'leaderGlyphs', fontId: ids[1], glyph: '.' },
        { kind: 'clipGroup', primitives: [{ kind: 'glyphRun', fontId: ids[2], text: 'b', glyphs: [], leaderGlyphs: { fontId: ids[3], glyph: '-' } }] },
      ],
    },
  });
  const pair = (left: string, right: string): [DisplayListDump, DisplayListDump] => [
    { engine: 'wasm', fonts: 'pinned', pages: { sample: [left] } },
    { engine: 'native', fonts: 'pinned', pages: { sample: [right] } },
  ];

  test('lists that differ only by a permutation of face ids compare equal', () => {
    const left = recorded({ body: [3, 8], heading: [8, 11] }, [8, 3, 11, 3]);
    const right = recorded({ heading: [2, 7], body: [5, 2] }, [2, 5, 7, 5]);
    expect(firstDifferences(...pair(left, right))).toEqual([]);
  });

  test('a primitive pointing at a different face of the same chains compares unequal', () => {
    const left = recorded({ body: [3, 8] }, [8, 3, 8, 3]);
    const right = recorded({ body: [3, 8] }, [8, 3, 3, 3]);
    const found = firstDifferences(...pair(left, right));
    expect(found).toHaveLength(1);
    expect(found[0]).toStartWith('sample page 0 primitive 2 of 3/3:');
  });

  test('a face id outside every chain keeps a number of its own', () => {
    const left = recorded({ body: [3, 8] }, [8, 3, 8, 21]);
    const right = recorded({ body: [3, 8] }, [8, 3, 8, 3]);
    const canonical = canonicalFaceIds(left);
    expect(canonical.fontChains).toEqual({ body: [0, 1] });
    expect(JSON.stringify(canonical.page)).toContain('"leaderGlyphs":{"fontId":2,');
    expect(firstDifferences(...pair(left, right))).toHaveLength(1);
  });
});
