import { expect, test } from 'bun:test';
import { dumpDisplayLists, firstDifferences } from './displayListParity';

test('with the same fonts, PNG export draws the same display list on the native and wasm engines', () => {
  const wasm = dumpDisplayLists('wasm', 'pinned');
  const native = dumpDisplayLists('native', 'pinned');
  expect(Object.keys(wasm.pages).sort()).toEqual(['demo', 'picture-corrupt', 'picture-drawn', 'picture-missing', 'picture-none', 'typography']);
  expect(Object.entries(wasm.pages).filter(([, pages]) => pages.length === 0).map(([name]) => name)).toEqual([]);
  expect(firstDifferences(wasm, native)).toEqual([]);
}, 240_000);
