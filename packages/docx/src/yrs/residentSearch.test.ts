import { expect, mock, test } from 'bun:test';
import type { YrsLoc, YrsStorySegment, YrsTextMatch } from './index';
import { readResidentSearch } from './residentSearch';

function reader() {
  const segments: YrsStorySegment[] = [
    { kind: 'text', text: 'cat cat', attributes: {} },
    { kind: 'pilcrow', paraId: 'p1', properties: {}, attributes: {} },
  ];
  const hits: YrsTextMatch[] = [
    { story: 'body', paraId: 'p1', start: 4, end: 7 },
    { story: 'hf:header', paraId: 'p1', start: 0, end: 3 },
    { story: 'fn:1', paraId: 'p1', start: 0, end: 3 },
    { story: 'body:unmapped', paraId: 'p1', start: 0, end: 3 },
    { story: 'body', paraId: 'missing', start: 0, end: 3 },
    { story: 'body', paraId: 'p1', start: 3, end: 3 },
    { story: 'body', paraId: 'p1', start: 3, end: 2 },
    { story: 'body', paraId: 'p1', start: 0, end: 3 },
  ].map((hit) => ({ ...hit, text: 'cat' }));
  return {
    hasStory: () => true,
    storySegments: () => segments,
    paragraphSpans: () => [{ paraId: 'p1', length: 7 }],
    searchText: mock(() => hits),
    encodeStickyPosition: (loc: YrsLoc) => ({ story: loc.story, encoded: Uint8Array.of(loc.offset) }),
    resolveStickyPosition: () => null as YrsLoc | null,
  };
}

test('resident search filters unsupported stories and ranges, sorts and caches anchors', () => {
  const source = reader();
  expect(readResidentSearch(source, 'cat', true)).toEqual({
    carried: 0,
    matches: [
      { story: 'body', paraId: 'p1', start: 0, displayFrom: 1, displayTo: 4,
        anchor: { story: 'body', encoded: Uint8Array.of(0) } },
      { story: 'body', paraId: 'p1', start: 4, displayFrom: 5, displayTo: 8,
        anchor: { story: 'body', encoded: Uint8Array.of(4) } },
    ],
  });
  expect(source.searchText).toHaveBeenCalledWith('cat', { caseSensitive: true });
});

test('resident search carries exact locations, the first at or after, and the last', () => {
  const source = reader();
  const carry = { story: 'body', encoded: Uint8Array.of(1) };
  for (const [offset, current] of [[0, 0], [4, 1], [2, 1], [7, 1]] as const) {
    source.resolveStickyPosition = () => ({ story: 'body', paraId: 'p1', offset });
    expect(readResidentSearch(source, 'cat', false, carry).carried).toBe(current);
  }
  source.resolveStickyPosition = () => null;
  expect(readResidentSearch(source, 'cat', false, carry).carried).toBe(0);
  source.resolveStickyPosition = () => { throw new Error('invalid sticky position'); };
  expect(readResidentSearch(source, 'cat', false, carry).carried).toBe(0);
});

test('resident search retains matches when anchors cannot be encoded and handles empty results', () => {
  const source = reader();
  source.encodeStickyPosition = () => { throw new Error('unavailable anchor'); };
  expect(readResidentSearch(source, 'cat', false).matches.map(({ anchor }) => anchor)).toEqual([null, null]);
  expect(readResidentSearch(source, '', false)).toEqual({ matches: [], carried: -1 });
  source.searchText.mockReturnValue([]);
  expect(readResidentSearch(source, 'missing', false)).toEqual({ matches: [], carried: -1 });
});
