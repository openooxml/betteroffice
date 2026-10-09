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
    encodeStickyPosition: mock((_loc: YrsLoc) => { throw new Error('unexpected anchor encoding'); }),
    resolveStickyPosition: () => null as YrsLoc | null,
  };
}

test('resident search filters unsupported stories and ranges and sorts without encoding anchors', () => {
  const source = reader();
  expect(readResidentSearch(source, 'cat', true)).toEqual({
    carried: 0,
    matches: [
      { story: 'body', paraId: 'p1', start: 0, displayFrom: 1, displayTo: 4 },
      { story: 'body', paraId: 'p1', start: 4, displayFrom: 5, displayTo: 8 },
    ],
  });
  expect(source.searchText).toHaveBeenCalledWith('cat', { caseSensitive: true });
  expect(source.encodeStickyPosition).not.toHaveBeenCalled();
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

test('resident search carries an old match offset encoded in the current paragraph to the following match', () => {
  const source = reader();
  const carry = { story: 'body', encoded: Uint8Array.of(1) };
  source.resolveStickyPosition = () => ({ story: 'body', paraId: 'p1', offset: 4 });
  expect(readResidentSearch(source, 'cat', false, carry).carried).toBe(1);
  source.storySegments()[0] = { kind: 'text', text: 'x cat cat', attributes: {} };
  source.paragraphSpans = () => [{ paraId: 'p1', length: 9 }];
  source.searchText.mockReturnValue([
    { story: 'body', paraId: 'p1', start: 2, end: 5, text: 'cat' },
    { story: 'body', paraId: 'p1', start: 6, end: 9, text: 'cat' },
  ]);
  const result = readResidentSearch(source, 'cat', false, carry);
  expect(result.carried).toBe(1);
  expect(result.matches[result.carried]).toEqual({
    story: 'body', paraId: 'p1', start: 6, displayFrom: 7, displayTo: 10,
  });
  expect(source.encodeStickyPosition).not.toHaveBeenCalled();
});

test('resident search handles empty results without encoding anchors', () => {
  const source = reader();
  expect(readResidentSearch(source, '', false)).toEqual({ matches: [], carried: -1 });
  source.searchText.mockReturnValue([]);
  expect(readResidentSearch(source, 'missing', false)).toEqual({ matches: [], carried: -1 });
  expect(source.encodeStickyPosition).not.toHaveBeenCalled();
});

test('resident search uses the position outline without reading segments and preserves legacy results', () => {
  const source = reader();
  const expected = readResidentSearch(source, 'cat', false);
  const segments = mock(() => { throw new Error('unexpected segment export'); });
  const fast = {
    ...source,
    storySegments: segments,
    positionOutline: () => ({ body: { contentStart: 0, size: 9, paragraphs: [
      { paraId: 'p1', displayStart: 0, length: 7, leading: 0 },
    ] } }),
  };
  const actual = readResidentSearch(fast, 'cat', false);
  expect(actual).toEqual(expected);
  expect(JSON.stringify(actual)).toBe(JSON.stringify(expected));
  expect(segments).not.toHaveBeenCalled();
  expect(readResidentSearch({ ...source, positionOutline: () => null }, 'cat', false)).toEqual(expected);
  expect(readResidentSearch({ ...fast, hasStory: (story) => story !== 'body' }, 'cat', false)).toEqual({ matches: [], carried: -1 });
});
