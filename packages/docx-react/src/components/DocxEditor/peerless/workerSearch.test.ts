import { expect, mock, test } from 'bun:test';
import { createWorkerSearch } from './workerSearch';

function fixture() {
  const native = { story: 'body', start: { paraId: 'p', offset: 7 }, end: { paraId: 'p', offset: 10 } };
  const accepted = { story: 'body', view: 'accepted', start: { paraId: 'p', offset: 3 }, end: { paraId: 'p', offset: 6 } };
  let version = 0;
  const listeners = new Set<() => void>();
  const geometry = { frameEpoch: 1, layoutRevision: 1, pages: [{ page: { pageId: '1', pageIndex: 0 }, bounds: { x: 10, y: 20 }, displayPage: { unbuilt: true, positionSpan: [1, 20] } }] };
  const read = mock(async (request: { kind: string }) => request.kind === 'readParagraphs'
    ? { ok: true, version: 'native-token', paragraphs: [{ story: 'body', paraId: 'p', text: 'xx cat' }] }
    : request.kind === 'searchText' ? { matches: [{ story: 'body', paraId: 'p', start: 7, displayFrom: 8, displayTo: 11 }], carried: 0 }
    : request.kind === 'findMatches' ? [{ yrsRange: native, text: 'cat', startOffset: 7, endOffset: 10 }]
    : request.kind === 'findText' ? { ok: true, version: 'native-token', matches: [{ text: 'cat', range: accepted }], truncated: false }
    : request.kind === 'rangeHighlights' ? { version: 'native-token', frameEpoch: 1, layoutRevision: 1,
      rects: [[{ pageId: '1', pageIndex: 0, x: 4, y: 5, width: 30, height: 12 }]], unbuiltPages: [] } : []);
  const submit = mock(async () => ({ ok: true, status: 'executed' }));
  const controller = { read, readAt: read, submit, environment: () => ({ mode: 'editing', readOnly: false }),
    getSnapshot: () => ({ status: 'ready', view: { identity: { session: 's', epoch: 1 }, version, geometry } }),
    subscribe: (listener: () => void) => { listeners.add(listener); return () => listeners.delete(listener); },
    ref: { readParagraphs: (request: unknown) => read({ kind: 'readParagraphs', request } as never), getCurrentPage: () => 1, flushPendingInput: async () => {} },
  };
  const reveal = mock(() => {});
  const search = createWorkerSearch(controller as never, reveal);
  return { search, read, submit, reveal, native, accepted, change() { version++; listeners.forEach((listener) => listener()); } };
}

test('host search reads worker matches, highlights without selecting, navigates and clears', async () => {
  const { search, submit, reveal } = fixture();
  expect(await search.host.search('cat')).toMatchObject({ query: 'cat', total: 1, current: 0 });
  expect(search.getSnapshot().highlights[0]).toMatchObject({ x: 14, y: 25, width: 30 });
  expect(submit).not.toHaveBeenCalled();
  expect(search.host.searchNext()?.current).toBe(0);
  expect(reveal).toHaveBeenCalled();
  search.host.clearSearch();
  expect(search.host.getSearchState()).toBeNull();
  expect(search.getSnapshot().highlights).toEqual([]);
  search.dispose();
});

test('replacement uses native document token and accepted targets rather than native highlight offsets', async () => {
  const { search, submit, accepted } = fixture();
  await search.find('cat', { matchCase: true, matchWholeWord: false });
  expect(await search.replace('dog', true)).toBe(1);
  expect(submit).toHaveBeenCalledWith({ kind: 'replaceMatches', payload: { targets: [accepted], text: 'dog', expectVersion: 'native-token' } });
  search.dispose();
});

test('stale replacements refresh search and do not retry the captured target', async () => {
  const { search, submit, read } = fixture();
  await search.find('cat', { matchCase: true, matchWholeWord: false });
  submit.mockImplementation(async () => ({ ok: false, failure: { code: 'command-failed', message: 'stale-version' } } as never));
  const before = read.mock.calls.length;
  expect(await search.replace('dog', true)).toBe(0);
  expect(submit).toHaveBeenCalledTimes(1);
  expect(read.mock.calls.length).toBeGreaterThan(before);
  search.dispose();
});

test('navigation reveals a never-built match using retained page position spans', async () => {
  const { search, read, reveal } = fixture();
  const original = read.getMockImplementation()!;
  read.mockImplementation(async (request) => request.kind === 'rangeHighlights'
    ? { version: 'native-token', frameEpoch: 1, layoutRevision: 1, rects: [[]], unbuiltPages: [0] } as never
    : original(request));
  expect(await search.host.search('cat')).toMatchObject({ total: 1 });
  expect(reveal).toHaveBeenCalled();
  search.dispose();
});

test('changed documents immediately discard old highlight rectangles', async () => {
  const { search, change } = fixture();
  await search.host.search('cat');
  expect(search.getSnapshot().highlights).toHaveLength(1);
  change();
  expect(search.getSnapshot().highlights).toEqual([]);
  await new Promise<void>(resolve => setImmediate(resolve));
  search.dispose();
});

test('a mismatch between native and accepted matches refuses replacement', async () => {
  const { search, submit, read } = fixture();
  await search.find('cat', { matchCase: true, matchWholeWord: false });
  const original = read.getMockImplementation()!;
  read.mockImplementation(async request => request.kind === 'readParagraphs'
    ? { ok: true, version: 'native-token', paragraphs: [{ story: 'body', paraId: 'p', text: 'cat cat' }] } as never
    : original(request));
  expect(await search.replace('dog', true)).toBe(0);
  expect(submit).not.toHaveBeenCalled();
  search.dispose();
});

test('highlight frame mismatch leaves the overlay empty', async () => {
  const { search, read } = fixture();
  const original = read.getMockImplementation()!;
  read.mockImplementation(async request => request.kind === 'rangeHighlights'
    ? { ...await original(request), frameEpoch: 2 } as never : original(request));
  await search.host.search('cat');
  expect(search.getSnapshot().highlights).toEqual([]);
  search.dispose();
});
