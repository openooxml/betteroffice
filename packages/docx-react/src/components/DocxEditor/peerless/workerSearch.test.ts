import { expect, mock, test } from 'bun:test';
import { createWorkerSearch } from './workerSearch';

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>(done => { resolve = done; });
  return { promise, resolve };
}

function fixture(starts = [7]) {
  const native = { story: 'body', start: { paraId: 'p', offset: 7 }, end: { paraId: 'p', offset: 10 } };
  const accepted = { story: 'body', view: 'accepted', start: { paraId: 'p', offset: 3 }, end: { paraId: 'p', offset: 6 } };
  let version = 0;
  const listeners = new Set<() => void>();
  const geometry = { frameEpoch: 1, layoutRevision: 1, pages: starts.map((_, pageIndex) => ({ page: { pageId: String(pageIndex + 1), pageIndex },
    bounds: { x: 10, y: 20 + pageIndex * 1000 }, displayPage: { unbuilt: true, positionSpan: [1, 20] } })) };
  const read = mock(async (request: { kind: string }) => request.kind === 'readParagraphs'
    ? { ok: true, version: 'native-token', paragraphs: [{ story: 'body', paraId: 'p', text: 'xx cat' }] }
    : request.kind === 'searchText' ? { matches: starts.map(start => ({ story: 'body', paraId: 'p', start, displayFrom: start + 1, displayTo: start + 4 })), carried: 0 }
    : request.kind === 'findMatches' ? [{ yrsRange: native, text: 'cat', startOffset: 7, endOffset: 10 }]
    : request.kind === 'findText' ? { ok: true, version: 'native-token', matches: [{ text: 'cat', range: accepted }], truncated: false }
    : request.kind === 'rangeHighlights' ? { version: 'native-token', frameEpoch: 1, layoutRevision: 1,
      rects: starts.map((_, pageIndex) => [{ pageId: String(pageIndex + 1), pageIndex, x: 4, y: 5, width: 30, height: 12 }]), unbuiltPages: [] } : []);
  const submit = mock(async () => ({ ok: true, status: 'executed' }));
  const controller = { read, readAt: read, submit, environment: () => ({ mode: 'editing', readOnly: false }),
    getSnapshot: () => ({ status: 'ready', view: { identity: { session: 's', epoch: 1 }, version, geometry } }),
    subscribe: (listener: () => void) => { listeners.add(listener); return () => listeners.delete(listener); },
    ref: { readParagraphs: (request: unknown) => read({ kind: 'readParagraphs', request } as never), getCurrentPage: () => 1, flushPendingInput: async () => {} },
  };
  const reveal = mock(() => {});
  const search = createWorkerSearch(controller as never, reveal);
  return { search, read, submit, reveal, native, accepted, change(notify = true) { version++; if (notify) listeners.forEach((listener) => listener()); } };
}

test('host search reads worker matches, highlights without selecting, navigates and clears', async () => {
  const { search, submit, reveal } = fixture();
  expect(await search.host.search('cat')).toMatchObject({ query: 'cat', total: 1, current: 0 });
  expect(search.getSnapshot().highlights[0]).toMatchObject({ x: 14, y: 25, width: 30 });
  expect(submit).not.toHaveBeenCalled();
  expect(search.host.searchNext()?.current).toBe(0);
  expect(submit).not.toHaveBeenCalled();
  expect(reveal).toHaveBeenCalled();
  search.host.clearSearch();
  expect(search.host.getSearchState()).toBeNull();
  expect(search.getSnapshot().highlights).toEqual([]);
  search.dispose();
});

test('closing find selects the current worker match before clearing search', async () => {
  const { search, submit } = fixture([0, 4]);
  search.open(false);
  await search.find('cat', { matchCase: false, matchWholeWord: false });
  search.host.searchNext();
  expect(submit).not.toHaveBeenCalled();
  search.close('cat', { matchCase: false, matchWholeWord: false });
  expect(submit).toHaveBeenCalledTimes(1);
  expect(submit).toHaveBeenCalledWith({ kind: 'select', payload: { selection: {
    anchor: { story: 'body', paraId: 'p', offset: 4 }, focus: { story: 'body', paraId: 'p', offset: 7 },
  } } });
  expect(search.getSnapshot()).toMatchObject({ isOpen: false, result: null, highlights: [] });
  expect(search.host.getSearchState()).toBeNull();
  search.dispose();
});

test.each([
  ['editing the query', 'dog', false, false],
  ['clearing the query', '', false, false],
  ['toggling Match case', 'cat', true, false],
  ['toggling Whole words', 'cat', false, true],
] as const)('closing find after %s before the next search does not select the old match', async (_, query, matchCase, matchWholeWord) => {
  const { search, submit } = fixture();
  search.open(false);
  await search.find('cat', { matchCase: false, matchWholeWord: false });
  search.close(query, { matchCase, matchWholeWord });
  expect(submit).not.toHaveBeenCalled();
  expect(search.getSnapshot()).toMatchObject({ isOpen: false, result: null, highlights: [] });
  expect(search.host.getSearchState()).toBeNull();
  search.dispose();
});

test('closing find after the document version changes does not select the old match', async () => {
  const { search, submit, change } = fixture();
  search.open(false);
  await search.find('cat', { matchCase: false, matchWholeWord: false });
  change(false);
  search.close('cat', { matchCase: false, matchWholeWord: false });
  expect(submit).not.toHaveBeenCalled();
  expect(search.getSnapshot()).toMatchObject({ isOpen: false, result: null, highlights: [] });
  search.dispose();
});

test('dialog find and navigation reveal matches on later pages without selecting', async () => {
  const { search, submit, reveal } = fixture([0, 4, 8]);
  search.open(false);
  await search.find('cat', { matchCase: false, matchWholeWord: false });
  const expectCurrent = async (pageIndex: number) => {
    await new Promise<void>(resolve => setImmediate(resolve));
    const target = { x: 14, y: 25 + pageIndex * 1000, width: 30, height: 12, current: true };
    expect(search.getSnapshot().highlights.filter(rect => rect.current)).toEqual([target]);
    expect(reveal).toHaveBeenLastCalledWith(target);
    expect(submit).not.toHaveBeenCalled();
  };
  await expectCurrent(0);
  search.host.searchNext();
  await expectCurrent(1);
  search.host.searchNext();
  await expectCurrent(2);
  search.host.searchNext();
  await expectCurrent(0);
  search.host.searchPrevious();
  await expectCurrent(2);
  search.host.searchPrevious();
  await expectCurrent(1);
  search.host.searchGoTo(2);
  await expectCurrent(2);
  search.host.searchGoTo(1);
  await expectCurrent(1);
  search.dispose();
});

test('closing find without matches leaves the worker selection unchanged', async () => {
  const { search, submit } = fixture([]);
  search.open(false);
  await search.find('cat', { matchCase: false, matchWholeWord: false });
  search.close('cat', { matchCase: false, matchWholeWord: false });
  expect(submit).not.toHaveBeenCalled();
  search.dispose();
});

test('replacement uses native document token and accepted targets rather than native highlight offsets', async () => {
  const { search, submit, accepted } = fixture();
  await search.host.search('cat', { caseSensitive: true });
  expect(await search.replace('dog', true)).toBe(1);
  expect(submit).toHaveBeenCalledWith({ kind: 'replaceMatches', payload: { targets: [accepted], text: 'dog', expectVersion: 'native-token' } });
  search.dispose();
});

test.each([
  ['closing', false],
  ['closing', true],
  ['clearing', false],
  ['clearing', true],
] as const)('%s find while replacement is pending does not restore search after acknowledgement (all: %s)', async (action, all) => {
  const { search, submit, read, reveal } = fixture();
  search.open(true);
  await search.find('cat', { matchCase: false, matchWholeWord: false });
  const started = deferred<void>();
  const acknowledgement = deferred<{ ok: boolean; status: string }>();
  submit.mockImplementation(() => { started.resolve(); return acknowledgement.promise; });
  const replacement = search.replace('dog', all);
  await started.promise;
  if (action === 'closing') search.close('cat', { matchCase: false, matchWholeWord: false });
  else search.host.clearSearch();
  expect(submit).toHaveBeenCalledTimes(1);
  expect(search.getSnapshot()).toMatchObject({ result: null, highlights: [] });
  const published = mock(() => {});
  const hostPublished = mock(() => {});
  search.subscribe(published);
  search.host.onSearchChange(hostPublished);
  const reads = read.mock.calls.length;
  const reveals = reveal.mock.calls.length;
  acknowledgement.resolve({ ok: true, status: 'executed' });
  expect(await replacement).toBe(1);
  expect(submit).toHaveBeenCalledTimes(1);
  expect(read).toHaveBeenCalledTimes(reads);
  expect(reveal).toHaveBeenCalledTimes(reveals);
  expect(published).not.toHaveBeenCalled();
  expect(hostPublished).not.toHaveBeenCalled();
  expect(search.getSnapshot()).toMatchObject({ result: null, highlights: [] });
  expect(search.host.getSearchState()).toBeNull();
  if (action === 'closing') expect(search.getSnapshot().isOpen).toBe(false);
  search.dispose();
});

test('dialog find after a pending replacement resolves does not restore a closed search', async () => {
  const { search, submit, read, reveal } = fixture();
  const options = { matchCase: false, matchWholeWord: false };
  search.open(true);
  await search.find('cat', options);
  const started = deferred<void>();
  const acknowledgement = deferred<{ ok: boolean; status: string }>();
  submit.mockImplementation(() => { started.resolve(); return acknowledgement.promise; });
  const replacement = search.replace('dog');
  await started.promise;
  search.close('cat', options);
  const closed = search.getSnapshot();
  const published = mock(() => {});
  const hostPublished = mock(() => {});
  search.subscribe(published);
  search.host.onSearchChange(hostPublished);
  const reads = read.mock.calls.length;
  const reveals = reveal.mock.calls.length;
  acknowledgement.resolve({ ok: true, status: 'executed' });
  expect(await replacement).toBe(1);
  expect(await search.find('cat', options)).toMatchObject({ total: 0, current: -1 });
  expect(search.getSnapshot()).toBe(closed);
  expect(search.getSnapshot()).toMatchObject({ isOpen: false, result: null, highlights: [] });
  expect(search.host.getSearchState()).toBeNull();
  expect(published).not.toHaveBeenCalled();
  expect(hostPublished).not.toHaveBeenCalled();
  expect(read).toHaveBeenCalledTimes(reads);
  expect(reveal).toHaveBeenCalledTimes(reveals);
  expect(submit).toHaveBeenCalledTimes(1);
  search.dispose();
});

test.each([false, true])('replacement refreshes an open search and preserves navigation and settled close (all: %s)', async all => {
  const { search, submit, read, reveal } = fixture();
  search.open(true);
  await search.find('cat', { matchCase: false, matchWholeWord: false });
  const started = deferred<void>();
  const acknowledgement = deferred<{ ok: boolean; status: string }>();
  submit.mockImplementationOnce(() => { started.resolve(); return acknowledgement.promise; });
  const searches = read.mock.calls.filter(([request]) => request.kind === 'searchText').length;
  const reveals = reveal.mock.calls.length;
  const replacement = search.replace('dog', all);
  await started.promise;
  acknowledgement.resolve({ ok: true, status: 'executed' });
  expect(await replacement).toBe(1);
  expect(read.mock.calls.filter(([request]) => request.kind === 'searchText')).toHaveLength(searches + 1);
  expect(reveal).toHaveBeenCalledTimes(reveals + 1);
  expect(search.getSnapshot()).toMatchObject({ isOpen: true, result: { totalCount: 1 } });
  expect(search.getSnapshot().highlights).toHaveLength(1);
  expect(search.host.searchNext()).toMatchObject({ total: 1, current: 0 });
  search.close('cat', { matchCase: false, matchWholeWord: false });
  expect(submit).toHaveBeenCalledTimes(2);
  expect(submit).toHaveBeenLastCalledWith({ kind: 'select', payload: { selection: {
    anchor: { story: 'body', paraId: 'p', offset: 7 }, focus: { story: 'body', paraId: 'p', offset: 10 },
  } } });
  expect(search.getSnapshot()).toMatchObject({ isOpen: false, result: null, highlights: [] });
  search.dispose();
});

test('stale replacements refresh search and do not retry the captured target', async () => {
  const { search, submit, read } = fixture();
  await search.host.search('cat', { caseSensitive: true });
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
  await search.host.search('cat', { caseSensitive: true });
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
