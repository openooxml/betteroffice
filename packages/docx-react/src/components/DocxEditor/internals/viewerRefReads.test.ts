import { expect, mock, spyOn, test } from 'bun:test';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import type { DocxFindDisplayMatch, ResidentDocumentRead, ResidentEngineWorkerClient } from '@betteroffice/docx/yrs';
import { markPresented, stampWorkerFrameVersion } from './layoutProvenance';
import { readViewerFindMatches, type ViewerRefReadAccess } from './viewerRefReads';

const options = { matchCase: false, matchWholeWord: true };
const matches: DocxFindDisplayMatch[] = [{
  paragraphIndex: 0, contentIndex: 0, startOffset: 0, endOffset: 4, text: 'word',
  displayFrom: 1, displayTo: 5,
  yrsRange: { story: 'body', start: { paraId: 'p', offset: 0 }, end: { paraId: 'p', offset: 4 } },
}];

function accessFor(read: unknown) {
  const host = {} as HTMLElement;
  let queries: DisplayListQueries;
  const present = (version: string) => {
    queries = { displayList: {} } as DisplayListQueries;
    stampWorkerFrameVersion(queries, version);
    markPresented(host, queries.displayList);
    return queries;
  };
  present('A');
  const access: ViewerRefReadAccess = {
    read: read as ResidentEngineWorkerClient['documentRead'],
    story: 'body', host: () => host, queries: () => queries,
    awaitFrame: mock(async () => present('B')),
    current: () => true, selection: () => null,
  };
  return { access, present };
}

test('find reads matches at the presented worker version without a selection', async () => {
  const read = mock(async () => ({ version: 'A', value: matches }));
  const { access } = accessFor(read);
  expect(await readViewerFindMatches(access, 'word', options)).toEqual({ version: 'A', matches });
  expect(read).toHaveBeenCalledWith({ kind: 'findMatches', searchText: 'word', options, expectVersion: 'A' });
  expect(access.awaitFrame).not.toHaveBeenCalled();
});

test('a superseded find retries against the next presented frame', async () => {
  const read = mock(async (_request: ResidentDocumentRead) => ({ version: 'B', value: matches }));
  const { access } = accessFor(read);
  expect(await readViewerFindMatches(access, 'word', options)).toEqual({ version: 'B', matches });
  expect(read.mock.calls).toEqual([
    [{ kind: 'findMatches', searchText: 'word', options, expectVersion: 'A' }],
    [{ kind: 'findMatches', searchText: 'word', options, expectVersion: 'B' }],
  ]);
  expect(access.awaitFrame).toHaveBeenCalledTimes(1);
});

test('find retries when the frame changes before an otherwise current reply', async () => {
  const { access, present } = accessFor(null);
  let reads = 0;
  access.read = (async () => {
    if (reads++ === 0) {
      present('B');
      return { version: 'A', value: matches };
    }
    return { version: 'B', value: matches };
  }) as ResidentEngineWorkerClient['documentRead'];
  expect(await readViewerFindMatches(access, 'word', options)).toEqual({ version: 'B', matches });
  expect(reads).toBe(2);
});

test('find returns null at the deadline for an unanswered read', async () => {
  const { access } = accessFor(() => new Promise(() => {}));
  const now = spyOn(Date, 'now').mockReturnValue(10_000).mockReturnValueOnce(0);
  try {
    expect(await readViewerFindMatches(access, 'word', options)).toBeNull();
  } finally {
    now.mockRestore();
  }
});

test('find stops after five superseded attempts', async () => {
  const read = mock(async () => ({ version: 'C', value: matches }));
  const { access } = accessFor(read);
  expect(await readViewerFindMatches(access, 'word', options)).toBeNull();
  expect(read).toHaveBeenCalledTimes(5);
  expect(access.awaitFrame).toHaveBeenCalledTimes(4);
});

test('find returns null when the viewer stops being current', async () => {
  const { access } = accessFor(async () => {
    access.current = () => false;
    return { version: 'A', value: matches };
  });
  expect(await readViewerFindMatches(access, 'word', options)).toBeNull();
});

test('find returns null without a presented frame and accepts empty matches', async () => {
  const read = mock(async () => ({ version: 'A', value: [] }));
  const { access } = accessFor(read);
  expect(await readViewerFindMatches(access, 'word', options)).toEqual({ version: 'A', matches: [] });
  access.host = () => null;
  expect(await readViewerFindMatches(access, 'word', options)).toBeNull();
  expect(read).toHaveBeenCalledTimes(1);
});
