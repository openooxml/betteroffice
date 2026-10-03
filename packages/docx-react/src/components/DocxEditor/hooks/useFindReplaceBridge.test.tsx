import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
import type { DocxFindDisplayMatch, YrsSession } from '@betteroffice/docx/yrs';
import type { useFindReplace } from '../../../hooks/useFindReplace';
import type { PagedEditorRef } from '../PagedEditor';
import { useFindReplaceBridge } from './useFindReplaceBridge';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const options = { matchCase: false, matchWholeWord: true };
function match(startOffset: number, text = 'word'): DocxFindDisplayMatch {
  const endOffset = startOffset + text.length;
  return {
    paragraphIndex: 0, contentIndex: 0, startOffset, endOffset, text,
    displayFrom: startOffset + 1, displayTo: endOffset + 1,
    yrsRange: { story: 'body', start: { paraId: 'p', offset: startOffset }, end: { paraId: 'p', offset: endOffset } },
  };
}
const matches = [match(0), match(10)];
type Read = { version: string; matches: DocxFindDisplayMatch[] } | null;

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((yes) => { resolve = yes; });
  return { promise, resolve };
}

function mount(viewer = true) {
  const paragraphs = mock(() => {
    if (viewer) throw new Error('unexpected replica read');
    return [{ paraId: 'p', text: 'word Word wording' }];
  });
  const version = mock(() => 'v1');
  const session = { paragraphs, setSelection: mock(() => {}), version } as unknown as YrsSession;
  const editor = {
    isWorkerViewer: mock(() => viewer),
    getYrsSession: mock(() => session),
    readViewerFindMatches: mock(async (): Promise<Read> => ({ version: 'v1', matches })),
    setSelection: mock(() => {}),
    scrollToPosition: mock(() => {}),
    syncYrsInputState: mock(() => true),
    yrsLocToDisplayPosition: mock((loc: { offset: number }) => loc.offset + 1),
  };
  const pagedEditorRef = { current: editor as unknown as PagedEditorRef | null };
  const findReplace = {
    setMatches: mock(() => {}), goToMatch: mock((_index: number) => {}), state: { isOpen: true },
  };
  const complete = mock(async () => ({ ok: true, status: 'executed' } as const));
  const hook = renderHook(() => useFindReplaceBridge({
    pagedEditorRef, findReplace: findReplace as unknown as ReturnType<typeof useFindReplace>, complete,
  }));
  return { hook, editor, session, version, paragraphs, pagedEditorRef, findReplace, complete };
}

test('viewer find publishes worker matches and selects and scrolls without reading the replica', async () => {
  const { hook, editor, session, paragraphs, findReplace } = mount();
  await act(async () => {
    expect(hook.result.current.handleFind('word', options)).toBeNull();
  });
  expect(editor.readViewerFindMatches).toHaveBeenCalledWith('word', options);
  expect(hook.result.current.findResultRef.current).toEqual({ matches, totalCount: 2, currentIndex: 0 });
  expect(findReplace.setMatches).toHaveBeenLastCalledWith(matches, 0);
  expect(editor.setSelection).toHaveBeenCalledWith(1, 5);
  expect(editor.scrollToPosition).toHaveBeenCalledWith(1);
  expect(paragraphs).not.toHaveBeenCalled();
  expect(session.setSelection).not.toHaveBeenCalled();
  expect(editor.syncYrsInputState).not.toHaveBeenCalled();
});

test('viewer find ignores a result superseded by a second search', async () => {
  const { hook, editor, findReplace } = mount();
  const first = deferred<Read>();
  const second = deferred<Read>();
  editor.readViewerFindMatches.mockReturnValueOnce(first.promise).mockReturnValueOnce(second.promise);
  act(() => {
    hook.result.current.handleFind('word', options);
    hook.result.current.handleFind('next', options);
  });
  const next = [match(20, 'next')];
  await act(async () => { second.resolve({ version: 'v1', matches: next }); });
  await act(async () => { first.resolve({ version: 'v1', matches }); });
  expect(findReplace.setMatches).toHaveBeenLastCalledWith(next, 0);
  expect(hook.result.current.findResultRef.current?.matches).toEqual(next);
  expect(editor.setSelection).toHaveBeenCalledTimes(1);
  expect(editor.setSelection).toHaveBeenCalledWith(21, 25);
});

test('viewer next and previous wrap through the worker matches', async () => {
  const { hook, editor, findReplace } = mount();
  await act(async () => { hook.result.current.handleFind('word', options); });
  act(() => {
    expect(hook.result.current.handleFindPrevious()).toEqual(matches[1]);
    expect(hook.result.current.handleFindNext()).toEqual(matches[0]);
    expect(hook.result.current.handleFindNext()).toEqual(matches[1]);
    expect(hook.result.current.handleFindNext()).toEqual(matches[0]);
  });
  expect(findReplace.goToMatch.mock.calls).toEqual([[0], [1], [0], [1], [0]]);
  expect(editor.setSelection).toHaveBeenLastCalledWith(1, 5);
});

test('viewer replacements refuse before completing or reading the replica', async () => {
  const { hook, complete, paragraphs } = mount();
  expect(await hook.result.current.handleReplace('new')).toBe(false);
  expect(await hook.result.current.handleReplaceAll('word', 'new', options)).toBe(0);
  expect(complete).not.toHaveBeenCalled();
  expect(paragraphs).not.toHaveBeenCalled();
});

test('empty search clears results and invalidates a pending viewer search', async () => {
  const { hook, editor, findReplace } = mount();
  const pending = deferred<Read>();
  editor.readViewerFindMatches.mockReturnValueOnce(pending.promise);
  act(() => {
    hook.result.current.handleFind('word', options);
    hook.result.current.handleFind('  ', options);
  });
  await act(async () => { pending.resolve({ version: 'v1', matches }); });
  expect(hook.result.current.findResultRef.current).toBeNull();
  expect(findReplace.setMatches).toHaveBeenLastCalledWith([], 0);
  expect(editor.setSelection).not.toHaveBeenCalled();
});

test('null worker result leaves the viewer results empty', async () => {
  const { hook, editor, findReplace } = mount();
  editor.readViewerFindMatches.mockResolvedValueOnce(null);
  await act(async () => { hook.result.current.handleFind('word', options); });
  expect(hook.result.current.findResultRef.current).toBeNull();
  expect(findReplace.setMatches).toHaveBeenLastCalledWith([], 0);
  expect(editor.setSelection).not.toHaveBeenCalled();
});

for (const change of ['editor', 'session', 'viewer', 'closed', 'reopened', 'unmount'] as const) {
  test(`viewer result is ignored after changing ${change}`, async () => {
    const { hook, editor, pagedEditorRef, findReplace } = mount();
    const pending = deferred<Read>();
    editor.readViewerFindMatches.mockReturnValueOnce(pending.promise);
    act(() => { hook.result.current.handleFind('word', options); });
    if (change === 'editor') pagedEditorRef.current = null;
    if (change === 'session') editor.getYrsSession.mockReturnValue({} as YrsSession);
    if (change === 'viewer') editor.isWorkerViewer.mockReturnValue(false);
    if (change === 'closed' || change === 'reopened') {
      findReplace.state.isOpen = false;
      hook.rerender();
    }
    if (change === 'reopened') {
      findReplace.state.isOpen = true;
      hook.rerender();
    }
    if (change === 'unmount') hook.unmount();
    await act(async () => { pending.resolve({ version: 'v1', matches }); });
    expect(editor.setSelection).not.toHaveBeenCalled();
    expect(hook.result.current.findResultRef.current).toBeNull();
  });
}

test('a viewer result still applies after the editor ref is rebuilt for the same session', async () => {
  const { hook, editor, pagedEditorRef } = mount();
  const pending = deferred<Read>();
  editor.readViewerFindMatches.mockReturnValueOnce(pending.promise);
  act(() => { hook.result.current.handleFind('word', options); });
  pagedEditorRef.current = { ...editor } as unknown as PagedEditorRef;
  await act(async () => { pending.resolve({ version: 'v1', matches }); });
  expect(editor.setSelection).toHaveBeenCalledWith(1, 5);
  expect(hook.result.current.findResultRef.current).toEqual({ matches, totalCount: 2, currentIndex: 0 });
});

test('viewer next and previous search again when the document version changed', async () => {
  const { hook, editor, version, findReplace } = mount();
  await act(async () => { hook.result.current.handleFind('word', options); });
  const shifted = [match(4), match(14), match(24)];
  version.mockReturnValue('v2');
  editor.readViewerFindMatches.mockResolvedValueOnce({ version: 'v2', matches: shifted });
  await act(async () => { expect(hook.result.current.handleFindNext()).toBeNull(); });
  expect(editor.readViewerFindMatches).toHaveBeenCalledTimes(2);
  expect(hook.result.current.findResultRef.current).toEqual({ matches: shifted, totalCount: 3, currentIndex: 1 });
  expect(findReplace.setMatches).toHaveBeenLastCalledWith(shifted, 1);
  expect(editor.setSelection).toHaveBeenLastCalledWith(15, 19);
  act(() => { expect(hook.result.current.handleFindPrevious()).toEqual(shifted[0]); });
  expect(editor.readViewerFindMatches).toHaveBeenCalledTimes(2);
});

test('viewer next searches again after an empty result when the document version changed', async () => {
  const { hook, editor, version } = mount();
  editor.readViewerFindMatches.mockResolvedValueOnce({ version: 'v1', matches: [] });
  await act(async () => { hook.result.current.handleFind('word', options); });
  act(() => { expect(hook.result.current.handleFindNext()).toBeNull(); });
  expect(editor.readViewerFindMatches).toHaveBeenCalledTimes(1);
  version.mockReturnValue('v2');
  await act(async () => { expect(hook.result.current.handleFindNext()).toBeNull(); });
  expect(editor.readViewerFindMatches).toHaveBeenCalledTimes(2);
  expect(hook.result.current.findResultRef.current).toEqual({ matches, totalCount: 2, currentIndex: 0 });
  expect(editor.setSelection).toHaveBeenLastCalledWith(1, 5);
});

test('editor find remains synchronous and maps and selects Yrs ranges', () => {
  const { hook, editor, session, findReplace } = mount(false);
  const expected = [match(0), match(5, 'Word')];
  act(() => {
    expect(hook.result.current.handleFind('word', options)).toEqual({
      matches: expected, totalCount: 2, currentIndex: 0,
    });
  });
  expect(findReplace.setMatches).toHaveBeenLastCalledWith(expected, 0);
  expect(session.setSelection).toHaveBeenCalledWith(
    { story: 'body', paraId: 'p', offset: 0 }, { story: 'body', paraId: 'p', offset: 4 }
  );
  expect(editor.syncYrsInputState).toHaveBeenCalledWith(false);
  expect(editor.scrollToPosition).toHaveBeenCalledWith(1);
  expect(editor.readViewerFindMatches).not.toHaveBeenCalled();
  expect(editor.setSelection).not.toHaveBeenCalled();
});
