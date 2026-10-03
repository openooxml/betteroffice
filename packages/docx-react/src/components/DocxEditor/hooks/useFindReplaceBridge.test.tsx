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
  const session = { paragraphs, setSelection: mock(() => {}) } as unknown as YrsSession;
  const editor = {
    isWorkerViewer: mock(() => viewer),
    getYrsSession: mock(() => session),
    readViewerFindMatches: mock(async () => matches as DocxFindDisplayMatch[] | null),
    setSelection: mock(() => {}),
    scrollToPosition: mock(() => {}),
    syncYrsInputState: mock(() => true),
    yrsLocToDisplayPosition: mock((loc: { offset: number }) => loc.offset + 1),
  };
  const pagedEditorRef = { current: editor as unknown as PagedEditorRef | null };
  const findReplace = { setMatches: mock(() => {}), goToMatch: mock((_index: number) => {}) };
  const complete = mock(async () => ({ ok: true, status: 'executed' } as const));
  const hook = renderHook(() => useFindReplaceBridge({
    pagedEditorRef, findReplace: findReplace as unknown as ReturnType<typeof useFindReplace>, complete,
  }));
  return { hook, editor, session, paragraphs, pagedEditorRef, findReplace, complete };
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
  const first = deferred<DocxFindDisplayMatch[] | null>();
  const second = deferred<DocxFindDisplayMatch[] | null>();
  editor.readViewerFindMatches.mockReturnValueOnce(first.promise).mockReturnValueOnce(second.promise);
  act(() => {
    hook.result.current.handleFind('word', options);
    hook.result.current.handleFind('next', options);
  });
  const next = [match(20, 'next')];
  await act(async () => { second.resolve(next); });
  await act(async () => { first.resolve(matches); });
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
  const pending = deferred<DocxFindDisplayMatch[] | null>();
  editor.readViewerFindMatches.mockReturnValueOnce(pending.promise);
  act(() => {
    hook.result.current.handleFind('word', options);
    hook.result.current.handleFind('  ', options);
  });
  await act(async () => { pending.resolve(matches); });
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

for (const change of ['editor', 'session', 'viewer', 'unmount'] as const) {
  test(`viewer result is ignored after changing ${change}`, async () => {
    const { hook, editor, pagedEditorRef } = mount();
    const pending = deferred<DocxFindDisplayMatch[] | null>();
    editor.readViewerFindMatches.mockReturnValueOnce(pending.promise);
    act(() => { hook.result.current.handleFind('word', options); });
    if (change === 'editor') pagedEditorRef.current = null;
    if (change === 'session') editor.getYrsSession.mockReturnValue({} as YrsSession);
    if (change === 'viewer') editor.isWorkerViewer.mockReturnValue(false);
    if (change === 'unmount') hook.unmount();
    await act(async () => { pending.resolve(matches); });
    expect(editor.setSelection).not.toHaveBeenCalled();
    expect(hook.result.current.findResultRef.current).toBeNull();
  });
}

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
