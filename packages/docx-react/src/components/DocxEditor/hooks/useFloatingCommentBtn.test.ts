import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, describe, expect, test } from 'bun:test';
import type { SetStateAction } from 'react';
import type { PagedEditorRef } from '../PagedEditor';
import { useFloatingCommentBtn } from './useFloatingCommentBtn';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');

afterEach(cleanup);
afterAll(() => {
  if (ownsDom) GlobalRegistrator.unregister();
});

describe('useFloatingCommentBtn', () => {
  test('clears the body comment affordance while a part is open', () => {
    let selectionReads = 0;
    const updates: Array<{ top: number; left: number } | null> = [];
    const setFloatingCommentBtn = (
      next: SetStateAction<{ top: number; left: number } | null>
    ) => {
      const previous = updates.length > 0 ? updates[updates.length - 1] : null;
      updates.push(typeof next === 'function' ? next(previous) : next);
    };
    const pagedEditorRef = {
      current: {
        getSelectionRange: () => {
          selectionReads += 1;
          return { from: 1, to: 2 };
        },
      } as PagedEditorRef,
    };
    const { result, rerender } = renderHook(
      ({ partEditOpen }) =>
        useFloatingCommentBtn({
          pagedEditorRef,
          scrollContainerRef: { current: null },
          editorContentRef: { current: null },
          isAddingCommentRef: { current: false },
          setFloatingCommentBtn,
          partEditOpen,
          readOnly: false,
          isLoading: false,
          zoom: 1,
        }),
      { initialProps: { partEditOpen: false } }
    );
    const readsBeforePartOpen = selectionReads;

    rerender({ partEditOpen: true });
    expect(updates[updates.length - 1]).toBeNull();

    act(() => result.current.recomputeFloatingCommentBtn());
    expect(selectionReads).toBe(readsBeforePartOpen);
    expect(updates[updates.length - 1]).toBeNull();
  });
});
