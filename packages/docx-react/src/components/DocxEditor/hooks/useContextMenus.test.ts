import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, describe, expect, test } from 'bun:test';
import type { PagedEditorRef } from '../PagedEditor';
import { useContextMenus } from './useContextMenus';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');

function options(
  partEditOpen: boolean,
  onAddComment: (range: { from: number; to: number; yPos: number | null }) => void = () => {}
) {
  return {
    pagedEditorRef: {
      current: {
        getYrsSession: () => null,
        getSelectionRange: () => ({ from: 1, to: 2 }),
      } as PagedEditorRef,
    },
    focusActiveEditor: () => {},
    openSplitCellDialog: () => {},
    editorContentRef: { current: null },
    displayListQueries: null,
    interactionPageHostRef: { current: null },
    i18n: undefined,
    partEditOpen,
    onAddComment,
  };
}

function openSelectionMenu(result: { current: ReturnType<typeof useContextMenus> }): void {
  act(() => result.current.handleContextMenu({ x: 10, y: 20, hasSelection: true }));
}

afterEach(cleanup);
afterAll(() => {
  if (ownsDom) GlobalRegistrator.unregister();
});

describe('selection context menu', () => {
  test('offers Comment for a body selection', () => {
    const { result } = renderHook(() => useContextMenus(options(false)));
    openSelectionMenu(result);

    expect(result.current.contextMenuItems.map((item) => item.action)).toContain('addComment');
  });

  test('does not offer Comment for a selection in an open note', () => {
    const { result } = renderHook(() => useContextMenus(options(true)));
    openSelectionMenu(result);

    expect(result.current.contextMenuItems.map((item) => item.action)).not.toContain('addComment');
  });

  test('rejects a displayed Comment action when a part opens before dispatch', async () => {
    let added = 0;
    const onAddComment = () => {
      added += 1;
    };
    const { result, rerender } = renderHook(
      ({ partEditOpen }) => useContextMenus(options(partEditOpen, onAddComment)),
      { initialProps: { partEditOpen: false } }
    );
    openSelectionMenu(result);
    expect(result.current.contextMenuItems.map((item) => item.action)).toContain('addComment');

    rerender({ partEditOpen: true });
    await act(async () => {
      await result.current.handleContextMenuAction('addComment');
    });

    expect(added).toBe(0);
  });
});
