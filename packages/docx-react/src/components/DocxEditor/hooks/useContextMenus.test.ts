import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, describe, expect, test } from 'bun:test';
import type { PagedEditorRef } from '../PagedEditor';
import { useContextMenus } from './useContextMenus';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');

function options(partEditOpen: boolean) {
  return {
    pagedEditorRef: {
      current: { getYrsSession: () => null } as PagedEditorRef,
    },
    focusActiveEditor: () => {},
    openSplitCellDialog: () => {},
    editorContentRef: { current: null },
    displayListQueries: null,
    interactionPageHostRef: { current: null },
    i18n: undefined,
    partEditOpen,
    onAddComment: () => {},
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
});
