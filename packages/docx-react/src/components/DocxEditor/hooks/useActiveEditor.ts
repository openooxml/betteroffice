import { useCallback } from 'react';
import type { PagedEditorRef } from '../PagedEditor';

/** Stable focus/history callbacks for the sole yrs-backed editor. */
export function useActiveEditor({
  pagedEditorRef,
}: {
  pagedEditorRef: React.RefObject<PagedEditorRef | null>;
}) {
  const focusActiveEditor = useCallback(() => {
    pagedEditorRef.current?.focus();
  }, [pagedEditorRef]);

  const undoActiveEditor = useCallback(() => {
    pagedEditorRef.current?.undo();
  }, [pagedEditorRef]);

  const redoActiveEditor = useCallback(() => {
    pagedEditorRef.current?.redo();
  }, [pagedEditorRef]);

  return { focusActiveEditor, undoActiveEditor, redoActiveEditor };
}
