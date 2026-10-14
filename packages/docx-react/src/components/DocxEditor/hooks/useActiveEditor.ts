import { useCallback } from 'react';
import type { PagedEditorRef } from '../PagedEditor';

/** Stable focus callback for the sole yrs-backed editor. */
export function useActiveEditor({
  pagedEditorRef,
}: {
  pagedEditorRef: React.RefObject<PagedEditorRef | null>;
}) {
  const focusActiveEditor = useCallback(() => {
    pagedEditorRef.current?.focus();
  }, [pagedEditorRef]);

  return { focusActiveEditor };
}
