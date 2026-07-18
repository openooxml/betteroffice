import { useCallback, useEffect, useRef } from 'react';
import {
  resolveDisplayPageClientRect,
  type DisplayListQueries,
} from '@betteroffice/docx/layout/render';
import type { PagedEditorRef } from '../PagedEditor';

/**
 * Owns the floating "Add comment" button that hovers next to a
 * non-empty selection on the right edge of the page. Recomputes its
 * position whenever:
 *  - the selection changes (caller invokes `recomputeFloatingCommentBtn`)
 *  - the scroll container resizes (ResizeObserver here)
 *  - the window resizes (`resize` listener here)
 *  - zoom changes (effect on `zoom`)
 *
 * Why both `ResizeObserver` and the explicit `resize` listener: the
 * ResizeObserver covers container-size changes (sidebar toggle,
 * loading→ready transition) but doesn't fire on pure window resize when
 * the container is already at its max-width. The zoom effect handles
 * zoom changes that move page edges without changing PM selection — the
 * PagedEditor's `onSelectionChange` no longer fires on mere overlay
 * redraws after the state-identity dedup in #268.
 *
 * `readOnly` is mirrored to a ref so the recompute callback stays
 * stable across renders.
 */
export function useFloatingCommentBtn({
  pagedEditorRef,
  scrollContainerRef,
  editorContentRef,
  isAddingCommentRef,
  setFloatingCommentBtn,
  readOnly,
  isLoading,
  zoom,
  canvasHostRef,
  displayListQueries,
}: {
  pagedEditorRef: React.RefObject<PagedEditorRef | null>;
  scrollContainerRef: React.RefObject<HTMLDivElement | null>;
  editorContentRef: React.RefObject<HTMLDivElement | null>;
  isAddingCommentRef: React.RefObject<boolean>;
  setFloatingCommentBtn: React.Dispatch<React.SetStateAction<{ top: number; left: number } | null>>;
  readOnly: boolean;
  isLoading: boolean;
  zoom: number;
  /** `.canvas-pages` host — the canvas path measures against its live canvas rects */
  canvasHostRef?: React.RefObject<HTMLDivElement | null>;
  /** display-list queries — non-null exactly while the canvas renderer paints */
  displayListQueries?: DisplayListQueries | null;
}) {
  const readOnlyForFloatingBtnRef = useRef(false);
  readOnlyForFloatingBtnRef.current = readOnly;

  const recomputeFloatingCommentBtn = useCallback(() => {
    if (isAddingCommentRef.current || readOnlyForFloatingBtnRef.current) {
      setFloatingCommentBtn(null);
      return;
    }
    const selection = pagedEditorRef.current?.getSelectionRange();
    if (!selection) return;
    const { from, to } = selection;
    if (from === to) {
      setFloatingCommentBtn(null);
      return;
    }
    const parentEl = editorContentRef.current;
    if (!parentEl) return;

    const queries = displayListQueries;
    const host = canvasHostRef?.current;
    if (!queries || !host) return;
    const rect = queries.anchorRect(from);
    if (!rect) return;
    const pageRect = resolveDisplayPageClientRect(host, queries, rect.pageIndex);
    const size = queries.pageSize(rect.pageIndex);
    if (!pageRect || !size) return;
    const parentRect = parentEl.getBoundingClientRect();
    const scaleY = size.height > 0 ? pageRect.height / size.height : 1;
    const top = pageRect.top - parentRect.top + rect.y * scaleY;
    const left = pageRect.right - parentRect.left;
    setFloatingCommentBtn({ top, left });
  }, [
    pagedEditorRef,
    scrollContainerRef,
    editorContentRef,
    isAddingCommentRef,
    setFloatingCommentBtn,
    canvasHostRef,
    displayListQueries,
  ]);

  // Reposition on container resize (sidebar toggle, loading→ready, window
  // resize). Re-run on isLoading flip because the scroll container only
  // mounts once the doc is ready.
  useEffect(() => {
    const container = scrollContainerRef.current;
    if (!container) return;
    const ro = new ResizeObserver(() => recomputeFloatingCommentBtn());
    ro.observe(container);
    const onWinResize = () => recomputeFloatingCommentBtn();
    window.addEventListener('resize', onWinResize);
    return () => {
      ro.disconnect();
      window.removeEventListener('resize', onWinResize);
    };
  }, [isLoading, recomputeFloatingCommentBtn, scrollContainerRef]);

  // Reposition on zoom — page edges shift without a selection change.
  useEffect(() => {
    recomputeFloatingCommentBtn();
  }, [zoom, recomputeFloatingCommentBtn]);

  return { recomputeFloatingCommentBtn };
}
