import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { HeadingInfo } from '@betteroffice/docx/utils';
import type { ResidentEngineWorkerClient } from '@betteroffice/docx/yrs';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import { presentedWorkerVersion } from '../internals/layoutProvenance';
import { ViewerOutlineReads } from '../internals/viewerOutlineReads';
import type { PagedEditorRef } from '../PagedEditor';

/** The body's headings as the engine classifies them, with their display positions. */
export function collectYrsHeadings(editor: PagedEditorRef | null): HeadingInfo[] {
  if (!editor) return [];
  const session = editor.getYrsSession();
  if (!session) return [];
  const headings: HeadingInfo[] = [];
  const levels = new Map(
    session.headings('body').map((entry) => [entry.paraId, entry.heading.outlineLevel])
  );
  for (const paragraph of session.paragraphs('body')) {
    const level = levels.get(paragraph.paraId);
    if (level == null || !paragraph.text.trim()) continue;
    const contentPosition = editor.yrsLocToDisplayPosition({
      story: 'body',
      paraId: paragraph.paraId,
      offset: 0,
    });
    if (contentPosition == null) continue;
    headings.push({ text: paragraph.text.trim(), level, pmPos: Math.max(0, contentPosition - 1) });
  }
  return headings;
}

/**
 * Owns the document outline panel: visibility, headings, and chrome
 * measurements that position it (toolbar height + horizontal scroll
 * offset of the editor).
 */
export function useOutlineSidebar({
  showOutlineProp,
  pagedEditorRef,
  scrollContainerRef,
  isLoading,
  viewerRef,
  viewerRead,
  queries,
}: {
  showOutlineProp: boolean;
  pagedEditorRef: React.RefObject<PagedEditorRef | null>;
  scrollContainerRef: React.RefObject<HTMLDivElement | null>;
  isLoading: boolean;
  viewerRef?: React.RefObject<boolean>;
  viewerRead?: ResidentEngineWorkerClient['documentRead'];
  queries?: DisplayListQueries | null;
}) {
  const [showOutline, setShowOutline] = useState(showOutlineProp);
  const showOutlineRef = useRef(false);
  showOutlineRef.current = showOutline;
  const [outlineHeadings, setHeadingInfos] = useState<HeadingInfo[]>([]);
  const queriesRef = useRef(queries);
  queriesRef.current = queries;
  const viewerReads = useMemo(() => viewerRead ? new ViewerOutlineReads(viewerRead) : null, [viewerRead]);
  const activeReads = useRef(viewerReads);
  activeReads.current = viewerReads;
  const refreshHeadings = useCallback(() => {
    if (!viewerRef?.current) {
      setHeadingInfos(collectYrsHeadings(pagedEditorRef.current));
      return;
    }
    const version = presentedWorkerVersion(queriesRef.current);
    if (!viewerReads || version === null) return;
    void viewerReads.collect(version, () => presentedWorkerVersion(queriesRef.current)).then((headings) => {
      if (headings && viewerRef?.current && activeReads.current === viewerReads) setHeadingInfos(headings);
    });
  }, [pagedEditorRef, viewerRef, viewerReads]);
  const navigateViewerHeading = useCallback((pmPos: number) => {
    void viewerReads?.navigate(pmPos, () => presentedWorkerVersion(queriesRef.current), (position) => {
      if (viewerRef?.current && activeReads.current === viewerReads) {
        pagedEditorRef.current?.scrollToPosition(position);
      }
    });
  }, [viewerReads, viewerRef, pagedEditorRef]);
  const version = presentedWorkerVersion(queries);
  useEffect(() => {
    if (viewerRef?.current && showOutline) refreshHeadings();
  }, [version, showOutline, isLoading, viewerRef, refreshHeadings]);

  // Sync outline visibility when prop changes
  useEffect(() => {
    setShowOutline(showOutlineProp);
    if (showOutlineProp) refreshHeadings();
  }, [showOutlineProp, refreshHeadings]);

  // Horizontal scroll offset of the editor scroll container. Used to slide the
  // outline panel and toggle button with the doc instead of leaving them pinned
  // to the viewport. Scroll updates are coalesced to one per frame — scroll
  // events fire faster than React can re-render the whole editor tree.
  // Re-runs after isLoading flips because the scroll container only mounts once
  // the doc is ready.
  const [editorScrollLeft, setEditorScrollLeft] = useState(0);
  useEffect(() => {
    const el = scrollContainerRef.current;
    if (!el) return;
    let frame = 0;
    const update = () => {
      frame = 0;
      setEditorScrollLeft(el.scrollLeft);
    };
    const onScroll = () => {
      if (frame === 0) frame = requestAnimationFrame(update);
    };
    update();
    el.addEventListener('scroll', onScroll, { passive: true });
    return () => {
      el.removeEventListener('scroll', onScroll);
      if (frame !== 0) cancelAnimationFrame(frame);
    };
  }, [isLoading, scrollContainerRef]);

  return {
    showOutline,
    setShowOutline,
    showOutlineRef,
    outlineHeadings,
    setHeadingInfos,
    refreshHeadings,
    navigateViewerHeading,
    editorScrollLeft,
  };
}
