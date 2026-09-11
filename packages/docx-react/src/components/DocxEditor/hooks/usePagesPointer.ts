import { useCallback, useEffect, useRef, useState } from 'react';

import type { CaretPosition, SelectionRect } from '@betteroffice/docx/layout';
import type { WrapType } from '@betteroffice/docx/docx/wrapTypes';
import {
  captureInlinePositionEmuFromDisplayList,
  DISPLAY_LIST_TABLE_INSERT_HIDE_DELAY_MS as TABLE_INSERT_HIDE_DELAY,
  detectDisplayListTableInsertHover,
  findDisplayListHyperlinkAtPoint,
  resolveCanvasPoint,
  resolveDisplayPageClientRect,
  type CanvasPointHit,
  type DisplayListQueries,
  type DisplayListTableRegion,
} from '@betteroffice/docx/layout/render';
import { sanitizeHref } from '@betteroffice/docx/utils';
import type { YrsCellLoc, YrsSession } from '@betteroffice/docx/yrs';

import type { YrsInputRef } from '../YrsInput';
import { useDragAutoScroll } from '../../../hooks/useDragAutoScroll';
import {
  canvasHoverCursor,
  createCursorPainter,
  type CanvasHoverCursor,
} from '../hoverCursor';
import type { YrsPositionProjection } from '../internals/yrsPositionProjection';
import {
  hitBelongsToPart,
  isNoteAreaHit,
  noteEditFromHit,
  partEditStory,
  partImageRegion,
  type NoteEdit,
  type PartEdit,
} from '../partEdit';
import type { YrsEditorCommand } from '../yrsCommands';

interface TableInsertButtonState {
  type: 'row' | 'column';
  x: number;
  y: number;
  cellPmPos: number;
}

interface ImageInfo {
  pos: number;
  wrapType: WrapType;
  cssFloat?: 'left' | 'right' | 'none' | null;
  inlinePositionEmu?: { horizontalEmu: number; verticalEmu: number };
}

export interface UsePagesPointerOptions {
  pagesContainerRef: React.RefObject<HTMLDivElement | null>;
  yrsInputRef: React.RefObject<YrsInputRef | null>;
  yrsSession: YrsSession | null;
  yrsRootStory: string;
  getYrsPositionProjection: (rootStory: string) => YrsPositionProjection | null;
  applyYrsCommand: (command: YrsEditorCommand) => boolean;
  syncYrsInputState: (docChanged: boolean) => boolean;
  readOnly: boolean;
  /** the non-body part open for editing — the body is inert behind it */
  partEdit?: PartEdit | null;
  displayListQueries?: DisplayListQueries | null;
  canvasHostRef?: React.RefObject<HTMLDivElement | null>;
  canvasOverlayTarget?: HTMLElement | null;
  onBodyClick?: () => void;
  onContextMenu?: (data: {
    x: number;
    y: number;
    hasSelection: boolean;
    image?: ImageInfo | null;
  }) => void;
  onHyperlinkClick?: (data: {
    href: string;
    displayText: string;
    tooltip?: string;
    position: { top: number; left: number };
  }) => void;
  onHeaderFooterDoubleClick?: (position: 'header' | 'footer', pageNumber?: number) => void;
  /** A single click landed in a note area — the host opens it for editing. */
  onNoteClick?: (note: NoteEdit) => void;
  setSelectionRects: React.Dispatch<React.SetStateAction<SelectionRect[]>>;
  setCaretPosition: React.Dispatch<React.SetStateAction<CaretPosition | null>>;
  setIsFocused: React.Dispatch<React.SetStateAction<boolean>>;
  scrollToPositionImpl: (pmPos: number, forParaIdScroll?: boolean) => void;
}

export interface UsePagesPointerReturn {
  handlePagesMouseDown: (e: React.MouseEvent) => void;
  handlePagesMouseMove: (e: React.MouseEvent) => void;
  /** drops the hover cursor when the pointer leaves the pages */
  handlePagesMouseLeave: () => void;
  handlePagesClick: (e: React.MouseEvent) => void;
  handlePagesContextMenu: (e: React.MouseEvent) => void;
  handleTableInsertClick: (e: React.MouseEvent) => void;
  tableInsertButton: TableInsertButtonState | null;
  clearTableInsertTimer: () => void;
  hideTableInsertButton: () => void;
  getPositionFromMouse: (clientX: number, clientY: number) => number | null;
}

function sameYrsTable(a: YrsCellLoc, b: YrsCellLoc): boolean {
  return a.story === b.story && a.tableIndex === b.tableIndex;
}

function sameYrsCell(a: YrsCellLoc, b: YrsCellLoc): boolean {
  return sameYrsTable(a, b) && a.row === b.row && a.column === b.column;
}

function cellIsWithinYrsRange(
  cell: YrsCellLoc | undefined,
  range: ReturnType<YrsSession['cellSelection']>
): boolean {
  if (!cell || !range || !sameYrsTable(cell, range.anchor) || !sameYrsTable(cell, range.head)) {
    return false;
  }
  const top = Math.min(range.anchor.row, range.head.row);
  const bottom = Math.max(range.anchor.row, range.head.row);
  const left = Math.min(range.anchor.column, range.head.column);
  const right = Math.max(range.anchor.column, range.head.column);
  return cell.row >= top && cell.row <= bottom && cell.column >= left && cell.column <= right;
}

function createTableKeyResolver(
  projection: YrsPositionProjection
): (docStart: number | undefined) => string | null {
  const cache = new Map<number, string | null>();
  return (docStart) => {
    if (docStart == null) return null;
    const cached = cache.get(docStart);
    if (cached !== undefined) return cached;
    const table = projection.tableAtPosition(docStart);
    const key = table ? String(table.start) : null;
    cache.set(docStart, key);
    return key;
  };
}

function createCellPmPosResolver(
  projection: YrsPositionProjection
): (tableKey: string, row: number, col: number) => number | null {
  const cache = new Map<string, number | null>();
  return (tableKey, row, col) => {
    const key = `${tableKey}:${row}:${col}`;
    const cached = cache.get(key);
    if (cached !== undefined) return cached;
    const tableStart = Number(tableKey);
    const out = Number.isFinite(tableStart)
      ? projection.cellPosition(tableStart, row, col)
      : null;
    cache.set(key, out);
    return out;
  };
}

function projectPageLocalRectToClient(
  pageRect: { left: number; top: number; width: number; height: number },
  pageSize: { width: number; height: number },
  rect: { x: number; y: number; w: number; h: number }
): { left: number; top: number; right: number; bottom: number } {
  const scaleX = pageSize.width > 0 ? pageRect.width / pageSize.width : 1;
  const scaleY = pageSize.height > 0 ? pageRect.height / pageSize.height : 1;
  return {
    left: pageRect.left + rect.x * scaleX,
    top: pageRect.top + rect.y * scaleY,
    right: pageRect.left + (rect.x + rect.w) * scaleX,
    bottom: pageRect.top + (rect.y + rect.h) * scaleY,
  };
}

export function usePagesPointer(opts: UsePagesPointerOptions): UsePagesPointerReturn {
  const {
    pagesContainerRef,
    yrsInputRef,
    yrsSession,
    yrsRootStory,
    getYrsPositionProjection,
    applyYrsCommand,
    syncYrsInputState,
    readOnly,
    partEdit = null,
    displayListQueries,
    canvasHostRef,
    canvasOverlayTarget,
    onBodyClick,
    onContextMenu,
    onHyperlinkClick,
    onHeaderFooterDoubleClick,
    onNoteClick,
    setSelectionRects,
    setCaretPosition,
    setIsFocused,
    scrollToPositionImpl,
  } = opts;

  const isDraggingRef = useRef(false);
  const dragAnchorRef = useRef<number | null>(null);
  const pendingPartCaretRef = useRef<{
    session: YrsSession;
    story: string;
    position: number;
  } | null>(null);
  const [pendingPartCaretVersion, setPendingPartCaretVersion] = useState(0);
  const yrsCellDragAnchorRef = useRef<YrsCellLoc | null>(null);
  const yrsCellDraggingRef = useRef(false);
  const [tableInsertButton, setTableInsertButton] = useState<TableInsertButtonState | null>(null);
  const tableInsertHideTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  const clearTableInsertTimer = useCallback(() => {
    if (!tableInsertHideTimerRef.current) return;
    clearTimeout(tableInsertHideTimerRef.current);
    tableInsertHideTimerRef.current = null;
  }, []);

  useEffect(
    () => () => {
      if (tableInsertHideTimerRef.current) clearTimeout(tableInsertHideTimerRef.current);
    },
    []
  );

  const dragExtendRef = useRef<(cx: number, cy: number) => void>(() => {});
  const dragAutoScrollCallbackRef = useCallback((cx: number, cy: number) => {
    dragExtendRef.current(cx, cy);
  }, []);
  const { updateMousePosition: updateDragScroll, stopAutoScroll: stopDragAutoScroll } =
    useDragAutoScroll({
      pagesContainerRef,
      onScrollExtendSelection: dragAutoScrollCallbackRef,
    });

  const resolveCanvasHit = useCallback(
    (clientX: number, clientY: number, clampToNearestPage: boolean): CanvasPointHit | null => {
      if (!displayListQueries) return null;
      const host = canvasHostRef?.current ?? pagesContainerRef.current;
      return host
        ? resolveCanvasPoint(host, displayListQueries, clientX, clientY, { clampToNearestPage })
        : null;
    },
    [displayListQueries, canvasHostRef, pagesContainerRef]
  );

  // written straight to the DOM: hover fires on every pointer move, and
  // re-rendering the editor for it is not affordable
  const cursorPainterRef = useRef(createCursorPainter());
  const paintHoverCursor = useCallback(
    (cursor: CanvasHoverCursor) => {
      cursorPainterRef.current(canvasHostRef?.current ?? pagesContainerRef.current, cursor);
    },
    [canvasHostRef, pagesContainerRef]
  );
  // last point over the pages, so a cursor that outlived what it described can
  // be re-resolved without waiting for a move. Null while it is elsewhere.
  const hoverPointRef = useRef<{ x: number; y: number } | null>(null);
  const resolveHoverCursor = useCallback(
    (clientX: number, clientY: number) => {
      hoverPointRef.current = { x: clientX, y: clientY };
      const hit = readOnly ? null : (resolveCanvasHit(clientX, clientY, false)?.hit ?? null);
      paintHoverCursor(canvasHoverCursor({ readOnly, partEdit }, hit));
    },
    [paintHoverCursor, partEdit, readOnly, resolveCanvasHit]
  );
  // A mode flip changes what is typeable under a pointer that has not moved.
  // A gesture still holds its own cursor; the release resolves the new mode.
  useEffect(() => {
    const point = hoverPointRef.current;
    if (!point || isDraggingRef.current || yrsCellDraggingRef.current) return;
    resolveHoverCursor(point.x, point.y);
  }, [resolveHoverCursor]);
  // Clears the cursor on unmount only. Keying this on the painter would clear
  // it on every identity change too — the effect above re-resolves in the same
  // commit, so nothing shows, but the write is redundant and reads as an exit.
  const paintHoverCursorRef = useRef(paintHoverCursor);
  paintHoverCursorRef.current = paintHoverCursor;
  useEffect(() => () => paintHoverCursorRef.current('default'), []);

  const getPositionFromMouse = useCallback(
    (clientX: number, clientY: number): number | null => {
      const hit = resolveCanvasHit(clientX, clientY, isDraggingRef.current)?.hit;
      if (!hit) return null;
      if (partEdit) return hitBelongsToPart(partEdit, hit) ? hit.pos : null;
      return hit.region === 'body' ? hit.pos : null;
    },
    [partEdit, resolveCanvasHit]
  );

  const resolveTarget = useCallback(
    (position: number) => {
      const projection = getYrsPositionProjection(yrsRootStory);
      return projection ? projection.targetAt(position) : null;
    },
    [getYrsPositionProjection, yrsRootStory]
  );

  const setTextSelection = useCallback(
    (anchor: number, head = anchor): void => {
      const input = yrsInputRef.current;
      const anchorTarget = resolveTarget(anchor);
      const headTarget = resolveTarget(head);
      if (!input || !anchorTarget || !headTarget || anchorTarget.story !== headTarget.story) return;
      if (
        yrsSession &&
        anchorTarget.cell &&
        headTarget.cell &&
        sameYrsTable(anchorTarget.cell, headTarget.cell)
      ) {
        yrsSession.setCellSelection({ anchor: anchorTarget.cell, head: headTarget.cell });
      }
      input.setSelectionFromDisplay(
        anchorTarget.displayPosition,
        headTarget.displayPosition,
        anchorTarget.story
      );
    },
    [resolveTarget, yrsInputRef, yrsSession]
  );

  const focusInput = useCallback(() => yrsInputRef.current?.focus(), [yrsInputRef]);

  const extendCellSelection = useCallback(
    (pmPos: number): boolean => {
      const anchor = yrsCellDragAnchorRef.current;
      const head = resolveTarget(pmPos)?.cell;
      if (!anchor || !head || !yrsSession || !sameYrsTable(anchor, head)) return false;
      if (!yrsCellDraggingRef.current && sameYrsCell(anchor, head)) return false;
      yrsCellDraggingRef.current = true;
      yrsSession.setCellSelection({ anchor, head });
      // Cell ranges are session-owned ephemeral state, not document updates.
      // Publish the changed range explicitly so toolbar enablement follows
      // drag selection even though the sticky text caret stays in one cell.
      syncYrsInputState(false);
      setSelectionRects([]);
      setCaretPosition(null);
      return true;
    },
    [resolveTarget, setCaretPosition, setSelectionRects, syncYrsInputState, yrsSession]
  );

  const handlePagesMouseDown = useCallback(
    (e: React.MouseEvent) => {
      pendingPartCaretRef.current = null;
      if (e.button === 2) {
        e.preventDefault();
        return;
      }
      if (e.button !== 0) return;
      setTableInsertButton(null);
      clearTableInsertTimer();
      e.preventDefault();
      if (readOnly) return;

      const point = resolveCanvasHit(e.clientX, e.clientY, false);
      const hit = point?.hit ?? null;
      const region = hit?.region ?? null;
      // A note area opens on a single click, wherever the caret was before:
      // one part is open at a time, so this both leaves the old one and picks
      // the note up. The caret the click asked for is placed once the editor
      // is on that note's story.
      const note = noteEditFromHit(hit);
      if (isNoteAreaHit(hit) && !note) return;
      if (note && onNoteClick && !hitBelongsToPart(partEdit, hit)) {
        e.stopPropagation();
        const pending =
          hit?.pos == null || !yrsSession
            ? null
            : {
                session: yrsSession,
                story: partEditStory(note),
                position: hit.pos,
              };
        pendingPartCaretRef.current = pending;
        onNoteClick(note);
        if (pending) setPendingPartCaretVersion((version) => version + 1);
        return;
      }
      if (partEdit) {
        if (!hitBelongsToPart(partEdit, hit) && onBodyClick) {
          e.stopPropagation();
          const pending =
            hit?.region === 'body' && hit.pos != null && yrsSession
              ? { session: yrsSession, story: 'body', position: hit.pos }
              : null;
          pendingPartCaretRef.current = pending;
          onBodyClick();
          if (pending) setPendingPartCaretVersion((version) => version + 1);
          return;
        }
      } else if ((region === 'header' || region === 'footer') && e.detail !== 2) {
        return;
      }

      const projection = getYrsPositionProjection(yrsRootStory);
      if (!projection) return;

      const imageRegion = partImageRegion(partEdit);
      if (displayListQueries && point && imageRegion) {
        const image = displayListQueries.imageAtPoint(
          point.pageIndex,
          point.x,
          point.y,
          imageRegion,
          hit?.rId
        );
        if (image) {
          e.stopPropagation();
          setTextSelection(image.pos, image.pos + 1);
          setSelectionRects([]);
          setCaretPosition(null);
          focusInput();
          if (!partEdit) setIsFocused(true);
          return;
        }
      }

      const pmPos = getPositionFromMouse(e.clientX, e.clientY);
      const targetPos = pmPos ?? Math.max(0, projection.size - 1);
      yrsCellDragAnchorRef.current = resolveTarget(targetPos)?.cell ?? null;
      yrsCellDraggingRef.current = false;
      isDraggingRef.current = true;
      dragAnchorRef.current = targetPos;
      setTextSelection(targetPos);
      focusInput();
      if (!partEdit) setIsFocused(true);
    },
    [
      clearTableInsertTimer,
      displayListQueries,
      focusInput,
      getPositionFromMouse,
      getYrsPositionProjection,
      onBodyClick,
      onNoteClick,
      partEdit,
      readOnly,
      resolveCanvasHit,
      resolveTarget,
      setCaretPosition,
      setIsFocused,
      setSelectionRects,
      setTextSelection,
      yrsRootStory,
      yrsSession,
    ]
  );

  useEffect(() => {
    const pending = pendingPartCaretRef.current;
    if (!pending) return;
    pendingPartCaretRef.current = null;
    if (
      pending.session !== yrsSession ||
      pending.story !== yrsRootStory ||
      pending.story !== partEditStory(partEdit)
    ) {
      return;
    }
    const projection = getYrsPositionProjection(yrsRootStory);
    if (!projection) return;
    const position = Math.min(Math.max(0, pending.position), Math.max(0, projection.size - 1));
    setTextSelection(position);
    focusInput();
  }, [
    focusInput,
    getYrsPositionProjection,
    partEdit,
    pendingPartCaretVersion,
    setTextSelection,
    yrsRootStory,
    yrsSession,
  ]);

  dragExtendRef.current = (cx, cy) => {
    if (!isDraggingRef.current || dragAnchorRef.current == null) return;
    const pmPos = getPositionFromMouse(cx, cy);
    if (pmPos == null || extendCellSelection(pmPos)) return;
    setTextSelection(dragAnchorRef.current, pmPos);
  };

  const dragRafRef = useRef<number | null>(null);
  const pendingDragPointRef = useRef<{ x: number; y: number } | null>(null);
  useEffect(
    () => () => {
      if (dragRafRef.current != null) cancelAnimationFrame(dragRafRef.current);
    },
    []
  );

  const handleMouseMove = useCallback(
    (e: MouseEvent) => {
      if (!isDraggingRef.current || dragAnchorRef.current == null) return;
      updateDragScroll(e.clientX, e.clientY);
      pendingDragPointRef.current = { x: e.clientX, y: e.clientY };
      dragRafRef.current ??= requestAnimationFrame(() => {
        dragRafRef.current = null;
        const point = pendingDragPointRef.current;
        if (!point || dragAnchorRef.current == null) return;
        const pmPos = getPositionFromMouse(point.x, point.y);
        if (pmPos == null || extendCellSelection(pmPos)) return;
        setTextSelection(dragAnchorRef.current, pmPos);
      });
    },
    [extendCellSelection, getPositionFromMouse, setTextSelection, updateDragScroll]
  );

  const handleMouseUp = useCallback(
    (e: MouseEvent) => {
      const wasDragging = isDraggingRef.current || yrsCellDraggingRef.current;
      isDraggingRef.current = false;
      yrsCellDragAnchorRef.current = null;
      yrsCellDraggingRef.current = false;
      stopDragAutoScroll();
      // the gesture held its own cursor; a pointer that never moves again
      // gets its answer here
      if (wasDragging) resolveHoverCursor(e.clientX, e.clientY);
    },
    [resolveHoverCursor, stopDragAutoScroll]
  );

  useEffect(() => {
    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);
    return () => {
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
    };
  }, [handleMouseMove, handleMouseUp]);

  const handlePagesMouseMove = useCallback(
    (e: React.MouseEvent) => {
      // a live gesture keeps the cursor it started with, and the release is
      // what re-reads the point under it
      if (isDraggingRef.current || yrsCellDraggingRef.current) return;
      hoverPointRef.current = { x: e.clientX, y: e.clientY };
      const point = readOnly ? null : resolveCanvasHit(e.clientX, e.clientY, false);
      paintHoverCursor(canvasHoverCursor({ readOnly, partEdit }, point?.hit ?? null));
      if (readOnly) return;
      const scheduleHide = () => {
        tableInsertHideTimerRef.current ??= setTimeout(() => {
          setTableInsertButton(null);
          tableInsertHideTimerRef.current = null;
        }, TABLE_INSERT_HIDE_DELAY);
      };
      const queries = displayListQueries;
      const host = canvasHostRef?.current ?? pagesContainerRef.current;
      const overlayTarget = canvasOverlayTarget ?? pagesContainerRef.current?.parentElement;
      const projection = getYrsPositionProjection(yrsRootStory);
      if (!queries || !host || !overlayTarget || !projection) return;
      const pageSize = point ? queries.pageSize(point.pageIndex) : null;
      const pageRect = point ? resolveDisplayPageClientRect(host, queries, point.pageIndex) : null;
      if (!point || !pageSize || !pageRect) {
        scheduleHide();
        return;
      }
      // Table affordances are indexed by band; a note carries none, so an open
      // note has no table to offer a row/column on.
      let region: DisplayListTableRegion = { kind: 'body' };
      if (partEdit) {
        const inBand =
          (partEdit.kind === 'header' || partEdit.kind === 'footer') &&
          hitBelongsToPart(partEdit, point.hit ?? null);
        if (!inBand || !point.hit?.rId) {
          scheduleHide();
          return;
        }
        region = { kind: partEdit.kind, rId: point.hit.rId };
      } else if (point.hit && point.hit.region !== 'body') {
        scheduleHide();
        return;
      }
      const hit = detectDisplayListTableInsertHover({
        list: queries.displayList,
        pageIndex: point.pageIndex,
        x: point.x,
        y: point.y,
        canvasRect: pageRect,
        pageSize,
        tableKeyOf: createTableKeyResolver(projection),
        cellPmPosOf: createCellPmPosResolver(projection),
        region,
      });
      if (!hit) {
        scheduleHide();
        return;
      }
      const targetRect = overlayTarget.getBoundingClientRect();
      setTableInsertButton({
        type: hit.type,
        x: hit.clientX - targetRect.left,
        y: hit.clientY - targetRect.top,
        cellPmPos: hit.cellPmPos,
      });
      clearTableInsertTimer();
    },
    [
      canvasHostRef,
      canvasOverlayTarget,
      clearTableInsertTimer,
      displayListQueries,
      getYrsPositionProjection,
      pagesContainerRef,
      paintHoverCursor,
      partEdit,
      readOnly,
      resolveCanvasHit,
      yrsRootStory,
    ]
  );

  const handlePagesMouseLeave = useCallback(() => {
    // a drag that runs off the pages keeps its cursor, like the move path
    if (isDraggingRef.current || yrsCellDraggingRef.current) return;
    hoverPointRef.current = null;
    paintHoverCursor('default');
  }, [paintHoverCursor]);

  const handleTableInsertClick = useCallback(
    (e: React.MouseEvent) => {
      e.preventDefault();
      e.stopPropagation();
      if (!tableInsertButton) return;
      const at = resolveTarget(tableInsertButton.cellPmPos + 1)?.cell;
      if (!at) return;
      yrsSession?.setCellSelection({ anchor: at, head: at });
      setTextSelection(tableInsertButton.cellPmPos + 1);
      applyYrsCommand(
        tableInsertButton.type === 'row'
          ? { type: 'tableInsertRow', side: 'below', at }
          : { type: 'tableInsertColumn', side: 'right', at }
      );
      setTableInsertButton(null);
      focusInput();
    },
    [applyYrsCommand, focusInput, resolveTarget, setTextSelection, tableInsertButton, yrsSession]
  );

  const handlePagesClick = useCallback(
    (e: React.MouseEvent) => {
      // Native canvas clicks move focus to the document body after mousedown.
      // Reassert the hidden input from the terminal click event so pointer
      // selection remains keyboard-ready even when the canvas host is outside
      // PagedEditor's React subtree.
      focusInput();
      const projection = getYrsPositionProjection(yrsRootStory);
      const queries = displayListQueries;
      const host = canvasHostRef?.current ?? pagesContainerRef.current;
      const point = resolveCanvasHit(e.clientX, e.clientY, false);
      if (projection && queries && host && point) {
        // Hyperlink primitives are indexed by band, so an open note — whose
        // area the index does not cover — resolves none and falls through to
        // the multi-click selection below.
        const inOpenPart = hitBelongsToPart(partEdit, point.hit ?? null);
        const region: DisplayListTableRegion | null = !partEdit
          ? { kind: 'body' }
          : inOpenPart && (partEdit.kind === 'header' || partEdit.kind === 'footer')
            ? { kind: partEdit.kind, rId: point.hit?.rId }
            : null;
        const displayHit = region
          ? findDisplayListHyperlinkAtPoint(
              queries.displayList,
              point.pageIndex,
              point.x,
              point.y,
              region
            )
          : null;
        const href = sanitizeHref(displayHit?.href ?? '');
        if (href) {
          e.preventDefault();
          const linkPosition = getPositionFromMouse(e.clientX, e.clientY);
          if (linkPosition != null) setTextSelection(linkPosition);
          if (href.startsWith('#')) {
            const bookmarkName = href.slice(1);
            const targetPos = projection.bookmarkPosition(bookmarkName);
            if (targetPos != null) {
              scrollToPositionImpl(targetPos);
              setTextSelection(targetPos + 1);
            }
            return;
          }
          const selection = yrsInputRef.current?.displaySelection();
          if (onHyperlinkClick && selection?.anchor === selection?.head) {
            const pageSize = queries.pageSize(point.pageIndex);
            const pageRect = resolveDisplayPageClientRect(host, queries, point.pageIndex);
            const targetRect =
              (canvasOverlayTarget ?? host.closest('.oox-root.paged-editor'))?.getBoundingClientRect() ??
              null;
            let linkLeft = e.clientX;
            let linkBottom = e.clientY;
            if (displayHit && pageRect && pageSize) {
              const linkRect = projectPageLocalRectToClient(pageRect, pageSize, displayHit.rect);
              linkLeft = linkRect.left;
              linkBottom = linkRect.bottom;
            }
            if (targetRect) {
              onHyperlinkClick({
                href,
                displayText: displayHit?.displayText ?? href,
                tooltip: displayHit?.tooltip,
                position: {
                  top: linkBottom - targetRect.top + 4,
                  left: linkLeft - targetRect.left,
                },
              });
            }
          }
          return;
        }
      }

      if (e.detail === 2 && !partEdit && onHeaderFooterDoubleClick) {
        const region = point?.hit?.region;
        if (region === 'header' || region === 'footer') {
          e.preventDefault();
          e.stopPropagation();
          onHeaderFooterDoubleClick(region, (point?.pageIndex ?? 0) + 1);
          return;
        }
      }

      const pmPos = getPositionFromMouse(e.clientX, e.clientY);
      const target = pmPos != null ? resolveTarget(pmPos) : null;
      if (!target) return;
      if (e.detail === 2) {
        yrsInputRef.current?.selectWordAtDisplay(target.displayPosition, target.story);
        focusInput();
      } else if (e.detail === 3) {
        yrsInputRef.current?.selectParagraphAtDisplay(target.displayPosition, target.story);
        focusInput();
      }
    },
    [
      canvasHostRef,
      canvasOverlayTarget,
      displayListQueries,
      focusInput,
      getPositionFromMouse,
      getYrsPositionProjection,
      onHeaderFooterDoubleClick,
      onHyperlinkClick,
      pagesContainerRef,
      partEdit,
      resolveCanvasHit,
      resolveTarget,
      scrollToPositionImpl,
      setTextSelection,
      yrsInputRef,
      yrsRootStory,
    ]
  );

  const handlePagesContextMenu = useCallback(
    (e: React.MouseEvent) => {
      if (!onContextMenu) return;
      e.preventDefault();
      const projection = getYrsPositionProjection(yrsRootStory);
      if (!projection) return;
      const readImageNodeAt = (pos: number): ImageInfo | null => {
        const node = projection.nodeAt(pos);
        if (node?.kind !== 'image') return null;
        return {
          pos,
          wrapType: (node.attrs.wrapType as WrapType | undefined) ?? 'inline',
          cssFloat: node.attrs.cssFloat as ImageInfo['cssFloat'],
        };
      };
      let imageInfo: ImageInfo | null = null;
      const point = resolveCanvasHit(e.clientX, e.clientY, false);
      const imageRegion = partImageRegion(partEdit);
      if (displayListQueries && point && imageRegion) {
        const image = displayListQueries.imageAtPoint(
          point.pageIndex,
          point.x,
          point.y,
          imageRegion,
          point.hit?.rId
        );
        if (image) imageInfo = readImageNodeAt(image.pos);
      }
      const selection = yrsInputRef.current?.displaySelection();
      if (
        !imageInfo &&
        imageRegion &&
        selection &&
        Math.abs(selection.anchor - selection.head) === 1
      ) {
        imageInfo = readImageNodeAt(Math.min(selection.anchor, selection.head));
      }
      if (imageInfo?.wrapType === 'inline' && displayListQueries && !partEdit) {
        imageInfo.inlinePositionEmu = captureInlinePositionEmuFromDisplayList(
          displayListQueries,
          imageInfo.pos
        );
      }
      const pmPos = getPositionFromMouse(e.clientX, e.clientY);
      const contextCell = pmPos != null ? resolveTarget(pmPos)?.cell : undefined;
      const keepCellSelection = cellIsWithinYrsRange(contextCell, yrsSession?.cellSelection() ?? null);
      if (
        pmPos != null &&
        !keepCellSelection &&
        (!selection ||
          selection.anchor === selection.head ||
          pmPos < Math.min(selection.anchor, selection.head) ||
          pmPos > Math.max(selection.anchor, selection.head))
      ) {
        setTextSelection(pmPos);
        focusInput();
        if (!partEdit) setIsFocused(true);
      }
      const latest = yrsInputRef.current?.displaySelection();
      onContextMenu({
        x: e.clientX,
        y: e.clientY,
        hasSelection: !!latest && latest.anchor !== latest.head,
        image: imageInfo,
      });
    },
    [
      displayListQueries,
      focusInput,
      getPositionFromMouse,
      getYrsPositionProjection,
      onContextMenu,
      partEdit,
      resolveCanvasHit,
      resolveTarget,
      setIsFocused,
      setTextSelection,
      yrsInputRef,
      yrsRootStory,
      yrsSession,
    ]
  );

  const hideTableInsertButton = useCallback(() => setTableInsertButton(null), []);
  const canvasHandlersRef = useRef({
    mousedown: handlePagesMouseDown,
    mousemove: handlePagesMouseMove,
    mouseleave: handlePagesMouseLeave,
    click: handlePagesClick,
    contextmenu: handlePagesContextMenu,
  });
  canvasHandlersRef.current = {
    mousedown: handlePagesMouseDown,
    mousemove: handlePagesMouseMove,
    mouseleave: handlePagesMouseLeave,
    click: handlePagesClick,
    contextmenu: handlePagesContextMenu,
  };

  useEffect(() => {
    if (!displayListQueries) return;
    const asReactEvent = (e: MouseEvent) => e as unknown as React.MouseEvent;
    const onCurrentCanvas = (e: MouseEvent): boolean => {
      const target = e.target instanceof Element ? e.target.closest('.canvas-pages') : null;
      return target !== null && target === canvasHostRef?.current;
    };
    const onMouseDown = (e: MouseEvent) => {
      if (onCurrentCanvas(e)) canvasHandlersRef.current.mousedown(asReactEvent(e));
    };
    const onMouseMove = (e: MouseEvent) => {
      // no event fires once the pointer is off the pages, so the move that
      // leaves them is what drops the hover cursor
      if (onCurrentCanvas(e)) canvasHandlersRef.current.mousemove(asReactEvent(e));
      else canvasHandlersRef.current.mouseleave();
    };
    const onClick = (e: MouseEvent) => {
      if (onCurrentCanvas(e)) canvasHandlersRef.current.click(asReactEvent(e));
    };
    const onContextMenu = (e: MouseEvent) => {
      if (onCurrentCanvas(e)) canvasHandlersRef.current.contextmenu(asReactEvent(e));
    };
    // CanvasPagedArea briefly replaces its host while a new geometry cache
    // becomes ready. Delegate from the stable document so the interactive
    // surface never loses its listener during that ref-only handoff.
    document.addEventListener('mousedown', onMouseDown);
    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('click', onClick);
    document.addEventListener('contextmenu', onContextMenu);
    return () => {
      document.removeEventListener('mousedown', onMouseDown);
      document.removeEventListener('mousemove', onMouseMove);
      document.removeEventListener('click', onClick);
      document.removeEventListener('contextmenu', onContextMenu);
    };
  }, [canvasHostRef, displayListQueries]);

  return {
    handlePagesMouseDown,
    handlePagesMouseMove,
    handlePagesMouseLeave,
    handlePagesClick,
    handlePagesContextMenu,
    handleTableInsertClick,
    tableInsertButton,
    clearTableInsertTimer,
    hideTableInsertButton,
    getPositionFromMouse,
  };
}
