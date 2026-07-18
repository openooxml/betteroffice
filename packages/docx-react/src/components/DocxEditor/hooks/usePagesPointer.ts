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
import type { YrsPositionProjection } from '../internals/yrsPositionProjection';
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
  hfEditMode?: 'header' | 'footer' | null;
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
  setSelectionRects: React.Dispatch<React.SetStateAction<SelectionRect[]>>;
  setCaretPosition: React.Dispatch<React.SetStateAction<CaretPosition | null>>;
  setIsFocused: React.Dispatch<React.SetStateAction<boolean>>;
  scrollToPositionImpl: (pmPos: number, forParaIdScroll?: boolean) => void;
}

export interface UsePagesPointerReturn {
  handlePagesMouseDown: (e: React.MouseEvent) => void;
  handlePagesMouseMove: (e: React.MouseEvent) => void;
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
    hfEditMode,
    displayListQueries,
    canvasHostRef,
    canvasOverlayTarget,
    onBodyClick,
    onContextMenu,
    onHyperlinkClick,
    onHeaderFooterDoubleClick,
    setSelectionRects,
    setCaretPosition,
    setIsFocused,
    scrollToPositionImpl,
  } = opts;

  const isDraggingRef = useRef(false);
  const dragAnchorRef = useRef<number | null>(null);
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

  const getPositionFromMouse = useCallback(
    (clientX: number, clientY: number): number | null => {
      const hit = resolveCanvasHit(clientX, clientY, isDraggingRef.current)?.hit;
      if (!hit) return null;
      if (hfEditMode) return hit.region === hfEditMode ? hit.pos : null;
      return hit.region === 'body' ? hit.pos : null;
    },
    [hfEditMode, resolveCanvasHit]
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
      const projection = getYrsPositionProjection(yrsRootStory);
      if (!projection) return;
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
      const region = point?.hit?.region ?? null;
      if (hfEditMode) {
        if (region !== hfEditMode && onBodyClick) {
          e.stopPropagation();
          onBodyClick();
          return;
        }
      } else if ((region === 'header' || region === 'footer') && e.detail !== 2) {
        return;
      }

      if (displayListQueries && point) {
        const image = displayListQueries.imageAtPoint(
          point.pageIndex,
          point.x,
          point.y,
          hfEditMode ?? 'body',
          point.hit?.rId
        );
        if (image) {
          e.stopPropagation();
          setTextSelection(image.pos, image.pos + 1);
          setSelectionRects([]);
          setCaretPosition(null);
          focusInput();
          if (!hfEditMode) setIsFocused(true);
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
      if (!hfEditMode) setIsFocused(true);
    },
    [
      clearTableInsertTimer,
      displayListQueries,
      focusInput,
      getPositionFromMouse,
      getYrsPositionProjection,
      hfEditMode,
      onBodyClick,
      readOnly,
      resolveCanvasHit,
      resolveTarget,
      setCaretPosition,
      setIsFocused,
      setSelectionRects,
      setTextSelection,
      yrsRootStory,
    ]
  );

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

  const handleMouseUp = useCallback(() => {
    isDraggingRef.current = false;
    yrsCellDragAnchorRef.current = null;
    yrsCellDraggingRef.current = false;
    stopDragAutoScroll();
  }, [stopDragAutoScroll]);

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
      if (readOnly || isDraggingRef.current || yrsCellDraggingRef.current) return;
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
      const point = resolveCanvasHit(e.clientX, e.clientY, false);
      const pageSize = point ? queries.pageSize(point.pageIndex) : null;
      const pageRect = point ? resolveDisplayPageClientRect(host, queries, point.pageIndex) : null;
      if (!point || !pageSize || !pageRect) {
        scheduleHide();
        return;
      }
      let region: DisplayListTableRegion = { kind: 'body' };
      if (hfEditMode) {
        if (point.hit?.region !== hfEditMode || !point.hit.rId) {
          scheduleHide();
          return;
        }
        region = { kind: hfEditMode, rId: point.hit.rId };
      } else if (point.hit?.region === 'header' || point.hit?.region === 'footer') {
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
      hfEditMode,
      pagesContainerRef,
      readOnly,
      resolveCanvasHit,
      yrsRootStory,
    ]
  );

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
        const hitRegion = point.hit?.region;
        let region: DisplayListTableRegion = { kind: 'body' };
        if (hfEditMode && hitRegion === hfEditMode) {
          region = { kind: hfEditMode, rId: point.hit?.rId };
        }
        const displayHit =
          !hfEditMode || hitRegion === hfEditMode
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

      if (e.detail === 2 && !hfEditMode && onHeaderFooterDoubleClick) {
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
      hfEditMode,
      onHeaderFooterDoubleClick,
      onHyperlinkClick,
      pagesContainerRef,
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
      if (displayListQueries && point) {
        const image = displayListQueries.imageAtPoint(
          point.pageIndex,
          point.x,
          point.y,
          hfEditMode ?? 'body',
          point.hit?.rId
        );
        if (image) imageInfo = readImageNodeAt(image.pos);
      }
      const selection = yrsInputRef.current?.displaySelection();
      if (!imageInfo && selection && Math.abs(selection.anchor - selection.head) === 1) {
        imageInfo = readImageNodeAt(Math.min(selection.anchor, selection.head));
      }
      if (imageInfo?.wrapType === 'inline' && displayListQueries && !hfEditMode) {
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
        if (!hfEditMode) setIsFocused(true);
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
      hfEditMode,
      onContextMenu,
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
    click: handlePagesClick,
    contextmenu: handlePagesContextMenu,
  });
  canvasHandlersRef.current = {
    mousedown: handlePagesMouseDown,
    mousemove: handlePagesMouseMove,
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
      if (onCurrentCanvas(e)) canvasHandlersRef.current.mousemove(asReactEvent(e));
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
    handlePagesClick,
    handlePagesContextMenu,
    handleTableInsertClick,
    tableInsertButton,
    clearTableInsertTimer,
    hideTableInsertButton,
    getPositionFromMouse,
  };
}
