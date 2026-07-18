/**
 * Image-interaction handlers for PagedEditor.
 *
 * Owns the resize / drag callbacks the `ImageSelectionOverlay` invokes.
 * `isImageInteractingRef` is set during a drag or resize so the selection
 * hook can suppress the deferred image-info clear (image stays selected
 * mid-drag instead of dropping out under the mouse).
 *
 * Drag move handling forks on `displayMode === 'float'` (or any of
 * square/tight/through wrap types): floating images get an EMU offset
 * update under wp:positionH/V; inline images get a PM `delete + insert`
 * pair at the drop position.
 */

import { useCallback } from 'react';
import { pixelsToEmu } from '@betteroffice/docx/utils';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';

import { canvasPageScale } from '../internals/canvasProjection';
import type { YrsPositionProjection } from '../internals/yrsPositionProjection';
import type { YrsEditorCommand } from '../yrsCommands';

export interface UseImageInteractionsOptions {
  pagesContainerRef: React.RefObject<HTMLDivElement | null>;
  getPositionProjection: () => YrsPositionProjection | null;
  isImageInteractingRef: React.MutableRefObject<boolean>;
  getPositionFromMouse: (clientX: number, clientY: number) => number | null;
  /** `.canvas-pages` host — canvas float-drag resolves the content origin here (null on the DOM path). */
  canvasHostRef?: React.RefObject<HTMLDivElement | null>;
  /** Display-list queries — non-null exactly while the canvas renderer paints. */
  displayListQueries?: DisplayListQueries | null;
  applyYrsCommand: (command: YrsEditorCommand) => boolean;
}

/**
 * Page-local authored body content origin from display-list metadata.
 */
function canvasContentOrigin(
  queries: DisplayListQueries,
  pageIndex: number
): { left: number; top: number } | null {
  const bounds = queries.contentBounds(pageIndex);
  return bounds ? { left: bounds.x, top: bounds.y } : null;
}

export interface UseImageInteractionsReturn {
  handleImageResize: (pmPos: number, newWidth: number, newHeight: number) => void;
  handleImageResizeStart: () => void;
  handleImageResizeEnd: () => void;
  handleImageDragMove: (pmPos: number, clientX: number, clientY: number) => void;
  handleImageDragStart: () => void;
  handleImageDragEnd: () => void;
}

export function useImageInteractions(
  opts: UseImageInteractionsOptions
): UseImageInteractionsReturn {
  const {
    pagesContainerRef,
    getPositionProjection,
    isImageInteractingRef,
    getPositionFromMouse,
    canvasHostRef,
    displayListQueries,
    applyYrsCommand,
  } = opts;

  const handleImageResize = useCallback(
    (pmPos: number, newWidth: number, newHeight: number) => {
      applyYrsCommand({
        type: 'imageGeometry',
        pmPos,
        patch: { width: newWidth, height: newHeight },
      });
    },
    [applyYrsCommand]
  );

  const handleImageResizeStart = useCallback(() => {
    isImageInteractingRef.current = true;
  }, [isImageInteractingRef]);

  const handleImageResizeEnd = useCallback(() => {
    isImageInteractingRef.current = false;
  }, [isImageInteractingRef]);

  const handleImageDragMove = useCallback(
    (pmPos: number, clientX: number, clientY: number) => {
      const node = getPositionProjection()?.nodeAt(pmPos);
      if (!node || node.kind !== 'image') return;

      if (isFloatingImageAttributes(node.attrs)) {
        const host = canvasHostRef?.current ?? pagesContainerRef.current;
        if (displayListQueries && host) {
          const pageCount = displayListQueries.pageCount();
          let hit: { pageIndex: number; localX: number; localY: number } | null = null;
          for (let i = 0; i < pageCount; i++) {
            const scale = canvasPageScale(host, displayListQueries, i);
            if (!scale) continue;
            const { canvasRect, scaleX, scaleY } = scale;
            const withinY = clientY >= canvasRect.top && clientY <= canvasRect.bottom;
            const isLast = i === pageCount - 1;
            if (withinY || (isLast && hit === null)) {
              hit = {
                pageIndex: i,
                localX: scaleX > 0 ? (clientX - canvasRect.left) / scaleX : 0,
                localY: scaleY > 0 ? (clientY - canvasRect.top) / scaleY : 0,
              };
              if (withinY) break;
            }
          }
          if (!hit) return;
          const origin = canvasContentOrigin(displayListQueries, hit.pageIndex);
          if (!origin) return;
          const hOffsetEmu = pixelsToEmu(hit.localX - origin.left);
          const vOffsetEmu = pixelsToEmu(hit.localY - origin.top);
          applyYrsCommand({
            type: 'imageGeometry',
            pmPos,
            patch: {
              position: {
                horizontal: { posOffset: hOffsetEmu, relativeTo: 'margin' },
                vertical: { posOffset: vOffsetEmu, relativeTo: 'margin' },
              },
            },
          });
          return;
        }

        return;
      } else {
        // Inline image relocation is not represented by the geometry operation.
        getPositionFromMouse(clientX, clientY);
      }
    },
    [
      applyYrsCommand,
      canvasHostRef,
      displayListQueries,
      getPositionFromMouse,
      getPositionProjection,
      pagesContainerRef,
    ]
  );

  const handleImageDragStart = useCallback(() => {
    isImageInteractingRef.current = true;
  }, [isImageInteractingRef]);

  const handleImageDragEnd = useCallback(() => {
    isImageInteractingRef.current = false;
  }, [isImageInteractingRef]);

  return {
    handleImageResize,
    handleImageResizeStart,
    handleImageResizeEnd,
    handleImageDragMove,
    handleImageDragStart,
    handleImageDragEnd,
  };
}

function isFloatingImageAttributes(attrs: Record<string, unknown>): boolean {
  if (attrs.displayMode === 'float') return true;
  return (
    attrs.wrapType === 'square' ||
    attrs.wrapType === 'tight' ||
    attrs.wrapType === 'through' ||
    attrs.wrapType === 'topAndBottom' ||
    attrs.wrapType === 'behind' ||
    attrs.wrapType === 'inFront'
  );
}
