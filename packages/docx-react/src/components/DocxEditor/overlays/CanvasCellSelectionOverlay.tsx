import { useLayoutEffect, useMemo, useState } from 'react';
import { createPortal } from 'react-dom';
import type { YrsSession } from '@betteroffice/docx/yrs';
import {
  deriveDisplayListSelectedCellRects,
  type DisplayListQueries,
  type DisplayListSelectedCellSpec,
  type DisplayListTableRegion,
} from '@betteroffice/docx/layout/render';
import { projectPageLocalRect } from '../internals/canvasProjection';
import type { YrsPositionProjection } from '../internals/yrsPositionProjection';
import { yrsCellLocFromStory } from '../yrsCommands';

export interface CanvasCellSelectionOverlayProps {
  session: YrsSession | null;
  positionProjection: YrsPositionProjection | null;
  overlayTarget: HTMLElement;
  canvasHostRef: React.RefObject<HTMLDivElement | null>;
  displayListQueries: DisplayListQueries;
  region?: DisplayListTableRegion;
  sidebarOpen: boolean;
  zoom: number;
}

interface ProjectedCell {
  key: string;
  left: number;
  top: number;
  width: number;
  height: number;
}

const BODY_REGION: DisplayListTableRegion = { kind: 'body' };

export function CanvasCellSelectionOverlay({
  session,
  positionProjection,
  overlayTarget,
  canvasHostRef,
  displayListQueries,
  region = BODY_REGION,
  sidebarOpen,
  zoom,
}: CanvasCellSelectionOverlayProps): React.ReactPortal | null {
  const specs = useMemo(
    () => selectedCellSpecs(session, positionProjection),
    [session, positionProjection, displayListQueries]
  );
  const [projected, setProjected] = useState<ProjectedCell[]>([]);

  useLayoutEffect(() => {
    const host = canvasHostRef.current;
    if (!positionProjection || !host || specs.length === 0) {
      setProjected([]);
      return;
    }
    const recompute = () => {
      const next: ProjectedCell[] = [];
      const rects = deriveDisplayListSelectedCellRects(
        displayListQueries.displayList,
        specs,
        createTableKeyResolver(positionProjection),
        region
      );
      for (const rect of rects) {
        const value = projectPageLocalRect(
          host,
          overlayTarget,
          displayListQueries,
          rect.pageIndex,
          rect.x,
          rect.y,
          rect.w,
          rect.h
        );
        if (!value) continue;
        next.push({
          key: `${rect.region}-${rect.rId ?? 'body'}-${rect.pageIndex}-${rect.tableKey}-${rect.row}-${rect.col}`,
          left: value.left,
          top: value.top,
          width: value.width,
          height: value.height,
        });
      }
      setProjected(next);
    };
    recompute();
    const observer = new ResizeObserver(recompute);
    observer.observe(host);
    observer.observe(overlayTarget);
    window.addEventListener('resize', recompute);
    host.addEventListener('transitionend', recompute);
    return () => {
      observer.disconnect();
      window.removeEventListener('resize', recompute);
      host.removeEventListener('transitionend', recompute);
    };
  }, [positionProjection, specs, overlayTarget, canvasHostRef, displayListQueries, region, sidebarOpen, zoom]);

  if (projected.length === 0) return null;
  return createPortal(
    <div
      aria-hidden="true"
      className="canvas-cell-selection-overlay"
      style={{ position: 'absolute', inset: 0, pointerEvents: 'none', zIndex: 10, overflow: 'visible' }}
    >
      {projected.map((cell) => (
        <div
          key={cell.key}
          style={{
            position: 'absolute',
            left: cell.left,
            top: cell.top,
            width: cell.width,
            height: cell.height,
            background: 'var(--doc-selection)',
            outline: '2px solid var(--doc-focus-ring)',
            outlineOffset: -2,
            pointerEvents: 'none',
            boxSizing: 'border-box',
          }}
        />
      ))}
    </div>,
    overlayTarget
  );
}

function selectedCellSpecs(
  session: YrsSession | null,
  projection: YrsPositionProjection | null
): DisplayListSelectedCellSpec[] {
  let selection;
  try {
    const caret = session?.selection();
    if (!caret || !yrsCellLocFromStory(caret.head.story)) return [];
    selection = session?.cellSelection();
  } catch {
    // Structural undo/remote deletion can invalidate the ephemeral rectangle
    // before the next input event replaces it.
    return [];
  }
  if (!selection || !projection) return [];
  const tableKey = projection.tableStartForLoc(selection.anchor);
  if (tableKey == null || projection.tableStartForLoc(selection.head) !== tableKey) return [];
  const top = Math.min(selection.anchor.row, selection.head.row);
  const bottom = Math.max(selection.anchor.row, selection.head.row);
  const left = Math.min(selection.anchor.column, selection.head.column);
  const right = Math.max(selection.anchor.column, selection.head.column);
  const specs: DisplayListSelectedCellSpec[] = [];
  for (let row = top; row <= bottom; row += 1) {
    for (let col = left; col <= right; col += 1) {
      specs.push({ tableKey: String(tableKey), row, col, rowSpan: 1, colSpan: 1 });
    }
  }
  return specs;
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

export default CanvasCellSelectionOverlay;
