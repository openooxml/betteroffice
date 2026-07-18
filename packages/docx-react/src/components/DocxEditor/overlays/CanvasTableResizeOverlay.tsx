/**
 * Canvas-mode table resize handles.
 *
 * This component sources handle geometry from the display list and portals
 * interactive handle strips onto the
 * visible canvas pages (portal target = `editorContentRef`).
 *
 * Geometry SOURCE only moves — the commit logic is the shared core
 * `tableResize` helpers (`commitColumnResize` / `commitRowResize` /
 * `commitRightEdgeResize`), byte-for-byte the same transactions the DOM
 * `useTableResizeState` dispatches. Per table fragment:
 *  - `deriveDisplayListTableFragments` gives the fragment's page-local origin,
 *    bottom, and per-row content bands;
 *  - its `docStartSample` resolves against the PM doc to the enclosing table
 *    node → `pmStart` + the twips `columnWidths`;
 *  - column boundaries = `originX` + cumulative column widths (exact + robust to
 *    colspan, unlike counting interior border lines); row boundaries sit in the
 *    gaps between content bands (robust to vmerge).
 *
 * Handles reuse the painter's class names so the shared hover / dragging CSS
 * (transparent → blue) applies with no new tokens. Body tables only — HF table
 * resize handles remain a documented canvas gap (their doc positions live in a
 * separate PM doc, so `deriveDisplayListTableFragments` + the commit path would
 * need to bind to the active header/footer story and its region geometry; see the note
 * in `displayListTables.ts`). The HF *caret + selection* on canvas are covered
 * by `CanvasHfSelectionOverlay`.
 *
 * @experimental part of the rust-canvas-engine change; shape may evolve.
 */

import React, { useLayoutEffect, useMemo, useState } from 'react';
import { createPortal } from 'react-dom';
import {
  deriveDisplayListTableFragments,
  type DisplayListQueries,
} from '@betteroffice/docx/layout/render';
import { projectPageLocalRect } from '../internals/canvasProjection';
import type { YrsEditorCommand } from '../yrsCommands';
import type { YrsPositionProjection } from '../internals/yrsPositionProjection';

const HANDLE_THICKNESS = 6;
const HANDLE_HALF = HANDLE_THICKNESS / 2;
const TWIPS_PER_PIXEL = 15;
const MIN_CELL_WIDTH_TWIPS = 300;

/** page-local (px) geometry + commit params for one handle, before projection. */
interface HandleSpec {
  id: string;
  type: 'col' | 'right';
  pageIndex: number;
  x: number;
  y: number;
  w: number;
  h: number;
  pmStart: number;
  // column handles:
  colIdx?: number;
  leftTwips?: number;
  rightTwips?: number;
  widthTwips?: number;
}

interface ProjectedHandle {
  spec: HandleSpec;
  left: number;
  top: number;
  width: number;
  height: number;
  scaleX: number;
}

const HANDLE_CLASS: Record<HandleSpec['type'], string> = {
  col: 'layout-table-resize-handle',
  right: 'layout-table-edge-handle-right',
};

export interface CanvasTableResizeOverlayProps {
  overlayTarget: HTMLElement;
  canvasHostRef: React.RefObject<HTMLDivElement | null>;
  displayListQueries: DisplayListQueries;
  positionProjection: YrsPositionProjection | null;
  applyYrsCommand: (command: YrsEditorCommand) => boolean;
  readOnly?: boolean;
  sidebarOpen: boolean;
  zoom: number;
}

export function CanvasTableResizeOverlay({
  overlayTarget,
  canvasHostRef,
  displayListQueries,
  positionProjection,
  applyYrsCommand,
  readOnly = false,
  sidebarOpen,
  zoom,
}: CanvasTableResizeOverlayProps): React.ReactPortal | null {
  // Page-local handle specs + commit params, rebuilt whenever the display list
  // rebuilds (its identity changes per build, including after a resize commit).
  const specs = useMemo<HandleSpec[]>(() => {
    if (readOnly) return [];
    if (!positionProjection) return [];

    // Cell content primitives carry each cell-paragraph's block id, so the core
    // grouper needs a table identity: resolve an in-cell doc position to the
    // enclosing table node's `before` position (a stable per-table key), cached
    // per doc position. The same map memoizes the table node (widths, rowCount).
    const keyCache = new Map<number, string | null>();
    const tableAt = new Map<string, { pmStart: number; widthsTwips: number[]; rowCount: number }>();
    const tableKeyOf = (ds: number | undefined): string | null => {
      if (ds == null) return null;
      const cached = keyCache.get(ds);
      if (cached !== undefined) return cached;
      const table = positionProjection.tableAtPosition(ds);
      const key = table ? String(table.start) : null;
      if (table && key && !tableAt.has(key)) {
        tableAt.set(key, {
          pmStart: table.start,
          widthsTwips: table.widthsTwips,
          rowCount: table.rowCount,
        });
      }
      keyCache.set(ds, key);
      return key;
    };

    const out: HandleSpec[] = [];
    for (const frag of deriveDisplayListTableFragments(
      displayListQueries.displayList,
      tableKeyOf
    )) {
      const table = tableAt.get(frag.tableKey);
      if (!table || table.widthsTwips.length === 0) continue;
      const { pmStart, widthsTwips, rowCount } = table;

      const widthsPx = widthsTwips.map((w) => w / TWIPS_PER_PIXEL);
      const top = frag.originY;
      const height = Math.max(0, frag.bottomY - frag.originY);

      // Column-between handles at each interior boundary.
      let cum = 0;
      for (let col = 0; col < widthsPx.length - 1; col++) {
        cum += widthsPx[col];
        out.push({
          id: `${frag.tableKey}-p${frag.pageIndex}-col${col}`,
          type: 'col',
          pageIndex: frag.pageIndex,
          x: frag.originX + cum - HANDLE_HALF,
          y: top,
          w: HANDLE_THICKNESS,
          h: height,
          pmStart,
          colIdx: col,
          leftTwips: widthsTwips[col],
          rightTwips: widthsTwips[col + 1],
        });
      }

      const bands = frag.rowBands;

      // Edge handles only on the fragment that ends the table.
      const lastRowIndex = rowCount - 1;
      const endsTable = bands.length > 0 && bands[bands.length - 1].rowIndex === lastRowIndex;
      if (endsTable) {
        out.push({
          id: `${frag.tableKey}-p${frag.pageIndex}-right`,
          type: 'right',
          pageIndex: frag.pageIndex,
          x: frag.originX + cumWidth(widthsPx) - HANDLE_HALF,
          y: top,
          w: HANDLE_THICKNESS,
          h: height,
          pmStart,
          colIdx: widthsPx.length - 1,
          widthTwips: widthsTwips[widthsTwips.length - 1],
        });
      }
    }
    return out;
  }, [displayListQueries, positionProjection, readOnly]);

  // Project page-local specs onto the visible canvas via the live canvas rects.
  const [projected, setProjected] = useState<ProjectedHandle[]>([]);
  useLayoutEffect(() => {
    const host = canvasHostRef.current;
    if (!host) {
      setProjected([]);
      return;
    }
    const recompute = () => {
      const next: ProjectedHandle[] = [];
      for (const spec of specs) {
        const p = projectPageLocalRect(
          host,
          overlayTarget,
          displayListQueries,
          spec.pageIndex,
          spec.x,
          spec.y,
          spec.w,
          spec.h
        );
        if (!p) continue;
        next.push({
          spec,
          left: p.left,
          top: p.top,
          width: p.width,
          height: p.height,
          scaleX: p.scaleX,
        });
      }
      setProjected(next);
    };
    recompute();
    const ro = new ResizeObserver(recompute);
    ro.observe(host);
    ro.observe(overlayTarget);
    window.addEventListener('resize', recompute);
    host.addEventListener('transitionend', recompute);
    return () => {
      ro.disconnect();
      window.removeEventListener('resize', recompute);
      host.removeEventListener('transitionend', recompute);
    };
  }, [specs, canvasHostRef, overlayTarget, displayListQueries, sidebarOpen, zoom]);

  const startDrag = (ph: ProjectedHandle, e: React.MouseEvent) => {
    if (readOnly) return;
    e.preventDefault();
    e.stopPropagation();
    const { spec, scaleX } = ph;
    const handleEl = e.currentTarget as HTMLElement;
    handleEl.classList.add('dragging');
    const startClientX = e.clientX;
    const origLeft = parseFloat(handleEl.style.left) || 0;

    // Accumulated, min-clamped commit values (twips).
    let left = spec.leftTwips ?? 0;
    let right = spec.rightTwips ?? 0;
    let width = spec.widthTwips ?? 0;

    const onMove = (me: MouseEvent) => {
      const deltaPx = scaleX > 0 ? (me.clientX - startClientX) / scaleX : 0;
      const deltaTwips = Math.round(deltaPx * TWIPS_PER_PIXEL);
      if (spec.type === 'col') {
        const nl = (spec.leftTwips ?? 0) + deltaTwips;
        const nr = (spec.rightTwips ?? 0) - deltaTwips;
        if (nl >= MIN_CELL_WIDTH_TWIPS && nr >= MIN_CELL_WIDTH_TWIPS) {
          left = nl;
          right = nr;
        }
      } else {
        const nextWidth = (spec.widthTwips ?? 0) + deltaTwips;
        if (nextWidth >= MIN_CELL_WIDTH_TWIPS) width = nextWidth;
      }
      handleEl.style.left = `${origLeft + (me.clientX - startClientX)}px`;
    };
    const onUp = () => {
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
      handleEl.classList.remove('dragging');
      if (spec.type === 'col') {
        applyYrsCommand({
          type: 'tableColumnWidths',
          pmStart: spec.pmStart,
          widths: [
            { column: spec.colIdx!, widthTwips: left },
            { column: spec.colIdx! + 1, widthTwips: right },
          ],
        });
      } else {
        applyYrsCommand({
          type: 'tableColumnWidths',
          pmStart: spec.pmStart,
          widths: [{ column: spec.colIdx!, widthTwips: width }],
        });
      }
    };
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
  };

  if (readOnly || projected.length === 0) return null;

  return createPortal(
    <div
      className="canvas-table-resize-overlay"
      style={{
        position: 'absolute',
        inset: 0,
        pointerEvents: 'none',
        zIndex: 14,
        overflow: 'visible',
      }}
    >
      {projected.map((ph) => (
        <div
          key={ph.spec.id}
          className={HANDLE_CLASS[ph.spec.type]}
          data-handle-type={ph.spec.type}
          data-table-doc-start={ph.spec.pmStart}
          data-column-index={ph.spec.colIdx}
          onMouseDown={(e) => startDrag(ph, e)}
          style={{
            position: 'absolute',
            left: ph.left,
            top: ph.top,
            width: ph.width,
            height: ph.height,
            cursor: 'col-resize',
            pointerEvents: 'auto',
            zIndex: 15,
          }}
        />
      ))}
    </div>,
    overlayTarget
  );
}

function cumWidth(widthsPx: number[]): number {
  let s = 0;
  for (const w of widthsPx) s += w;
  return s;
}

export default CanvasTableResizeOverlay;
