import type { CSSProperties, PointerEvent as ReactPointerEvent } from 'react';
import type { TranslationKey } from '@betteroffice/xlsx-i18n';
import { useTranslation } from '../i18n';
import { CHART_RESIZE_HANDLES, type ChartResizeHandle } from './chartResize';

/** A chart frame already scaled into the overlay's CSS-pixel space. */
export interface ScaledChartRect {
  x: number;
  y: number;
  w: number;
  h: number;
}

const HANDLE_SIZE = 8;

const HANDLE_LABELS: Record<ChartResizeHandle, TranslationKey> = {
  nw: 'a11y.chartResizeNw',
  n: 'a11y.chartResizeN',
  ne: 'a11y.chartResizeNe',
  e: 'a11y.chartResizeE',
  se: 'a11y.chartResizeSe',
  s: 'a11y.chartResizeS',
  sw: 'a11y.chartResizeSw',
  w: 'a11y.chartResizeW',
};

/** Which arrow keys move an edge handle; corners accept both axes. */
function handleAxes(handle: ChartResizeHandle): { horizontal: boolean; vertical: boolean } {
  const horizontal = handle === 'nw' || handle === 'w' || handle === 'sw' ||
    handle === 'ne' || handle === 'e' || handle === 'se';
  const vertical = handle === 'nw' || handle === 'n' || handle === 'ne' ||
    handle === 'sw' || handle === 's' || handle === 'se';
  return { horizontal, vertical };
}

const handleStyle: CSSProperties = {
  position: 'absolute',
  width: HANDLE_SIZE,
  height: HANDLE_SIZE,
  marginLeft: -HANDLE_SIZE / 2,
  marginTop: -HANDLE_SIZE / 2,
  boxSizing: 'border-box',
  border: '1px solid #217346',
  background: '#ffffff',
  pointerEvents: 'auto',
  touchAction: 'none',
};

/**
 * The eight resize handles around a selected chart. The overlay host is
 * `pointer-events: none`, so each handle re-enables hit testing for itself; the
 * shared gesture hook owns the drag from pointerdown to release. A mousedown is
 * stopped here so the press can never fall through to the grid's cell selection
 * handler in browsers that still synthesize one. Each handle is a focusable
 * button: arrow keys move its edge through the same commit path, with Shift
 * taking a ten-pixel step.
 */
export function ChartResizeHandles({ rect, onStart, onNudge }: {
  rect: ScaledChartRect;
  onStart: (handle: ChartResizeHandle, event: ReactPointerEvent<HTMLElement>) => void;
  onNudge: (handle: ChartResizeHandle, dx: number, dy: number) => void;
}) {
  const { t } = useTranslation();
  return <>
    {CHART_RESIZE_HANDLES.map((spec) => {
      const axes = handleAxes(spec.handle);
      const shortcuts = [axes.horizontal ? 'ArrowLeft ArrowRight' : null,
        axes.vertical ? 'ArrowUp ArrowDown' : null].filter(Boolean).join(' ');
      const hint: TranslationKey = axes.horizontal && axes.vertical ? 'a11y.chartResizeHintBoth'
        : axes.horizontal ? 'a11y.chartResizeHintHorizontal' : 'a11y.chartResizeHintVertical';
      const label = t(HANDLE_LABELS[spec.handle]);
      return (
        <button key={spec.handle} type="button"
          data-testid="xlsx-chart-resize-handle" data-handle={spec.handle}
          aria-label={label} aria-keyshortcuts={shortcuts}
          title={t(hint, { label })}
          onPointerDown={(event) => onStart(spec.handle, event)}
          onMouseDown={(event) => event.stopPropagation()}
          onKeyDown={(event) => {
            if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight' &&
              event.key !== 'ArrowUp' && event.key !== 'ArrowDown') return;
            // the selected chart treats arrows as a move; every arrow pressed
            // on a focused handle belongs to the handle, even when this edge
            // does not run along that axis (it is then a no-op).
            event.preventDefault();
            event.stopPropagation();
            const step = event.shiftKey ? 10 : 1;
            const dx = (event.key === 'ArrowRight' ? step : event.key === 'ArrowLeft' ? -step : 0) *
              (axes.horizontal ? 1 : 0);
            const dy = (event.key === 'ArrowDown' ? step : event.key === 'ArrowUp' ? -step : 0) *
              (axes.vertical ? 1 : 0);
            if (dx === 0 && dy === 0) return;
            onNudge(spec.handle, dx, dy);
          }}
          style={{ ...handleStyle, left: rect.x + rect.w * spec.x, top: rect.y + rect.h * spec.y,
            cursor: spec.cursor, padding: 0 }} />
      );
    })}
  </>;
}
