/**
 * Mounts the interactive content-control overlay (core
 * `buildInteractiveOverlayPage`) 1:1 ABOVE one canvas page: same origin, same
 * page-local pixel space as the canvas and the a11y mirror. Unlike the mirror
 * (invisible, pointer-inert), this layer's buttons are visible, focusable and
 * clickable — they are the canvas path's `.layout-sdt-widget` /
 * `.layout-sdt-repeat-btn` triggers, so the existing delegated
 * ContentControlWidgets handlers pick their clicks up unchanged and route the
 * change through the active yrs content-control write path.
 *
 * The focus-steal guard lives INSIDE the core builder (a native mousedown
 * listener on the overlay root): a synthetic React handler here would fire at
 * the React root, after the canvas host's native pointer routing already moved
 * the caret. Rebuilt whenever the page's display list changes — the same
 * trigger that re-rasters the canvas.
 */

import { useEffect, useRef } from 'react';
import {
  buildInteractiveOverlayPage,
  displayPageRevision,
  type DisplayPage,
} from '@betteroffice/docx/layout/render';
import type { TFunction } from '@betteroffice/docx-i18n';
import { useTranslation } from '../../i18n';

export function CanvasInteractiveOverlay({
  page,
  zoom = 1,
  defer = false,
}: {
  page: DisplayPage;
  zoom?: number;
  /** See {@link CanvasPageMirror} — off-window pages build at idle time. */
  defer?: boolean;
}) {
  const hostRef = useRef<HTMLDivElement>(null);
  // Position-shift deltas mutate primitives in place — identity alone is stale.
  const builtForRef = useRef<{ page: DisplayPage; revision: number; t: TFunction } | null>(null);
  const { t } = useTranslation();

  useEffect(() => {
    const host = hostRef.current;
    if (!host) return;
    const built = builtForRef.current;
    if (built?.page === page && built.revision === displayPageRevision(page) && built.t === t) {
      return;
    }
    const build = (): void => {
      const overlay = buildInteractiveOverlayPage(page, {
        labels: {
          control: t('a11y.contentControl'),
          addRepeatingItem: t('a11y.addRepeatingItem'),
          removeRepeatingItem: t('a11y.removeRepeatingItem'),
        },
      });
      host.replaceChildren(overlay);
      builtForRef.current = { page, revision: displayPageRevision(page), t };
    };
    if (!defer) {
      build();
      return () => {
        host.replaceChildren();
        builtForRef.current = null;
      };
    }
    if (typeof requestIdleCallback === 'function') {
      const id = requestIdleCallback(build, { timeout: 1500 });
      return () => cancelIdleCallback(id);
    }
    const id = setTimeout(build, 150);
    return () => clearTimeout(id);
  }, [page, t, defer]);

  return (
    <div
      ref={hostRef}
      className="canvas-interactive-overlay"
      // Overlay content is built in page-local px; when the canvas is enlarged
      // for zoom (CSS size = page * zoom), CSS-scale the overlay by the same
      // factor from its top-left origin so the buttons stay on the painted
      // controls — identical to the mirror's transform.
      style={{
        position: 'absolute',
        left: 0,
        top: 0,
        pointerEvents: 'none',
        transform: zoom !== 1 ? `scale(${zoom})` : undefined,
        transformOrigin: '0 0',
      }}
    />
  );
}
