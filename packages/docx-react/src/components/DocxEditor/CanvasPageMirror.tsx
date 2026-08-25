/**
 * Mounts the accessibility mirror (core `buildMirrorPage`) 1:1 under one
 * canvas page: same origin, same page-local pixel space, so
 * `getBoundingClientRect` on mirror nodes returns the rects the canvas
 * painted. The mirror is invisible (opacity 0) and inert to the pointer
 * (pointer-events none) but deliberately NOT aria-hidden — it is the
 * accessible content of the canvas. Rebuilt, debounced, when the page's
 * display list changes; the canvas re-raster is never held back by it.
 *
 * Focus never lands here: the hidden input remains the editing surface.
 */

import { useRef } from 'react';
import { buildMirrorPage, type DisplayPage } from '@betteroffice/docx/layout/render';
import { useTranslation } from '../../i18n';
import { useCoalescedSubtreeMount } from './hooks/useCoalescedSubtreeMount';

export function CanvasPageMirror({ page, zoom = 1 }: { page: DisplayPage; zoom?: number }) {
  const hostRef = useRef<HTMLDivElement>(null);
  const { t } = useTranslation();

  useCoalescedSubtreeMount(
    hostRef,
    () =>
      buildMirrorPage(page, {
        labels: {
          page: t('a11y.pageLabel', { number: page.pageIndex + 1 }),
          header: t('a11y.headerLabel'),
          footer: t('a11y.footerLabel'),
        },
      }),
    [page, t]
  );

  return (
    <div
      ref={hostRef}
      className="canvas-page-mirror"
      // The mirror content is built in page-local px; when the canvas is
      // enlarged for zoom (CSS size = page * zoom), CSS-scale the mirror by the
      // same factor from its top-left origin so its nodes' `getBoundingClientRect`
      // still lands on the painted glyphs. At zoom = 1 this is an identity scale.
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
