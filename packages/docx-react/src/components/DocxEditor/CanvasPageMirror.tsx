/**
 * Mounts the accessibility mirror (core `buildMirrorPage`) 1:1 under one
 * canvas page: same origin, same page-local pixel space, so
 * `getBoundingClientRect` on mirror nodes returns the rects the canvas
 * painted. The mirror is invisible (opacity 0) and inert to the pointer
 * (pointer-events none) but deliberately NOT aria-hidden — it is the
 * accessible content of the canvas. Rebuilt whenever the page's display list
 * changes — the same trigger that re-rasters the canvas.
 *
 * Focus never lands here: the hidden input remains the editing surface.
 */

import { useRef } from 'react';
import { buildMirrorPage, type DisplayPage } from '@betteroffice/docx/layout/render';
import type { TFunction } from '@betteroffice/docx-i18n';
import { useTranslation } from '../../i18n';
import { usePageChrome } from './usePageChrome';

const makeMirror = (page: DisplayPage, t: TFunction): HTMLElement =>
  buildMirrorPage(page, {
    labels: {
      page: t('a11y.pageLabel', { number: page.pageIndex + 1 }),
      header: t('a11y.headerLabel'),
      footer: t('a11y.footerLabel'),
    },
  });

export function CanvasPageMirror({
  page,
  zoom = 1,
  active = true,
  defer = false,
  visible = true,
  registerBuild,
}: {
  page: DisplayPage;
  zoom?: number;
  /** Holds the mirror; an inactive page keeps only its empty host. */
  active?: boolean;
  /** The first build may wait for idle time. */
  defer?: boolean;
  /** The page is in the page window, so a rebuild never waits. */
  visible?: boolean;
  /** Receives a function that builds the mirror at once. */
  registerBuild?: (build: (() => void) | null) => void;
}) {
  const hostRef = useRef<HTMLDivElement>(null);
  const { t } = useTranslation();
  usePageChrome(hostRef, {
    page,
    t,
    active,
    defer,
    rebuildAtOnce: visible,
    registerBuild,
    make: makeMirror,
  });

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
