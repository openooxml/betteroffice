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
import {
  buildMirrorPage,
  buildMirrorPageText,
  reduceMirrorToText,
  type DisplayPage,
} from '@betteroffice/docx/layout/render';
import type { TFunction } from '@betteroffice/docx-i18n';
import { useTranslation } from '../../i18n';
import { usePageChrome, type PageChromeHandle } from './usePageChrome';

const mirrorLabels = (page: DisplayPage, t: TFunction) => ({
  labels: {
    page: t('a11y.pageLabel', { number: page.pageIndex + 1 }),
    header: t('a11y.headerLabel'),
    footer: t('a11y.footerLabel'),
  },
});
const makeMirror = (page: DisplayPage, t: TFunction): HTMLElement =>
  buildMirrorPage(page, mirrorLabels(page, t));
const makeMirrorText = (
  page: DisplayPage,
  t: TFunction,
  mirror: HTMLElement | null
): HTMLElement =>
  mirror ? reduceMirrorToText(mirror) : buildMirrorPageText(page, mirrorLabels(page, t));

export function CanvasPageMirror({
  page,
  zoom = 1,
  active = true,
  defer = false,
  visible = true,
  register,
  noteAnchorRevision = 0,
}: {
  page: DisplayPage;
  /**
   * `displayPageNoteAnchorRevision(page)`: an owned shift moves the note
   * anchors the mirror renders without replacing the page.
   */
  noteAnchorRevision?: number;
  zoom?: number;
  /** Holds the full mirror; an inactive page keeps readable text. */
  active?: boolean;
  /** The first build may wait for idle time. */
  defer?: boolean;
  /** In the page window: a rebuild after a content change never waits. */
  visible?: boolean;
  /** Receives the handle that builds the mirror at once. */
  register?: (handle: PageChromeHandle | null) => void;
}) {
  const hostRef = useRef<HTMLDivElement>(null);
  const { t } = useTranslation();
  usePageChrome(hostRef, {
    page,
    t,
    active,
    defer,
    rebuildAtOnce: visible,
    urgentRevision: noteAnchorRevision,
    register,
    make: makeMirror,
    fallback: makeMirrorText,
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
