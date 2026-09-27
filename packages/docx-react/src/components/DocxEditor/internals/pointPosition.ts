import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import type { PointPosition } from '@betteroffice/docx/plugin-api';
import {
  createCanvasHostProjector,
  createRenderedDomContext,
} from '@betteroffice/docx/plugin-api/RenderedDomContext';
import type { YrsLoc, YrsSession } from '@betteroffice/docx/yrs';
import type { DocxPointPosition } from '../types';
import { readSessionVersion, sourceVersionOf } from './layoutProvenance';

/** What resolving a hit needs from the paged editor. */
export interface PointPositionEditor {
  getYrsSession(): YrsSession | null;
  displayPositionToYrsLoc(position: PointPosition): YrsLoc | null;
}

/**
 * Resolves a hit from the layout of `layoutVersion` into a batch target, or null unless that
 * layout shows the session's current version: display positions shift with every edit. Rust
 * projects the live location into the accepted view, so text a pending deletion hides before
 * the point does not count.
 */
export function resolvePointPosition(
  editor: PointPositionEditor | null | undefined,
  hit: PointPosition | null,
  layoutVersion: string | null
): DocxPointPosition | null {
  const session = editor?.getYrsSession() ?? null;
  if (!editor || !session || !hit || layoutVersion === null) return null;
  if (readSessionVersion(session) !== layoutVersion) return null;
  const loc = editor.displayPositionToYrsLoc(hit);
  if (!loc) return null;
  let offset: number;
  try {
    offset = session.selectionText({ story: loc.story, start: loc, end: loc }).before.length;
  } catch {
    return null;
  }
  return {
    ...hit,
    version: layoutVersion,
    target: {
      kind: 'range',
      story: loc.story,
      start: { paraId: loc.paraId, offset },
      end: { paraId: loc.paraId, offset },
      view: 'accepted',
    },
  };
}

/** The text under a client point on the canvas pages `host` paints from `queries`. */
export function positionAtClientPoint(
  editor: PointPositionEditor | null | undefined,
  host: HTMLElement | null | undefined,
  queries: DisplayListQueries | null | undefined,
  zoom: number,
  clientX: number,
  clientY: number
): DocxPointPosition | null {
  if (!host || !queries) return null;
  const dom = createRenderedDomContext(host, zoom, {
    displayListQueries: queries,
    projector: createCanvasHostProjector(host, queries, zoom),
  });
  return resolvePointPosition(
    editor,
    dom.getPositionAtPoint(clientX, clientY),
    sourceVersionOf(queries)
  );
}
