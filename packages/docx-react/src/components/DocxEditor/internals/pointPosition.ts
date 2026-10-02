import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import type { PointPosition } from '@betteroffice/docx/plugin-api';
import {
  createCanvasHostProjector,
  createRenderedDomContext,
} from '@betteroffice/docx/plugin-api/RenderedDomContext';
import type { ResidentEngineWorkerClient, YrsLoc, YrsSession } from '@betteroffice/docx/yrs';
import type { DocxPointPosition } from '../types';
import {
  isPresented,
  presentedWorkerVersion,
  readSessionVersion,
  sourceVersionOf,
} from './layoutProvenance';

/** What resolving a hit needs from the paged editor. */
export interface PointPositionEditor {
  getYrsSession(): YrsSession | null;
  displayPositionToYrsLoc(position: PointPosition): YrsLoc | null;
  /** True while typed or composed input has yet to reach the session. */
  hasPendingInput(): boolean;
}

/**
 * Resolves a hit from the layout `queries` answer for into a batch target, or null unless `host`
 * shows that layout's pixels, it lays out the session's current version and no input is still on
 * its way to the session: display positions shift with every edit. Rust projects the live
 * location into the accepted view, so text a pending deletion hides before the point does not
 * count.
 */
export function resolvePointPosition(
  editor: PointPositionEditor | null | undefined,
  hit: PointPosition | null,
  host: object | null | undefined,
  queries: DisplayListQueries | null | undefined
): DocxPointPosition | null {
  const session = editor?.getYrsSession() ?? null;
  if (!editor || !session || !hit || !queries || !isPresented(host, queries.displayList)) {
    return null;
  }
  const version = sourceVersionOf(queries);
  if (version === null || readSessionVersion(session) !== version) return null;
  if (editor.hasPendingInput()) return null;
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
    version,
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
  return resolvePointPosition(editor, dom.getPositionAtPoint(clientX, clientY), host, queries);
}

/** A hit on the presented worker frame and the worker version that frame lays out. */
interface ViewerHit {
  hit: PointPosition;
  version: string;
}

function viewerHitAtClientPoint(
  host: HTMLElement | null | undefined,
  queries: DisplayListQueries | null | undefined,
  zoom: number,
  clientX: number,
  clientY: number
): ViewerHit | null {
  if (!host || !queries || !isPresented(host, queries.displayList)) return null;
  const version = presentedWorkerVersion(queries);
  if (version === null) return null;
  const dom = createRenderedDomContext(host, zoom, {
    displayListQueries: queries,
    projector: createCanvasHostProjector(host, queries, zoom),
  });
  const hit = dom.getPositionAtPoint(clientX, clientY);
  return hit ? { hit, version } : null;
}

/**
 * The text under a client point in a viewer session: the hit comes from the presented frame,
 * its text range from the document worker. Null when the frame changed version meanwhile.
 */
export async function readViewerPositionAtClientPoint(
  read: ResidentEngineWorkerClient['documentRead'],
  host: HTMLElement | null | undefined,
  queries: DisplayListQueries | null | undefined,
  zoom: number,
  clientX: number,
  clientY: number,
  shown: () => DisplayListQueries | null | undefined
): Promise<DocxPointPosition | null> {
  const at = viewerHitAtClientPoint(host, queries, zoom, clientX, clientY);
  if (!at) return null;
  const reply = await read({ kind: 'pointPosition', hit: at.hit, expectVersion: at.version }).catch(
    () => null
  );
  if (!reply?.value || reply.version !== at.version) return null;
  return presentedWorkerVersion(shown()) === at.version ? reply.value : null;
}

const VIEWER_POINT_CACHE_LIMIT = 256;

/**
 * Synchronous point reads for a viewer session: a point not read yet for the presented version
 * answers null and starts its read, so the caller's retry finds it.
 */
export class ViewerPointPositions {
  private version: string | null = null;
  private readonly answers = new Map<string, DocxPointPosition | null | 'pending'>();

  positionAt(
    read: ResidentEngineWorkerClient['documentRead'],
    host: HTMLElement | null | undefined,
    queries: DisplayListQueries | null | undefined,
    zoom: number,
    clientX: number,
    clientY: number
  ): DocxPointPosition | null {
    const at = viewerHitAtClientPoint(host, queries, zoom, clientX, clientY);
    if (!at) return null;
    if (at.version !== this.version || this.answers.size > VIEWER_POINT_CACHE_LIMIT) {
      this.version = at.version;
      this.answers.clear();
    }
    const { hit } = at;
    const key = `${hit.region}:${hit.rId ?? ''}:${hit.noteId ?? ''}:${hit.pageIndex}:${hit.position}`;
    const known = this.answers.get(key);
    if (known !== undefined) return known === 'pending' ? null : known;
    this.answers.set(key, 'pending');
    const settle = (value: DocxPointPosition | null | undefined): void => {
      if (this.version !== at.version) return;
      if (value === undefined) this.answers.delete(key);
      else this.answers.set(key, value);
    };
    read({ kind: 'pointPosition', hit, expectVersion: at.version }).then(
      (reply) => settle(reply.version === at.version ? reply.value : undefined),
      () => settle(undefined)
    );
    return null;
  }
}
