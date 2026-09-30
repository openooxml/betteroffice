import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import type { PointPosition } from '@betteroffice/docx/plugin-api';
import {
  createCanvasHostProjector,
  createRenderedDomContext,
} from '@betteroffice/docx/plugin-api/RenderedDomContext';
import type { YrsLoc, YrsSession } from '@betteroffice/docx/yrs';
import type { DocxPointPosition } from '../types';
import { workerProposalAuthority } from './workerProposalAuthority';
import { isPresented, readSessionVersion, sourceVersionOf } from './layoutProvenance';

/** What resolving a hit needs from the paged editor. */
export interface PointPositionEditor {
  getYrsSession(): YrsSession | null;
  displayPositionToYrsLoc(position: PointPosition): YrsLoc | null;
  /** True while typed or composed input has yet to reach the session. */
  hasPendingInput(): boolean;
}

function pointPositionVersion(
  editor: PointPositionEditor | null | undefined,
  hit: PointPosition | null,
  host: object | null | undefined,
  queries: DisplayListQueries | null | undefined,
  requireVersion = true
): string | null {
  const session = editor?.getYrsSession() ?? null;
  if (!editor || !session || !hit || !queries || !isPresented(host, queries.displayList)) return null;
  const version = sourceVersionOf(queries);
  return version !== null && (!requireVersion || readSessionVersion(session) === version) && !editor.hasPendingInput()
    ? version
    : null;
}

function resolvePointPositionLoc(
  editor: PointPositionEditor,
  session: YrsSession,
  hit: PointPosition,
  version: string
): DocxPointPosition | null {
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

/** Resolves a presented hit against the main session. */
export function resolvePointPosition(
  editor: PointPositionEditor | null | undefined,
  hit: PointPosition | null,
  host: object | null | undefined,
  queries: DisplayListQueries | null | undefined
): DocxPointPosition | null {
  const version = pointPositionVersion(editor, hit, host, queries);
  return version === null ? null : resolvePointPositionLoc(editor!, editor!.getYrsSession()!, hit!, version);
}

/** Resolves a presented hit through the current document authority. */
export async function readPointPosition(
  editor: PointPositionEditor | null | undefined,
  hit: PointPosition | null,
  host: object | null | undefined,
  queries: DisplayListQueries | null | undefined,
  current: () => boolean = () => true
): Promise<DocxPointPosition | null> {
  const session = editor?.getYrsSession();
  if (!session) return null;
  const authority = workerProposalAuthority(session);
  const version = pointPositionVersion(editor, hit, host, queries, !authority || authority.initialized);
  if (version === null || !current()) return null;
  const main = async (position: PointPosition, expectVersion: string) =>
    current() && editor!.getYrsSession() === session &&
      pointPositionVersion(editor, position, host, queries) === expectVersion
      ? resolvePointPositionLoc(editor!, session, position, expectVersion)
      : null;
  const result = authority ? await authority.pointPosition(hit!, version, main) : await main(hit!, version);
  if (!current() || editor!.getYrsSession() !== session ||
      pointPositionVersion(editor, hit, host, queries) !== result?.version) return null;
  return result;
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

/** Reads a client hit through the current document authority. */
export async function readPositionAtClientPoint(
  editor: (PointPositionEditor & { flushPendingInput(): Promise<void> }) | null | undefined,
  host: HTMLElement | null | undefined,
  queries: DisplayListQueries | null | undefined,
  zoom: number,
  clientX: number,
  clientY: number,
  current: () => boolean = () => true
): Promise<DocxPointPosition | null> {
  if (!editor || !host || !queries || !current()) return null;
  const session = editor.getYrsSession();
  if (editor.hasPendingInput()) {
    await editor.flushPendingInput();
    if (!current() || editor.getYrsSession() !== session) return null;
  }
  const dom = createRenderedDomContext(host, zoom, {
    displayListQueries: queries,
    projector: createCanvasHostProjector(host, queries, zoom),
  });
  return readPointPosition(editor, dom.getPositionAtPoint(clientX, clientY), host, queries, current);
}
