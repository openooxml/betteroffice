import type { Document } from '../types/document';
import type { YrsSession } from './index';
import { headerFooterStory } from './sessionInternals';
import { YrsPositionProjection } from './yrsPositionProjection';

export { isUnavailableTablePayload } from './yrsPositionProjection';

const TABLE_UNAVAILABLE =
  'a table is unavailable after a concurrent undo: redo on the collaborator who undid, or delete the table';

type TableScanSession = Pick<YrsSession, 'storyIds' | 'storySegments' | 'listComments'>;

/**
 * Whether `error` is the engine's refusal of a table operation or save over a table that a
 * concurrent undo left without its properties, grid or rows.
 * @internal
 */
export function isTableUnavailableRefusal(error: unknown): boolean {
  return error instanceof Error && 'code' in error && error.code === 'table-unavailable';
}

/** The coded refusal the engine throws for an unavailable table, built on this side. @internal */
export function tableUnavailableError(): Error & { code: string } {
  const error = new Error(TABLE_UNAVAILABLE) as Error & { code: string };
  error.name = 'TableUnavailableError';
  error.code = 'table-unavailable';
  return error;
}

/**
 * The root stories a save of `document` writes, as the save projection lists them: the body,
 * each header and footer by its relationship, each footnote and endnote, and each comment.
 */
function savedRootStories(session: TableScanSession, document: Document): string[] {
  const { package: pkg } = document;
  return [
    'body',
    ...[...(pkg.headers?.keys() ?? []), ...(pkg.footers?.keys() ?? [])].map((rId) => headerFooterStory(session, rId)),
    ...(pkg.footnotes ?? []).map((note) => `fn:${note.id}`),
    ...(pkg.endnotes ?? []).map((note) => `en:${note.id}`),
    ...session.listComments().map((comment) => `comment:${comment.id}`),
  ];
}

/**
 * Whether a save of `document` from `session` would meet a table a concurrent undo left without
 * its payload: the editor's position projection walk from each root story the save writes,
 * through the cells of complete tables and block controls' stories. A story the walk names but
 * the session lacks reads as empty. Reads every reachable story.
 * @internal
 */
export function hasUnavailableTable(session: TableScanSession, document: Document): boolean {
  const live = new Set(session.storyIds());
  const source = { segments: (story: string) => (live.has(story) ? session.storySegments(story) : []) };
  return [...new Set(savedRootStories(session, document))]
    .filter((root) => live.has(root))
    .some((root) => new YrsPositionProjection(session, root, source).allTables().some((table) => table.unavailable));
}

/**
 * Refuses to make or write `document` from `session` while a save would meet a table a
 * concurrent undo left without its payload: the written file would lose that table. Every call
 * scans; the session is unchanged.
 * @internal
 */
export function refuseUnavailableTables(session: TableScanSession, document: Document): void {
  if (hasUnavailableTable(session, document)) throw tableUnavailableError();
}
