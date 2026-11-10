import type { YrsSession } from './index';
import { YrsPositionProjection } from './yrsPositionProjection';

export { isUnavailableTablePayload } from './yrsPositionProjection';

const TABLE_UNAVAILABLE =
  'a table is unavailable after a concurrent undo: redo on the collaborator who undid, or delete the table';

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
 * Whether a reader of `session` meets a table a concurrent undo left without its payload: the
 * editor's position projection walk from each root story (a story that is no table cell's or
 * block control's), through the cells of complete tables and block controls' stories. A story
 * the walk names but the session lacks reads as empty. Reads every reachable story.
 * @internal
 */
export function hasUnavailableTable(session: Pick<YrsSession, 'storyIds' | 'storySegments'>): boolean {
  const stories = session.storyIds();
  const live = new Set(stories);
  const source = { segments: (story: string) => (live.has(story) ? session.storySegments(story) : []) };
  return stories
    .filter((story) => !story.includes(':t') && !story.includes(':sdt'))
    .some((root) => new YrsPositionProjection(session, root, source).allTables().some((table) => table.unavailable));
}

/**
 * Refuses to write `session` while a reader can meet a table a concurrent undo left without its
 * payload: the written file would lose that table. Every call scans; the session is unchanged.
 * @internal
 */
export function refuseUnavailableTables(session: Pick<YrsSession, 'storyIds' | 'storySegments'>): void {
  if (hasUnavailableTable(session)) throw tableUnavailableError();
}
