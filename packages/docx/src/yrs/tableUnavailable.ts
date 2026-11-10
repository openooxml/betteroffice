import type { YrsSession } from './index';

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

/** Whether a table embed's payload lacks its properties, grid or rows. @internal */
export function isUnavailableTablePayload(payload: Record<string, unknown>): boolean {
  return !('tblPr' in payload && 'grid' in payload && 'rows' in payload);
}

/** The coded refusal the engine throws for an unavailable table, built on this side. @internal */
export function tableUnavailableError(): Error & { code: string } {
  const error = new Error(TABLE_UNAVAILABLE) as Error & { code: string };
  error.name = 'TableUnavailableError';
  error.code = 'table-unavailable';
  return error;
}

/** Whether `id` names a table cell's story: a `t{n}` segment followed by an `r{n}c{n}` one. */
function isCellStory(id: string): boolean {
  return /(^|:)t\d+:r\d+c\d+(~\d+)?(:|$)/.test(id);
}

/** The cell stories a complete table payload's rows name, defaulted as the projection does. */
function namedCells(rows: unknown, story: string, tableIndex: number): string[] {
  if (!Array.isArray(rows)) return [];
  const named: string[] = [];
  rows.forEach((row: unknown, rowIndex) => {
    const cells = row && typeof row === 'object' ? (row as { cells?: unknown }).cells : undefined;
    if (!Array.isArray(cells)) return;
    cells.forEach((cell: unknown, cellIndex) => {
      if (!cell || typeof cell !== 'object') return;
      const id = (cell as { story?: unknown }).story;
      named.push(typeof id === 'string' ? id : `${story}:t${tableIndex}:r${rowIndex}c${cellIndex}`);
    });
  });
  return named;
}

/**
 * Whether a reader of `session` can reach a table a concurrent undo left without its payload:
 * one in a root story or in a cell of a complete table, at any depth. Reads every reachable story.
 * @internal
 */
export function hasUnavailableTable(session: Pick<YrsSession, 'storyIds' | 'storySegments'>): boolean {
  const live = new Set(session.storyIds());
  const queue = [...live].filter((id) => !isCellStory(id));
  const reached = new Set(queue);
  while (queue.length > 0) {
    const story = queue.pop()!;
    let tableIndex = 0;
    for (const segment of session.storySegments(story)) {
      if (segment.kind !== 'embed' || segment.embedKind !== 'table') continue;
      if (isUnavailableTablePayload(segment.payload)) return true;
      for (const cell of namedCells(segment.payload.rows, story, tableIndex)) {
        if (live.has(cell) && !reached.has(cell)) {
          reached.add(cell);
          queue.push(cell);
        }
      }
      tableIndex += 1;
    }
  }
  return false;
}

/**
 * Refuses to save while a reader of `session` can reach a table a concurrent undo left without
 * its payload: the saved file would lose that table. Every save scans; the session is unchanged.
 * @internal
 */
export function refuseUnavailableTables(session: Pick<YrsSession, 'storyIds' | 'storySegments'>): void {
  if (hasUnavailableTable(session)) throw tableUnavailableError();
}
