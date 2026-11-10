import { sessionUnavailableTables } from './sessionInternals';

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

/**
 * Refuses to save while `session` holds a table a concurrent undo left without its payload: the
 * saved file would lose that table. The session is unchanged.
 * @internal
 */
export function refuseUnavailableTables(session: object): void {
  if (sessionUnavailableTables(session).length > 0) throw tableUnavailableError();
}
