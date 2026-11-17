import type { Document } from '../types/document';
import type { YrsSession, YrsTableLoc } from './index';
import { headerFooterStory, sessionUnavailableTablesFrom } from './sessionInternals';
import { YrsPositionProjection } from './yrsPositionProjection';

export { isUnavailableTablePayload } from './yrsPositionProjection';

const TABLE_UNAVAILABLE =
  'a table is unavailable after a concurrent undo: redo on the collaborator who undid, or delete the table';

type TableScanSession = Pick<YrsSession, 'storyIds' | 'storySegments' | 'listComments' | 'version'>;

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
 * @internal
 */
export function savedRootStories(session: Pick<YrsSession, 'listComments'>, document: Document): string[] {
  const { package: pkg } = document;
  return [...new Set([
    'body',
    ...[...(pkg.headers?.keys() ?? []), ...(pkg.footers?.keys() ?? [])].map((rId) => headerFooterStory(session, rId)),
    ...(pkg.footnotes ?? []).map((note) => `fn:${note.id}`),
    ...(pkg.endnotes ?? []).map((note) => `en:${note.id}`),
    ...session.listComments().map((comment) => `comment:${comment.id}`),
  ])];
}

function sorted(tables: YrsTableLoc[]): YrsTableLoc[] {
  const seen = new Set<string>();
  return tables
    .filter((table) => {
      const key = `${table.story}\u0000${table.tableIndex}`;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    })
    .sort((left, right) => (left.story < right.story ? -1 : left.story > right.story ? 1 : left.tableIndex - right.tableIndex));
}

/**
 * The unavailable tables the editor's position projection walk meets from `roots`, through
 * the cells of complete tables and block controls' stories; a story the walk names but the
 * session lacks reads as empty. The fallback for a session without its engine.
 * @internal
 */
export function unavailableTablesByProjection(
  session: Pick<YrsSession, 'storyIds' | 'storySegments'>,
  roots: readonly string[]
): YrsTableLoc[] {
  const live = new Set(session.storyIds());
  const source = { segments: (story: string) => (live.has(story) ? session.storySegments(story) : []) };
  return sorted(roots
    .filter((root) => live.has(root))
    .flatMap((root) => new YrsPositionProjection(session, root, source).allTables()
      .filter((table) => table.unavailable)
      .map(({ story, tableIndex }) => ({ story, tableIndex }))));
}

/**
 * The unavailable tables a walk from `roots` meets: the engine's walk, or the projection walk
 * when the session has no engine or the engine walk could not look everywhere (a depth or
 * position limit). Throws the coded refusal when neither walk can finish: a save that cannot
 * look must not read that as nothing there.
 */
function scanFrom(session: TableScanSession, roots: readonly string[]): YrsTableLoc[] {
  const found = sessionUnavailableTablesFrom(session, roots);
  if (found) return found;
  try {
    return unavailableTablesByProjection(session, roots);
  } catch {
    throw tableUnavailableError();
  }
}

/**
 * The unavailable tables a save of `document` would meet, by [`scanFrom`] its saved root stories.
 * @internal
 */
export function unavailableTables(session: TableScanSession, document: Document): YrsTableLoc[] {
  return scanFrom(session, savedRootStories(session, document));
}

/** A gate pass: the session version and the root stories the scan covered. */
interface GateMark {
  version: string;
  roots: ReadonlySet<string>;
}

/** Documents made from a session after a scan, with what that scan covered. */
const gated = new WeakMap<Document, GateMark>();

/**
 * Refuses to make or write `document` from `session` while a save would meet a table a
 * concurrent undo left without its payload: the written file would lose that table. A document
 * this gate already passed at the session's current version is not scanned again, unless it
 * now lists a root story that scan did not cover. A pass marks `document`, so the writer it is
 * handed to next does not scan it again. Returns the root stories it covered. The session is
 * unchanged.
 * @internal
 */
export function refuseUnavailableTables(session: TableScanSession, document: Document): string[] {
  const roots = savedRootStories(session, document);
  const version = session.version();
  const mark = gated.get(document);
  if (mark?.version === version && roots.every((root) => mark.roots.has(root))) return roots;
  if (scanFrom(session, roots).length > 0) throw tableUnavailableError();
  gated.set(document, { version, roots: new Set(roots) });
  return roots;
}

/**
 * Records that `document`, made from `session` just now, passed the gate at its version with a
 * scan of `roots`.
 * @internal
 */
export function markGated(session: Pick<YrsSession, 'version'>, document: Document, roots: readonly string[]): void {
  gated.set(document, { version: session.version(), roots: new Set(roots) });
}

/** Carries `from`'s gate record to `to`, a copy of it made without reading the session. @internal */
export function carryGate(from: Document, to: Document): void {
  const mark = gated.get(from);
  if (mark !== undefined) gated.set(to, mark);
}
