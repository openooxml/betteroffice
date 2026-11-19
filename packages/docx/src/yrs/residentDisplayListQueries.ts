import type { DisplayList } from '../layout/render/displayList';
import {
  createDisplayListQueries,
  displayListQueriesHoldParsedList,
  endDisplayListQueriesLine,
  type DisplayListQueries,
} from '../layout/render/displayListQueries';
import { loadedRustDisplayListQueryEngine, type RustDisplayListQueryEngine } from '../layout/render/rustDisplayList';

/** Pages a session's query engine keeps parsed in the wasm store; the others are parsed again when a query reads them. */
export const RESIDENT_QUERY_STORE_PAGES = 12;

interface HeldQueries {
  list: DisplayList;
  engine: RustDisplayListQueryEngine | null;
  queries: DisplayListQueries;
}

const held = new WeakMap<object, HeldQueries>();

/**
 * The query engine over `list` for the session `owner`: one per display list, built by `create`
 * from the previous one so the parsed pages both lists share stay parsed. One that recorded a source
 * error is closed and the next is built fresh.
 */
export function residentDisplayListQueries(
  owner: object, list: DisplayList, create: typeof createDisplayListQueries = createDisplayListQueries
): DisplayListQueries {
  const engine = loadedRustDisplayListQueryEngine();
  const current = held.get(owner);
  const failed = current?.queries.sourceState?.().status === 'error';
  if (current && !failed && current.list === list && current.engine === engine) return current.queries;
  const donor = current && !failed && current.engine === engine ? current.queries : null;
  const queries = create(list, engine ?? undefined, donor, owner, RESIDENT_QUERY_STORE_PAGES);
  if (donor && displayListQueriesHoldParsedList(donor)) queries.prime();
  current?.queries.dispose();
  held.set(owner, { list, engine, queries });
  return queries;
}

/** Closes the query engine of the session `owner` and the wasm store pages it holds. */
export function releaseResidentDisplayListQueries(owner: object): void {
  held.get(owner)?.queries.dispose();
  held.delete(owner);
  endDisplayListQueriesLine(owner);
}
