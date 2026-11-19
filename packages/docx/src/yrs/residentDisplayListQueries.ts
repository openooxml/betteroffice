import type { DisplayList } from '../layout/render/displayList';
import {
  createDisplayListQueries,
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
 * The query engine over `list` for the session `owner`: one per display list, built from the
 * previous one so the parsed pages both lists share stay parsed.
 */
export function residentDisplayListQueries(owner: object, list: DisplayList): DisplayListQueries {
  const engine = loadedRustDisplayListQueryEngine();
  const current = held.get(owner);
  if (current && current.list === list && current.engine === engine) return current.queries;
  const queries = createDisplayListQueries(list, engine ?? undefined,
    current && current.engine === engine ? current.queries : null, owner, RESIDENT_QUERY_STORE_PAGES);
  queries.prime();
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
