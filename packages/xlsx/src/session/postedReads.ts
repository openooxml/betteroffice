import type { WorkbookEditSession } from './editSession';

type PostedRunner = <T>(call: () => Promise<T>) => Promise<T>;
const postedReads = new WeakMap<object, PostedRunner>();

export function registerPostedReads(session: object, run: PostedRunner): void {
  postedReads.set(session, run);
}

/**
 * Like runAfterEdits, but later edits wait only until the read is posted, not until it replies.
 * For reads the worker captures when it runs them and orders before later edits (non-reorderable).
 */
export function runAfterEditsUntilPosted<T>(session: WorkbookEditSession, call: () => Promise<T>): Promise<T> {
  const run = postedReads.get(session);
  return run ? run(call) : session.runAfterEdits(call);
}
