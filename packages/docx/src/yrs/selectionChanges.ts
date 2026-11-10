import type { YrsSession } from './index';

export type YrsSelectionCause = 'remote' | 'seed';

interface Publication {
  version: string;
  selection: string;
  cause: YrsSelectionCause | undefined;
}

const publications = new WeakMap<object, Publication>();
const scopes = new WeakMap<object, { cause: YrsSelectionCause | undefined }>();

function selectionKey(selection: ReturnType<YrsSession['selection']>): string {
  const point = (loc: NonNullable<typeof selection>['anchor']) => [loc.story, loc.paraId, loc.offset];
  return JSON.stringify(selection ? [point(selection.anchor), point(selection.head)] : null);
}

export function yrsSelectionCause(session: YrsSession): YrsSelectionCause | undefined {
  const scope = scopes.get(session);
  if (scope) return scope.cause;
  const publication = publications.get(session);
  return publication?.version === session.version() && publication.selection === selectionKey(session.selection())
    ? publication.cause : undefined;
}

export function withYrsSelectionCause<T>(session: YrsSession, cause: YrsSelectionCause | undefined, run: () => T): T {
  const previous = scopes.get(session);
  scopes.set(session, { cause });
  try { return run(); }
  finally {
    if (previous) scopes.set(session, previous);
    else scopes.delete(session);
  }
}

export function recordYrsSelectionCause(session: YrsSession, cause = scopes.get(session)?.cause,
  selection = session.selection()): void {
  publications.set(session, { version: session.version(), selection: selectionKey(selection), cause });
}
