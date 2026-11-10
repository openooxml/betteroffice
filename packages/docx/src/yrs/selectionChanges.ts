import type { YrsSession } from './index';

export type YrsSelectionCause = 'remote' | 'seed';

interface Publication {
  version: string;
  cause: YrsSelectionCause | undefined;
}

const publications = new WeakMap<object, Publication>();
const scopes = new WeakMap<object, { cause: YrsSelectionCause | undefined }>();
const listeners = new WeakMap<object, Set<(cause: YrsSelectionCause | undefined) => void>>();

export function yrsSelectionCause(session: YrsSession): YrsSelectionCause | undefined {
  const scope = scopes.get(session);
  if (scope) return scope.cause;
  const publication = publications.get(session);
  return publication?.version === session.version() ? publication.cause : undefined;
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

export function notifyYrsSelection(session: YrsSession, cause = scopes.get(session)?.cause): void {
  publications.set(session, { version: session.version(), cause });
  for (const listener of listeners.get(session) ?? []) listener(cause);
}

export function subscribeYrsSelection(session: YrsSession, listener: (cause: YrsSelectionCause | undefined) => void): () => void {
  const subscribed = listeners.get(session) ?? new Set<(cause: YrsSelectionCause | undefined) => void>();
  listeners.set(session, subscribed);
  subscribed.add(listener);
  return () => { subscribed.delete(listener); };
}
