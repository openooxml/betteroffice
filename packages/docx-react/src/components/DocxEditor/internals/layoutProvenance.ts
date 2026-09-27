const sourceVersions = new WeakMap<object, string>();

/** Records the document version a layout or query facade was built from. */
export function stampSourceVersion(target: object, version: string | null): void {
  if (version !== null) sourceVersions.set(target, version);
}

/** The document version `target` shows, or null when unknown. */
export function sourceVersionOf(target: object | null | undefined): string | null {
  return target ? (sourceVersions.get(target) ?? null) : null;
}

/** A session's current version; null for a session that cannot report one. */
export function readSessionVersion(
  session: { version?: () => string } | null | undefined
): string | null {
  try {
    return typeof session?.version === 'function' ? session.version() : null;
  } catch {
    return null;
  }
}
