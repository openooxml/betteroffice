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

const presentedLists = new WeakMap<object, object>();

/** Records that the canvas pages under `host` finished painting `displayList`. */
export function markPresented(host: object, displayList: object): void {
  presentedLists.set(host, displayList);
}

/** Forgets what `host` shows, while its canvas pages repaint for a new surface or zoom. */
export function clearPresented(host: object): void {
  presentedLists.delete(host);
}

/** Whether the canvas pages under `host` show the pixels of `displayList`. */
export function isPresented(host: object | null | undefined, displayList: object): boolean {
  return !!host && presentedLists.get(host) === displayList;
}
