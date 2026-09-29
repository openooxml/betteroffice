import type { YrsRenderEnv } from '@betteroffice/docx/yrs';

const sourceVersions = new WeakMap<object, string>();

/** Records the document version a layout or query facade was built from. */
export function stampSourceVersion(target: object, version: string | null): void {
  if (version !== null) sourceVersions.set(target, version);
}

/** The document version `target` shows, or null when unknown. */
export function sourceVersionOf(target: object | null | undefined): string | null {
  return target ? (sourceVersions.get(target) ?? null) : null;
}

const revisionPreviewKeys = new WeakMap<object, string>();

/** A canonical key for a revision preview; '' when nothing is previewed. */
export function revisionPreviewKey(preview?: YrsRenderEnv['revisionPreview']): string {
  const entries = Object.entries(preview ?? {})
    .filter(([, decision]) => decision === 'accepted' || decision === 'rejected')
    .sort(([left], [right]) => (left < right ? -1 : 1));
  return entries.length === 0 ? '' : JSON.stringify(entries);
}

/** Stamped on a frame whose revision preview is not known; it matches no preview. */
export const UNKNOWN_REVISION_PREVIEW_KEY = '?';

/** Records the {@link revisionPreviewKey} a layout or query facade was rendered with. */
export function stampRevisionPreviewKey(target: object, key: string): void {
  revisionPreviewKeys.set(target, key);
}

/** The {@link revisionPreviewKey} `target` was rendered with, or null when unstamped. */
export function revisionPreviewKeyOf(target: object | null | undefined): string | null {
  return target ? (revisionPreviewKeys.get(target) ?? null) : null;
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
