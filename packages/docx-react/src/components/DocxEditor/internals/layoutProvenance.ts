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

export interface WorkerFrameProvenance {
  version: string;
  preview: boolean;
  asOpened: boolean;
}

const workerFrames = new WeakMap<object, WorkerFrameProvenance>();

/** Records the worker document version a worker frame's display list and queries show. */
export function stampWorkerFrameVersion(
  target: object,
  version: string | null | undefined,
  preview = false,
  asOpened = false
): void {
  if (version != null) workerFrames.set(target, { version, preview, asOpened });
}

/** The worker document version `target` shows, or null for a frame the worker did not lay out. */
export function workerFrameVersionOf(target: object | null | undefined): string | null {
  return target ? (workerFrames.get(target)?.version ?? null) : null;
}

/** The provenance of the presented worker frame. */
export function presentedWorkerFrame(
  queries: { readonly displayList: object } | null | undefined
): WorkerFrameProvenance | null {
  return queries
    ? (workerFrames.get(queries) ?? workerFrames.get(queries.displayList) ?? null)
    : null;
}

/** The worker document version the frame `queries` query lays out, or null. */
export function presentedWorkerVersion(
  queries: { readonly displayList: object } | null | undefined
): string | null {
  return presentedWorkerFrame(queries)?.version ?? null;
}

const supersededLayouts = new WeakSet<object>();

/** Records that the document changed past `layout` before it was shown. */
export function markSupersededLayout(layout: object, superseded = true): void {
  if (superseded) supersededLayouts.add(layout);
  else supersededLayouts.delete(layout);
}

/** Whether the document changed past `layout` before it was shown; such a layout never settles. */
export function isSupersededLayout(layout: object | null | undefined): boolean {
  return layout ? supersededLayouts.has(layout) : false;
}

const queuedLayoutSessions = new WeakSet<object>();

/** Records whether a pass for `session` that waited behind a worker pass has yet to start. */
export function markLayoutQueued(session: object, queued: boolean): void {
  if (queued) queuedLayoutSessions.add(session);
  else queuedLayoutSessions.delete(session);
}

/** Whether a layout pass for `session` is queued; nothing shown meanwhile settles. */
export function isLayoutQueued(session: object | null | undefined): boolean {
  return session ? queuedLayoutSessions.has(session) : false;
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

interface PresentationOptions {
  worker?: boolean;
}

const presentedLists = new WeakMap<object, object>();
const presentListeners = new Set<(displayList: object, options?: PresentationOptions) => void>();

/** Records that the canvas pages under `host` finished painting `displayList`. */
export function markPresented(
  host: object,
  displayList: object,
  options?: PresentationOptions
): void {
  presentedLists.set(host, displayList);
  for (const listener of [...presentListeners]) listener(displayList, options);
}

/** Calls `listener` with each display list whose pages finish painting. */
export function onPresented(
  listener: (displayList: object, options?: PresentationOptions) => void
): () => void {
  presentListeners.add(listener);
  return () => presentListeners.delete(listener);
}

const replayFailureListeners = new Set<(displayList: object, error: unknown) => void>();

/** Records that the canvas pages failed to paint `displayList`. */
export function markReplayFailed(displayList: object, error: unknown): void {
  for (const listener of [...replayFailureListeners]) listener(displayList, error);
}

/** Calls `listener` with each display list whose pages fail to paint. */
export function onReplayFailed(
  listener: (displayList: object, error: unknown) => void
): () => void {
  replayFailureListeners.add(listener);
  return () => replayFailureListeners.delete(listener);
}

/** Forgets what `host` shows, while its canvas pages repaint for a new surface or zoom. */
export function clearPresented(host: object): void {
  presentedLists.delete(host);
}

/** Whether the canvas pages under `host` show the pixels of `displayList`. */
export function isPresented(host: object | null | undefined, displayList: object): boolean {
  return !!host && presentedLists.get(host) === displayList;
}
