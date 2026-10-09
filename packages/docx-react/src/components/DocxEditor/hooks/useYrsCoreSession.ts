import { useCallback, useEffect, useRef, useState } from 'react';
import type { LayoutBlock } from '@betteroffice/docx/layout/pagination';
import type { Document } from '@betteroffice/docx/types/document';
import type {
  YrsDocxHost,
  YrsInputPositionMap,
  YrsLoc,
  YrsRenderEnv,
  YrsSession,
} from '@betteroffice/docx/yrs';
import {
  EditorDirtyStories,
  ResidentWorkerSaveUnavailableError,
  hostSaveMetadata,
  mergeDocxHostMetadata,
  serialWorkerSaves,
} from '@betteroffice/docx/yrs';
import type { DocxEditorCollaborationOptions } from '../types';
import type {
  OpenInWorker,
  OpenPreviewInWorker,
  UseRustDisplayListResult,
  WorkerOpenedDocument,
} from './useDisplayList';
import { markLayoutQueued } from '../internals/layoutProvenance';
import {
  adoptWorkerOpenHandoverVersion,
  adoptWorkerOpenMirrorVersion,
  deferWorkerOpenReplica,
  holdWorkerOpenDocument,
  releaseWorkerOpenDocument,
  requestWorkerOpenReplica,
  workerOpenDocumentHeld,
  workerOpenReplicaPending,
} from '../internals/workerOpenReplica';
import { DocxWorkerError } from '../internals/docxWorkerError';
import {
  beginWorkerProposalHandover,
  installEditorWorkerProposalActivation,
  hasEditorWorkerProposalRounds,
  registerWorkerProposalAuthority,
  registeredWorkerProposalAuthority,
  workerProposalFailure,
} from '../internals/workerProposalAuthority';
import { registerWorkerOpenSave } from '../internals/workerOpenSave';
import { registerWorkerOpenExport } from '../internals/workerOpenExport';
import { registerQueuedOpeningInput } from '../internals/queuedOpeningInput';
import { bootstrapWorkerOpenPeer } from '../internals/bootstrapWorkerOpenPeer';

export { dirtyProjectionStory, mergeDocxHostMetadata } from '@betteroffice/docx/yrs';

type YrsFacadeModule = typeof import('@betteroffice/docx/yrs');

/** The React editor's sole mutable document session. */
export interface YrsCoreSession {
  session: YrsSession | null;
  /** The seed generation `session` was created for. */
  sessionGeneration: number | null;
  replicaReady: boolean;
  /** Starts loading the main-thread replica when needed. */
  requestReplica(): void;
  workerProposalsReady: boolean;
  replicaReadyRef?: React.RefObject<boolean>;
  experimentalWorkerOpen?: boolean;
  storyBlocks(storyId: string, env: YrsRenderEnv): LayoutBlock[] | null;
  bodyBlocks(env: YrsRenderEnv): LayoutBlock[] | null;
  inputPositionMap(storyId?: string): YrsInputPositionMap | null;
  displayPositionToLoc(position: number, storyId?: string): YrsLoc | null;
  locToDisplayPosition(loc: YrsLoc): number | null;
  documentFromYrs(baseDocument?: Document | null): Document | null;
  /** Marks stories changed outside the save projection; the live selection's story by default. */
  publishDirectInput(stories?: string | readonly string[]): void;
  /**
   * `session` is a display-only preview of the document's first pages; the
   * full session replaces it once they have painted.
   */
  previewing: boolean;
  /** The preview session whose pages the renderer keeps until the full session's replace them. */
  handoffFrom: YrsSession | null;
  /** From the preview until a frame of the full session is presented: nothing reads or edits. */
  opening: boolean;
  /** A frame of `engine`'s layout was published for display. */
  notifyFramePresented(engine: unknown): void;
  /**
   * Fails the load with `error` if its full session has yet to show a frame,
   * dropping that session and the preview. Returns whether it did. An error
   * of another session than `session`, when given, fails nothing.
   */
  failOpening(error: Error, session?: unknown): boolean;
  /**
   * Materializes the save projection base when the main thread is next idle,
   * for a host that projects the document on every change.
   */
  scheduleCompatibilityWarm(): void;
  cancelCompatibilityWarm(): void;
}

interface YrsCoreSessionCallbacks {
  isCurrentLoad?: (generation: number) => boolean;
  /** A session was created for the current load, before it is seeded. */
  onSession?: (session: YrsSession) => void;
  onHostDocument?: (
    host: YrsDocxHost,
    generation: number,
    session: YrsSession,
    options?: { preview: boolean }
  ) => void;
  /** `opened`: the load's full document was already accepted. */
  onError?: (error: Error, generation: number, options?: { opened: boolean }) => void;
  onReplicaError?: (error: Error, generation: number) => void;
}

interface WorkerOpenOptions {
  openInWorker: OpenInWorker;
  /** Opens the first-page preview in the worker too, so this thread runs none of it. */
  openPreviewInWorker?: OpenPreviewInWorker;
  renderedFrame: object | null;
  settledDisplayList?: UseRustDisplayListResult['settledDisplayList'];
  workerProposals?: boolean;
  viewer?: boolean;
  refreshWorkerLayout?: () => void;
  /** The engine whose provisional layout is shown with the rest not yet asked of the worker. */
  pendingCompletion?: unknown;
  layoutCompleteSession?: unknown;
  /** A worker-held proposal changed document content. */
  onWorkerContentChange?: () => void;
  onPeerUpdate?: (stories: readonly string[]) => void;
  /** The worker found tracked changes of its own in the opened document. */
  onWorkerRevisions?: () => void;
}

export interface YrsCoreSessionOptions {
  /** Open a display-only preview of the first pages before the full document. */
  previewFirstPage?: boolean;
  /** How long a preview waits for the full document to open; see {@link FULL_OPEN_TIMEOUT_MS}. */
  fullOpenTimeoutMs?: number;
  /** The engine the renderer still builds with; a replaced session it names lives on. */
  heldEngine?: unknown;
  /** The engine whose frame is on screen; a replaced session it names lives on. */
  shownEngine?: unknown;
  workerOpen?: WorkerOpenOptions;
  /** Open images as `media:{n}` tokens the canvas resolves from the session. */
  mediaTokens?: boolean;
}

/** Body blocks a first-page preview parses. */
const PREVIEW_BODY_BLOCKS = 200;
/** Paragraph weight budget for a first-page preview. */
const PREVIEW_PARAGRAPH_BUDGET = 256;
/** How long the full open waits for the preview's pages to paint. */
const PREVIEW_PAINT_TIMEOUT_MS = 2000;
/** Bounds the wait for the painted preview to reach the screen; hidden tabs get no frames. */
const PREVIEW_FRAME_WAIT_MS = 100;
const REPLICA_FRAME_WAIT_MS = 1000;
/**
 * How long a preview waits for the full session, from the end of its own
 * paint. A full open that has not produced one by then fails the load; once
 * it exists, the preview stays until the full session's first frame shows.
 */
const FULL_OPEN_TIMEOUT_MS = 10_000;

/** A display-only session of the first pages of `bytes`, or null when it cannot open. */
async function openPreview(
  yrs: YrsFacadeModule,
  bytes: Uint8Array,
  clientId: number | undefined
): Promise<{ session: YrsSession; host: YrsDocxHost } | null> {
  const session = await yrs.createYrsSession({ clientId });
  try {
    const host = session.openDocxPreview(bytes, PREVIEW_BODY_BLOCKS, PREVIEW_PARAGRAPH_BUDGET);
    if (host) return { session, host };
    session.destroy();
    return null;
  } catch (error) {
    console.warn('[yrs] the first-page preview could not open; opening in full', error);
    session.destroy();
    return null;
  }
}

/**
 * A display-only session of the first pages of `bytes` that the resident worker opened and lays
 * out, or null when no worker takes it. The session holds no document of its own: it loads the
 * preview on this thread only if something needs it here.
 */
async function openWorkerPreview(
  yrs: YrsFacadeModule,
  bytes: Uint8Array,
  clientId: number | undefined,
  openPreviewInWorker: OpenPreviewInWorker,
  onPosted: () => void,
  viewer: boolean
): Promise<{ session: YrsSession; host: YrsDocxHost; laidOut: Promise<void> } | null> {
  const session = await yrs.createYrsSession({ clientId });
  session.markDisplayOnly();
  let release = (): void => {};
  if (!viewer) {
    const loadHere = (): void => {
      const host = session.openDocxPreview(bytes, PREVIEW_BODY_BLOCKS, PREVIEW_PARAGRAPH_BUDGET);
      if (!host) throw new Error('The first-page preview cannot open');
      if (host.wholeBody) session.setPartialDocument(false);
    };
    deferWorkerOpenReplica(session, async () => loadHere, loadHere, () => release());
  }
  try {
    const pending = openPreviewInWorker(session, bytes, PREVIEW_BODY_BLOCKS, PREVIEW_PARAGRAPH_BUDGET);
    onPosted();
    const opened = await pending;
    if (opened) {
      release = opened.release;
      const host = yrs.decodeDocxHostJson(opened.hostJson, bytes);
      if (host.wholeBody) session.setPartialDocument(false);
      return { session, host, laidOut: opened.bootstrapPosted };
    }
  } catch (error) {
    console.warn('[yrs] the worker could not open the first-page preview', error);
  }
  session.destroy();
  return null;
}

export interface YrsSeedSources {
  bytes: Uint8Array | null;
  document: Document | null;
  initialUpdate?: Uint8Array;
  mediaTokens?: boolean;
}

/**
 * Hydrates a fresh session. Shared collaboration state wins over both seed
 * shapes so a client joining a room never seeds an independent replica.
 */
export function seedYrsSession(
  session: Pick<YrsSession, 'openDocx' | 'loadState'>,
  seedDocumentIntoYrs: (document: Document) => void,
  seed: YrsSeedSources
): YrsDocxHost | null {
  const { bytes, document, initialUpdate, mediaTokens } = seed;
  if (bytes) {
    const host = session.openDocx(bytes, !initialUpdate, mediaTokens ? { mediaTokens } : undefined);
    if (initialUpdate) session.loadState(initialUpdate.slice());
    return host;
  }
  if (initialUpdate) {
    session.loadState(initialUpdate.slice());
    return null;
  }
  if (document) seedDocumentIntoYrs(document);
  return null;
}

/**
 * Materializes the save-projection base once. `materializeDocx` re-parses the
 * retained source and ships the full envelope (every media entry, twice, as
 * JSON), so a host projecting every change warms it before the first edit.
 */
export function warmCompatibilityBase(
  session: Pick<YrsSession, 'materializeDocx'>,
  compatibilityBase: { current: Document | null }
): void {
  if (compatibilityBase.current) return;
  try {
    compatibilityBase.current = session.materializeDocx();
  } catch (error) {
    console.error('[yrs] failed to warm the save projection base', error);
  }
}

/**
 * Frees sessions the editor let go of. Consumers' effects in the commit that replaces a session
 * still run with the session they rendered, and `held` names sessions still shown after it (a
 * handed-off preview, the renderer's layout engine), so `retire` keeps such a session until
 * nothing renders with it; any other session is freed at once, and unmounting frees every
 * retired one.
 */
function useRetiredSessions(
  session: YrsSession | null,
  /** A fixed number of holders on every render: the effect compares them slot by slot. */
  held: readonly unknown[]
): (replaced: YrsSession | null) => void {
  const renderedRef = useRef({ session, held });
  renderedRef.current = { session, held };
  const retiredRef = useRef(new Set<YrsSession>());
  const unmountedRef = useRef(false);
  useEffect(() => {
    for (const retired of retiredRef.current) {
      if (retired === session || held.includes(retired)) continue;
      retiredRef.current.delete(retired);
      retired.destroy();
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [session, ...held]);
  useEffect(() => {
    unmountedRef.current = false;
    return () => {
      unmountedRef.current = true;
      for (const retired of retiredRef.current) retired.destroy();
      retiredRef.current.clear();
    };
  }, []);
  return useCallback((replaced: YrsSession | null): void => {
    if (!replaced) return;
    const rendered = renderedRef.current;
    if (
      !unmountedRef.current &&
      (replaced === rendered.session || rendered.held.includes(replaced))
    ) {
      retiredRef.current.add(replaced);
    } else {
      replaced.destroy();
    }
  }, []);
}

export function useYrsCoreSession(
  enabled: boolean,
  document: Document | null,
  seedDocument: Document | null,
  seedBytes: Uint8Array | null,
  seedGeneration: number,
  collaboration?: DocxEditorCollaborationOptions,
  callbacks?: YrsCoreSessionCallbacks,
  options?: YrsCoreSessionOptions
): YrsCoreSession {
  const workerOpen = options?.workerOpen;
  const collaborationClientId = collaboration?.clientId;
  const collaborationInitialUpdate = workerOpen ? undefined : collaboration?.initialUpdate;
  const sessionRef = useRef<YrsSession | null>(null);
  const facadeRef = useRef<YrsFacadeModule | null>(null);
  const documentRef = useRef(document);
  documentRef.current = document;
  const callbacksRef = useRef(callbacks);
  callbacksRef.current = callbacks;
  const compatibilityBaseRef = useRef<Document | null>(null);
  const compatibilitySourceBufferRef = useRef<ArrayBuffer | undefined>(undefined);
  const cancelCompatibilityWarmRef = useRef<(() => void) | null>(null);
  const seedBytesRef = useRef(seedBytes);
  seedBytesRef.current = seedBytes;
  const mediaTokensRef = useRef(workerOpen ? false : options?.mediaTokens);
  mediaTokensRef.current = workerOpen ? false : options?.mediaTokens;
  const inputPositionMapsRef = useRef(new Map<string, YrsInputPositionMap>());
  const dirtyStoriesRef = useRef(new EditorDirtyStories());
  const markProjectionStories = useCallback((stories: readonly string[]): void => {
    for (const story of stories) dirtyStoriesRef.current.add(story);
  }, []);
  const enabledRef = useRef(enabled);
  enabledRef.current = enabled;
  const [session, setSession] = useState<YrsSession | null>(null);
  const [sessionGeneration, setSessionGeneration] = useState<number | null>(null);
  const [replicaReady, setReplicaReady] = useState(true);
  const [replicaRequestVersion, setReplicaRequestVersion] = useState(0);
  const [workerProposalsReady, setWorkerProposalsReady] = useState(false);
  const workerOpenRef = useRef(workerOpen);
  workerOpenRef.current = workerOpen;
  const replicaReadyRef = useRef(true);
  const openInWorker = workerOpen?.openInWorker;
  const openPreviewInWorkerRef = useRef(workerOpen?.openPreviewInWorker);
  openPreviewInWorkerRef.current = workerOpen?.openPreviewInWorker;
  const workerOpenEnabledRef = useRef(Boolean(openInWorker));
  workerOpenEnabledRef.current = Boolean(openInWorker);
  const pendingReplicaRef = useRef<ReturnType<typeof deferWorkerOpenReplica> | null>(null);
  const startReplicaRef = useRef<(() => void) | null>(null);
  const replicaStatePrefetchRef = useRef<{ start(): void; clear(): void } | null>(null);
  // Asks the worker whether the document has tracked changes, once per session.
  const revisionQueryRef = useRef<(() => void) | null>(null);
  const workerLaidOutRef = useRef<(() => void) | null>(null);
  const replicaGateRef = useRef<{ reached: boolean } | null>(null);
  const requestReplicaRef = useRef<(() => void) | null>(null);
  const openingInputSessionRef = useRef<YrsSession | null>(null);
  const eagerPeerStartRef = useRef<{ session: YrsSession; start(): void } | null>(null);
  const openReplicaGate = useCallback((): void => {
    const gate = replicaGateRef.current;
    const start = startReplicaRef.current;
    if (!gate || !start) return;
    gate.reached = true;
    start();
  }, []);
  const inheritedFrameRef = useRef<object | null>(null);
  const renderedFrameRef = useRef(workerOpen?.renderedFrame ?? null);
  renderedFrameRef.current = workerOpen?.renderedFrame ?? null;
  const [previewing, setPreviewing] = useState(false);
  const [handoffFrom, setHandoffFrom] = useState<YrsSession | null>(null);
  const paintWaitRef = useRef<{ session: YrsSession; resolve: () => void } | null>(null);
  const previewingRef = useRef(false);
  previewingRef.current = previewing;
  const retiringRef = useRef<YrsSession | null>(null);
  const failOpeningRef = useRef<((error: Error, session?: unknown) => boolean) | null>(null);
  // Collaboration shares one replica from the start, so it never previews.
  // Read once per load: a later change of the option does not reopen it.
  const previewFirstPageRef = useRef(false);
  previewFirstPageRef.current =
    options?.previewFirstPage === true && !collaboration && !collaborationInitialUpdate;
  const fullOpenTimeoutRef = useRef(FULL_OPEN_TIMEOUT_MS);
  fullOpenTimeoutRef.current = options?.fullOpenTimeoutMs ?? FULL_OPEN_TIMEOUT_MS;
  const retire = useRetiredSessions(session, [
    handoffFrom,
    options?.heldEngine ?? null,
    options?.shownEngine ?? null,
  ]);
  // The handoff and the renderer's layout hold the preview until the full session replaces both.
  const retirePreview = useCallback(
    (retiring: YrsSession): void => {
      if (retiringRef.current === retiring) {
        retiringRef.current = null;
        setHandoffFrom(null);
      }
      retire(retiring);
    },
    [retire]
  );

  useEffect(() => {
    setSession(null);
    setWorkerProposalsReady(false);
    setPreviewing(false);
    setHandoffFrom(null);
    if (openInWorker) {
      replicaReadyRef.current = true;
      setReplicaReady(true);
    }
    if (!enabled || (!seedDocument && !seedBytes)) return;
    let cancelled = false;
    let openedWorker: WorkerOpenedDocument | null = null;
    let unregisterSave: (() => void) | null = null;
    let unregisterExport: (() => void) | null = null;
    let unregisterOpeningInput: (() => void) | null = null;
    inputPositionMapsRef.current.clear();
    dirtyStoriesRef.current.clear();
    compatibilityBaseRef.current = null;
    compatibilitySourceBufferRef.current = undefined;

    let abandoned = false;
    let shown: { session: YrsSession; host: YrsDocxHost } | null = null;
    let fullOpenTimer: ReturnType<typeof setTimeout> | null = null;
    // Disposes a full open started alongside the preview that the load never took.
    let dropEarly = (): void => {};
    let earlyTaken = false;
    const stale = () =>
      cancelled ||
      abandoned ||
      callbacksRef.current?.isCurrentLoad?.(seedGeneration) === false;
    const previewFirstPage = previewFirstPageRef.current;
    // A media-token load keeps its full open on the main thread.
    const openWorker = mediaTokensRef.current ? undefined : openInWorker;
    // A failed full open takes the preview down with it, as a failed open
    // without one would leave no session.
    const dropPreview = (preview: YrsSession): void => {
      if (sessionRef.current === preview) {
        sessionRef.current = null;
        setSession(null);
      }
      previewingRef.current = false;
      setPreviewing(false);
      retirePreview(preview);
    };
    const fail = (error: unknown, options?: { opened: boolean }): void => {
      // A replaced load's preview was retired with it; the next load owns the preview state.
      if (shown && !cancelled) dropPreview(shown.session);
      if (!cancelled && callbacksRef.current?.isCurrentLoad?.(seedGeneration) !== false) {
        callbacksRef.current?.onError?.(
          error instanceof Error ? error : new Error(String(error)),
          seedGeneration,
          options
        );
      }
    };

    failOpeningRef.current = (error: Error, session?: unknown): boolean => {
      const full = sessionRef.current;
      if (!shown || stale() || !full || full === shown.session) return false;
      if (session !== undefined && session !== full) return false;
      abandoned = true;
      pendingReplicaRef.current?.cancel();
      pendingReplicaRef.current = null;
      openedWorker?.destroy();
      openedWorker = null;
      sessionRef.current = null;
      setSession(null);
      retire(full);
      retiringRef.current = null;
      setHandoffFrom(null);
      fail(error, { opened: true });
      return true;
    };

    void import('@betteroffice/docx/yrs')
      .then(async (yrs) => {
        // A copy hashed with Web Crypto keeps the package's hash off this
        // thread; it hashes while the preview opens.
        let preparing: Promise<Uint8Array | null> | null = null;
        const prepare = (): Promise<Uint8Array | null> => {
          if (!preparing) {
            preparing = seedBytes ? yrs.prepareDocxBytes(seedBytes) : Promise.resolve(null);
            // Awaited below unless the load ends first.
            preparing.catch(() => {});
          }
          return preparing;
        };
        // The full session, opened in the worker when one takes it.
        const openFull = async (): Promise<{
          bytes: Uint8Array | null;
          next: YrsSession;
          host: YrsDocxHost | null;
        } | null> => {
          const bytes = await prepare();
          if (stale()) return null;
          const next = await yrs.createYrsSession({ clientId: collaborationClientId });
          if (stale()) {
            next.destroy();
            return null;
          }
          callbacksRef.current?.onSession?.(next);
          let host: YrsDocxHost | null = null;
          if (openWorker && bytes && !collaborationInitialUpdate) {
            try {
              openedWorker = await openWorker(
                next,
                bytes,
                yrs.preparedDocxDigest(bytes),
                seedGeneration
              );
              if (!stale()) {
                host = openedWorker ? yrs.decodeDocxHostJson(openedWorker.hostJson, bytes) : null;
              }
            } catch (error) {
              openedWorker?.destroy();
              openedWorker = null;
              next.destroy();
              if (error instanceof yrs.ResidentWorkerOutOfMemoryError) {
                throw error;
              }
              throw error instanceof DocxWorkerError ? error : new DocxWorkerError('open', error);
            }
            if (stale()) {
              openedWorker?.destroy();
              openedWorker = null;
              next.destroy();
              return null;
            }
          }
          return { bytes, next, host };
        };
        // The worker opens the full document while the preview opens and paints. A preview the
        // worker opened lays out there first: the full open queues right behind that layout.
        const openPreviewInWorker =
          openWorker &&
          previewFirstPage &&
          seedBytes &&
          !collaborationInitialUpdate &&
          yrs.canUseResidentEngineWorker()
            ? openPreviewInWorkerRef.current
            : undefined;
        // The worker's preview open goes first; the hash then runs while the worker parses.
        if (!openPreviewInWorker) void prepare();
        let releaseFull = (): void => {};
        const fullQueued = new Promise<void>((resolve) => {
          releaseFull = resolve;
        });
        const early =
          openWorker && previewFirstPage && seedBytes && !collaborationInitialUpdate
            ? openPreviewInWorker
              ? fullQueued.then(openFull)
              : openFull()
            : null;
        early?.catch(() => {});
        dropEarly = (): void => {
          releaseFull();
          if (earlyTaken) return;
          earlyTaken = true;
          void early?.then((full) => {
            if (!full) return;
            openedWorker?.destroy();
            openedWorker = null;
            full.next.destroy();
          }, () => {});
        };
        // A preview paints the first pages before the full open begins.
        const inWorker =
          openPreviewInWorker && seedBytes
            ? await openWorkerPreview(
                yrs,
                seedBytes,
                collaborationClientId,
                openPreviewInWorker,
                () => void prepare(),
                workerOpenRef.current?.viewer === true
              )
            : null;
        const opened =
          inWorker ??
          (previewFirstPage && seedBytes && !stale() && !workerOpenRef.current?.viewer
            ? await openPreview(yrs, seedBytes, collaborationClientId)
            : null);
        shown = opened;
        if (!inWorker) releaseFull();
        if (opened && stale()) {
          opened.session.destroy();
          dropEarly();
          return;
        }
        if (opened) {
          callbacksRef.current?.onSession?.(opened.session);
          const painted = new Promise<void>((resolve) => {
            paintWaitRef.current = { session: opened.session, resolve };
            setTimeout(resolve, PREVIEW_PAINT_TIMEOUT_MS);
          });
          if (inWorker) void Promise.race([inWorker.laidOut, painted]).then(releaseFull);
          sessionRef.current = opened.session;
          facadeRef.current = yrs;
          previewingRef.current = true;
          setSession(opened.session);
          setPreviewing(true);
          setSessionGeneration(seedGeneration);
          callbacksRef.current?.onHostDocument?.(opened.host, seedGeneration, opened.session, {
            preview: true,
          });
          await painted;
          if (paintWaitRef.current?.session === opened.session) paintWaitRef.current = null;
          // Two frames: a worker canvas's commit can reach the screen a frame
          // after the presentation, and a main-thread full open blocks this thread.
          await new Promise<void>((resolve) => {
            const bound = setTimeout(resolve, PREVIEW_FRAME_WAIT_MS);
            requestAnimationFrame(() =>
              requestAnimationFrame(() =>
                setTimeout(() => {
                  clearTimeout(bound);
                  resolve();
                }, 0)
              )
            );
          });
          if (stale()) {
            dropEarly();
            return;
          }
          const preview = opened.session;
          fullOpenTimer = setTimeout(() => {
            fullOpenTimer = null;
            if (sessionRef.current !== preview) return;
            abandoned = true;
            fail(new Error('The document did not finish opening in time'));
          }, fullOpenTimeoutRef.current);
        }
        const full = await (early ?? openFull());
        earlyTaken = true;
        if (!full) return;
        if (stale()) {
          openedWorker?.destroy();
          openedWorker = null;
          full.next.destroy();
          return;
        }
        const { bytes, next } = full;
        let host = full.host;
        try {
          if (openedWorker && bytes) {
            inheritedFrameRef.current = renderedFrameRef.current;
            const worker = openedWorker;
            const source = bytes;
            const workerHost = full.host!;
            unregisterExport = registerWorkerOpenExport(next, {
              export: (options, context) => {
                if (stale()) return Promise.reject(new Error('The document changed while exporting'));
                return worker.exportStructuredWithPages(next, options, context);
              },
            });
            const saveInOrder = serialWorkerSaves(dirtyStoriesRef.current);
            unregisterSave = registerWorkerOpenSave(next, {
              available: () => !stale() && worker.canSave(),
              sourceReplaced: () => compatibilitySourceBufferRef.current !== undefined &&
                compatibilityBaseRef.current?.originalBuffer !== compatibilitySourceBufferRef.current,
              save: (comments, peer) => {
                const task = () => saveInOrder(async (stories) => {
                  if (stale()) throw new Error('The document changed while saving');
                  if (!worker.canSave()) throw new ResidentWorkerSaveUnavailableError('No document worker');
                  const currentHost = documentRef.current ?? host?.document;
                  return worker.save({
                    comments,
                    ...(currentHost ? { host: hostSaveMetadata(currentHost) } : {}),
                    ...(peer ? { stories } : {}),
                  }, peer, (apply) => dirtyStoriesRef.current.adoptWorkerSaveUpdates(apply));
                });
                const authority = hasEditorWorkerProposalRounds(next) ? registeredWorkerProposalAuthority(next) : null;
                return authority ? authority.save(task) : task();
              },
            });
            let revisionsQueried = false;
            const queryRevisions = (): void => {
              if (revisionsQueried) return;
              revisionsQueried = true;
              void worker.revisionCount().then(
                (count) => {
                  if (stale() || sessionRef.current !== next || count === 0) return;
                  workerOpenRef.current?.onWorkerRevisions?.();
                },
                () => {}
              );
            };
            const activateEditorProposals = () => {
              if (stale() || !worker.canSave()) return null;
              let laidOut = renderedFrameRef.current && renderedFrameRef.current !== inheritedFrameRef.current
                ? Promise.resolve()
                : new Promise<void>((resolve) => {
                    workerLaidOutRef.current = resolve;
                  });
              const authority = registerWorkerProposalAuthority(next, worker, {
                editorPeer: true,
                relayout: () => {
                  if (!authority.initialized) {
                    laidOut = new Promise<void>((resolve) => {
                      workerLaidOutRef.current = resolve;
                    });
                  }
                  markLayoutQueued(next, true);
                  workerOpenRef.current?.refreshWorkerLayout?.();
                },
                current: () => !stale(),
                laidOut: () => worker.whenBootstrapSent
                  ? Promise.race([laidOut, worker.whenBootstrapSent()])
                  : laidOut,
                contentChanged: () => workerOpenRef.current?.onWorkerContentChange?.(),
                projectionChanged: (stories) => {
                  inputPositionMapsRef.current.clear();
                  markProjectionStories(stories);
                },
                peerUpdated: (stories) => {
                  if (!workerOpenReplicaPending(next)) workerOpenRef.current?.onPeerUpdate?.(stories);
                },
                adopted: (version) => {
                  adoptWorkerOpenMirrorVersion(next, version);
                  worker.mirrorReady();
                },
              });
              authority.subscribe(() => {
                if (!stale()) setWorkerProposalsReady(authority.initialized);
              });
              return authority;
            };
            const deferEditorReplica = (): ReturnType<typeof deferWorkerOpenReplica> => {
              type StateRevision = NonNullable<ReturnType<NonNullable<WorkerOpenedDocument['stateRevision']>>>;
              type StateSnapshot = Awaited<ReturnType<NonNullable<WorkerOpenedDocument['encodeVersionedState']>>>;
              type PrefetchedState = {
                revision: StateRevision;
                result: Promise<StateSnapshot>;
              };
              let prefetchedState: PrefetchedState | null = null;
              const clearPrefetchedState = (): void => {
                prefetchedState = null;
              };
              const sameRevision = (revision: StateRevision): boolean => {
                const current = worker.stateRevision?.();
                return current?.owner === revision.owner && current.sequence === revision.sequence;
              };
              const encodeState = async (prefetch?: boolean): Promise<StateSnapshot> =>
                worker.encodeVersionedState
                  ? worker.encodeVersionedState(prefetch)
                  : { state: await worker.encodeState(prefetch), version: undefined };
              const encodeReplicaState = async (): Promise<StateSnapshot> => {
                const cached = prefetchedState;
                clearPrefetchedState();
                let update: StateSnapshot;
                if (cached && sameRevision(cached.revision)) {
                  update = await cached.result;
                  if (!sameRevision(cached.revision)) update = await encodeState();
                } else {
                  update = await encodeState();
                }
                if (stale() || sessionRef.current !== next || !pending.pending) {
                  throw new Error('The document changed while opening the replica');
                }
                return update;
              };
              const gate = { reached: false };
              const request = (): void => {
                pending.requestReady();
                if (gate.reached) startReplicaRef.current?.();
              };
              const pending = deferWorkerOpenReplica(
                next,
                async () => {
                  const handover = beginWorkerProposalHandover(next);
                  const handedOver = handover ? await handover : null;
                  const update = handedOver ?? await encodeReplicaState();
                  let bootstrapped = false;
                  return [
                    () => {
                      bootstrapped = bootstrapWorkerOpenPeer(next, update, source, workerHost);
                    },
                    () => { if (!bootstrapped) next.loadState(update.state); },
                    () => {
                      if (handedOver) handedOver.complete();
                      else if ('version' in update && update.version !== undefined) {
                        adoptWorkerOpenHandoverVersion(next, update.version);
                      }
                    },
                  ];
                },
                () => {
                  clearPrefetchedState();
                  if (registeredWorkerProposalAuthority(next)?.holdsWorkerState()) {
                    throw new Error('The resident worker holds proposals the main thread cannot rebuild');
                  }
                  worker.fallback();
                  if (hasEditorWorkerProposalRounds(next) && registeredWorkerProposalAuthority(next)?.retire('source-fallback')) {
                    console.warn('[yrs] the source fallback retired worker proposal authority to the peer');
                  }
                  next.openDocx(source, true);
                  if (registeredWorkerProposalAuthority(next)) next.mirrorWorkerDocument(null);
                },
                () => {
                  if (stale()) return;
                  inputPositionMapsRef.current.clear();
                  replicaReadyRef.current = true;
                  worker.replicaReady();
                  setReplicaReady(true);
                },
                {
                  current: () => !stale() && sessionRef.current === next &&
                    pendingReplicaRef.current === pending,
                  cancel: () => {
                    clearPrefetchedState();
                    worker.destroy();
                  },
                  waitForLayout: true,
                }
              );
              pendingReplicaRef.current = pending;
              replicaStatePrefetchRef.current = {
                clear: clearPrefetchedState,
                start: () => {
                  if (
                    stale() || sessionRef.current !== next ||
                    !pending.pending || pending.started || prefetchedState ||
                    registeredWorkerProposalAuthority(next)
                  ) return;
                  const revision = worker.stateRevision?.();
                  if (!revision) return;
                  const result = encodeState(true);
                  prefetchedState = { revision, result };
                  void result.catch(() => {});
                },
              };
              replicaGateRef.current = gate;
              requestReplicaRef.current = request;
              unregisterOpeningInput = registerQueuedOpeningInput(next, () => {
                if (stale() || sessionRef.current !== next || !pending.pending) return;
                openingInputSessionRef.current = next;
                if (eagerPeerStartRef.current?.session === next) eagerPeerStartRef.current.start();
              });
              startReplicaRef.current = () => {
                if (
                  stale() ||
                  sessionRef.current !== next ||
                  pendingReplicaRef.current !== pending
                ) return;
                requestWorkerOpenReplica(next);
              };
              replicaReadyRef.current = false;
              setReplicaReady(false);
              void pending.ready.catch((error: unknown) => {
                clearPrefetchedState();
                if (!stale() && sessionRef.current === next && pendingReplicaRef.current === pending) {
                  const onError = callbacksRef.current?.onReplicaError ?? callbacksRef.current?.onError;
                  onError?.(
                    error instanceof Error ? error : new Error(String(error)),
                    seedGeneration
                  );
                }
              });
              revisionQueryRef.current = queryRevisions;
              return pending;
            };
            if (workerOpenRef.current?.viewer) {
              holdWorkerOpenDocument(next, deferEditorReplica);
              revisionQueryRef.current = queryRevisions;
              replicaReadyRef.current = false;
              setReplicaReady(false);
            } else {
              deferEditorReplica();
              if (!workerOpenRef.current?.workerProposals) installEditorWorkerProposalActivation(next, activateEditorProposals);
            }
            if (workerOpenRef.current?.workerProposals) {
              let laidOut = new Promise<void>((resolve) => {
                workerLaidOutRef.current = resolve;
              });
              const authority = registerWorkerProposalAuthority(next, worker, {
                relayout: () => {
                  if (!authority.initialized) {
                    laidOut = new Promise<void>((resolve) => {
                      workerLaidOutRef.current = resolve;
                    });
                  }
                  markLayoutQueued(next, true);
                  workerOpenRef.current?.refreshWorkerLayout?.();
                },
                current: () => !stale(),
                laidOut: () => laidOut,
                contentChanged: () => workerOpenRef.current?.onWorkerContentChange?.(),
                projectionChanged: markProjectionStories,
                adopted: (version) => {
                  adoptWorkerOpenMirrorVersion(next, version);
                  worker.mirrorReady();
                },
                handedOver: (version) => adoptWorkerOpenHandoverVersion(next, version),
              });
              authority.subscribe(() => {
                if (!stale()) setWorkerProposalsReady(authority.initialized);
              });
            }
          } else {
            if (workerOpenRef.current?.viewer) {
              throw new DocxWorkerError('open', new Error('The document worker is unavailable'));
            }
            host = seedYrsSession(next, (document) => yrs.documentToYrs(next, document), {
              bytes,
              document: seedDocument,
              initialUpdate: collaborationInitialUpdate,
              mediaTokens: mediaTokensRef.current,
            });
          }
        } catch (error) {
          openedWorker?.destroy();
          next.destroy();
          throw error;
        }
        sessionRef.current = next;
        facadeRef.current = yrs;
        if (fullOpenTimer !== null) {
          clearTimeout(fullOpenTimer);
          fullOpenTimer = null;
        }
        if (opened) {
          // Maps and projections of the preview do not describe this session.
          inputPositionMapsRef.current.clear();
          dirtyStoriesRef.current.clear();
          compatibilityBaseRef.current = null;
          compatibilitySourceBufferRef.current = undefined;
          previewingRef.current = false;
          retiringRef.current = opened.session;
          setHandoffFrom(opened.session);
        }
        setSession(next);
        setPreviewing(false);
        setSessionGeneration(seedGeneration);
        if (host) callbacksRef.current?.onHostDocument?.(host, seedGeneration, next);
      })
      .catch((error) => {
        dropEarly();
        console.error('[yrs] failed to start the editing session', error);
        if (abandoned) return;
        fail(error);
      });

    return () => {
      cancelled = true;
      replicaStatePrefetchRef.current?.clear();
      replicaStatePrefetchRef.current = null;
      unregisterSave?.();
      unregisterExport?.();
      unregisterOpeningInput?.();
      openingInputSessionRef.current = null;
      pendingReplicaRef.current?.cancel();
      pendingReplicaRef.current = null;
      startReplicaRef.current = null;
      revisionQueryRef.current = null;
      workerLaidOutRef.current?.();
      workerLaidOutRef.current = null;
      replicaGateRef.current = null;
      requestReplicaRef.current = null;
      openedWorker?.destroy();
      failOpeningRef.current = null;
      if (fullOpenTimer !== null) clearTimeout(fullOpenTimer);
      paintWaitRef.current?.resolve();
      paintWaitRef.current = null;
      if (retiringRef.current !== sessionRef.current) retire(retiringRef.current);
      retiringRef.current = null;
      cancelCompatibilityWarmRef.current?.();
      cancelCompatibilityWarmRef.current = null;
      retire(sessionRef.current);
      sessionRef.current = null;
      facadeRef.current = null;
      inputPositionMapsRef.current.clear();
      dirtyStoriesRef.current.clear();
    };
  }, [
    enabled,
    seedDocument,
    seedBytes,
    seedGeneration,
    collaborationClientId,
    collaborationInitialUpdate,
    openInWorker,
    markProjectionStories,
    retire,
    retirePreview,
  ]);

  useEffect(() => {
    if (
      workerOpen?.viewer ||
      !openInWorker ||
      !session ||
      session !== sessionRef.current ||
      sessionGeneration !== seedGeneration ||
      !workerOpenDocumentHeld(session)
    ) return;
    const pending = releaseWorkerOpenDocument(session);
    if (pending?.pending) setReplicaRequestVersion((version) => version + 1);
  }, [workerOpen?.viewer, openInWorker, session, sessionGeneration, seedGeneration]);

  const hasOwnWorkerFrame = workerOpen?.renderedFrame != null &&
    workerOpen.renderedFrame !== inheritedFrameRef.current;
  useEffect(() => {
    if (
      !openInWorker ||
      !session ||
      session !== sessionRef.current ||
      !hasOwnWorkerFrame ||
      (!workerOpenDocumentHeld(session) &&
        ((!pendingReplicaRef.current?.pending && !hasEditorWorkerProposalRounds(session)) || !startReplicaRef.current)) ||
      previewing ||
      (handoffFrom && options?.shownEngine !== session)
    ) return;
    workerLaidOutRef.current?.();
    const authority = registeredWorkerProposalAuthority(session);
    if (authority) {
      void authority.initialize().catch((error) => {
        console.error('[yrs] failed to initialize worker proposals', error);
      });
    }
  }, [
    openInWorker,
    session,
    hasOwnWorkerFrame,
    workerOpen?.renderedFrame,
    previewing,
    handoffFrom,
    options?.shownEngine,
  ]);

  useEffect(() => {
    if (!openInWorker) return;
    if (
      !session ||
      session !== sessionRef.current ||
      !hasOwnWorkerFrame
    ) return;
    if (previewing || (handoffFrom && options?.shownEngine !== session)) return;
    if (workerOpen?.pendingCompletion === session) return;
    if (workerOpen?.viewer || workerOpenDocumentHeld(session)) revisionQueryRef.current?.();
  }, [
    openInWorker,
    session,
    hasOwnWorkerFrame,
    workerOpen?.pendingCompletion,
    workerOpen?.viewer,
    previewing,
    handoffFrom,
    options?.shownEngine,
  ]);

  useEffect(() => {
    if (!openInWorker) return;
    const pending = pendingReplicaRef.current;
    const prefetch = replicaStatePrefetchRef.current;
    if (
      !session ||
      session !== sessionRef.current ||
      workerOpen?.viewer ||
      workerOpenDocumentHeld(session) ||
      !pending?.pending ||
      pending.started ||
      !startReplicaRef.current ||
      previewing
    ) return;
    const visibilityDocument = globalThis.document;
    const controller = new AbortController();
    let frameId: number | null = null;
    let idleId: number | null = null;
    let timer: ReturnType<typeof setTimeout> | null = null;
    const cleanup = (): void => {
      controller.abort();
      if (eagerPeerStartRef.current === eagerStart) eagerPeerStartRef.current = null;
      clearTimeout(fallbackTimer);
      if (timer !== null) clearTimeout(timer);
      if (frameId !== null) cancelAnimationFrame(frameId);
      if (idleId !== null) cancelIdleCallback(idleId);
      visibilityDocument?.removeEventListener('visibilitychange', onVisibilityChange);
    };
    const startPeer = (): void => {
      if (controller.signal.aborted) return;
      cleanup();
      if (sessionRef.current === session && pendingReplicaRef.current === pending) openReplicaGate();
    };
    const onVisibilityChange = (): void => {
      if (visibilityDocument?.visibilityState === 'hidden') startPeer();
    };
    const fallbackTimer = setTimeout(startPeer, 10_000);
    const eagerStart = { session, start: startPeer };
    eagerPeerStartRef.current = eagerStart;
    visibilityDocument?.addEventListener('visibilitychange', onVisibilityChange);
    if (visibilityDocument?.visibilityState === 'hidden') {
      startPeer();
    } else if (openingInputSessionRef.current === session) {
      startPeer();
    } else if (
      workerOpen?.layoutCompleteSession === session &&
      hasOwnWorkerFrame && retiringRef.current === null &&
      workerOpen?.pendingCompletion !== session &&
      (!handoffFrom || options?.shownEngine === session)
    ) {
      const settled = workerOpenRef.current?.settledDisplayList?.(
        null, null, 'window', controller.signal, () => {
          if (!controller.signal.aborted && retiringRef.current === null) prefetch?.start();
        }
      ) ?? Promise.resolve();
      const authority = registeredWorkerProposalAuthority(session) !== null;
      if (authority) {
        timer = setTimeout(startPeer, REPLICA_FRAME_WAIT_MS);
        if (typeof requestAnimationFrame === 'function') {
          frameId = requestAnimationFrame(() => {
            frameId = requestAnimationFrame(startPeer);
          });
        } else {
          clearTimeout(timer);
          timer = setTimeout(startPeer, 0);
        }
      }
      void settled.then(() => {
        if (controller.signal.aborted || authority) return;
        if (typeof requestIdleCallback === 'function') {
          idleId = requestIdleCallback(startPeer, { timeout: 2000 });
        } else {
          timer = setTimeout(startPeer, 0);
        }
      }, () => {});
    }
    return cleanup;
  }, [
    openInWorker,
    session,
    hasOwnWorkerFrame,
    openReplicaGate,
    replicaRequestVersion,
    workerOpen?.pendingCompletion,
    workerOpen?.layoutCompleteSession,
    workerOpen?.viewer,
    previewing,
    handoffFrom,
    options?.shownEngine,
  ]);

  const requestReplica = useCallback((): void => {
    requestReplicaRef.current?.();
  }, []);

  const notifyFramePresented = useCallback((engine: unknown): void => {
    const waiting = paintWaitRef.current;
    if (waiting && waiting.session === engine) {
      paintWaitRef.current = null;
      waiting.resolve();
    }
    const retiring = retiringRef.current;
    if (retiring && engine === sessionRef.current && engine !== retiring) retirePreview(retiring);
  }, [retirePreview]);

  const failOpening = useCallback(
    (error: Error, session?: unknown): boolean => {
      if (retiringRef.current !== null && (failOpeningRef.current?.(error, session) ?? false)) {
        return true;
      }
      const owner = session ?? sessionRef.current;
      if (owner && workerOpenDocumentHeld(owner as YrsSession)) return false;
      if (owner === sessionRef.current && owner &&
        workerProposalFailure(owner as YrsSession) === error) return false;
      if (session === undefined || session === sessionRef.current) {
        pendingReplicaRef.current?.requestReady();
        startReplicaRef.current?.();
      }
      return false;
    },
    []
  );

  // Save, export and getDocument materialize the base on first use; only a
  // host projecting every change asks for it ahead of the first edit.
  const scheduleCompatibilityWarm = useCallback((): void => {
    const live = sessionRef.current;
    if (
      !enabledRef.current ||
      !live ||
      (workerOpenEnabledRef.current && workerOpenReplicaPending(live)) ||
      live.isDisplayOnly() ||
      !seedBytesRef.current ||
      compatibilityBaseRef.current ||
      cancelCompatibilityWarmRef.current
    ) {
      return;
    }
    const warm = (): void => {
      cancelCompatibilityWarmRef.current = null;
      if (sessionRef.current === live) warmCompatibilityBase(live, compatibilityBaseRef);
    };
    if (typeof requestIdleCallback === 'function') {
      const id = requestIdleCallback(warm);
      cancelCompatibilityWarmRef.current = () => cancelIdleCallback(id);
      return;
    }
    const id = setTimeout(warm, 200);
    cancelCompatibilityWarmRef.current = () => clearTimeout(id);
  }, []);

  const cancelCompatibilityWarm = useCallback((): void => {
    cancelCompatibilityWarmRef.current?.();
    cancelCompatibilityWarmRef.current = null;
  }, []);

  useEffect(() => {
    const onReplica = collaboration?.onReplica;
    // A preview is display-only, never a replica; the full session follows.
    if (
      !onReplica ||
      !session ||
      session.isDisplayOnly() ||
      (openInWorker && !replicaReady)
    ) {
      return;
    }
    onReplica(session);
    return () => onReplica(null);
  }, [collaboration?.onReplica, openInWorker, replicaReady, session]);

  const storyBlocks = useCallback((storyId: string, env: YrsRenderEnv): LayoutBlock[] | null => {
    if (!enabledRef.current) return null;
    try {
      const live = sessionRef.current;
      if (!live || !live.hasStory(storyId)) return null;
      return live.yrsBlocksForStory(storyId, env) as LayoutBlock[];
    } catch (error) {
      console.error(`[yrs] failed to lower story ${storyId}`, error);
      return null;
    }
  }, []);

  const bodyBlocks = useCallback(
    (env: YrsRenderEnv): LayoutBlock[] | null => storyBlocks('body', env),
    [storyBlocks]
  );

  const inputPositionMap = useCallback((storyId = 'body'): YrsInputPositionMap | null => {
    const live = sessionRef.current;
    const facade = facadeRef.current;
    if (!enabledRef.current || !live || !facade || !live.hasStory(storyId)) return null;
    const cached = inputPositionMapsRef.current.get(storyId);
    if (cached) return cached;
    const map = facade.createYrsInputPositionMap(storyId, live.paragraphSpans(storyId));
    inputPositionMapsRef.current.set(storyId, map);
    return map;
  }, []);

  const displayPositionToLoc = useCallback(
    (position: number, storyId = 'body'): YrsLoc | null => {
      const facade = facadeRef.current;
      const map = inputPositionMap(storyId);
      return facade && map ? facade.displayPositionToYrsLoc(map, position) : null;
    },
    [inputPositionMap]
  );

  const locToDisplayPosition = useCallback(
    (loc: YrsLoc): number | null => {
      const facade = facadeRef.current;
      const map = inputPositionMap(loc.story);
      return facade && map ? facade.yrsLocToDisplayPosition(map, loc) : null;
    },
    [inputPositionMap]
  );

  const documentFromYrs = useCallback((baseDocument?: Document | null): Document | null => {
    const live = sessionRef.current;
    const facade = facadeRef.current;
    const host = baseDocument === undefined ? documentRef.current : baseDocument;
    let base = host;
    // A preview holds only the first pages: nothing saves or exports it.
    if (!enabledRef.current || previewingRef.current || !live || !facade || !base ||
      workerOpenDocumentHeld(live)) return null;
    try {
      if (workerOpenEnabledRef.current && workerOpenReplicaPending(live)) return null;
      const compatibilityBase = compatibilityBaseRef.current ?? live.materializeDocx();
      if (compatibilityBase) {
        compatibilitySourceBufferRef.current ??= compatibilityBase.originalBuffer;
        base = mergeDocxHostMetadata(compatibilityBase, base);
      }
      const dirtyStories = dirtyStoriesRef.current;
      const projected = facade.yrsToDocument(
        live,
        base,
        dirtyStories.projection.projectionOptions()
      );
      dirtyStories.projected();
      if (compatibilityBase) compatibilityBaseRef.current = projected;
      return projected;
    } catch (error) {
      console.error('[yrs] failed to project the document for save', error);
      return null;
    }
  }, []);

  const publishDirectInput = useCallback((stories?: string | readonly string[]): void => {
    const live = sessionRef.current;
    if (!live || !live.hasStory('body')) return;
    inputPositionMapsRef.current.clear();
    const dirty =
      stories === undefined
        ? [live.selection()?.head.story ?? 'body']
        : typeof stories === 'string'
          ? [stories]
          : stories;
    markProjectionStories(dirty);
  }, [markProjectionStories]);

  return {
    session,
    sessionGeneration,
    replicaReady: !openInWorker || replicaReady,
    requestReplica,
    workerProposalsReady,
    replicaReadyRef: openInWorker ? replicaReadyRef : undefined,
    experimentalWorkerOpen: Boolean(openInWorker),
    previewing,
    handoffFrom,
    opening: previewing || handoffFrom !== null,
    notifyFramePresented,
    failOpening,
    storyBlocks,
    bodyBlocks,
    inputPositionMap,
    displayPositionToLoc,
    locToDisplayPosition,
    documentFromYrs,
    publishDirectInput,
    scheduleCompatibilityWarm,
    cancelCompatibilityWarm,
  };
}

/**
 * Warms the compatibility base for a host that projects every change, once
 * the session's own first display list is on screen. A replacement session
 * can inherit the previous session's frame until its own layout lands, so the
 * frame shown when the session changed never qualifies it. A renderer leaving
 * readiness, or the host losing its last content listener, cancels a pending
 * warm.
 */
export function useCompatibilityWarm(
  session: YrsSession | null,
  renderedFrame: object | null,
  projectsEveryChange: boolean,
  schedule: () => void,
  cancel: () => void
): void {
  const inheritedRef = useRef<{ session: YrsSession | null; frame: object | null }>({
    session: null,
    frame: null,
  });
  if (inheritedRef.current.session !== session) {
    inheritedRef.current = { session, frame: renderedFrame };
  }
  const ownFrame = renderedFrame !== null && renderedFrame !== inheritedRef.current.frame;
  useEffect(() => {
    if (!ownFrame || !projectsEveryChange) {
      cancel();
      return;
    }
    if (session) schedule();
  }, [cancel, ownFrame, projectsEveryChange, schedule, session]);
}
