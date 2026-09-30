import { useCallback, useEffect, useRef, useState } from 'react';
import type { LayoutBlock } from '@betteroffice/docx/layout/pagination';
import type {
  Document,
  Endnote,
  Footnote,
  HeaderFooter,
  Section,
} from '@betteroffice/docx/types/document';
import type {
  YrsDocxHost,
  YrsInputPositionMap,
  YrsLoc,
  YrsRenderEnv,
  YrsSession,
} from '@betteroffice/docx/yrs';
import type { DocxEditorCollaborationOptions } from '../types';

type YrsFacadeModule = typeof import('@betteroffice/docx/yrs');

/** The React editor's sole mutable document session. */
export interface YrsCoreSession {
  session: YrsSession | null;
  /** The seed generation `session` was created for. */
  sessionGeneration: number | null;
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
}

/** Body blocks a first-page preview parses. */
const PREVIEW_BODY_BLOCKS = 200;
/** How long the full open waits for the preview's pages to paint. */
const PREVIEW_PAINT_TIMEOUT_MS = 2000;
/** Bounds the wait for the painted preview to reach the screen; hidden tabs get no frames. */
const PREVIEW_FRAME_WAIT_MS = 100;
/**
 * How long a preview waits for the full session, from the end of its own
 * paint. A full open that has not produced one by then fails the load; once
 * it exists, the preview stays until the full session's first frame shows.
 */
const FULL_OPEN_TIMEOUT_MS = 10_000;

function mergeHeaderFooterMaps(
  full: Map<string, HeaderFooter> | undefined,
  host: Map<string, HeaderFooter> | undefined
): Map<string, HeaderFooter> | undefined {
  if (host === undefined) return undefined;
  return new Map(
    [...host].map(([relationshipId, metadata]) => {
      const existing = full?.get(relationshipId);
      return [relationshipId, existing ? { ...metadata, content: existing.content } : metadata];
    })
  );
}

function mergeNotes<T extends Footnote | Endnote>(
  full: T[] | undefined,
  host: T[] | undefined
): T[] | undefined {
  if (host === undefined) return undefined;
  return host.map((metadata) => {
    const existing = full?.find((note) => note.id === metadata.id);
    return existing ? { ...existing, ...metadata, content: existing.content } : metadata;
  });
}

function mergeSections(
  full: Section[] | undefined,
  host: Section[] | undefined
): Section[] | undefined {
  if (host === undefined) return undefined;
  return host.map((metadata, index) => {
    const existing =
      full?.find((section) => section.id !== undefined && section.id === metadata.id) ??
      full?.[index];
    return existing ? { ...metadata, content: existing.content } : metadata;
  });
}

export function mergeDocxHostMetadata(full: Document, host: Document): Document {
  const fullPackage = full.package;
  const hostPackage = host.package;
  return {
    ...full,
    contractVersion: host.contractVersion ?? full.contractVersion,
    originalBuffer: full.originalBuffer ?? host.originalBuffer,
    warnings: host.warnings,
    package: {
      ...fullPackage,
      contractVersion: hostPackage.contractVersion ?? fullPackage.contractVersion,
      styles: hostPackage.styles,
      theme: hostPackage.theme,
      settings: hostPackage.settings,
      fontTable: hostPackage.fontTable,
      relationships: hostPackage.relationships,
      headers: mergeHeaderFooterMaps(fullPackage.headers, hostPackage.headers),
      footers: mergeHeaderFooterMaps(fullPackage.footers, hostPackage.footers),
      footnotes: mergeNotes(fullPackage.footnotes, hostPackage.footnotes),
      endnotes: mergeNotes(fullPackage.endnotes, hostPackage.endnotes),
      document: {
        ...fullPackage.document,
        sections: mergeSections(fullPackage.document.sections, hostPackage.document.sections),
        finalSectionProperties: hostPackage.document.finalSectionProperties,
        comments: hostPackage.document.comments,
      },
    },
  };
}

/** A display-only session of the first pages of `bytes`, or null when it cannot open. */
async function openPreview(
  yrs: YrsFacadeModule,
  bytes: Uint8Array,
  clientId: number | undefined
): Promise<{ session: YrsSession; host: YrsDocxHost } | null> {
  const session = await yrs.createYrsSession({ clientId });
  try {
    const host = session.openDocxPreview(bytes, PREVIEW_BODY_BLOCKS);
    if (host) return { session, host };
    session.destroy();
    return null;
  } catch (error) {
    console.warn('[yrs] the first-page preview could not open; opening in full', error);
    session.destroy();
    return null;
  }
}

export interface YrsSeedSources {
  bytes: Uint8Array | null;
  document: Document | null;
  initialUpdate?: Uint8Array;
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
  const { bytes, document, initialUpdate } = seed;
  if (bytes) {
    const host = session.openDocx(bytes, !initialUpdate);
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

/** Story a direct-input edit dirties: the hf/note root it sits in, everything else the body. */
export function dirtyProjectionStory(activeStory: string): string {
  return ['hf:', 'fn:', 'en:'].some((prefix) => activeStory.startsWith(prefix))
    ? activeStory.split(':', 2).join(':')
    : 'body';
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
  const collaborationClientId = collaboration?.clientId;
  const collaborationInitialUpdate = collaboration?.initialUpdate;
  const sessionRef = useRef<YrsSession | null>(null);
  const facadeRef = useRef<YrsFacadeModule | null>(null);
  const documentRef = useRef(document);
  documentRef.current = document;
  const callbacksRef = useRef(callbacks);
  callbacksRef.current = callbacks;
  const compatibilityBaseRef = useRef<Document | null>(null);
  const cancelCompatibilityWarmRef = useRef<(() => void) | null>(null);
  const seedBytesRef = useRef(seedBytes);
  seedBytesRef.current = seedBytes;
  const inputPositionMapsRef = useRef(new Map<string, YrsInputPositionMap>());
  const projectionStoriesRef = useRef(new Set<string>());
  const enabledRef = useRef(enabled);
  enabledRef.current = enabled;
  const [session, setSession] = useState<YrsSession | null>(null);
  const [sessionGeneration, setSessionGeneration] = useState<number | null>(null);
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
    setPreviewing(false);
    setHandoffFrom(null);
    if (!enabled || (!seedDocument && !seedBytes)) return;
    let cancelled = false;
    inputPositionMapsRef.current.clear();
    projectionStoriesRef.current.clear();
    compatibilityBaseRef.current = null;

    let abandoned = false;
    let shown: { session: YrsSession; host: YrsDocxHost } | null = null;
    let fullOpenTimer: ReturnType<typeof setTimeout> | null = null;
    const stale = () =>
      cancelled ||
      abandoned ||
      callbacksRef.current?.isCurrentLoad?.(seedGeneration) === false;
    const previewFirstPage = previewFirstPageRef.current;
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
      if (shown) dropPreview(shown.session);
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
        const prepared = seedBytes ? yrs.prepareDocxBytes(seedBytes) : Promise.resolve(null);
        // Awaited below unless the load ends first.
        prepared.catch(() => {});
        // A preview paints the first pages first; the full open, which
        // blocks this thread for the whole package, waits until they have.
        const opened =
          previewFirstPage && seedBytes
            ? await openPreview(yrs, seedBytes, collaborationClientId)
            : null;
        shown = opened;
        if (opened && stale()) {
          opened.session.destroy();
          return;
        }
        if (opened) {
          callbacksRef.current?.onSession?.(opened.session);
          const painted = new Promise<void>((resolve) => {
            paintWaitRef.current = { session: opened.session, resolve };
            setTimeout(resolve, PREVIEW_PAINT_TIMEOUT_MS);
          });
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
          // after the presentation, and the full open blocks this thread.
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
          if (stale()) return;
          const preview = opened.session;
          fullOpenTimer = setTimeout(() => {
            fullOpenTimer = null;
            if (sessionRef.current !== preview) return;
            abandoned = true;
            fail(new Error('The document did not finish opening in time'));
          }, fullOpenTimeoutRef.current);
        }
        const bytes = await prepared;
        if (stale()) return;
        const next = await yrs.createYrsSession({ clientId: collaborationClientId });
        if (stale()) {
          next.destroy();
          return;
        }
        callbacksRef.current?.onSession?.(next);
        let host: YrsDocxHost | null;
        try {
          host = seedYrsSession(next, (document) => yrs.documentToYrs(next, document), {
            bytes,
            document: seedDocument,
            initialUpdate: collaborationInitialUpdate,
          });
        } catch (error) {
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
          projectionStoriesRef.current.clear();
          compatibilityBaseRef.current = null;
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
        console.error('[yrs] failed to start the editing session', error);
        if (abandoned) return;
        fail(error);
      });

    return () => {
      cancelled = true;
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
      projectionStoriesRef.current.clear();
    };
  }, [
    enabled,
    seedDocument,
    seedBytes,
    seedGeneration,
    collaborationClientId,
    collaborationInitialUpdate,
    retire,
    retirePreview,
  ]);

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
    (error: Error, session?: unknown): boolean =>
      retiringRef.current !== null && (failOpeningRef.current?.(error, session) ?? false),
    []
  );

  // Save, export and getDocument materialize the base on first use; only a
  // host projecting every change asks for it ahead of the first edit.
  const scheduleCompatibilityWarm = useCallback((): void => {
    const live = sessionRef.current;
    if (
      !enabledRef.current ||
      !live ||
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
    if (!onReplica || !session || session.isDisplayOnly()) return;
    onReplica(session);
    return () => onReplica(null);
  }, [collaboration?.onReplica, session]);

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
    if (!enabledRef.current || previewingRef.current || !live || !facade || !base) return null;
    try {
      const compatibilityBase = compatibilityBaseRef.current ?? live.materializeDocx();
      if (compatibilityBase) {
        base = mergeDocxHostMetadata(compatibilityBase, base);
      }
      const dirtyStories = projectionStoriesRef.current;
      const projected = facade.yrsToDocument(
        live,
        base,
        dirtyStories.size > 0 ? { storyIds: new Set(dirtyStories) } : undefined
      );
      dirtyStories.clear();
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
    for (const story of dirty) projectionStoriesRef.current.add(dirtyProjectionStory(story));
  }, []);

  return {
    session,
    sessionGeneration,
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
