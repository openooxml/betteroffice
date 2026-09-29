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
   * Materializes the save projection base when the main thread is next idle,
   * for a host that projects the document on every change.
   */
  scheduleCompatibilityWarm(): void;
  cancelCompatibilityWarm(): void;
  /**
   * `session` is a display-only preview of the document's first pages; the
   * full session replaces it once they have painted.
   */
  previewing: boolean;
  /** The preview session whose pages the renderer keeps until the full session's replace them. */
  handoffFrom: YrsSession | null;
  /** A frame of `engine`'s layout was published for display. */
  notifyFramePresented(engine: unknown): void;
}

interface YrsCoreSessionCallbacks {
  isCurrentLoad?: (generation: number) => boolean;
  onHostDocument?: (
    host: YrsDocxHost,
    generation: number,
    options?: { preview: boolean }
  ) => void;
  onError?: (error: Error, generation: number) => void;
}

export interface YrsCoreSessionOptions {
  /** Open a display-only preview of the first pages before the full document. */
  previewFirstPage?: boolean;
}

/** Body blocks a first-page preview parses. */
const PREVIEW_BODY_BLOCKS = 200;
/** How long the full open waits for the preview's pages to paint. */
const PREVIEW_PAINT_TIMEOUT_MS = 2000;

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
    return { session, host: session.openDocxPreview(bytes, PREVIEW_BODY_BLOCKS) };
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
  const retiringRef = useRef<YrsSession | null>(null);
  // Collaboration shares one replica from the start, so it never previews.
  const previewFirstPage =
    options?.previewFirstPage === true && !collaboration && !collaborationInitialUpdate;

  useEffect(() => {
    setSession(null);
    setPreviewing(false);
    setHandoffFrom(null);
    if (!enabled || (!seedDocument && !seedBytes)) return;
    let cancelled = false;
    inputPositionMapsRef.current.clear();
    projectionStoriesRef.current.clear();
    compatibilityBaseRef.current = null;

    const stale = () =>
      cancelled || callbacksRef.current?.isCurrentLoad?.(seedGeneration) === false;

    void import('@betteroffice/docx/yrs')
      .then(async (yrs) => {
        // A preview paints the first pages first; the full open, which
        // blocks this thread for the whole package, waits until they have.
        const shown =
          previewFirstPage && seedBytes
            ? await openPreview(yrs, seedBytes, collaborationClientId)
            : null;
        if (shown && stale()) {
          shown.session.destroy();
          return;
        }
        if (shown) {
          const painted = new Promise<void>((resolve) => {
            paintWaitRef.current = { session: shown.session, resolve };
            setTimeout(resolve, PREVIEW_PAINT_TIMEOUT_MS);
          });
          sessionRef.current = shown.session;
          facadeRef.current = yrs;
          setSession(shown.session);
          setPreviewing(true);
          setSessionGeneration(seedGeneration);
          callbacksRef.current?.onHostDocument?.(shown.host, seedGeneration, { preview: true });
          await painted;
          await new Promise<void>((resolve) =>
            requestAnimationFrame(() => setTimeout(resolve, 0))
          );
          if (stale()) return;
        }
        const next = await yrs.createYrsSession({ clientId: collaborationClientId });
        if (stale()) {
          next.destroy();
          return;
        }
        const host = seedYrsSession(next, (document) => yrs.documentToYrs(next, document), {
          bytes: seedBytes,
          document: seedDocument,
          initialUpdate: collaborationInitialUpdate,
        });
        sessionRef.current = next;
        facadeRef.current = yrs;
        if (shown) {
          retiringRef.current = shown.session;
          setHandoffFrom(shown.session);
        }
        setSession(next);
        setPreviewing(false);
        setSessionGeneration(seedGeneration);
        if (host) callbacksRef.current?.onHostDocument?.(host, seedGeneration);
      })
      .catch((error) => {
        console.error('[yrs] failed to start the editing session', error);
        if (
          !cancelled &&
          callbacksRef.current?.isCurrentLoad?.(seedGeneration) !== false
        ) {
          callbacksRef.current?.onError?.(
            error instanceof Error ? error : new Error(String(error)),
            seedGeneration
          );
        }
      });

    return () => {
      cancelled = true;
      cancelCompatibilityWarmRef.current?.();
      cancelCompatibilityWarmRef.current = null;
      paintWaitRef.current?.resolve();
      paintWaitRef.current = null;
      if (retiringRef.current !== sessionRef.current) retiringRef.current?.destroy();
      retiringRef.current = null;
      sessionRef.current?.destroy();
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
    previewFirstPage,
  ]);

  const notifyFramePresented = useCallback((engine: unknown): void => {
    const waiting = paintWaitRef.current;
    if (waiting && waiting.session === engine) {
      paintWaitRef.current = null;
      waiting.resolve();
    }
    const retiring = retiringRef.current;
    if (retiring && engine === sessionRef.current && engine !== retiring) {
      retiringRef.current = null;
      setHandoffFrom(null);
      // Components let go of the preview on this commit.
      setTimeout(() => retiring.destroy(), 0);
    }
  }, []);

  // Save, export and getDocument materialize the base on first use; only a
  // host projecting every change asks for it ahead of the first edit.
  const scheduleCompatibilityWarm = useCallback((): void => {
    const live = sessionRef.current;
    if (
      !enabledRef.current ||
      !live ||
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
    if (!onReplica || !session) return;
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
    if (!enabledRef.current || !live || !facade || !base) return null;
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
    notifyFramePresented,
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
