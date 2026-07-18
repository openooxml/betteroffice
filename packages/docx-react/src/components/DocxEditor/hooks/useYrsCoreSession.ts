import { useCallback, useEffect, useRef, useState } from 'react';
import type { LayoutBlock } from '@betteroffice/docx/layout/pagination';
import type { Document } from '@betteroffice/docx/types/document';
import type {
  YrsInputPositionMap,
  YrsLoc,
  YrsRenderEnv,
  YrsSession,
} from '@betteroffice/docx/yrs';
import { tracePerfSync } from '../internals/perfTrace';

type YrsFacadeModule = typeof import('@betteroffice/docx/yrs');

/** The React editor's sole mutable document session. */
export interface YrsCoreSession {
  session: YrsSession | null;
  storyBlocks(storyId: string, env: YrsRenderEnv): LayoutBlock[] | null;
  bodyBlocks(env: YrsRenderEnv): LayoutBlock[] | null;
  inputPositionMap(storyId?: string): YrsInputPositionMap | null;
  displayPositionToLoc(position: number, storyId?: string): YrsLoc | null;
  locToDisplayPosition(loc: YrsLoc): number | null;
  documentFromYrs(): Document | null;
  publishDirectInput(): void;
}

interface YrsWatchdogGlobal {
  divergences: number;
  checksum: string | null;
}

interface YrsRenderGlobal {
  fromYrs: number;
  fallback: number;
}

declare global {
  // eslint-disable-next-line no-var
  var __yrsWatchdog: YrsWatchdogGlobal | undefined;
  // eslint-disable-next-line no-var
  var __yrsRender: YrsRenderGlobal | undefined;
}

export function useYrsCoreSession(
  enabled: boolean,
  document: Document | null,
  seedDocument: Document | null
): YrsCoreSession {
  const sessionRef = useRef<YrsSession | null>(null);
  const facadeRef = useRef<YrsFacadeModule | null>(null);
  const documentRef = useRef(document);
  documentRef.current = document;
  const inputPositionMapsRef = useRef(new Map<string, YrsInputPositionMap>());
  const projectionStoriesRef = useRef(new Set<string>());
  const enabledRef = useRef(enabled);
  enabledRef.current = enabled;
  const [session, setSession] = useState<YrsSession | null>(null);

  useEffect(() => {
    if (!enabled || !seedDocument) return;
    let cancelled = false;
    setSession(null);
    inputPositionMapsRef.current.clear();
    projectionStoriesRef.current.clear();
    globalThis.__yrsRender = { fromYrs: 0, fallback: 0 };
    globalThis.__yrsWatchdog = { divergences: 0, checksum: null };

    void import('@betteroffice/docx/yrs')
      .then(async (yrs) => {
        const next = await yrs.createYrsSession();
        if (cancelled) {
          next.destroy();
          return;
        }
        yrs.documentToYrs(next, seedDocument);
        sessionRef.current = next;
        facadeRef.current = yrs;
        globalThis.__yrsWatchdog = {
          divergences: 0,
          checksum: next.storyChecksum('body').toString(),
        };
        setSession(next);
      })
      .catch((error) => {
        console.error('[yrs] failed to start the editing session', error);
      });

    return () => {
      cancelled = true;
      sessionRef.current?.destroy();
      sessionRef.current = null;
      facadeRef.current = null;
      inputPositionMapsRef.current.clear();
      projectionStoriesRef.current.clear();
      globalThis.__yrsWatchdog = undefined;
      globalThis.__yrsRender = undefined;
    };
  }, [enabled, seedDocument]);

  const storyBlocks = useCallback((storyId: string, env: YrsRenderEnv): LayoutBlock[] | null => {
    if (!enabledRef.current) return null;
    try {
      const live = sessionRef.current;
      if (!live || !live.storyIds().includes(storyId)) {
        bumpRenderCounter('fallback');
        return null;
      }
      const blocks = live.yrsBlocksForStory(storyId, env) as LayoutBlock[];
      bumpRenderCounter('fromYrs');
      return blocks;
    } catch (error) {
      bumpRenderCounter('fallback');
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
    if (!enabledRef.current || !live || !facade || !live.storyIds().includes(storyId)) return null;
    const cached = inputPositionMapsRef.current.get(storyId);
    if (cached) return cached;
    return tracePerfSync(
      'selectionReads',
      () => {
        const map = facade.createYrsInputPositionMap(storyId, live.paragraphSpans(storyId));
        inputPositionMapsRef.current.set(storyId, map);
        return map;
      },
      { calls: 1, detail: 'inputPositionMapMiss' }
    );
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

  const documentFromYrs = useCallback((): Document | null => {
    const live = sessionRef.current;
    const facade = facadeRef.current;
    const base = documentRef.current;
    if (!enabledRef.current || !live || !facade || !base) return null;
    try {
      const dirtyStories = projectionStoriesRef.current;
      const projected = facade.yrsToDocument(
        live,
        base,
        dirtyStories.size > 0 ? { storyIds: new Set(dirtyStories) } : undefined
      );
      dirtyStories.clear();
      return projected;
    } catch (error) {
      console.error('[yrs] failed to project the document for save', error);
      return null;
    }
  }, []);

  const publishDirectInput = useCallback((): void => {
    const live = sessionRef.current;
    const watchdog = globalThis.__yrsWatchdog;
    if (!live || !watchdog || !live.storyIds().includes('body')) return;
    inputPositionMapsRef.current.clear();
    const activeStory = live.selection()?.head.story ?? 'body';
    projectionStoriesRef.current.add(
      activeStory.startsWith('hf:') || activeStory.startsWith('fn:') ? activeStory : 'body'
    );
    watchdog.checksum = live.storyChecksum('body').toString();
  }, []);

  return {
    session,
    storyBlocks,
    bodyBlocks,
    inputPositionMap,
    displayPositionToLoc,
    locToDisplayPosition,
    documentFromYrs,
    publishDirectInput,
  };
}

function bumpRenderCounter(key: keyof YrsRenderGlobal): void {
  const counter = globalThis.__yrsRender ?? { fromYrs: 0, fallback: 0 };
  counter[key] += 1;
  globalThis.__yrsRender = counter;
}
