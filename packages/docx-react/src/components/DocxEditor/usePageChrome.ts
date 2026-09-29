import { useCallback, useEffect, useRef, type RefObject } from 'react';
import { displayPageRevision, type DisplayPage } from '@betteroffice/docx/layout/render';
import type { TFunction } from '@betteroffice/docx-i18n';

export interface PageChromeOptions {
  page: DisplayPage;
  t: TFunction;
  /** Holds the chrome; an inactive page keeps only its empty host. */
  active: boolean;
  /** A first build may wait for idle time. */
  defer: boolean;
  /** A rebuild of chrome already shown never waits. */
  rebuildAtOnce: boolean;
  /** Receives a function that builds the chrome at once, for a page needed before it is active. */
  registerBuild?: (build: (() => void) | null) => void;
  make: (page: DisplayPage, t: TFunction) => HTMLElement;
}

/** Builds one page's mirror or overlay into `hostRef`, and rebuilds it with the page. */
export function usePageChrome(
  hostRef: RefObject<HTMLDivElement | null>,
  { page, t, active, defer, rebuildAtOnce, registerBuild, make }: PageChromeOptions
): void {
  // Read when a build is scheduled: a change of scheduling alone keeps the
  // built DOM (and any focus inside it) in place.
  const scheduling = useRef({ defer, rebuildAtOnce });
  scheduling.current = { defer, rebuildAtOnce };
  const latest = useRef({ page, t, make });
  latest.current = { page, t, make };
  // Position-shift deltas mutate primitives in place — identity alone is stale.
  const builtForRef = useRef<{ page: DisplayPage; revision: number; t: TFunction } | null>(null);
  const current = (): boolean => {
    const built = builtForRef.current;
    const { page, t } = latest.current;
    return built?.page === page && built.revision === displayPageRevision(page) && built.t === t;
  };
  const build = useCallback((): void => {
    const host = hostRef.current;
    if (!host || current()) return;
    const { page, t, make } = latest.current;
    // Keep the previous chrome connected until its replacement is ready.
    host.replaceChildren(make(page, t));
    builtForRef.current = { page, revision: displayPageRevision(page), t };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [hostRef]);

  useEffect(() => {
    registerBuild?.(build);
    return () => registerBuild?.(null);
  }, [build, registerBuild]);

  useEffect(() => {
    const host = hostRef.current;
    if (!host) return;
    if (!active) {
      if (builtForRef.current) {
        host.replaceChildren();
        builtForRef.current = null;
      }
      return;
    }
    if (current()) return;
    const { defer, rebuildAtOnce } = scheduling.current;
    if (builtForRef.current ? rebuildAtOnce : !defer) {
      build();
      return;
    }
    if (typeof requestIdleCallback === 'function') {
      const id = requestIdleCallback(build, { timeout: 1500 });
      return () => cancelIdleCallback(id);
    }
    const id = setTimeout(build, 150);
    return () => clearTimeout(id);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [page, t, active, build]);
}
