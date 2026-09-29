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
  /** A rebuild after a content change never waits. */
  rebuildAtOnce: boolean;
  /**
   * Counts in-place changes that are content changes for this chrome. Other
   * in-place changes (position shifts, see `displayPageRevision`) rebuild it
   * at idle time.
   */
  urgentRevision: number;
  /** Receives the handle that builds the chrome at once, for a page needed before it is active. */
  register?: (handle: PageChromeHandle | null) => void;
  make: (page: DisplayPage, t: TFunction) => HTMLElement;
}

/** Builds one page's chrome outside its render cycle. */
export interface PageChromeHandle {
  /** Builds the chrome for the current page now, unless it already shows it. */
  build(): void;
  /** Whether the chrome shows the current page. */
  current(): boolean;
  /** Clears chrome built while the page is inactive. */
  release(): void;
}

interface BuiltFor {
  page: DisplayPage;
  revision: number;
  urgentRevision: number;
  t: TFunction;
}

/** Builds one page's mirror or overlay into `hostRef`, and rebuilds it with the page. */
export function usePageChrome(
  hostRef: RefObject<HTMLDivElement | null>,
  { page, t, active, defer, rebuildAtOnce, urgentRevision, register, make }: PageChromeOptions
): void {
  // Owned deltas shift primitive positions in place: identity alone is stale.
  const revision = displayPageRevision(page);
  // Read when a build is scheduled: a change of scheduling alone keeps the
  // built DOM (and any focus inside it) in place.
  const scheduling = useRef({ defer, rebuildAtOnce });
  scheduling.current = { defer, rebuildAtOnce };
  const latest = useRef<BuiltFor & { make: PageChromeOptions['make'] }>({
    page,
    revision,
    urgentRevision,
    t,
    make,
  });
  latest.current = { page, revision, urgentRevision, t, make };
  const activeRef = useRef(active);
  activeRef.current = active;
  const builtForRef = useRef<BuiltFor | null>(null);
  const current = (): boolean => {
    const built = builtForRef.current;
    const wanted = latest.current;
    return (
      built?.page === wanted.page &&
      built.revision === displayPageRevision(wanted.page) &&
      built.urgentRevision === wanted.urgentRevision &&
      built.t === wanted.t
    );
  };
  const build = useCallback((): void => {
    const host = hostRef.current;
    if (!host || current()) return;
    const { page, urgentRevision, t, make } = latest.current;
    // Keep the previous chrome connected until its replacement is ready.
    host.replaceChildren(make(page, t));
    builtForRef.current = { page, revision: displayPageRevision(page), urgentRevision, t };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [hostRef]);

  useEffect(() => {
    register?.({
      build,
      current,
      release: () => {
        const host = hostRef.current;
        if (!host || activeRef.current || !builtForRef.current) return;
        host.replaceChildren();
        builtForRef.current = null;
      },
    });
    return () => register?.(null);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [build, hostRef, register]);

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
    const built = builtForRef.current;
    const { defer, rebuildAtOnce } = scheduling.current;
    const changed =
      built !== null &&
      (built.page !== page || built.t !== t || built.urgentRevision !== urgentRevision);
    if (built ? changed && rebuildAtOnce : !defer) {
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
  }, [page, t, revision, urgentRevision, active, build]);
}
