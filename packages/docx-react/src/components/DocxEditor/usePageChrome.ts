import { useCallback, useEffect, useRef, type RefObject } from 'react';
import { displayPageRevision, type DisplayPage } from '@betteroffice/docx/layout/render';
import type { TFunction } from '@betteroffice/docx-i18n';

type MakeChrome = (page: DisplayPage, t: TFunction) => HTMLElement;
/** `chrome` is the page's current chrome, if built for this page, which it may take over. */
type MakeFallback = (page: DisplayPage, t: TFunction, chrome: HTMLElement | null) => HTMLElement;

export interface PageChromeOptions {
  page: DisplayPage;
  t: TFunction;
  /** Holds the chrome; an inactive page holds only its `fallback`, if any. */
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
  make: MakeChrome;
  /** What an inactive page shows, built at idle time; null for nothing. */
  fallback?: MakeFallback | null;
}

/** Builds one page's chrome outside its render cycle. */
export interface PageChromeHandle {
  /** Builds the chrome for the current page now, unless it already shows it. */
  build(): void;
  /** Takes chrome built while the page is inactive back to its fallback. */
  release(): void;
}

interface BuiltFor {
  kind: 'chrome' | 'fallback';
  page: DisplayPage;
  revision: number;
  urgentRevision: number;
  t: TFunction;
}

const TAB_STOPS = 'a[href], button, input, select, textarea, [tabindex]';

/** The elements under `root` that Tab stops at, in document order. */
function tabStops(root: ParentNode): HTMLElement[] {
  return Array.from(root.querySelectorAll<HTMLElement>(TAB_STOPS)).filter(
    (element) => element.tabIndex >= 0 && !(element as HTMLButtonElement).disabled
  );
}

/**
 * What identifies a tab stop across a rebuild, whatever changed around it;
 * null when only its place does. A content control's button shows its state
 * and its group is its position, so only its control id names it.
 */
const stopKey = (stop: HTMLElement): string | null => {
  const { sdtGroupId, sdtControlId, sdtWidget, sdtRepeat } = stop.dataset;
  if (stop.tagName === 'BUTTON' && sdtGroupId !== undefined) {
    return sdtControlId === undefined
      ? null
      : ['control', sdtControlId, sdtWidget, sdtRepeat].join('\u0000');
  }
  return [stop.tagName, stop.getAttribute('href'), stop.id, stop.textContent].join('\u0000');
};

/**
 * Replaces `host`'s content with `next`. Focus inside it moves to the same
 * stop in `next`, or to the stop at the same place when that one is gone.
 */
function replaceKeepingFocus(host: HTMLElement, next: HTMLElement | null): void {
  if (next && host.childNodes.length === 1 && host.firstChild === next) return;
  const focused = host.ownerDocument.activeElement;
  const stops = focused instanceof HTMLElement && host.contains(focused) ? tabStops(host) : [];
  const at = stops.indexOf(focused as HTMLElement);
  const key = at >= 0 ? stopKey(stops[at]!) : null;
  const nth = stops.slice(0, at).filter((stop) => stopKey(stop) === key).length;
  if (next) host.replaceChildren(next);
  else host.replaceChildren();
  if (at < 0) return;
  const rebuilt = tabStops(host);
  const same = key === null ? [] : rebuilt.filter((stop) => stopKey(stop) === key);
  (same[nth] ?? rebuilt[at])?.focus({ preventScroll: true });
}

/** Builds one page's mirror or overlay into `hostRef`, and rebuilds it with the page. */
export function usePageChrome(
  hostRef: RefObject<HTMLDivElement | null>,
  {
    page,
    t,
    active,
    defer,
    rebuildAtOnce,
    urgentRevision,
    register,
    make,
    fallback = null,
  }: PageChromeOptions
): void {
  // Owned deltas shift primitive positions in place: identity alone is stale.
  const revision = displayPageRevision(page);
  // Read when a build is scheduled: a change of scheduling alone keeps the
  // built DOM (and any focus inside it) in place.
  const scheduling = useRef({ defer, rebuildAtOnce });
  scheduling.current = { defer, rebuildAtOnce };
  const latest = useRef({ page, urgentRevision, t, make, fallback });
  latest.current = { page, urgentRevision, t, make, fallback };
  const activeRef = useRef(active);
  activeRef.current = active;
  const builtForRef = useRef<BuiltFor | null>(null);
  const shows = (kind: BuiltFor['kind']): boolean => {
    const built = builtForRef.current;
    const wanted = latest.current;
    return (
      built?.kind === kind &&
      built.page === wanted.page &&
      built.t === wanted.t &&
      // A fallback's links do not move with a position shift.
      (kind === 'fallback' ||
        (built.revision === displayPageRevision(wanted.page) &&
          built.urgentRevision === wanted.urgentRevision))
    );
  };
  const build = useCallback((): void => {
    const host = hostRef.current;
    if (!host || shows('chrome')) return;
    const { page, urgentRevision, t, make } = latest.current;
    // Keep the previous chrome connected until its replacement is ready.
    replaceKeepingFocus(host, make(page, t));
    builtForRef.current = {
      kind: 'chrome',
      page,
      revision: displayPageRevision(page),
      urgentRevision,
      t,
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [hostRef]);
  const showFallback = useCallback((): void => {
    const host = hostRef.current;
    if (!host || shows('fallback')) return;
    const { page, urgentRevision, t, fallback } = latest.current;
    if (!fallback) {
      if (builtForRef.current) replaceKeepingFocus(host, null);
      builtForRef.current = null;
      return;
    }
    const chrome = shows('chrome') ? (host.firstElementChild as HTMLElement | null) : null;
    replaceKeepingFocus(host, fallback(page, t, chrome));
    builtForRef.current = {
      kind: 'fallback',
      page,
      revision: displayPageRevision(page),
      urgentRevision,
      t,
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [hostRef]);

  useEffect(() => {
    register?.({
      build,
      release: () => {
        if (!activeRef.current && builtForRef.current?.kind === 'chrome') showFallback();
      },
    });
    return () => register?.(null);
  }, [build, register, showFallback]);

  useEffect(() => {
    const host = hostRef.current;
    if (!host) return;
    const idle = (work: () => void): (() => void) => {
      if (typeof requestIdleCallback === 'function') {
        const id = requestIdleCallback(work, { timeout: 1500 });
        return () => cancelIdleCallback(id);
      }
      const id = setTimeout(work, 150);
      return () => clearTimeout(id);
    };
    if (!active) {
      if (shows('fallback')) return;
      // Leaving the window frees the chrome at once; a first fallback waits.
      if (builtForRef.current?.kind === 'chrome' || !fallback) {
        showFallback();
        return;
      }
      return idle(showFallback);
    }
    if (shows('chrome')) return;
    const built = builtForRef.current?.kind === 'chrome' ? builtForRef.current : null;
    const { defer, rebuildAtOnce } = scheduling.current;
    const changed =
      built !== null &&
      (built.page !== page || built.t !== t || built.urgentRevision !== urgentRevision);
    if (built ? changed && rebuildAtOnce : !defer) {
      build();
      return;
    }
    return idle(build);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [page, t, revision, urgentRevision, active, defer, build, showFallback, fallback]);
}
