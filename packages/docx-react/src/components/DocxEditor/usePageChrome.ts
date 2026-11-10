import { useCallback, useEffect, useRef, type RefObject } from 'react';
import { displayPageIsShiftOf, displayPageRevision, type DisplayPage } from '@betteroffice/docx/layout/render';
import type { TFunction } from '@betteroffice/docx-i18n';
import { scheduleIdleWork } from './hooks/pageBuildScheduler';

type MakeChrome = (page: DisplayPage, t: TFunction) => HTMLElement;
/** Brings `chrome`, built by `make`, up to `page` in place; false when it cannot. */
type UpdateChrome = (chrome: HTMLElement, page: DisplayPage, t: TFunction) => boolean;
/** `chrome` is the page's current chrome, if built for this page, which it may take over. */
type MakeFallback = (page: DisplayPage, t: TFunction, chrome: HTMLElement | null) => HTMLElement;

export interface PageChromeOptions {
  page: DisplayPage;
  t: TFunction;
  /** Holds the chrome; an inactive page holds only its `fallback`, if any. */
  active: boolean;
  /** A first build may wait for idle time. */
  defer: boolean;
  /** While held, first builds and leaving the window wait until a frame after release. */
  hold?: PageChromeHold | null;
  /** A rebuild after a content change never waits. */
  rebuildAtOnce: boolean;
  /** Counts urgent changes; other position shifts rebuild at idle time. */
  urgentRevision: number;
  /** Receives the handle that builds the chrome at once, for a page needed before it is active. */
  register?: (handle: PageChromeHandle | null) => void;
  make: MakeChrome;
  /** Updates built chrome in place instead of replacing it, when it can. */
  update?: UpdateChrome | null;
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

const IDLE_SLICE_MS = 8;
const fallbackQueue: { work: () => void; expiresAt: number }[] = [];
let cancelFallbackDrain: (() => void) | null = null;
let drainingFallbacks = false;

function scheduleFallbackDrain(): void {
  if (cancelFallbackDrain || drainingFallbacks || fallbackQueue.length === 0) return;
  const drain = (deadline?: IdleDeadline): void => {
    const start = performance.now();
    cancelFallbackDrain = null;
    drainingFallbacks = true;
    try {
      fallbackQueue.shift()?.work();
      while (
        deadline &&
        !deadline.didTimeout &&
        deadline.timeRemaining() > 1 &&
        performance.now() - start < IDLE_SLICE_MS &&
        fallbackQueue.length > 0
      ) {
        fallbackQueue.shift()!.work();
      }
    } finally {
      drainingFallbacks = false;
      scheduleFallbackDrain();
    }
  };
  cancelFallbackDrain = scheduleIdleWork(drain, fallbackQueue[0]!.expiresAt, 50).cancel;
}

function enqueueFallback(work: () => void): () => void {
  const queued = {
    work,
    expiresAt: performance.now() + (typeof requestIdleCallback === 'function' ? 5000 : 50),
  };
  fallbackQueue.push(queued);
  scheduleFallbackDrain();
  return () => {
    const index = fallbackQueue.indexOf(queued);
    if (index >= 0) fallbackQueue.splice(index, 1);
    if (fallbackQueue.length === 0) {
      cancelFallbackDrain?.();
      cancelFallbackDrain = null;
    }
  };
}

/** Holds back page chrome work, for example while the pages scroll. */
export interface PageChromeHold {
  readonly held: boolean;
  /** Runs `work` once the hold ends; returns its cancel. */
  whenReleased(work: () => void): () => void;
}

export function createPageChromeHold(): PageChromeHold & { set(held: boolean): void } {
  let held = false;
  const waiting = new Set<() => void>();
  return {
    get held() {
      return held;
    },
    set(next) {
      held = next;
      if (held) return;
      const released = [...waiting];
      waiting.clear();
      for (const work of released) work();
    },
    whenReleased(work) {
      if (!held) {
        work();
        return () => {};
      }
      const entry = (): void => work();
      waiting.add(entry);
      return () => {
        waiting.delete(entry);
      };
    },
  };
}

/** Runs held work in a frame after release; other work waits for idle time. */
function afterHold(hold: PageChromeHold | null, work: () => void): () => void {
  let cancel = (): void => {};
  let cancelled = false;
  const run = (): void => {
    if (cancelled) return;
    if (hold?.held) wait();
    else work();
  };
  const frame = (): void => {
    if (cancelled) return;
    const id = requestAnimationFrame(run);
    cancel = () => cancelAnimationFrame(id);
  };
  const wait = (): void => {
    cancel = hold!.whenReleased(frame);
  };
  if (hold?.held) wait();
  else cancel = enqueueFallback(run);
  return () => {
    cancelled = true;
    cancel();
  };
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

const BLOCK = '[data-block-id]';

/** A place in the text of a mirror block, by block id, which of the blocks with that id, and text offset. */
interface BlockPoint {
  id: string;
  nth: number;
  offset: number;
}

const blocksWithId = (host: HTMLElement, id: string): HTMLElement[] =>
  Array.from(host.querySelectorAll<HTMLElement>(BLOCK)).filter((block) => block.dataset.blockId === id);

function blockPoint(host: HTMLElement, node: Node | null, offset: number): BlockPoint | null {
  const element = node instanceof Element ? node : node?.parentElement;
  const block = element?.closest<HTMLElement>(BLOCK);
  if (!node || !block || !host.contains(block)) return null;
  const range = host.ownerDocument.createRange();
  range.selectNodeContents(block);
  range.setEnd(node, offset);
  const id = block.dataset.blockId!;
  return { id, nth: blocksWithId(host, id).indexOf(block), offset: range.toString().length };
}

const blockAt = (host: HTMLElement, point: BlockPoint): HTMLElement | null =>
  blocksWithId(host, point.id)[point.nth] ?? null;

function textPoint(block: HTMLElement, offset: number): [Node, number] {
  const walker = block.ownerDocument.createTreeWalker(block, NodeFilter.SHOW_TEXT);
  let seen = 0;
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const length = node.textContent?.length ?? 0;
    if (offset <= seen + length) return [node, offset - seen];
    seen += length;
  }
  return [block, block.childNodes.length];
}

/** Focuses a mirror block that is not a tab stop; it stops being focusable once it loses focus. */
function focusBlock(block: HTMLElement): void {
  if (!block.hasAttribute('tabindex')) {
    block.tabIndex = -1;
    block.addEventListener('blur', () => {
      if (block.getAttribute('tabindex') === '-1') block.removeAttribute('tabindex');
    }, { once: true });
  }
  block.focus({ preventScroll: true });
}

/**
 * Runs `change` on `host`'s content. When it loses the focused stop, focus
 * returns to that stop if it survived, else moves to the same stop in the new
 * content, or to the stop at the same place when that one is gone. A focused
 * mirror block that is not a stop, and a selection that lay wholly in `host`,
 * return to the same block and text offset when that block still exists.
 * Focus and selection elsewhere are left alone. Returns what `change` returns.
 */
function changeKeepingFocus(host: HTMLElement, change: () => boolean): boolean {
  const doc = host.ownerDocument;
  const focused = doc.activeElement;
  const inside = focused instanceof HTMLElement && host.contains(focused) && focused !== host;
  const stops = inside ? tabStops(host) : [];
  const at = stops.indexOf(focused as HTMLElement);
  const key = at >= 0 ? stopKey(stops[at]!) : null;
  const nth = stops.slice(0, at).filter((stop) => stopKey(stop) === key).length;
  const focusedBlock = inside && at < 0 && focused.matches(BLOCK) ? blockPoint(host, focused, 0) : null;
  const selection = doc.getSelection();
  const before = selection && selection.rangeCount === 1 && host.contains(selection.anchorNode) &&
    host.contains(selection.focusNode)
    ? {
      anchorNode: selection.anchorNode, anchorOffset: selection.anchorOffset,
      focusNode: selection.focusNode, focusOffset: selection.focusOffset,
      anchor: blockPoint(host, selection.anchorNode, selection.anchorOffset),
      focus: blockPoint(host, selection.focusNode, selection.focusOffset),
    }
    : null;
  if (!change()) return false;
  const active = doc.activeElement;
  if (focusedBlock && active !== focused && (!active || active === doc.body || host.contains(active))) {
    const block = blockAt(host, focusedBlock);
    if (block) focusBlock(block);
  }
  if (before?.anchor && before.focus && selection) {
    const kept = selection.rangeCount === 1 && host.contains(before.anchorNode) && host.contains(before.focusNode) &&
      selection.anchorNode === before.anchorNode && selection.anchorOffset === before.anchorOffset &&
      selection.focusNode === before.focusNode && selection.focusOffset === before.focusOffset;
    const anchor = kept ? null : blockAt(host, before.anchor);
    const extent = kept ? null : blockAt(host, before.focus);
    if (anchor && extent) {
      selection.setBaseAndExtent(...textPoint(anchor, before.anchor.offset), ...textPoint(extent, before.focus.offset));
    }
  }
  if (at < 0 || doc.activeElement === focused) return true;
  const rebuilt = tabStops(host);
  if (rebuilt.includes(focused as HTMLElement)) {
    (focused as HTMLElement).focus({ preventScroll: true });
    return true;
  }
  const same = key === null ? [] : rebuilt.filter((stop) => stopKey(stop) === key);
  (same[nth] ?? rebuilt[at])?.focus({ preventScroll: true });
  return true;
}

/** Replaces `host`'s content with `next`, keeping focus as `changeKeepingFocus` does. */
function replaceKeepingFocus(host: HTMLElement, next: HTMLElement | null): void {
  if (next && host.childNodes.length === 1 && host.firstChild === next) return;
  changeKeepingFocus(host, () => {
    if (next) host.replaceChildren(next);
    else host.replaceChildren();
    return true;
  });
}

/** Marks which built page revision `host` shows; an unbuilt page's chrome is an empty shell and carries no mark. */
function markBuiltFrom(host: HTMLElement, page: DisplayPage): void {
  if (page.unbuilt) delete host.dataset.pageChromeBuilt;
  else host.dataset.pageChromeBuilt = String(displayPageRevision(page));
}

/** Builds one page's mirror or overlay into `hostRef`, and rebuilds it with the page. */
export function usePageChrome(
  hostRef: RefObject<HTMLDivElement | null>,
  {
    page,
    t,
    active,
    defer,
    hold = null,
    rebuildAtOnce,
    urgentRevision,
    register,
    make,
    update = null,
    fallback = null,
  }: PageChromeOptions
): void {
  // Owned deltas shift primitive positions in place: identity alone is stale.
  const revision = displayPageRevision(page);
  // Read when a build is scheduled: a change of scheduling alone keeps the
  // built DOM (and any focus inside it) in place.
  const scheduling = useRef({ defer, rebuildAtOnce });
  scheduling.current = { defer, rebuildAtOnce };
  const latest = useRef({ page, urgentRevision, t, make, update, fallback });
  latest.current = { page, urgentRevision, t, make, update, fallback };
  const activeRef = useRef(active);
  activeRef.current = active;
  const builtForRef = useRef<BuiltFor | null>(null);
  const shows = (kind: BuiltFor['kind']): boolean => {
    const built = builtForRef.current;
    const wanted = latest.current;
    return (
      built?.kind === kind &&
      (built.page === wanted.page ||
        (kind === 'fallback' && displayPageIsShiftOf(wanted.page, built.page))) &&
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
    const { page, urgentRevision, t, make, update } = latest.current;
    const built = builtForRef.current;
    const chrome = built?.kind === 'chrome' && built.t === t ? host.firstElementChild : null;
    const updated =
      update !== null &&
      chrome instanceof HTMLElement &&
      host.childNodes.length === 1 &&
      changeKeepingFocus(host, () => update(chrome, page, t));
    // Keep the previous chrome connected until its replacement is ready.
    if (!updated) replaceKeepingFocus(host, make(page, t));
    host.dataset.pageChrome = 'chrome';
    markBuiltFrom(host, page);
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
      delete host.dataset.pageChrome;
      delete host.dataset.pageChromeBuilt;
      builtForRef.current = null;
      return;
    }
    const chrome = shows('chrome') ? (host.firstElementChild as HTMLElement | null) : null;
    replaceKeepingFocus(host, fallback(page, t, chrome));
    host.dataset.pageChrome = 'fallback';
    markBuiltFrom(host, page);
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
      if (shows('fallback')) {
        builtForRef.current!.page = page;
        markBuiltFrom(host, page);
        return;
      }
      // Leaving the window frees the chrome at once unless held; a first fallback waits.
      if ((builtForRef.current?.kind === 'chrome' || !fallback) && !hold?.held) {
        showFallback();
        return;
      }
      return afterHold(hold, showFallback);
    }
    if (shows('chrome')) return;
    const built = builtForRef.current?.kind === 'chrome' ? builtForRef.current : null;
    const { defer, rebuildAtOnce } = scheduling.current;
    const changed =
      built !== null &&
      ((built.page !== page && !displayPageIsShiftOf(page, built.page)) ||
        built.t !== t || built.urgentRevision !== urgentRevision);
    if (built ? changed && rebuildAtOnce : !defer) {
      if (!built && hold?.held) return afterHold(hold, build);
      build();
      return;
    }
    return hold?.held ? afterHold(hold, build) : idle(build);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [page, t, revision, urgentRevision, active, defer, hold, build, showFallback, fallback]);
}
