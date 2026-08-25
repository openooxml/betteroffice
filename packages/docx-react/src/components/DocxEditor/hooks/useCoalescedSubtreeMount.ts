import { useEffect, useRef } from 'react';

const REBUILD_DEBOUNCE_MS = 150;
const REBUILD_MAX_WAIT_MS = 1000;

/**
 * Mounts an imperatively built DOM subtree under `hostRef`, coalescing rapid
 * successive rebuilds: the first mount builds synchronously, later dependency
 * changes are trailing-debounced with a max wait. Sustained typing re-renders
 * the owning component per keystroke, but the subtree (thousands of nodes for
 * an a11y mirror page) is rebuilt at most once per pause or per max-wait
 * window instead of once per keystroke. The painted canvas carries the visual
 * update, so deferring these DOM consumers (screen readers, plugin DOM
 * queries) is not user-visible.
 */
export function useCoalescedSubtreeMount(
  hostRef: React.RefObject<HTMLElement | null>,
  build: () => Node,
  deps: readonly unknown[]
): void {
  const timerRef = useRef<number | null>(null);
  const deadlineRef = useRef<number | null>(null);
  const buildRef = useRef(build);
  buildRef.current = build;

  // eslint-disable-next-line react-hooks/exhaustive-deps
  useEffect(() => {
    const host = hostRef.current;
    if (!host) return;
    const run = () => {
      timerRef.current = null;
      deadlineRef.current = null;
      // Keep the previous subtree connected until the replacement is ready.
      host.replaceChildren(buildRef.current());
    };
    if (!host.hasChildNodes() && timerRef.current === null) {
      run();
      return;
    }
    const now = performance.now();
    deadlineRef.current ??= now + REBUILD_MAX_WAIT_MS;
    if (timerRef.current !== null) window.clearTimeout(timerRef.current);
    timerRef.current = window.setTimeout(
      run,
      Math.min(REBUILD_DEBOUNCE_MS, Math.max(0, deadlineRef.current - now))
    );
    // The pending rebuild intentionally survives dependency changes — each
    // change reschedules it above; only unmount cancels.
  }, deps);

  useEffect(
    () => () => {
      if (timerRef.current !== null) window.clearTimeout(timerRef.current);
    },
    []
  );
}
