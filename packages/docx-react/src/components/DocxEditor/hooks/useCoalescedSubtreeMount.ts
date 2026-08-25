import { useEffect, useRef } from 'react';

const REBUILD_DEBOUNCE_MS = 150;
const REBUILD_MAX_WAIT_MS = 1000;

/**
 * Mounts an imperatively built DOM subtree under `hostRef`: the first mount
 * builds synchronously, later dependency changes are trailing-debounced with
 * a max wait, so sustained typing rebuilds the (thousands-of-nodes) subtree
 * once per pause instead of once per keystroke. Only for consumers whose
 * staleness is invisible while the canvas carries the visual update.
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
