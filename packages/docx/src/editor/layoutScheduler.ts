/**
 * rAF-coalescing layout scheduler shared by the React and Vue adapters
 * (issue #696 Tier 2). Rapid doc-changing transactions (a burst of
 * keystrokes) collapse to a single layout pass per animation frame: while a
 * frame is pending, later `schedule` calls just replace the target state, so
 * only the final state lays out.
 *
 * `scheduleFrame`/`cancelFrame` are injected so a headless/test host can pass
 * a synchronous stub; they default to requestAnimationFrame.
 */

export interface LayoutScheduler<TState = unknown> {
  /** Request a layout for `state`, coalesced into the pending frame. */
  schedule(state: TState): void;
  /** Cancel any pending frame (call on teardown). */
  cancel(): void;
}

export function createLayoutScheduler<TState>(
  run: (state: TState) => void,
  scheduleFrame: (cb: () => void) => number = (cb) => requestAnimationFrame(cb),
  cancelFrame: (handle: number) => void = (handle) => cancelAnimationFrame(handle)
): LayoutScheduler<TState> {
  let pending: { handle: number; state: TState } | null = null;

  return {
    schedule(state: TState): void {
      // A frame is already queued — just update the target state so the
      // pending pass uses the latest doc (coalescing).
      if (pending) {
        pending.state = state;
        return;
      }
      // Set `pending` BEFORE scheduling so a synchronous `scheduleFrame`
      // (test / headless host) sees it when its callback fires inline.
      pending = { handle: 0, state };
      const handle = scheduleFrame(() => {
        const p = pending;
        pending = null;
        if (p) run(p.state);
      });
      // A synchronous frame already cleared `pending`; only record the handle
      // when the frame is still genuinely pending (async case).
      if (pending) pending.handle = handle;
    },

    cancel(): void {
      if (pending) {
        cancelFrame(pending.handle);
        pending = null;
      }
    },
  };
}
