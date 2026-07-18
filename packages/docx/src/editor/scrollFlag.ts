/**
 * Strip ProseMirror's internal `UPDATED_SCROLL` flag from a transaction so the
 * hidden, off-screen editor's `updateState` does not force-scroll its ancestors
 * to the selection — the paginated painter owns scroll, not the PM view.
 *
 * Shared by both adapters' `dispatchTransaction` (issue #696 Tier 2). React has
 * always done this; Vue did not, so its hidden editor could yank an ancestor's
 * scroll on edit. Centralizing the magic bit + the drift canary here keeps the
 * two in lockstep.
 *
 * `Transaction.updated` is a private bitfield in `prosemirror-state` whose
 * `UPDATED_SCROLL` bit is not exported. It is `4` in current PM
 * (state/src/transaction.ts). If a future PM release renumbers the bits before
 * SCROLL, `assertScrollFlagShape` logs once so the constant can be updated.
 */

/** Minimal transaction shape used by the legacy scroll-flag compatibility seam. */
export interface ScrollFlagTransaction {
  scrollIntoView(): ScrollFlagTransaction;
}

const PM_UPDATED_SCROLL = 4;
let pmScrollFlagAsserted = false;

function assertScrollFlagShape(probeTr: ScrollFlagTransaction): void {
  if (pmScrollFlagAsserted) return;
  pmScrollFlagAsserted = true;
  try {
    const probe = probeTr.scrollIntoView() as unknown as { updated?: number };
    if (typeof probe.updated !== 'number' || (probe.updated & PM_UPDATED_SCROLL) === 0) {
      console.warn(
        '[docx-editor] prosemirror-state UPDATED_SCROLL bit shape changed; ' +
          'paginated scroll suppression may be stale. Update PM_UPDATED_SCROLL.'
      );
    }
  } catch {
    // Probe failed (e.g. PM mocked in tests) — skip silently.
  }
}

/**
 * Clear the scroll-into-view flag on `transaction` in place. Call inside
 * `dispatchTransaction` before `state.apply(transaction)`. `probeTr` should be
 * a fresh `view.state.tr` used only for the one-shot drift canary (so the real
 * transaction is never mutated by the probe).
 */
export function stripScrollFlag(
  transaction: ScrollFlagTransaction,
  probeTr: ScrollFlagTransaction
): void {
  assertScrollFlagShape(probeTr);
  // `updated` is `private` on PM's Transaction, so a plain intersection
  // collapses to `never`. The double-cast is the documented escape hatch.
  const trWithUpdated = transaction as unknown as { updated?: number };
  if (typeof trWithUpdated.updated === 'number') {
    trWithUpdated.updated &= ~PM_UPDATED_SCROLL;
  }
}
