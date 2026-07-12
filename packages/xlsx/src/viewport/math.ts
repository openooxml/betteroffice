/**
 * Pure viewport math: visible-range computation against cumulative track
 * offsets, and scroll clamping. No DOM, no side effects — every function is a
 * deterministic transform of plain numbers so it is golden-testable and can be
 * mirrored in Rust.
 */

import type { TrackOffsets, ViewportState, VisibleRange, VisibleCells } from './types';

/**
 * Number of tracks described by an offsets array.
 */
export function trackCount(offsets: TrackOffsets): number {
  return Math.max(0, offsets.length - 1);
}

/**
 * Total pixel extent of a track axis (position of the trailing edge).
 */
export function totalExtent(offsets: TrackOffsets): number {
  return offsets.length > 0 ? offsets[offsets.length - 1] : 0;
}

function clampInt(value: number, lo: number, hi: number): number {
  return Math.max(lo, Math.min(hi, value));
}

// largest track index whose leading edge is <= target, searched in [lo, hi].
function trackAt(offsets: TrackOffsets, target: number, lo: number, hi: number): number {
  let left = lo;
  let right = hi;
  while (left < right) {
    const mid = (left + right + 1) >> 1;
    if (offsets[mid] <= target) left = mid;
    else right = mid - 1;
  }
  return left;
}

/**
 * The inclusive range of scrollable (non-frozen) tracks intersecting the
 * viewport. `frozen` tracks at the head are pinned and excluded here — the
 * scrollable window starts after them and is shortened by their pinned width.
 */
export function visibleRange(
  offsets: TrackOffsets,
  frozen: number,
  scroll: number,
  viewportExtent: number
): VisibleRange {
  const count = trackCount(offsets);
  if (count === 0) return { first: 0, last: -1 };

  const f = clampInt(Math.trunc(frozen), 0, count);
  if (f >= count) return { first: 0, last: -1 };

  const frozenSize = offsets[f] - offsets[0];
  const windowStart = offsets[f] + Math.max(0, scroll);
  const windowExtent = Math.max(0, viewportExtent - frozenSize);
  const windowEnd = windowStart + windowExtent;

  const first = trackAt(offsets, windowStart, f, count - 1);
  // trailing edge strictly past windowStart guarantees a nonzero sliver shows.
  if (offsets[first + 1] <= windowStart || windowExtent === 0) {
    return { first: 0, last: -1 };
  }
  const last = trackAt(offsets, windowEnd, first, count - 1);
  return { first, last };
}

/**
 * Visible column and row ranges for a frame, frozen tracks excluded.
 */
export function visibleCells(
  state: ViewportState,
  colOffsets: TrackOffsets,
  rowOffsets: TrackOffsets
): VisibleCells {
  return {
    cols: visibleRange(colOffsets, state.frozenCols, state.scrollX, state.width),
    rows: visibleRange(rowOffsets, state.frozenRows, state.scrollY, state.height),
  };
}

// furthest a scroll offset can move before the last track sits flush with the
// viewport edge; never negative when content is smaller than the viewport.
function maxScroll(offsets: TrackOffsets, frozen: number, viewportExtent: number): number {
  const count = trackCount(offsets);
  const f = clampInt(Math.trunc(frozen), 0, count);
  const scrollableExtent = totalExtent(offsets) - offsets[f];
  const visibleExtent = Math.max(0, viewportExtent - (offsets[f] - offsets[0]));
  return Math.max(0, scrollableExtent - visibleExtent);
}

/**
 * Clamp a viewport's scroll offsets into the valid range for the given content.
 * Returns the same state when already in range, a corrected copy otherwise.
 */
export function clampScroll(
  state: ViewportState,
  colOffsets: TrackOffsets,
  rowOffsets: TrackOffsets
): ViewportState {
  const scrollX = clampInt(state.scrollX, 0, maxScroll(colOffsets, state.frozenCols, state.width));
  const scrollY = clampInt(state.scrollY, 0, maxScroll(rowOffsets, state.frozenRows, state.height));
  if (scrollX === state.scrollX && scrollY === state.scrollY) return state;
  return { ...state, scrollX, scrollY };
}

/**
 * Build a uniform offsets array of `count` tracks each `size` px wide.
 */
export function uniformOffsets(count: number, size: number): TrackOffsets {
  const offsets: TrackOffsets = new Array(count + 1);
  for (let i = 0; i <= count; i++) offsets[i] = i * size;
  return offsets;
}
