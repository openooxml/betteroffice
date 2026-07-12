/**
 * Pure viewport types. No DOM — this is the compute seam the a11y grid, the
 * canvas painter, and a future Rust port all read from.
 */

/**
 * Cumulative pixel offsets for a track axis (columns or rows). `offsets[i]` is
 * the leading edge of track `i`; `offsets.length === trackCount + 1`, so the
 * last element is the total extent. Frozen tracks are the first `frozen` of
 * these — they are pinned and always visible.
 */
export type TrackOffsets = number[];

/**
 * Scroll + geometry state driving virtualization. `scrollX/scrollY` are in
 * content pixels of the scrollable (non-frozen) region; `width/height` is the
 * visible viewport in css pixels; `dpr` scales to device pixels at paint time.
 */
export interface ViewportState {
  scrollX: number;
  scrollY: number;
  width: number;
  height: number;
  dpr: number;
  frozenRows: number;
  frozenCols: number;
}

/**
 * An inclusive track index range. `last < first` (e.g. `{ first: 0, last: -1 }`)
 * denotes an empty range.
 */
export interface VisibleRange {
  first: number;
  last: number;
}

/**
 * The visible track window on both axes, frozen tracks excluded (they always
 * render). Feed to the renderer to build the display list for a frame.
 */
export interface VisibleCells {
  cols: VisibleRange;
  rows: VisibleRange;
}
