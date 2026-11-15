import { useCallback, useEffect, useRef, useState } from 'react';
import type { CSSProperties, MouseEvent as ReactMouseEvent, PointerEvent as ReactPointerEvent, ReactNode, WheelEvent as ReactWheelEvent } from 'react';
import { normalizeRange, type DisplayList, type GridMeta, type Selection } from '@betteroffice/xlsx';
import { XLSX_MAX_COLUMNS, XLSX_MAX_ROWS } from '../sheetLimits';
import { columnLetter } from '../columnLetter';

/** Fixed header band sizes; column widths and row heights inside them follow the sheet zoom. */
export const COLUMN_HEADER_HEIGHT = 24;
const ROW_HEADER_WIDTH = 44;
/** Where a boundary resize strip starts across the band on each axis. */
const COLUMN_RESIZE_STRIP_TOP = 9;
const ROW_RESIZE_STRIP_LEFT = 10;

/**
 * The freeze split chrome, mirrored from `crates/xlsx-render` (`PANE_DIVIDER_COLOR`
 * and `PANE_DIVIDER_WIDTH`) so the band line lands on the canvas line. The
 * canvas strokes centred on the divider coordinate, hence the half-width offset.
 */
const FREEZE_DIVIDER_COLOR = '#8a8a8a';
const FREEZE_DIVIDER_WIDTH = 2;

/** How far a pointer must drag inside a resize strip before it resizes instead of clicking. */
export const HEADER_RESIZE_SLOP_PX = 4;
/** Half-width of the resize strip at a track boundary, in screen pixels. */
export const HEADER_RESIZE_ZONE_PX = 4;
/** How far a pointer must drag a label before it reorders instead of pressing. */
export const HEADER_REORDER_SLOP_PX = 6;
/**
 * A column width is stored in characters of the default font's max digit width
 * and a row height in points; the frame is in CSS pixels. These are the same
 * conversions the renderer applies (ECMA-376 §18.3.1.13, 96 dpi), so a drag in
 * pixels can be written back to the engine exactly.
 */
const MAX_DIGIT_WIDTH_PX = 7;
const COLUMN_WIDTH_PAD_PX = 5;
/** Excel's limits, converted to the pixels a drag is measured in. */
export const MAX_COLUMN_WIDTH_PX = 255 * MAX_DIGIT_WIDTH_PX + COLUMN_WIDTH_PAD_PX;
export const MAX_ROW_HEIGHT_PX = (409.5 * 4) / 3;

/** A dragged column width in pixels to the characters `setColWidth` stores. */
export function columnWidthFromPx(px: number): number {
  return Math.max(0, (px - COLUMN_WIDTH_PAD_PX) / MAX_DIGIT_WIDTH_PX);
}

/** A dragged row height in pixels to the points `setRowHeight` stores. */
export function rowHeightFromPx(px: number): number {
  return Math.max(0, (px * 3) / 4);
}

interface HeaderTrack {
  /** Absolute sheet row or column index. */
  index: number;
  /** Leading edge in the frame, before zoom. */
  offset: number;
  /** Track size in the frame, before zoom. */
  size: number;
}

/**
 * What a surface supplies to make the visible header bands behave like real
 * spreadsheet chrome. Every member is optional: a surface that only paints
 * labels (the read-only viewer) passes nothing and the bands stay inert.
 */
export interface SheetGridHeaderHandlers {
  /**
   * Primary press on a column or row header: select the whole track. A press
   * inside a selected run is held back to arm that run's reorder; if the
   * pointer comes up without a drag, the same press is replayed here, so a
   * click still collapses the selection to the pressed track.
   */
  onHeaderMouseDown?(axis: 'col' | 'row', index: number, event: ReactMouseEvent): void;
  /**
   * A committed boundary drag: `size` is the new width or height in sheet
   * units. When the boundary's leading track sits inside a whole-column (or
   * whole-row) selection, `run` names that selection in sheet indices
   * (`at` = first track, `count` = tracks) and the surface sets every track in
   * it to `size`, the way a spreadsheet resizes a selected block from one
   * boundary. Otherwise `run` is undefined and only the leading track resizes.
   */
  onHeaderResize?(axis: 'col' | 'row', index: number, size: number, run?: { at: number; count: number }): void;
  /**
   * A committed label drag: move the `count` tracks starting at sheet index
   * `at` so they land immediately before sheet index `before`. All three are
   * sheet indices in pre-move coordinates, ready for the engine's
   * `moveCols`/`moveRows` op. `at` and `count` come from the selection range,
   * not from the painted band, so a run whose leading tracks are hidden or
   * scrolled off still moves whole; `before` is the sheet index of the first
   * painted track after the drop, or one past the last painted track when the
   * drop lies beyond the window.
   *
   * The band drops a release that names the run's leading track, its trailing
   * successor, or any track strictly inside it, so a commit always moves
   * something. Supplying this makes a label drag reorder *when the pressed
   * track is already selected in full*; a drag from any other track extends
   * the header selection instead (`onHeaderDragSelect`). A press on one track
   * of a whole-column (or whole-row) selection moves the whole selected run.
   *
   * The surface commits it with the engine's `moveCols`/`moveRows` op, which
   * carries widths, styles, merges and formula references with the run and
   * refuses a move that would tear a merged range.
   */
  onHeaderReorder?(axis: 'col' | 'row', at: number, before: number, count: number): void;
  /**
   * A press on the corner above the row numbers and left of the letters: the
   * spreadsheet gesture that selects every cell.
   */
  onSelectAll?(): void;
  /**
   * A label drag reached another track: extend the selection to `index`, the
   * sheet index of the track now under the pointer. Fired as the pointer
   * crosses tracks, never for the track it was pressed on, and never when the
   * drag became a reorder. Supplying this makes a drag on an unselected track
   * select instead of doing nothing.
   */
  onHeaderDragSelect?(axis: 'col' | 'row', index: number): void;
  /** Pointer down anywhere on the header chrome (label, resize edge, corner), any button. */
  onHeaderPress?(): void;
}

/** The track a pointer at `offset` inside the band lands in, `null` outside it. */
export function trackAtOffset(
  tracks: HeaderTrack[],
  offset: number,
): HeaderTrack | null {
  for (const track of tracks) {
    if (offset >= track.offset && offset < track.offset + track.size) return track;
  }
  return null;
}

/**
 * The resize gesture a pointer at `offset` inside the band starts, if any: the
 * boundary between two tracks, resolved to the track on the left of (or above)
 * it. The last track's trailing edge is never a boundary, so it never steals a
 * click. Hidden tracks produce no boundary because `headerTracks` drops them.
 */
export function resizeGestureAt(
  tracks: HeaderTrack[],
  offset: number,
  zoom = 1,
): HeaderTrack | null {
  const zone = HEADER_RESIZE_ZONE_PX / (zoom || 1);
  for (let i = 0; i < tracks.length - 1; i += 1) {
    const boundary = tracks[i].offset + tracks[i].size;
    if (Math.abs(offset - boundary) <= zone) return tracks[i];
  }
  return null;
}

/** Clamp a dragged track size, in pixels, to what the engine accepts. */
export function clampTrackSize(axis: 'col' | 'row', size: number): number {
  const maximum = axis === 'col' ? MAX_COLUMN_WIDTH_PX : MAX_ROW_HEIGHT_PX;
  return Math.min(Math.max(size, 0), maximum);
}

/**
 * The tracks as a live drag renders them: the dragged track takes `delta` more
 * (or fewer) pixels and every track after it shifts by the same amount, so the
 * band follows the pointer instead of only drawing a guide.
 */
export function resizePreview(tracks: HeaderTrack[], index: number, delta: number): HeaderTrack[] {
  let passed = false;
  return tracks.map((track) => {
    if (track.index === index) {
      passed = true;
      return { ...track, size: Math.max(0, track.size + delta) };
    }
    return passed ? { ...track, offset: track.offset + delta } : track;
  });
}

/**
 * The visible tracks of one axis in a painted frame. Hidden tracks are already
 * compacted out by the renderer; a zero-sized track is skipped so its label
 * cannot overlap the next one.
 */
export function headerTracks(
  grid: GridMeta | undefined,
  axis: 'row' | 'col',
): HeaderTrack[] {
  if (!grid) return [];
  const offsets = axis === 'col' ? grid.colOffsets : grid.rowOffsets;
  const indices = axis === 'col' ? grid.colIndices : grid.rowIndices;
  const sizes = axis === 'col' ? grid.colSizes : grid.rowSizes;
  const start = axis === 'col' ? grid.startCol : grid.startRow;
  const tracks: HeaderTrack[] = [];
  for (let local = 0; local < offsets.length - 1; local += 1) {
    // the renderer's size is the full track; the offset gap is narrower on a
    // frozen track whose leading edge the pane clips. A resize gesture starts
    // from the full size, or dragging a clipped frozen boundary collapses it.
    const size = sizes?.[local] ?? offsets[local + 1] - offsets[local];
    if (size <= 0) continue;
    tracks.push({ index: indices?.[local] ?? start + local, offset: offsets[local], size });
  }
  return tracks;
}

/**
 * The visible geometry of one axis as a comparable string: one
 * `index:offset:size` entry per painted track. Gestures are measured against
 * this, not against the frame object, which is a new identity on every
 * repaint — a hover or a selection echo must not cancel a drag, while a
 * scroll, a resize, a hide toggle or a sheet switch moves a track and does.
 */
export function tracksSignature(tracks: HeaderTrack[]): string {
  return tracks.map((track) => `${track.index}:${track.offset}:${track.size}`).join('|');
}

/** A compacted hidden run as the band marks it: the boundary it collapsed onto. */
export interface HiddenSpanHandle {
  /** Position in the visible tracks whose leading boundary stands for the run. */
  position: number;
  /** First hidden sheet index of the run. */
  start: number;
  /** Last hidden sheet index of the run. */
  end: number;
}

/**
 * The boundaries a compacted hidden span collapsed onto, with the span itself:
 * the first visible track after each span's end. The band marks that boundary
 * with a doubled separator, so the collapse reads as a gap rather than as an
 * ordinary edge. A span with no visible track after it in the window has no
 * boundary to mark, and one with no visible track before it has no left
 * neighbour to separate from, so both are skipped. Spans that share a boundary
 * are merged into one run, so a single separator marks the whole gap.
 *
 * A frozen lead pins its tracks above the scrolling body, so a span that ends
 * before the body's first track collapsed onto a boundary between the frozen
 * lead and the body: marking it there would name a gap the reader cannot
 * scroll to. Such a span is dropped, separator and all; the body window is
 * bounded by the first visible track after the frozen lead and the last
 * visible track.
 */
export function hiddenSpanHandles(
  grid: GridMeta | undefined,
  axis: 'row' | 'col',
  tracks: HeaderTrack[],
): HiddenSpanHandle[] {
  const spans = axis === 'col' ? grid?.hiddenColSpans : grid?.hiddenRowSpans;
  if (!spans || spans.length === 0) return [];
  const frozen = axis === 'col' ? grid?.frozenCols ?? 0 : grid?.frozenRows ?? 0;
  const bodyStart = tracks.findIndex((track) => track.index >= frozen);
  if (bodyStart < 0) return [];
  const handles: HiddenSpanHandle[] = [];
  for (const [start, end] of spans) {
    if (end < tracks[bodyStart].index) continue;
    if (!tracks.some((track) => track.index < start)) continue;
    const following = tracks.findIndex((track) => track.index > end);
    if (following <= 0) continue;
    const shared = handles.find((handle) => handle.position === following);
    if (shared) {
      shared.start = Math.min(shared.start, start);
      shared.end = Math.max(shared.end, end);
    } else {
      handles.push({ position: following, start, end });
    }
  }
  return handles;
}

/** The visible-track positions `hiddenSpanHandles` marks, one per boundary. */
export function hiddenBoundaries(
  grid: GridMeta | undefined,
  axis: 'row' | 'col',
  tracks: HeaderTrack[],
): number[] {
  return hiddenSpanHandles(grid, axis, tracks).map((handle) => handle.position);
}

/**
 * The insertion gap a pointer at `offset` (in band units) lands in: the number of
 * tracks whose midpoint lies before it. Track geometry only — the selection, and
 * therefore any merged cell inside it, never feeds this.
 */
export function reorderGapAt(tracks: HeaderTrack[], offset: number): number {
  let gap = 0;
  for (const track of tracks) {
    if (track.offset + track.size / 2 < offset) gap += 1;
  }
  return gap;
}

/**
 * The run a label drag moves, in sheet indices: the whole selection when it is
 * a whole-column (or whole-row) selection, otherwise just the pressed track.
 * The band captures it at the press, before the surface's own press handler
 * collapses the selection to the pressed track.
 */
export function reorderRun(
  axis: 'col' | 'row',
  index: number,
  range: { top: number; bottom: number; left: number; right: number } | null,
): { from: number; to: number } {
  if (!range) return { from: index, to: index };
  const wholeTrack = axis === 'col'
    ? range.top === 0 && range.bottom === XLSX_MAX_ROWS - 1
    : range.left === 0 && range.right === XLSX_MAX_COLUMNS - 1;
  if (!wholeTrack) return { from: index, to: index };
  return axis === 'col'
    ? { from: range.left, to: range.right }
    : { from: range.top, to: range.bottom };
}

/**
 * The whole selected run a boundary drag resizes, in sheet indices, when the
 * boundary's leading track lies inside it: a spreadsheet resizes every column
 * of a multi-column selection from one boundary. `undefined` means the drag
 * resizes the single leading track, which is today's behaviour.
 */
export function resizeRun(
  axis: 'col' | 'row',
  index: number,
  range: { top: number; bottom: number; left: number; right: number } | null,
): { at: number; count: number } | undefined {
  const whole = reorderRun(axis, index, range);
  if (!range || (whole.from === index && whole.to === index)) return undefined;
  if (index < whole.from || index > whole.to) return undefined;
  return { at: whole.from, count: whole.to - whole.from + 1 };
}

/**
 * Whether the selection reaches a track, for the light header fill. A whole-row
 * selection does not tint column labels (and the other way round): its extent on
 * the other axis is the whole sheet, which would light the entire band.
 */
export function trackTouched(
  axis: 'col' | 'row',
  index: number,
  range: { top: number; bottom: number; left: number; right: number; } | null,
): boolean {
  if (!range) return false;
  const fullCols = range.top === 0 && range.bottom === XLSX_MAX_ROWS - 1;
  const fullRows = range.left === 0 && range.right === XLSX_MAX_COLUMNS - 1;
  return axis === 'col'
    ? !fullRows && range.left <= index && index <= range.right
    : !fullCols && range.top <= index && index <= range.bottom;
}

/**
 * Whether a track is selected in full: the selection spans every row (columns)
 * or every column (rows), and this track lies inside its other axis.
 */
export function trackSelected(
  axis: 'col' | 'row',
  index: number,
  range: { top: number; bottom: number; left: number; right: number } | null,
): boolean {
  if (!range) return false;
  return axis === 'col'
    ? range.top === 0 && range.bottom === XLSX_MAX_ROWS - 1 && range.left <= index && index <= range.right
    : range.left === 0 && range.right === XLSX_MAX_COLUMNS - 1 && range.top <= index && index <= range.bottom;
}

const bandBase: CSSProperties = {
  background: '#f8f9fa',
  color: '#5f6368',
  font: '11px/1 ui-sans-serif, system-ui, sans-serif',
  userSelect: 'none',
};

const cellBase: CSSProperties = {
  position: 'absolute',
  boxSizing: 'border-box',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  overflow: 'hidden',
  whiteSpace: 'nowrap',
};

/**
 * The band snaps a line's coordinate to the CSS pixel grid with `Math.round`,
 * exactly as `selectionOutline` snaps a marquee rect: the canvas paints a
 * gridline at `b` as `[round(b) - 1, round(b)]`, so the band line shares that
 * device column at any zoom.
 */
const freezeDividerStyle = (axis: 'col' | 'row', at: number, zoom: number): CSSProperties =>
  axis === 'col'
    ? {
        position: 'absolute',
        left: Math.round(at * zoom) - FREEZE_DIVIDER_WIDTH / 2,
        top: 0,
        width: FREEZE_DIVIDER_WIDTH,
        height: COLUMN_HEADER_HEIGHT,
        background: FREEZE_DIVIDER_COLOR,
        pointerEvents: 'none',
        zIndex: 1,
      }
    : {
        position: 'absolute',
        left: 0,
        top: Math.round(at * zoom) - FREEZE_DIVIDER_WIDTH / 2,
        width: ROW_HEADER_WIDTH,
        height: FREEZE_DIVIDER_WIDTH,
        background: FREEZE_DIVIDER_COLOR,
        pointerEvents: 'none',
        zIndex: 1,
      };

/** What the band painted when a gesture armed; a commit must not outlive it. */
interface GestureOrigin {
  /** The frame object the gesture was measured against. */
  frame: DisplayList | null;
  /** The sheet the frame belonged to, when the surface names it. */
  sheet: string | number | null;
  /**
   * {@link tracksSignature} of the gesture's axis as painted at the arm. A
   * repaint that leaves the tracks where they were keeps the gesture alive; a
   * scroll, a resize, a hide toggle or a sheet switch moves one and cancels.
   */
  tracks: string;
}

/** The frame the band paints right now: what a release reads, not the arming render. */
type PaintedFrame = Pick<GestureOrigin, 'frame' | 'sheet'>;

/** A boundary drag in progress; `host` is the band that captured the pointer. */
interface ResizeDrag extends GestureOrigin {
  axis: 'col' | 'row';
  index: number;
  size: number;
  /** Whole selected run this drag resizes, captured at the press; undefined for one track. */
  run?: { at: number; count: number };
  host: HTMLElement;
  pointerId: number;
  start: number;
  startSize: number;
  moved: boolean;
}

/** A label drag in progress; `fromIndex` is the dragged track's sheet index. */
interface ReorderDrag extends GestureOrigin {
  axis: 'col' | 'row';
  /** Sheet index of the dragged track, stable across a re-render. */
  fromIndex: number;
  /** The selected run this drag moves, in sheet indices, captured at the press. */
  run: { from: number; to: number };
  host: HTMLElement;
  pointerId: number;
  /** Pointer position in client coordinates when the press landed. */
  start: number;
  /** Insertion gap under the pointer, `null` until the slop is passed. */
  gap: number | null;
  /**
   * The press the band swallowed to arm this drag. Replayed as the surface's
   * press handler when the pointer comes up without a drop, so a click inside a
   * selection still collapses it to the pressed track.
   */
  press: { index: number; event: ReactMouseEvent } | null;
}

/** A selection drag across a band; `index` is the last track the pointer reached. */
interface TrackDrag extends GestureOrigin {
  axis: 'col' | 'row';
  host: HTMLElement;
  pointerId: number;
  index: number;
}

/**
 * The visible sheet chrome around the scrolling cell canvas: the column letters
 * above it and the row numbers to its left. The bands sit outside the scroller,
 * so the canvas viewport shrinks by their size and rows and columns never sit
 * underneath them. Labels come from the painted frame's grid metadata, which
 * already handles hidden, resized and frozen tracks and a scrolled viewport.
 *
 * `headerHandlers` turns the bands into chrome: a press selects the whole
 * track, a right-press opens a header menu, a drag near a boundary resizes that
 * track, and a drag on a label either reorders it or extends the header
 * selection. The band follows the spreadsheet rule for that last pair: a press
 * on a track the selection already covers in full arms a reorder of the whole
 * selected run, and a release without a drag is handed back as the surface's
 * press, so a click still collapses the run to the pressed track. A press
 * anywhere else drags the selection instead. Bands stay
 * `aria-hidden` either way; the accessible grid is A11yGridTable, which no
 * pointer gesture here touches.
 *
 * `selection` paints: it highlights the headers of fully selected tracks, and
 * it is what tells a fully selected track (a reorder) from a cell range that
 * merely reaches one (a selection drag).
 */
export function SheetGrid({ frame, zoom, leadingInset = 0, headerHandlers, selection, sheet, selectAllLabel, showSheetHeaders = false, children }: {
  frame?: DisplayList | null;
  zoom: number;
  /** Space left of the row header for a plugin gutter. */
  leadingInset?: number;
  /** Whether to paint the column and row header bands. */
  showSheetHeaders?: boolean;
  /** Selection, header menu, resize and reorder callbacks; absent leaves the bands inert. */
  headerHandlers?: SheetGridHeaderHandlers;
  /** Current selection, used only to highlight fully selected tracks. */
  selection?: Selection | null;
  /** Tooltip for the select-all corner; with `onSelectAll` it makes the corner live. */
  selectAllLabel?: string;
  /**
   * Identity of the sheet this frame paints. A gesture armed on one sheet never
   * commits on another, and it also needs the band's visible tracks where the
   * press left them: a scroll, a resize or a hide toggle cancels. Without this
   * the frame object stands in for the sheet, so any repaint mid-drag cancels.
   */
  sheet?: string | number | null;
  children: ReactNode;
}) {
  const grid = frame?.grid;
  const columns = headerTracks(grid, 'col');
  const rows = headerTracks(grid, 'row');
  const dividerX = grid?.dividerX;
  const dividerY = grid?.dividerY;
  // commit-time guards read what is painted now, not the render that armed the
  // gesture: a release can arrive after a repaint or a sheet switch.
  const paintedRef = useRef<PaintedFrame>({ frame: frame ?? null, sheet: sheet ?? null });
  paintedRef.current = { frame: frame ?? null, sheet: sheet ?? null };
  const tracksRef = useRef<{ col: HeaderTrack[]; row: HeaderTrack[] }>({ col: columns, row: rows });
  tracksRef.current = { col: columns, row: rows };
  /**
   * Whether a gesture still describes what the band paints: the same sheet (or,
   * unnamed, the same frame object) and the same visible track geometry. A
   * drag armed on one sheet and released on another, or after the window
   * scrolled or a track was resized or hidden, has nothing left to commit.
   */
  const sameOrigin = useCallback(
    (active: GestureOrigin & { axis: 'col' | 'row' }): boolean => {
      const painted = paintedRef.current;
      if (active.sheet === null) return painted.frame === active.frame;
      return painted.sheet === active.sheet &&
        tracksSignature(tracksRef.current[active.axis]) === active.tracks;
    },
    []
  );
  const drag = useRef<ResizeDrag | null>(null);
  const reorder = useRef<ReorderDrag | null>(null);
  const trackDrag = useRef<TrackDrag | null>(null);
  const [resizing, setResizing] = useState<{ axis: 'col' | 'row'; index: number; size: number; startSize: number } | null>(null);
  // `pointer` is the drag position along the band in css px; the dragged run
  // follows it so the gesture reads as lifting the tracks out of the band.
  // `runFrom`/`runTo` are the captured run in sheet indices, so every label of
  // the run lifts and a scrolled band lifts the right ones.
  const [reordering, setReordering] = useState<
    { axis: 'col' | 'row'; gap: number; fromIndex: number; runFrom: number; runTo: number; pointer: number } | null
  >(null);
  const [hover, setHover] = useState<{ axis: 'col' | 'row'; index: number } | null>(null);
  const range = selection ? normalizeRange(selection) : null;
  // the band follows the pointer while dragging: the dragged track takes the new
  // size and every later track shifts by the same delta, as in a spreadsheet.
  const preview = (axis: 'col' | 'row', tracks: HeaderTrack[]) => resizing?.axis === axis
    ? resizePreview(tracks, resizing.index, resizing.size - resizing.startSize)
    : tracks;

  // A gesture that outlives its band (the pointer left, or the tracks were
  // repainted mid-drag) must still end, so release runs on the window too.
  const endDrag = useCallback((commit: boolean) => {
    const active = drag.current;
    if (!active) return;
    drag.current = null;
    setResizing(null);
    if (active.host.hasPointerCapture?.(active.pointerId)) {
      active.host.releasePointerCapture(active.pointerId);
    }
    if (!commit || !active.moved || !active.host.isConnected) return;
    // `isConnected` only says the band is still in the document; the sheet it
    // paints, and the tracks inside it, may have moved under the press. A
    // scroll would otherwise write this drag's size onto whatever track now
    // sits at the armed slot.
    if (!sameOrigin(active)) return;
    headerHandlers?.onHeaderResize?.(active.axis, active.index, clampTrackSize(active.axis, active.size), active.run);
  }, [headerHandlers, sameOrigin]);

  const endReorder = useCallback((commit: boolean) => {
    const active = reorder.current;
    if (!active) return;
    reorder.current = null;
    setReordering(null);
    if (active.host.hasPointerCapture?.(active.pointerId)) {
      active.host.releasePointerCapture(active.pointerId);
    }
    // a press that never passed the slop is a click: hand it back to the
    // surface so it collapses the selection to the pressed track, the way a
    // plain header press does. A cancelled gesture never selects.
    if (active.gap === null) {
      // the click belongs to the sheet and band that took the press: a switch
      // or a repaint that moved a track would select the wrong thing.
      if (commit && active.press && sameOrigin(active)) {
        headerHandlers?.onHeaderMouseDown?.(active.axis, active.press.index, active.press.event);
      }
      return;
    }
    if (!commit) return;
    // the gap was measured against the armed tracks; a sheet switch, or a
    // repaint that moved them, no longer maps that slot, so the move is
    // dropped instead of landing on whatever track now sits under the pointer.
    if (!sameOrigin(active)) return;
    const tracks = tracksRef.current[active.axis === 'col' ? 'col' : 'row'];
    if (tracks.length === 0) return;
    // the gap counts the tracks behind the pointer, so the track painted at
    // that slot is the sheet track the run lands immediately before; past the
    // last painted track the run lands after every visible one. Both are sheet
    // indices, which is what the move op takes.
    const before = active.gap < tracks.length
      ? tracks[active.gap].index
      : tracks[tracks.length - 1].index + 1;
    // the run is the selection's sheet span, captured at the press, so its
    // leading tracks being hidden or scrolled off changes nothing here.
    const at = active.run.from;
    const count = active.run.to - active.run.from + 1;
    // a drop on the run's leading edge, on its trailing edge, or anywhere
    // strictly inside it leaves the sheet unchanged, so no move is posted.
    if (at <= before && before <= at + count) return;
    headerHandlers?.onHeaderReorder?.(active.axis, at, before, count);
  }, [headerHandlers, sameOrigin]);

  // a selection drag commits nothing: the last extend the surface saw is the
  // selection, so the release only drops the capture.
  const endTrackDrag = useCallback(() => {
    const active = trackDrag.current;
    if (!active) return;
    trackDrag.current = null;
    if (active.host.hasPointerCapture?.(active.pointerId)) {
      active.host.releasePointerCapture(active.pointerId);
    }
  }, []);

  // pointerup can land outside the band (or on a band that unmounted), and a
  // capture the browser dropped silently would leave the last drag armed.
  useEffect(() => {
    const release = (event: PointerEvent) => {
      const active = drag.current;
      if (active && event.pointerId === active.pointerId) endDrag(true);
      const moving = reorder.current;
      if (moving && event.pointerId === moving.pointerId) endReorder(true);
      const selecting = trackDrag.current;
      if (selecting && event.pointerId === selecting.pointerId) endTrackDrag();
    };
    const cancel = () => {
      endDrag(false);
      endReorder(false);
      endTrackDrag();
    };
    window.addEventListener('pointerup', release);
    window.addEventListener('pointercancel', cancel);
    return () => {
      window.removeEventListener('pointerup', release);
      window.removeEventListener('pointercancel', cancel);
    };
  }, [endDrag, endReorder, endTrackDrag]);

  const bandPointerDown = useCallback(
    (axis: 'col' | 'row', tracks: HeaderTrack[], event: ReactPointerEvent<HTMLDivElement>) => {
      headerHandlers?.onHeaderPress?.();
      if (event.button !== 0) return;
      const band = event.currentTarget;
      const box = band.getBoundingClientRect();
      const along = (axis === 'col' ? event.clientX - box.left : event.clientY - box.top) / zoom;
      // a boundary is only a resize when the surface can commit one: with no
      // resize handler a viewer press near the slop must fall through and
      // select the track instead of starting a gesture nothing can finish.
      const boundary = headerHandlers?.onHeaderResize ? resizeGestureAt(tracks, along, zoom) : null;
      if (boundary) {
        // a boundary press is a resize, not a track press: stop it here so the
        // label underneath never sees it, and capture the pointer so the drag
        // survives leaving the band.
        event.preventDefault();
        event.stopPropagation();
        drag.current = {
          ...paintedRef.current,
          tracks: tracksSignature(tracks),
          axis,
          index: boundary.index,
          size: boundary.size,
          // read off `selection` at the press: a boundary inside a whole-track
          // run resizes the whole run, like a spreadsheet.
          run: resizeRun(axis, boundary.index, range),
          host: band,
          pointerId: event.pointerId,
          start: axis === 'col' ? event.clientX : event.clientY,
          startSize: boundary.size,
          moved: false,
        };
        if (band.setPointerCapture) band.setPointerCapture(event.pointerId);
        return;
      }
      // away from a boundary, a label press either reorders or drags the header
      // selection, the way a spreadsheet does: only a track the selection
      // already covers in full can be picked up, and a modifier press always
      // extends instead. Anything else drags the selection. The press is kept
      // off the band's ancestors and captured here either way; on a track the
      // selection does not already cover, the label's own mousedown still runs
      // and is what establishes the selection anchor.
      const track = trackAtOffset(tracks, along);
      if (!track) return;
      const extendsSelection = event.shiftKey || event.metaKey || event.ctrlKey;
      if (!extendsSelection && headerHandlers?.onHeaderReorder &&
          trackSelected(axis, track.index, range)) {
        event.stopPropagation();
        // the run is read off `selection` here, at the press, before anything
        // else can change it: the label's mousedown is suppressed while this
        // gesture is armed, so the painted run stays the drag's subject.
        reorder.current = {
          ...paintedRef.current,
          tracks: tracksSignature(tracks),
          axis,
          fromIndex: track.index,
          run: reorderRun(axis, track.index, range),
          host: band,
          pointerId: event.pointerId,
          start: axis === 'col' ? event.clientX : event.clientY,
          gap: null,
          press: null,
        };
        if (band.setPointerCapture) band.setPointerCapture(event.pointerId);
        return;
      }
      if (!headerHandlers?.onHeaderDragSelect) return;
      event.stopPropagation();
      trackDrag.current = {
        ...paintedRef.current,
        tracks: tracksSignature(tracks),
        axis, host: band, pointerId: event.pointerId, index: track.index,
      };
      if (band.setPointerCapture) band.setPointerCapture(event.pointerId);
    },
    [headerHandlers, range, zoom]
  );

  const bandPointerMove = useCallback(
    (axis: 'col' | 'row', tracks: HeaderTrack[], event: ReactPointerEvent<HTMLDivElement>) => {
      const active = drag.current;
      if (active && active.axis === axis) {
        const along = axis === 'col' ? event.clientX : event.clientY;
        if (!active.moved && Math.abs(along - active.start) < HEADER_RESIZE_SLOP_PX) return;
        active.moved = true;
        active.size = active.startSize + (along - active.start) / zoom;
        setResizing({ axis, index: active.index, size: active.size, startSize: active.startSize });
        return;
      }
      const moving = reorder.current;
      if (moving && moving.axis === axis) {
        // the sheet must still be the one the press landed on and its tracks
        // must sit where they did: a repaint alone (hover, a selection echo)
        // keeps the drag, a scroll or a sheet switch drops it rather than
        // resolving the lifted slot against new tracks.
        if (!sameOrigin(moving)) {
          endReorder(false);
          return;
        }
        // the dragged label must still be painted: a hide toggle or a scroll
        // inside the same sheet leaves no track to lift.
        if (tracks.findIndex((track) => track.index === moving.fromIndex) < 0) {
          endReorder(false);
          return;
        }
        const along = axis === 'col' ? event.clientX : event.clientY;
        if (moving.gap === null && Math.abs(along - moving.start) < HEADER_REORDER_SLOP_PX) return;
        const box = moving.host.getBoundingClientRect();
        const along_px = axis === 'col' ? event.clientX - box.left : event.clientY - box.top;
        moving.gap = reorderGapAt(tracks, along_px / zoom);
        setReordering({ axis, gap: moving.gap, fromIndex: moving.fromIndex,
          runFrom: moving.run.from, runTo: moving.run.to, pointer: along_px });
        return;
      }
      // a selection drag extends track by track: the same index again is not a
      // new extend, and the callback only fires where the pointer crosses into
      // a different track. Once the band scrolls under the press the index the
      // pointer reaches is no longer the one the drag was measured against.
      const selecting = trackDrag.current;
      if (!selecting || selecting.axis !== axis) return;
      if (!sameOrigin(selecting)) {
        endTrackDrag();
        return;
      }
      const box = selecting.host.getBoundingClientRect();
      const along_px = axis === 'col' ? event.clientX - box.left : event.clientY - box.top;
      const track = trackAtOffset(tracks, along_px / zoom);
      if (!track || track.index === selecting.index) return;
      selecting.index = track.index;
      headerHandlers?.onHeaderDragSelect?.(axis, track.index);
    },
    [endReorder, endTrackDrag, headerHandlers, sameOrigin, zoom]
  );

  const bandPointerUp = useCallback(
    (axis: 'col' | 'row', event: ReactPointerEvent<HTMLDivElement>) => {
      const active = drag.current;
      if (active && active.axis === axis && event.pointerId === active.pointerId) endDrag(true);
      const moving = reorder.current;
      if (moving && moving.axis === axis && event.pointerId === moving.pointerId) endReorder(true);
      const selecting = trackDrag.current;
      if (selecting && selecting.axis === axis && event.pointerId === selecting.pointerId) endTrackDrag();
    },
    [endDrag, endReorder, endTrackDrag]
  );

  const labelHandlers = (axis: 'col' | 'row', index: number) => {
    const mouseDown = headerHandlers?.onHeaderMouseDown;
    return {
      onPointerEnter: () => setHover({ axis, index }),
      onPointerLeave: () => setHover((current) =>
        current?.axis === axis && current.index === index ? null : current),
      // a press already armed as a reorder must not also select the track: the
      // surface's press collapses a multi-track selection to the pressed track,
      // and the run the drag is moving would lose its highlight mid-gesture.
      // The armed run stays as it is, so the pressed label reads as the handle
      // of the run and the band keeps showing everything the drag will move. A
      // press on a track outside a selected run still selects normally.
      onMouseDown: mouseDown
        ? (event: ReactMouseEvent) => {
            const armed = reorder.current;
            if (armed) {
              // the reorder owns this press; remember it so a release without a
              // drag can replay it as the plain click it was.
              if (!armed.press) armed.press = { index, event };
              return;
            }
            mouseDown(axis, index, event);
          }
        : undefined,
    };
  };

  // The canvas strokes a gridline centred on the track boundary; a border inside
  // the label would end half a pixel short and read as a one-pixel seam against
  // the cells. These dividers sit on the same coordinates as the canvas stroke.
  // A boundary a compacted hidden span collapsed onto is left to `hiddenDividers`
  // below, which draws the doubled separator in its place.
  const trackDividers = (axis: 'col' | 'row', tracks: HeaderTrack[]) => {
    const marked = hiddenBoundaries(grid, axis, tracks);
    return tracks.map((track, position) => {
      if (marked.includes(position + 1)) return null;
      const at = Math.round((track.offset + track.size) * zoom) - 1;
      return <span key={`${axis}:divider:${track.index}`} aria-hidden data-track-divider
        style={axis === 'col'
          ? { position: 'absolute', left: at, top: 0, width: 1, height: COLUMN_HEADER_HEIGHT,
              background: '#e0e0e0', pointerEvents: 'none' }
          : { position: 'absolute', left: 0, top: at, width: ROW_HEADER_WIDTH, height: 1,
              background: '#e0e0e0', pointerEvents: 'none' }} />;
    });
  };

  // Excel marks a compacted hidden run with two hairlines straddling the boundary
  // the run collapsed onto: the pair, not a single line, tells a gap apart from an
  // ordinary edge. Drawn after the track dividers (so no gridline sits between the
  // pair) but below the split and the drag guides.
  const hiddenDividers = (axis: 'col' | 'row', tracks: HeaderTrack[]) =>
    hiddenBoundaries(grid, axis, tracks).flatMap((position) => {
      const at = Math.round(tracks[position].offset * zoom);
      return ([-1, 1] as const).map((side) => <span
        key={`${axis}:hidden:${tracks[position].index}:${side}`}
        aria-hidden data-testid="xlsx-hidden-divider" data-axis={axis}
        style={axis === 'col'
          ? { position: 'absolute', left: at + side * 2, top: 0, width: 1,
              height: COLUMN_HEADER_HEIGHT, background: '#9aa0a6', pointerEvents: 'none' }
          : { position: 'absolute', left: 0, top: at + side * 2, width: ROW_HEADER_WIDTH,
              height: 1, background: '#9aa0a6', pointerEvents: 'none' }} />);
    });

  // A fully selected track marks itself with a two-pixel accent on the edge that
  // faces the cells. It is drawn over the band edge (which stays a hairline for
  // every other track) so the accent replaces that edge instead of stacking a
  // second rule under the label.
  const trackUnderlines = (axis: 'col' | 'row', tracks: HeaderTrack[]) =>
    tracks.filter((track) => trackSelected(axis, track.index, range)).map((track) => <span
      key={`${axis}:underline:${track.index}`} aria-hidden data-testid="xlsx-header-underline"
      data-axis={axis}
      style={axis === 'col'
        ? { position: 'absolute', left: track.offset * zoom, bottom: 0, width: track.size * zoom,
            height: 2, background: '#9aa0a6', pointerEvents: 'none', zIndex: 2 }
        : { position: 'absolute', top: track.offset * zoom, right: 0, height: track.size * zoom,
            width: 2, background: '#9aa0a6', pointerEvents: 'none', zIndex: 2 }} />);

  const resizeStrips = (axis: 'col' | 'row', tracks: HeaderTrack[]) => {
    if (!headerHandlers?.onHeaderResize) return null;
    const vertical = axis === 'col';
    // the last track has no successor, so it has no boundary strip.
    return tracks.slice(0, -1).map((track) => <div
      key={`${axis}:${track.index}`}
      data-resize-axis={axis}
      data-resize-index={track.index}
      aria-hidden
      style={vertical
        ? { position: 'absolute', left: (track.offset + track.size) * zoom - HEADER_RESIZE_ZONE_PX,
            top: COLUMN_RESIZE_STRIP_TOP, width: HEADER_RESIZE_ZONE_PX * 2,
            height: COLUMN_HEADER_HEIGHT - COLUMN_RESIZE_STRIP_TOP, cursor: 'col-resize' }
        : { position: 'absolute', left: ROW_RESIZE_STRIP_LEFT,
            top: (track.offset + track.size) * zoom - HEADER_RESIZE_ZONE_PX,
            width: ROW_HEADER_WIDTH - ROW_RESIZE_STRIP_LEFT, height: HEADER_RESIZE_ZONE_PX * 2,
            cursor: 'row-resize' }}
    />);
  };

  // the bands sit outside the scroller, so a wheel over them would otherwise be
  // a dead zone: hand it to the sheet's own scroll viewport.
  const bandWheel = (event: ReactWheelEvent<HTMLDivElement>) => {
    const scroller = event.currentTarget.parentElement?.querySelector<HTMLElement>('[data-testid="xlsx-scroll"]');
    if (!scroller) return;
    scroller.scrollTop += event.deltaY;
    scroller.scrollLeft += event.deltaX;
  };

  /** Boundary of the dragged track while previewing, in band pixels. */
  const resizeGuideAt = (axis: 'col' | 'row') => {
    if (!resizing || resizing.axis !== axis) return null;
    const track = preview(axis, axis === 'col' ? columns : rows).find((entry) => entry.index === resizing.index);
    return track ? Math.round((track.offset + track.size) * zoom) : null;
  };

  /** The insertion gap while dragging a label, in band pixels. */
  const reorderGuideAt = (axis: 'col' | 'row') => {
    if (reordering?.axis !== axis) return null;
    const tracks = preview(axis, axis === 'col' ? columns : rows);
    const before = reordering.gap > 0 ? tracks[reordering.gap - 1] : null;
    return before ? Math.round((before.offset + before.size) * zoom) : 0;
  };

  // Three levels, all neutral: a whole-track selection is filled and underlined,
  // a selection that merely reaches the track is filled lightly, hover is filled
  // lighter still. The cells themselves stay unfilled; only the bands mark them.
  const labelStyle = (axis: 'col' | 'row', index: number, placed: CSSProperties): CSSProperties => {
    const full = trackSelected(axis, index, range);
    const touched = !full && trackTouched(axis, index, range);
    const hovered = hover?.axis === axis && hover.index === index;
    // the captured run lifts, not just the pressed track: the labels are known
    // by sheet index, so every painted track inside `runFrom..runTo` rises.
    const lifting = reordering?.axis === axis ? reordering : null;
    const lifted = lifting !== null && lifting.runFrom <= index && index <= lifting.runTo;
    const pressed = lifting && lifted
      ? (axis === 'col' ? columns : rows).find((track) => track.index === lifting.fromIndex)
      : undefined;
    // the run rides the pointer as one block: the pressed track centres under it
    // and every label of the run keeps its spacing around that track. The labels
    // sit above their neighbours without a shadow of their own — adjacent
    // labels would cast onto each other — while the line names the drop.
    const travel = lifting && pressed
      ? lifting.pointer - (pressed.offset + pressed.size / 2) * zoom
      : 0;
    return {
      ...placed,
      backgroundColor: lifted ? '#ffffff' : full ? '#e3e7ea' : hovered ? '#f1f3f4' : touched ? '#eceff1' : undefined,
      fontWeight: full || lifted ? 600 : undefined,
      color: full || lifted ? '#3c4043' : undefined,
      zIndex: lifted ? 3 : undefined,
      transform: lifted ? `translate${axis === 'col' ? 'X' : 'Y'}(${travel}px)` : undefined,
      cursor: lifted ? 'grabbing' : undefined,
    };
  };

  const bandEvents = (axis: 'col' | 'row', tracks: HeaderTrack[]) =>
    headerHandlers?.onHeaderResize || headerHandlers?.onHeaderReorder || headerHandlers?.onHeaderDragSelect ||
      headerHandlers?.onHeaderPress
      ? {
          onPointerDown: (event: ReactPointerEvent<HTMLDivElement>) => bandPointerDown(axis, tracks, event),
          onPointerMove: (event: ReactPointerEvent<HTMLDivElement>) => bandPointerMove(axis, tracks, event),
          onPointerUp: (event: ReactPointerEvent<HTMLDivElement>) => bandPointerUp(axis, event),
          onPointerCancel: () => {
            if (drag.current?.axis === axis) endDrag(false);
            if (reorder.current?.axis === axis) endReorder(false);
            if (trackDrag.current?.axis === axis) endTrackDrag();
          },
        }
      : {};

  const columnResizeGuide = resizeGuideAt('col');
  const rowResizeGuide = resizeGuideAt('row');
  const columnReorderGuide = reorderGuideAt('col');
  const rowReorderGuide = reorderGuideAt('row');

  if (!showSheetHeaders) return <>{children}</>;

  return <div
    data-testid="xlsx-grid"
    style={{
      display: 'grid',
      gridTemplateColumns: `${ROW_HEADER_WIDTH}px minmax(0, 1fr)`,
      gridTemplateRows: `${COLUMN_HEADER_HEIGHT}px minmax(0, 1fr)`,
      flex: 1,
      minWidth: 0,
      minHeight: 0,
      marginLeft: leadingInset === 0 ? undefined : leadingInset,
    }}
  >
    <div
      data-testid="xlsx-grid-corner"
      aria-hidden={headerHandlers?.onSelectAll ? undefined : true}
      role={headerHandlers?.onSelectAll ? 'button' : undefined}
      title={headerHandlers?.onSelectAll ? selectAllLabel : undefined}
      aria-label={headerHandlers?.onSelectAll ? selectAllLabel : undefined}
      onPointerDown={headerHandlers?.onSelectAll
        ? (event) => { event.stopPropagation(); headerHandlers.onHeaderPress?.(); }
        : undefined}
      onMouseDown={headerHandlers?.onSelectAll
        ? (event) => { event.stopPropagation(); }
        : undefined}
      onClick={headerHandlers?.onSelectAll ? () => headerHandlers.onSelectAll?.() : undefined}
      style={{
        ...bandBase,
        gridColumn: 1,
        gridRow: 1,
        position: 'relative',
        borderRight: '1px solid #e0e0e0',
        borderBottom: '1px solid #e0e0e0',
        cursor: headerHandlers?.onSelectAll ? 'pointer' : undefined,
      }}
    >
      {headerHandlers?.onSelectAll && <span
        aria-hidden
        style={{
          position: 'absolute', right: 5, bottom: 4, width: 0, height: 0,
          borderTop: '4px solid transparent', borderLeft: '4px solid transparent',
          borderRight: '4px solid #70757a', borderBottom: '4px solid #70757a',
        }}
      />}
    </div>
    <div
      data-testid="xlsx-column-headers"
      aria-hidden
      style={{ ...bandBase, gridColumn: 2, gridRow: 1, position: 'relative', overflow: 'hidden' }}
      onWheel={bandWheel}
      {...bandEvents('col', columns)}
    >
      {preview('col', columns).map((column) => <div
        key={column.index}
        data-col={column.index}
         {...labelHandlers('col', column.index)}
        style={labelStyle('col', column.index, {
          ...cellBase,
          left: column.offset * zoom,
          top: 0,
          width: column.size * zoom,
          height: COLUMN_HEADER_HEIGHT,
        })}
      >
        {column.size * zoom >= 12 ? columnLetter(column.index) : null}
      </div>)}
      {trackDividers('col', preview('col', columns))}
      {hiddenDividers('col', preview('col', columns))}
      {/* the split replaces the boundary line it sits on: drawn after the track
          dividers so the two-pixel split covers the one-pixel gridline, as the
          canvas does. */}
      {dividerX !== undefined && <div
        data-testid="xlsx-freeze-divider"
        data-axis="col"
        aria-hidden
        style={freezeDividerStyle('col', dividerX, zoom)}
      />}
      {resizeStrips('col', columns)}
      {columnResizeGuide !== null && <div
        data-testid="xlsx-resize-guide"
        aria-hidden
        style={{ position: 'absolute', left: columnResizeGuide, top: 0, width: 1,
          height: COLUMN_HEADER_HEIGHT, background: '#5f6368', pointerEvents: 'none' }}
      />}
      {columnReorderGuide !== null && <div
        data-testid="xlsx-reorder-guide"
        aria-hidden
        style={{ position: 'absolute', left: columnReorderGuide, top: 0, width: 2,
          height: COLUMN_HEADER_HEIGHT, background: '#5f6368', pointerEvents: 'none', zIndex: 4 }}
      />}
      {/* the band's edge facing the cells: the canvas starts its own grid inside
          the viewport, so this line belongs to the chrome. */}
      <span aria-hidden data-testid="xlsx-header-edge" data-axis="col"
        style={{ position: 'absolute', left: 0, right: 0, bottom: 0, height: 1,
          background: '#e0e0e0', pointerEvents: 'none' }} />
      {trackUnderlines('col', preview('col', columns))}
    </div>
    <div
      data-testid="xlsx-row-headers"
      aria-hidden
      style={{ ...bandBase, gridColumn: 1, gridRow: 2, position: 'relative', overflow: 'hidden' }}
      onWheel={bandWheel}
      {...bandEvents('row', rows)}
    >
      {preview('row', rows).map((row) => <div
        key={row.index}
        data-row={row.index}
         {...labelHandlers('row', row.index)}
        style={labelStyle('row', row.index, {
          ...cellBase,
          left: 0,
          top: row.offset * zoom,
          width: ROW_HEADER_WIDTH,
          height: row.size * zoom,
        })}
      >
        {row.size * zoom >= 10 ? row.index + 1 : null}
      </div>)}
      {trackDividers('row', preview('row', rows))}
      {hiddenDividers('row', preview('row', rows))}
      {dividerY !== undefined && <div
        data-testid="xlsx-freeze-divider"
        data-axis="row"
        aria-hidden
        style={freezeDividerStyle('row', dividerY, zoom)}
      />}
      {resizeStrips('row', rows)}
      {rowResizeGuide !== null && <div
        data-testid="xlsx-resize-guide"
        aria-hidden
        style={{ position: 'absolute', left: 0, top: rowResizeGuide, width: ROW_HEADER_WIDTH,
          height: 1, background: '#5f6368', pointerEvents: 'none' }}
      />}
      {rowReorderGuide !== null && <div
        data-testid="xlsx-reorder-guide"
        aria-hidden
        style={{ position: 'absolute', left: 0, top: rowReorderGuide, width: ROW_HEADER_WIDTH,
          height: 2, background: '#5f6368', pointerEvents: 'none', zIndex: 4 }}
      />}
      <span aria-hidden data-testid="xlsx-header-edge" data-axis="row"
        style={{ position: 'absolute', left: 'auto', right: 0, top: 0, bottom: 0, width: 1,
          background: '#e0e0e0', pointerEvents: 'none' }} />
      {trackUnderlines('row', preview('row', rows))}
    </div>
    <div style={{
      gridColumn: 2,
      gridRow: 2,
      position: 'relative',
      display: 'flex',
      flexDirection: 'column',
      minWidth: 0,
      minHeight: 0,
    }}>
      {children}
    </div>
  </div>;
}
