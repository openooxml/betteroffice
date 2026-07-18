import type { WrapTextDirection } from '../pagination/types';

/**
 * Minimum horizontal room a side must offer before we treat it as usable for
 * text wrapping. Below this, the would-be segment is treated as a no-go and
 * the float falls back to single-side wrap. Without this guard, an image
 * flush with the right margin produces a 2-px right segment that the painter
 * cannot fit text into, and the segments path then bypasses leftMargin /
 * rightMargin composition with co-occurring floats — text overlaps the image.
 *
 * Reused by `layoutFloatingTable` (decide if a floating table is effectively
 * block-like) and by the paragraph measurer (decide if a line should be bumped
 * past obstructing floats). Keep these usages in sync.
 */
export const MIN_WRAP_SEGMENT_WIDTH = 24;

/**
 * Clamp floating-wrap margins so at least a usable strip of the line
 * remains. Unchecked margins from near-full-width tables/images can exceed
 * contentWidth and collapse every line to ~1 glyph (the "single character
 * per line after a wide floating table" bug).
 *
 * Returned margins are zeroed when:
 * - either side alone is >= contentWidth (no strip on that side at all), or
 * - their sum is >= contentWidth (no strip exists between the two sides).
 */
export function clampFloatingWrapMargins(
  leftMargin: number,
  rightMargin: number,
  contentWidth: number
): { leftMargin: number; rightMargin: number } {
  const cw = Math.max(1, contentWidth);
  const lm = Math.max(0, leftMargin);
  const rm = Math.max(0, rightMargin);
  if (lm >= cw || rm >= cw || lm + rm >= cw) {
    return { leftMargin: 0, rightMargin: 0 };
  }
  return { leftMargin: lm, rightMargin: rm };
}

export interface FloatingExclusionRect {
  /** which margin absorbs the exclusion when wrapText alone doesn't pick a side. */
  side: 'left' | 'right';
  /** object's left edge in content-area px. */
  x: number;
  /** object's top edge in content-area px. */
  y: number;
  width: number;
  height: number;
  distTop: number;
  distBottom: number;
  distLeft: number;
  distRight: number;
  wrapText?: WrapTextDirection;
  wrapType?: string;
}

export interface FloatingImageZone {
  leftMargin: number;
  rightMargin: number;
  topY: number;
  bottomY: number;
  segments?: FloatingLineSegmentZone[];
  /**
   * Full-width vertical band (OOXML `topAndBottom` wrap): no text fits beside
   * it, so any line overlapping `[topY, bottomY]` is pushed below the band.
   */
  fullWidthBlock?: boolean;
}

export interface FloatingLineSegmentZone {
  leftOffset: number;
  availableWidth: number;
}

export interface FloatingLineMargins {
  leftMargin: number;
  rightMargin: number;
  segments?: FloatingLineSegmentZone[];
}

export function rectsToFloatingZones(
  rects: FloatingExclusionRect[],
  contentWidth: number
): FloatingImageZone[] {
  return rects.map((rect) => {
    const rectLeft = rect.x - rect.distLeft;
    const rectRight = rect.x + rect.width + rect.distRight;
    const rectTop = rect.y - rect.distTop;
    const rectBottom = rect.y + rect.height + rect.distBottom;

    // topAndBottom: full-width band, no side wrap. Lines overlapping the band
    // flow above/below it (handled in getFloatingMargins / findClearLineY).
    if (rect.wrapType === 'topAndBottom') {
      return {
        leftMargin: 0,
        rightMargin: 0,
        topY: rectTop,
        bottomY: rectBottom,
        fullWidthBlock: true,
      };
    }

    let leftMargin = 0;
    let rightMargin = 0;
    let segments: FloatingLineSegmentZone[] | undefined;

    const wrapText = rect.wrapText ?? 'bothSides';

    if (wrapText === 'right') {
      leftMargin = leftObjectMargin(rectRight);
    } else if (wrapText === 'left') {
      rightMargin = rightObjectMargin(rectLeft, contentWidth);
    } else if (wrapText === 'largest') {
      ({ leftMargin, rightMargin } = largestSideMargins(rectLeft, rectRight, contentWidth));
    } else if (canSplitCenteredBothSidesWrap(rectLeft, rectRight, contentWidth)) {
      segments = centeredWrapSegments(rectLeft, rectRight, contentWidth);
    } else if (rect.side === 'left') {
      leftMargin = leftObjectMargin(rectRight);
    } else {
      rightMargin = rightObjectMargin(rectLeft, contentWidth);
    }

    // Clamp margins that exceed contentWidth (near-full-width floats whose
    // outer edge sits past the content area). Without this, body text after
    // the float collapses to ~1 glyph per line. Segments-based wrapping
    // (centered both-sides) already keeps leftMargin/rightMargin at 0, so
    // the clamp is a no-op there.
    const clamped = clampFloatingWrapMargins(leftMargin, rightMargin, contentWidth);
    return {
      leftMargin: clamped.leftMargin,
      rightMargin: clamped.rightMargin,
      topY: rectTop,
      bottomY: rectBottom,
      segments,
    };
  });
}

export function getFloatingAvailableWidth(margins: FloatingLineMargins, baseWidth: number): number {
  const segmentWidth = margins.segments?.reduce((sum, segment) => sum + segment.availableWidth, 0);
  return segmentWidth ?? baseWidth - margins.leftMargin - margins.rightMargin;
}

export function getFloatingMargins(
  lineY: number,
  lineHeight: number,
  zones: FloatingImageZone[] | undefined,
  paragraphYOffset: number
): FloatingLineMargins {
  if (!zones || zones.length === 0) {
    return { leftMargin: 0, rightMargin: 0 };
  }

  let leftMargin = 0;
  let rightMargin = 0;
  let segments: FloatingLineSegmentZone[] | undefined;

  const absoluteLineTop = paragraphYOffset + lineY;
  const absoluteLineBottom = absoluteLineTop + lineHeight;

  for (const zone of zones) {
    if (absoluteLineBottom <= zone.topY || absoluteLineTop >= zone.bottomY) continue;
    if (zone.fullWidthBlock) {
      // No room beside a full-width band: a zero-width segment forces the
      // available width to 0, which findClearLineY uses to push the line below.
      return { leftMargin: 0, rightMargin: 0, segments: [{ leftOffset: 0, availableWidth: 0 }] };
    }
    if (zone.segments?.length) {
      segments = segments ? intersectSegments(segments, zone.segments) : zone.segments;
      continue;
    }
    leftMargin = Math.max(leftMargin, zone.leftMargin);
    rightMargin = Math.max(rightMargin, zone.rightMargin);
  }

  return { leftMargin, rightMargin, segments };
}

/**
 * Find the next vertical position at or below `startY` where the available
 * text width is at least `minWidth`. Used to skip lines past stacked floats
 * when there is no horizontal room for meaningful text at the current Y.
 *
 * Returns `startY` if the current position already has enough room, otherwise
 * the lowest `bottomY` of any zone currently obstructing the line. The caller
 * is expected to re-query margins at the returned Y.
 */
export function findClearLineY(
  startY: number,
  lineHeight: number,
  zones: FloatingImageZone[] | undefined,
  contentWidth: number,
  minWidth: number
): number {
  if (!zones || zones.length === 0) return startY;

  let y = startY;
  // Bounded loop — at most one step per zone the line currently overlaps,
  // plus a safety cushion. Prevents pathological re-entry while keeping the
  // happy path O(zones).
  for (let i = 0; i < zones.length + 2; i++) {
    const margins = getFloatingMargins(y, lineHeight, zones, 0);
    const width = getFloatingAvailableWidth(margins, contentWidth);
    if (width >= minWidth) return y;

    const lineBottom = y + lineHeight;
    let nextY = Infinity;
    for (const zone of zones) {
      if (lineBottom <= zone.topY || y >= zone.bottomY) continue;
      if (zone.bottomY > y && zone.bottomY < nextY) {
        nextY = zone.bottomY;
      }
    }
    if (!Number.isFinite(nextY) || nextY <= y) return y;
    y = nextY;
  }
  return y;
}

function intersectSegments(
  a: FloatingLineSegmentZone[],
  b: FloatingLineSegmentZone[]
): FloatingLineSegmentZone[] {
  const result: FloatingLineSegmentZone[] = [];
  for (const left of a) {
    for (const right of b) {
      const start = Math.max(left.leftOffset, right.leftOffset);
      const end = Math.min(
        left.leftOffset + left.availableWidth,
        right.leftOffset + right.availableWidth
      );
      if (end > start) {
        result.push({ leftOffset: start, availableWidth: end - start });
      }
    }
  }
  return result;
}

function canSplitCenteredBothSidesWrap(
  rectLeft: number,
  rectRight: number,
  contentWidth: number
): boolean {
  return rectLeft > MIN_WRAP_SEGMENT_WIDTH && rectRight + MIN_WRAP_SEGMENT_WIDTH < contentWidth;
}

function centeredWrapSegments(
  rectLeft: number,
  rectRight: number,
  contentWidth: number
): FloatingLineSegmentZone[] {
  return [
    { leftOffset: 0, availableWidth: Math.max(0, rectLeft) },
    {
      leftOffset: Math.max(0, rectRight),
      availableWidth: Math.max(0, contentWidth - rectRight),
    },
  ].filter((segment) => segment.availableWidth > 1);
}

function largestSideMargins(
  rectLeft: number,
  rectRight: number,
  contentWidth: number
): Pick<FloatingLineMargins, 'leftMargin' | 'rightMargin'> {
  const leftWidth = Math.max(0, rectLeft);
  const rightWidth = Math.max(0, contentWidth - rectRight);
  return rightWidth >= leftWidth
    ? { leftMargin: leftObjectMargin(rectRight), rightMargin: 0 }
    : { leftMargin: 0, rightMargin: rightObjectMargin(rectLeft, contentWidth) };
}

function leftObjectMargin(rectRight: number): number {
  return Math.max(0, rectRight);
}

function rightObjectMargin(rectLeft: number, contentWidth: number): number {
  return Math.max(0, contentWidth - rectLeft);
}
