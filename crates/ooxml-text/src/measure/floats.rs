//! Per-line float exclusion-zone geometry — a port of the measurement
//! subset of `packages/core/src/layout/measure/floatingZones.ts`
//! (`getFloatingMargins`, `getFloatingAvailableWidth`, `findClearLineY`,
//! `intersectSegments`), so the Rust line filler resolves the same margins,
//! segment strips, and skip-below-float hops the TS wrap loop does.
//!
//! Zone *extraction* (objects → zones) stays on the TS side
//! (`measureBlocksPipeline.ts`); this module only consumes the extracted
//! zones the host passes through in the [`super::MeasureInput`] envelope.

use super::input::{FloatSegmentIn, FloatZoneIn};

/// TS `MIN_WRAP_SEGMENT_WIDTH` (`layout/pagination/floatingObjects.ts`): the
/// minimum horizontal room a strip must offer before a line uses it; below
/// this the line hops past the obstructing floats instead.
pub(super) const MIN_WRAP_SEGMENT_WIDTH: f32 = 24.0;

/// TS `FloatingLineMargins`: the zone-resolved context for one line probe.
pub(super) struct LineMargins {
    pub left: f32,
    pub right: f32,
    /// `Some` mirrors TS `segments !== undefined` — it may hold an empty
    /// vec (an intersection that collapsed), which still wins over the
    /// margin math in [`available_width`] exactly like TS's `?? ` chain.
    pub segments: Option<Vec<FloatSegmentIn>>,
}

/// TS `getFloatingMargins` (floatingZones.ts): probe every zone against the
/// line's absolute Y interval `[paragraphYOffset + lineY, … + lineHeight)`.
/// Edges are exclusive on both sides (`lineBottom <= topY || lineTop >=
/// bottomY` skips). A `fullWidthBlock` zone returns early with the synthetic
/// zero-width segment; a segment zone intersects into the running strip set
/// and contributes no margins; plain zones max their margins in.
pub(super) fn floating_margins(
    line_y: f32,
    line_height: f32,
    zones: &[FloatZoneIn],
    paragraph_y_offset: f32,
) -> LineMargins {
    let mut left = 0.0f32;
    let mut right = 0.0f32;
    let mut segments: Option<Vec<FloatSegmentIn>> = None;

    let line_top = paragraph_y_offset + line_y;
    let line_bottom = line_top + line_height;

    for zone in zones {
        if line_bottom <= zone.top_y || line_top >= zone.bottom_y {
            continue;
        }
        if zone.full_width_block {
            // No room beside a full-width band: a zero-width segment forces
            // the available width to 0, which find_clear_line_y uses to push
            // the line below (early return, like TS).
            return LineMargins {
                left: 0.0,
                right: 0.0,
                segments: Some(vec![FloatSegmentIn {
                    left_offset: 0.0,
                    available_width: 0.0,
                }]),
            };
        }
        // TS `zone.segments?.length` — an empty segment list falls through
        // to the margin path.
        if let Some(zone_segments) = zone.segments.as_deref().filter(|s| !s.is_empty()) {
            segments = Some(match segments {
                Some(acc) => intersect_segments(&acc, zone_segments),
                None => zone_segments.to_vec(),
            });
            continue;
        }
        left = left.max(zone.left_margin);
        right = right.max(zone.right_margin);
    }

    LineMargins {
        left,
        right,
        segments,
    }
}

/// TS `getFloatingAvailableWidth`: the segment-strip sum wins whenever
/// segments are present (0 for an empty set), else base − left − right.
pub(super) fn available_width(margins: &LineMargins, base_width: f32) -> f32 {
    match &margins.segments {
        Some(segments) => segments.iter().map(|s| s.available_width).sum(),
        None => base_width - margins.left - margins.right,
    }
}

/// TS `findClearLineY`: the next Y at or below `start_y` where the available
/// width reaches `min_width`, stepping to the lowest `bottomY` of any zone
/// obstructing the current probe. Bounded to `zones.len() + 2` steps like
/// the TS loop. `start_y` is already absolute (the TS call site passes
/// `paragraphYOffset + cumulativeHeight` and probes with offset 0).
pub(super) fn find_clear_line_y(
    start_y: f32,
    line_height: f32,
    zones: &[FloatZoneIn],
    content_width: f32,
    min_width: f32,
) -> f32 {
    if zones.is_empty() {
        return start_y;
    }

    let mut y = start_y;
    for _ in 0..zones.len() + 2 {
        let margins = floating_margins(y, line_height, zones, 0.0);
        if available_width(&margins, content_width) >= min_width {
            return y;
        }

        let line_bottom = y + line_height;
        let mut next_y = f32::INFINITY;
        for zone in zones {
            if line_bottom <= zone.top_y || y >= zone.bottom_y {
                continue;
            }
            if zone.bottom_y > y && zone.bottom_y < next_y {
                next_y = zone.bottom_y;
            }
        }
        if !next_y.is_finite() || next_y <= y {
            return y;
        }
        y = next_y;
    }
    y
}

/// TS `intersectSegments`: pairwise strip overlaps, in input order.
fn intersect_segments(a: &[FloatSegmentIn], b: &[FloatSegmentIn]) -> Vec<FloatSegmentIn> {
    let mut result = Vec::new();
    for left in a {
        for right in b {
            let start = left.left_offset.max(right.left_offset);
            let end = (left.left_offset + left.available_width)
                .min(right.left_offset + right.available_width);
            if end > start {
                result.push(FloatSegmentIn {
                    left_offset: start,
                    available_width: end - start,
                });
            }
        }
    }
    result
}
