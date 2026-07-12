use super::input::{FloatSegmentIn, FloatZoneIn};

pub(super) const MIN_WRAP_SEGMENT_WIDTH: f32 = 24.0;

pub(super) struct LineMargins {
    pub left: f32,
    pub right: f32,
    pub segments: Option<Vec<FloatSegmentIn>>,
}

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
            return LineMargins {
                left: 0.0,
                right: 0.0,
                segments: Some(vec![FloatSegmentIn {
                    left_offset: 0.0,
                    available_width: 0.0,
                }]),
            };
        }
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

pub(super) fn available_width(margins: &LineMargins, base_width: f32) -> f32 {
    match &margins.segments {
        Some(segments) => segments.iter().map(|s| s.available_width).sum(),
        None => base_width - margins.left - margins.right,
    }
}

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
