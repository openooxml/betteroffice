//! Registry of float exclusion zones — Rust port of the pagination-side
//! floating-object geometry.
//!
//! TS source: `packages/core/src/layout/pagination/floatingObjects.ts`
//!
//! Exported symbols (1:1 with the TS module):
//! - `MIN_WRAP_SEGMENT_WIDTH` (const)
//! - `BlockedRegion`, `Rect`, `Clearance`, `WrapSide` (parameter shapes)
//! - `FreeSpan`
//! - `FloatRegionRegistry` with `set_layout_context`, `clear`,
//!   `register_floating_object`, `resolve_free_span`, `zones_for_page`
//! - `create_float_region_registry`
//!
//! Anchored (floating) images carve horizontal space out of the lines they
//! overlap; paragraph layout asks this registry, per line, how much width
//! remains and where the line starts.
//!
//! INTEGRATION
//! ===========
//! Measurement happens host-side before the JSON seam: the measure
//! pipeline pre-scans the block list, extracts `FloatingImageZone`s
//! (`{ leftMargin, rightMargin, topY, bottomY, segments?, fullWidthBlock? }`,
//! content-area coordinates relative to the anchor block) from anchored
//! images, floating tables, and floating text boxes, groups co-located floats
//! (paragraph-relative zones merge when Y ranges overlap AND anchors are
//! within 4 blocks; margin/page-relative zones group by identical `topY`;
//! page-pinned full-width bands re-anchor to block 0), and threads the active
//! zones plus cumulative Y into each `measureBlock` call. Consequently the
//! measured-block input JSON arrives with float effects ALREADY BAKED INTO
//! THE LINE MEASURES — per-line widths/offsets reflect the exclusion zones;
//! no `FloatingImageZone` data crosses the seam.
//!
//! What this module needs at pagination time is the page-space claim of each
//! float: `BlockedRegion { id, pageNumber, rect, clearance, wrapSide }`. The
//! integrator registers those from anchored-object placement inside the place
//! loop (rect in page coordinates, clearance = the wp:effectExtent-style wrap
//! distances, wrapSide from the wrap mode) and calls `resolve_free_span` per
//! line / `zones_for_page` when composing a page. `mod` is declared but not
//! yet wired into the spine (`lib.rs` / `place.rs`); the integrator swaps
//! these into place-loop hooks later.
//!
//! Numeric fidelity: all math is f64 with JS `Math.min`/`Math.max` semantics
//! (NaN-propagating, -0-aware) and the exact TS iteration order.
//!
//! Note on tests: the TS module has no dedicated unit-test file (repo tests
//! only re-import `MIN_WRAP_SEGMENT_WIDTH` through `floatingZones.ts`), so
//! the `#[cfg(test)]` suite below mirrors the TS implementation semantics
//! branch by branch instead of porting an existing spec.

use serde::{Deserialize, Serialize};

/// A text segment beside a float narrower than this (px) is not worth
/// wrapping into.
pub const MIN_WRAP_SEGMENT_WIDTH: f64 = 24.0;

/// JS `Math.min`: NaN-propagating, and `Math.min(0, -0) === -0`.
fn js_min(a: f64, b: f64) -> f64 {
    if a.is_nan() || b.is_nan() {
        f64::NAN
    } else if a < b {
        a
    } else if b < a {
        b
    } else if a.is_sign_negative() {
        a
    } else {
        b
    }
}

/// JS `Math.max`: NaN-propagating, and `Math.max(-0, 0) === 0`.
fn js_max(a: f64, b: f64) -> f64 {
    if a.is_nan() || b.is_nan() {
        f64::NAN
    } else if a > b {
        a
    } else if b > a {
        b
    } else if a.is_sign_positive() {
        a
    } else {
        b
    }
}

/// The float's box on the page.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Rect {
    pub x: f64,
    pub y: f64,
    pub width: f64,
    pub height: f64,
}

/// Wrap gap kept clear on each side of the box (wp:effectExtent-style
/// padding).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Clearance {
    pub top: f64,
    pub bottom: f64,
    pub left: f64,
    pub right: f64,
}

/// Which side text wraps on: `Left` = text on left, `Right` = text on right,
/// `Both` = both sides.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum WrapSide {
    Left,
    Right,
    Both,
    None,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct BlockedRegion {
    /// Unique ID for the floating object.
    pub id: String,
    /// Page number (1-indexed).
    pub page_number: u32,
    pub rect: Rect,
    pub clearance: Clearance,
    pub wrap_side: WrapSide,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FreeSpan {
    /// Available width for text.
    pub span: f64,
    /// X offset from normal start position.
    pub shift_x: f64,
}

/// Manages floating objects and computes text wrapping exclusions.
#[derive(Debug, Default)]
pub struct FloatRegionRegistry {
    claimed: Vec<BlockedRegion>,
    content_width: f64,
    left_margin: f64,
}

impl FloatRegionRegistry {
    pub fn new() -> Self {
        Self {
            claimed: Vec::new(),
            content_width: 0.0,
            left_margin: 0.0,
        }
    }

    /// Set the layout context (content width and margins).
    pub fn set_layout_context(&mut self, content_width: f64, left_margin: f64) {
        self.content_width = content_width;
        self.left_margin = left_margin;
    }

    /// Forget every registered region (call when starting a new layout).
    pub fn clear(&mut self) {
        self.claimed = Vec::new();
    }

    /// Register a floating object's claimed region.
    pub fn register_floating_object(&mut self, zone: BlockedRegion) {
        self.claimed.push(zone);
    }

    /// Width and start offset left for a line once the floats it crosses have
    /// taken their share.
    ///
    /// `line_y` — Y position of the line (relative to page top);
    /// `line_height` — height of the line; `page_number` — current page
    /// number. Returns available width and X offset for the line.
    pub fn resolve_free_span(&self, line_y: f64, line_height: f64, page_number: u32) -> FreeSpan {
        let band_top = line_y;
        let band_bottom = line_y + line_height;

        // keep the floats on this page whose padded box intersects the line's
        // vertical band — rejecting the two disjoint cases (entirely above,
        // entirely below) is equivalent and cheap
        let crossing_floats: Vec<&BlockedRegion> = self
            .claimed
            .iter()
            .filter(|zone| {
                if zone.page_number != page_number {
                    return false;
                }
                if zone.wrap_side == WrapSide::None {
                    return false;
                }

                let cleared_top = zone.rect.y - zone.clearance.top;
                let cleared_bottom = zone.rect.y + zone.rect.height + zone.clearance.bottom;

                let misses_band = band_bottom <= cleared_top || band_top >= cleared_bottom;
                !misses_band
            })
            .collect();

        if crossing_floats.is_empty() {
            return FreeSpan {
                span: self.content_width,
                shift_x: 0.0,
            };
        }

        // squeeze the writable interval [textStart, textEnd) — in content-area
        // coordinates — from both directions as each float claims its side
        let mut text_start = 0.0_f64;
        let mut text_end = self.content_width;

        for zone in crossing_floats {
            let float_left = zone.rect.x - self.left_margin;
            let float_right = float_left + zone.rect.width;

            if zone.wrap_side == WrapSide::Left || zone.wrap_side == WrapSide::Both {
                // text runs on the left of the float, ending before its cleared edge
                let fence = float_left - zone.clearance.left;
                text_end = js_min(text_end, fence);
            }

            if zone.wrap_side == WrapSide::Right || zone.wrap_side == WrapSide::Both {
                // text runs on the right of the float, starting past its cleared edge
                let fence = float_right + zone.clearance.right;
                text_start = js_max(text_start, fence);
            }
        }

        FreeSpan {
            span: js_max(0.0, text_end - text_start),
            shift_x: text_start,
        }
    }

    /// Every registered region on one page (for rendering).
    pub fn zones_for_page(&self, page_number: u32) -> Vec<&BlockedRegion> {
        self.claimed
            .iter()
            .filter(|z| z.page_number == page_number)
            .collect()
    }
}

/// Create a new `FloatRegionRegistry` instance.
pub fn create_float_region_registry() -> FloatRegionRegistry {
    FloatRegionRegistry::new()
}

#[cfg(test)]
mod tests {
    use super::*;

    fn region(
        id: &str,
        page_number: u32,
        rect: (f64, f64, f64, f64),
        clearance: (f64, f64, f64, f64),
        wrap_side: WrapSide,
    ) -> BlockedRegion {
        BlockedRegion {
            id: id.to_string(),
            page_number,
            rect: Rect {
                x: rect.0,
                y: rect.1,
                width: rect.2,
                height: rect.3,
            },
            clearance: Clearance {
                top: clearance.0,
                bottom: clearance.1,
                left: clearance.2,
                right: clearance.3,
            },
            wrap_side,
        }
    }

    fn registry(content_width: f64, left_margin: f64) -> FloatRegionRegistry {
        let mut r = create_float_region_registry();
        r.set_layout_context(content_width, left_margin);
        r
    }

    #[test]
    fn min_wrap_segment_width_matches_ts() {
        assert_eq!(MIN_WRAP_SEGMENT_WIDTH, 24.0);
    }

    #[test]
    fn empty_registry_returns_full_content_width() {
        let r = registry(500.0, 96.0);
        let fs = r.resolve_free_span(100.0, 16.0, 1);
        assert_eq!(
            fs,
            FreeSpan {
                span: 500.0,
                shift_x: 0.0
            }
        );
    }

    #[test]
    fn float_on_another_page_is_ignored() {
        let mut r = registry(500.0, 0.0);
        r.register_floating_object(region(
            "img1",
            2,
            (100.0, 50.0, 100.0, 100.0),
            (0.0, 0.0, 0.0, 0.0),
            WrapSide::Right,
        ));
        let fs = r.resolve_free_span(60.0, 16.0, 1);
        assert_eq!(
            fs,
            FreeSpan {
                span: 500.0,
                shift_x: 0.0
            }
        );
    }

    #[test]
    fn wrap_side_none_is_ignored() {
        let mut r = registry(500.0, 0.0);
        r.register_floating_object(region(
            "img1",
            1,
            (100.0, 50.0, 100.0, 100.0),
            (0.0, 0.0, 0.0, 0.0),
            WrapSide::None,
        ));
        let fs = r.resolve_free_span(60.0, 16.0, 1);
        assert_eq!(
            fs,
            FreeSpan {
                span: 500.0,
                shift_x: 0.0
            }
        );
    }

    #[test]
    fn line_entirely_above_or_below_cleared_box_is_unaffected() {
        let mut r = registry(500.0, 0.0);
        // cleared band with top/bottom clearance: [50-10, 150+20] = [40, 170]
        r.register_floating_object(region(
            "img1",
            1,
            (100.0, 50.0, 100.0, 100.0),
            (10.0, 20.0, 0.0, 0.0),
            WrapSide::Right,
        ));
        // entirely above: bandBottom == clearedTop → misses (<= boundary)
        assert_eq!(r.resolve_free_span(24.0, 16.0, 1).span, 500.0);
        // entirely below: bandTop == clearedBottom → misses (>= boundary)
        assert_eq!(r.resolve_free_span(170.0, 16.0, 1).span, 500.0);
        // crossing the clearance-extended band (not the raw rect) is blocked
        let fs = r.resolve_free_span(41.0, 8.0, 1);
        assert_eq!(
            fs,
            FreeSpan {
                span: 300.0,
                shift_x: 200.0
            }
        );
    }

    #[test]
    fn wrap_left_fences_text_end_before_cleared_left_edge() {
        let mut r = registry(500.0, 0.0);
        r.register_floating_object(region(
            "img1",
            1,
            (300.0, 0.0, 150.0, 100.0),
            (0.0, 0.0, 12.0, 12.0),
            WrapSide::Left,
        ));
        // text runs left of the float: textEnd = 300 - 12 = 288
        let fs = r.resolve_free_span(10.0, 16.0, 1);
        assert_eq!(
            fs,
            FreeSpan {
                span: 288.0,
                shift_x: 0.0
            }
        );
    }

    #[test]
    fn wrap_right_shifts_text_start_past_cleared_right_edge() {
        let mut r = registry(500.0, 0.0);
        r.register_floating_object(region(
            "img1",
            1,
            (0.0, 0.0, 150.0, 100.0),
            (0.0, 0.0, 12.0, 12.0),
            WrapSide::Right,
        ));
        // text runs right of the float: textStart = 150 + 12 = 162
        let fs = r.resolve_free_span(10.0, 16.0, 1);
        assert_eq!(
            fs,
            FreeSpan {
                span: 338.0,
                shift_x: 162.0
            }
        );
    }

    #[test]
    fn wrap_both_applies_both_fences_from_one_float() {
        let mut r = registry(500.0, 0.0);
        r.register_floating_object(region(
            "img1",
            1,
            (200.0, 0.0, 100.0, 100.0),
            (0.0, 0.0, 5.0, 7.0),
            WrapSide::Both,
        ));
        // textEnd = 200 - 5 = 195; textStart = 300 + 7 = 307
        // span = max(0, 195 - 307) = 0; shiftX = 307
        let fs = r.resolve_free_span(10.0, 16.0, 1);
        assert_eq!(
            fs,
            FreeSpan {
                span: 0.0,
                shift_x: 307.0
            }
        );
    }

    #[test]
    fn left_margin_translates_page_x_into_content_coordinates() {
        let mut r = registry(500.0, 96.0);
        // rect.x is page-absolute; floatLeft = 296 - 96 = 200
        r.register_floating_object(region(
            "img1",
            1,
            (296.0, 0.0, 100.0, 100.0),
            (0.0, 0.0, 0.0, 0.0),
            WrapSide::Left,
        ));
        let fs = r.resolve_free_span(10.0, 16.0, 1);
        assert_eq!(
            fs,
            FreeSpan {
                span: 200.0,
                shift_x: 0.0
            }
        );
    }

    #[test]
    fn multiple_floats_squeeze_from_both_directions() {
        let mut r = registry(500.0, 0.0);
        // right-wrapping float on the left edge: textStart = 100
        r.register_floating_object(region(
            "left-img",
            1,
            (0.0, 0.0, 100.0, 100.0),
            (0.0, 0.0, 0.0, 0.0),
            WrapSide::Right,
        ));
        // left-wrapping float on the right edge: textEnd = 400
        r.register_floating_object(region(
            "right-img",
            1,
            (400.0, 0.0, 100.0, 100.0),
            (0.0, 0.0, 0.0, 0.0),
            WrapSide::Left,
        ));
        let fs = r.resolve_free_span(10.0, 16.0, 1);
        assert_eq!(
            fs,
            FreeSpan {
                span: 300.0,
                shift_x: 100.0
            }
        );
    }

    #[test]
    fn tightest_fence_wins_across_overlapping_floats() {
        let mut r = registry(500.0, 0.0);
        r.register_floating_object(region(
            "a",
            1,
            (0.0, 0.0, 100.0, 100.0),
            (0.0, 0.0, 0.0, 0.0),
            WrapSide::Right,
        ));
        r.register_floating_object(region(
            "b",
            1,
            (0.0, 20.0, 150.0, 100.0),
            (0.0, 0.0, 0.0, 10.0),
            WrapSide::Right,
        ));
        // textStart = max(100, 160) = 160
        let fs = r.resolve_free_span(30.0, 16.0, 1);
        assert_eq!(
            fs,
            FreeSpan {
                span: 340.0,
                shift_x: 160.0
            }
        );
    }

    #[test]
    fn span_is_clamped_at_zero_when_fences_cross() {
        let mut r = registry(200.0, 0.0);
        r.register_floating_object(region(
            "wide",
            1,
            (-50.0, 0.0, 400.0, 100.0),
            (0.0, 0.0, 12.0, 12.0),
            WrapSide::Right,
        ));
        // textStart = 362 > textEnd = 200 → span 0, shiftX stays 362
        let fs = r.resolve_free_span(10.0, 16.0, 1);
        assert_eq!(
            fs,
            FreeSpan {
                span: 0.0,
                shift_x: 362.0
            }
        );
    }

    #[test]
    fn clear_forgets_every_region() {
        let mut r = registry(500.0, 0.0);
        r.register_floating_object(region(
            "img1",
            1,
            (0.0, 0.0, 100.0, 100.0),
            (0.0, 0.0, 0.0, 0.0),
            WrapSide::Right,
        ));
        r.clear();
        assert_eq!(
            r.resolve_free_span(10.0, 16.0, 1),
            FreeSpan {
                span: 500.0,
                shift_x: 0.0
            }
        );
        assert!(r.zones_for_page(1).is_empty());
    }

    #[test]
    fn zones_for_page_filters_by_page_and_keeps_wrap_none() {
        let mut r = registry(500.0, 0.0);
        let z1 = region(
            "p1-a",
            1,
            (0.0, 0.0, 10.0, 10.0),
            (0.0, 0.0, 0.0, 0.0),
            WrapSide::Right,
        );
        let z2 = region(
            "p1-b",
            1,
            (0.0, 0.0, 10.0, 10.0),
            (0.0, 0.0, 0.0, 0.0),
            WrapSide::None,
        );
        let z3 = region(
            "p2-a",
            2,
            (0.0, 0.0, 10.0, 10.0),
            (0.0, 0.0, 0.0, 0.0),
            WrapSide::Left,
        );
        r.register_floating_object(z1.clone());
        r.register_floating_object(z2.clone());
        r.register_floating_object(z3.clone());

        let page1 = r.zones_for_page(1);
        assert_eq!(page1.len(), 2);
        assert_eq!(*page1[0], z1);
        assert_eq!(*page1[1], z2);
        assert_eq!(r.zones_for_page(2), vec![&z3]);
        assert!(r.zones_for_page(3).is_empty());
    }

    #[test]
    fn blocked_region_deserializes_from_ts_camel_case_json() {
        let json = r#"{
            "id": "img1",
            "pageNumber": 1,
            "rect": { "x": 296.0, "y": 50.0, "width": 100.0, "height": 80.0 },
            "clearance": { "top": 0.0, "bottom": 0.0, "left": 12.0, "right": 12.0 },
            "wrapSide": "both"
        }"#;
        let zone: BlockedRegion = serde_json::from_str(json).unwrap();
        assert_eq!(zone.page_number, 1);
        assert_eq!(zone.wrap_side, WrapSide::Both);
        assert_eq!(zone.rect.x, 296.0);
        assert_eq!(zone.clearance.left, 12.0);

        let fs = FreeSpan {
            span: 300.0,
            shift_x: 100.0,
        };
        let out = serde_json::to_value(&fs).unwrap();
        assert_eq!(out["span"], 300.0);
        assert_eq!(out["shiftX"], 100.0);
    }

    #[test]
    fn js_min_max_mirror_math_semantics() {
        assert!(js_min(1.0, f64::NAN).is_nan());
        assert!(js_max(f64::NAN, 1.0).is_nan());
        assert!(js_min(0.0, -0.0).is_sign_negative());
        assert!(js_max(-0.0, 0.0).is_sign_positive());
        assert_eq!(js_min(2.0, 3.0), 2.0);
        assert_eq!(js_max(2.0, 3.0), 3.0);
    }
}
