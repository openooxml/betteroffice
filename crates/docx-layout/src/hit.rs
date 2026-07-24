//! Display-list hit-testing: point -> PM position and PM range -> rects.
//!
//! Behavioral port of the painted-DOM resolvers (`clickToPositionDom` /
//! `getSelectionRectsFromDom` in `packages/core/src/layout/geometry`): a click
//! first tries a direct hit on a text primitive's box, then snaps to the
//! nearest line by vertical center distance and the nearest primitive on it,
//! choosing docStart or docEnd by which side the pointer is on. Range rects
//! emit one rect per overlapped text primitive (proportional sub-span), a
//! 4px sliver for blank-line markers, and the full box for images — the same
//! shapes the DOM version reads off `span[data-doc-start]` elements.
//!
//! Region awareness: [`hit_test_regions`] first tests the page's HF bands
//! (`DisplayPage.header` / `.footer`) and resolves within the winning band's
//! primitives, identifying the region and its `rId` so callers can route the
//! click to that HF PM doc; [`range_rects_in_region`] mirrors that scoping for
//! selection geometry — given a region + rId it resolves the from/to inside
//! that HF band (the same HF doc is painted on every page carrying the part,
//! so it emits one rect-set per such page). [`range_rects`] is the body-only
//! convenience wrapper.
//!
//! Text primitives carry no explicit line box, so vertical bands derive from
//! the CSS font shorthand's px size: ascent ≈ 1.0em above the baseline,
//! descent ≈ 0.25em below. Character positions interpolate proportionally
//! over the run width (the v0 display list distributes advances the same way,
//! so hit-testing and painting agree by construction).

use crate::display_list::{DisplayList, DocAttrs, HfRegion, Primitive};
use serde::Serialize;

/// vertical slack when matching a pointer to a span's band, mirrors the ±4px
/// tolerance in the DOM HF fallback resolver
const BAND_SLACK: f64 = 4.0;

/// width of the selection sliver drawn for a blank line (BLANK_LINE_SELECTION_WIDTH_PX)
const BLANK_LINE_SELECTION_WIDTH: f64 = 4.0;

const LINE_CENTER_EPSILON: f64 = 0.5;

/// parse the px size out of a CSS font shorthand ("700 16px Calibri, ...")
fn font_px(font: &str) -> f64 {
    for token in font.split_whitespace() {
        if let Some(num) = token.strip_suffix("px")
            && let Ok(v) = num.parse::<f64>()
        {
            return v;
        }
    }
    14.666667 // 11pt default
}

/// a positioned text-bearing primitive (a `TextRunPrimitive` or a
/// `GlyphRunPrimitive`) flattened to the geometry the resolvers need. Both
/// resolve through the same proportional-over-width model so hit geometry is
/// invariant to whether the run painted via `fillText` or shaped glyphs — the
/// display-list flag flips rendering, never caret placement.
struct TextHit<'a> {
    text: &'a str,
    attrs: &'a DocAttrs,
    x: f64,
    width: f64,
    baseline: f64,
    top: f64,
    bottom: f64,
    doc_start: i64,
    doc_end: i64,
}

fn text_hits(prims: &[Primitive]) -> Vec<TextHit<'_>> {
    let mut out = Vec::new();
    for p in prims {
        match p {
            Primitive::Text(t) => {
                let (Some(ds), Some(de)) = (t.attrs.doc_start, t.attrs.doc_end) else {
                    continue; // unpositioned re-paints (vmerge continuations) never win a hit
                };
                let fp = font_px(&t.font);
                let baseline = t.baseline_y.as_f64().unwrap_or(0.0);
                out.push(TextHit {
                    text: &t.text,
                    attrs: &t.attrs,
                    x: t.x.as_f64().unwrap_or(0.0),
                    width: t.width.as_f64().unwrap_or(0.0),
                    baseline,
                    top: baseline - fp,
                    bottom: baseline + fp * 0.25,
                    doc_start: ds,
                    doc_end: de,
                });
            }
            Primitive::GlyphRun(g) => {
                let (Some(ds), Some(de)) = (g.attrs.doc_start, g.attrs.doc_end) else {
                    continue;
                };
                if g.glyphs.is_empty() {
                    continue;
                }
                let fp = g.size;
                // baseline / extent are derived from the real glyph geometry:
                // each PlacedGlyph carries its pen advance, so the run's left edge
                // is the min glyph x and its right edge is the trailing glyph's
                // `x + advance` (F3 — no more uniform trailing-advance estimate
                // that drifted ~3px on mixed-font lines). Marks sit above the
                // baseline (smaller y), so the max glyph y is the base baseline.
                let baseline = g
                    .glyphs
                    .iter()
                    .map(|gl| gl.y)
                    .fold(f64::NEG_INFINITY, f64::max);
                let min_x = g.glyphs.iter().map(|gl| gl.x).fold(f64::INFINITY, f64::min);
                let right = g
                    .glyphs
                    .iter()
                    .map(|gl| gl.x + gl.advance)
                    .fold(f64::NEG_INFINITY, f64::max);
                let width = (right - min_x).max(0.0);
                out.push(TextHit {
                    text: &g.text,
                    attrs: &g.attrs,
                    x: min_x,
                    width,
                    baseline,
                    top: baseline - fp,
                    bottom: baseline + fp * 0.25,
                    doc_start: ds,
                    doc_end: de,
                });
            }
            _ => {}
        }
    }
    out
}

/// interpolate a character boundary inside a run: equal per-char advances over
/// the run width (the deterministic stand-in for the DOM's per-glyph bisection)
fn position_in_run(hit: &TextHit<'_>, x: f64) -> i64 {
    let chars = hit.text.chars().count() as i64;
    let span = hit.doc_end - hit.doc_start;
    let units = if chars > 0 { chars } else { span };
    if units <= 0 || hit.width <= 0.0 {
        return hit.doc_start;
    }
    let ratio = ((x - hit.x) / hit.width).clamp(0.0, 1.0);
    let unit = (ratio * units as f64).round() as i64;
    // map char index to PM offset (1:1 for text runs; clamp into the span)
    (hit.doc_start + unit).min(hit.doc_end)
}

fn same_line_owner(left: &TextHit<'_>, right: &TextHit<'_>) -> bool {
    match (&left.attrs.table, &right.attrs.table) {
        (Some(left), Some(right)) => return left.table_id == right.table_id,
        (Some(_), None) | (None, Some(_)) => return false,
        (None, None) => {}
    }
    if left.attrs.cell != right.attrs.cell {
        return false;
    }
    if left.attrs.para_id.is_some() || right.attrs.para_id.is_some() {
        return left.attrs.para_id == right.attrs.para_id;
    }
    if left.attrs.block_key.is_some() || right.attrs.block_key.is_some() {
        return left.attrs.block_key == right.attrs.block_key;
    }
    if left.attrs.block_id.is_some() || right.attrs.block_id.is_some() {
        return left.attrs.block_id == right.attrs.block_id;
    }
    left.doc_end == right.doc_start
}

struct VisualLine<'a> {
    page_index: usize,
    column_index: usize,
    baseline: f64,
    hits: Vec<TextHit<'a>>,
}

impl VisualLine<'_> {
    fn contains_position(&self, position: i64) -> bool {
        self.hits
            .iter()
            .any(|hit| position >= hit.doc_start && position <= hit.doc_end)
    }

    fn center(&self) -> f64 {
        let top = self
            .hits
            .iter()
            .map(|hit| hit.top)
            .fold(f64::INFINITY, f64::min);
        let bottom = self
            .hits
            .iter()
            .map(|hit| hit.bottom)
            .fold(f64::NEG_INFINITY, f64::max);
        (top + bottom) / 2.0
    }

    fn position_at_x(&self, x: f64) -> i64 {
        let mut best: Option<(f64, i64)> = None;
        for hit in &self.hits {
            let (distance, position) = if x < hit.x {
                (hit.x - x, hit.doc_start)
            } else if x > hit.x + hit.width {
                (x - (hit.x + hit.width), hit.doc_end)
            } else {
                (0.0, position_in_run(hit, x))
            };
            if best.is_none_or(|current| distance < current.0) {
                best = Some((distance, position));
            }
        }
        best.expect("a visual line has at least one hit").1
    }
}

fn visual_lines(dl: &DisplayList) -> Vec<VisualLine<'_>> {
    let mut lines: Vec<VisualLine<'_>> = Vec::new();
    for (page_index, page) in dl.pages.iter().enumerate() {
        for hit in text_hits(&page.primitives) {
            let center_x = hit.x + hit.width / 2.0;
            let column_index = page
                .column_bounds
                .iter()
                .position(|bounds| {
                    let x = bounds.x.as_f64().unwrap_or(0.0);
                    let width = bounds.width.as_f64().unwrap_or(0.0);
                    center_x >= x && center_x <= x + width
                })
                .unwrap_or(0);
            let matching_line = lines.iter().position(|line| {
                line.page_index == page_index
                    && line.column_index == column_index
                    && (line.baseline - hit.baseline).abs() <= LINE_CENTER_EPSILON
                    && line
                        .hits
                        .first()
                        .is_some_and(|previous| same_line_owner(previous, &hit))
            });
            if let Some(index) = matching_line {
                lines[index].hits.push(hit);
            } else {
                lines.push(VisualLine {
                    page_index,
                    column_index,
                    baseline: hit.baseline,
                    hits: vec![hit],
                });
            }
        }
    }
    lines.sort_by(|left, right| {
        left.page_index
            .cmp(&right.page_index)
            .then(left.column_index.cmp(&right.column_index))
            .then(left.baseline.total_cmp(&right.baseline))
    });
    lines
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VerticalDirection {
    Up,
    Down,
}

impl VerticalDirection {
    fn parse(direction: &str) -> Result<Self, String> {
        match direction {
            "up" => Ok(Self::Up),
            "down" => Ok(Self::Down),
            other => Err(format!("unknown vertical direction {other:?}")),
        }
    }
}

#[derive(Serialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct VerticalMove {
    pub position: i64,
    pub goal_x: f64,
}

pub fn vertical_move(
    dl: &DisplayList,
    position: i64,
    direction: VerticalDirection,
    goal_x: Option<f64>,
) -> Option<VerticalMove> {
    let caret = caret_rect(dl, position)?;
    let goal_x = goal_x.filter(|x| x.is_finite()).unwrap_or(caret.x);
    let lines = visual_lines(dl);
    let caret_center = caret.y + caret.height / 2.0;
    let current = lines
        .iter()
        .enumerate()
        .filter(|(_, line)| line.page_index == caret.page_index && line.contains_position(position))
        .min_by(|(_, left), (_, right)| {
            (left.center() - caret_center)
                .abs()
                .total_cmp(&(right.center() - caret_center).abs())
        })
        .map(|(index, _)| index)?;
    let target = match direction {
        VerticalDirection::Up => current.checked_sub(1),
        VerticalDirection::Down => current.checked_add(1).filter(|index| *index < lines.len()),
    };
    Some(VerticalMove {
        position: target.map_or(position, |index| lines[index].position_at_x(goal_x)),
        goal_x,
    })
}

/// PM position under a page-local point, or None when the page has no
/// positioned content. Ports the clickToPositionDom resolution order:
/// direct span hit -> image hit -> nearest line -> nearest span on it.
/// Body-only; region-aware callers use [`hit_test_regions`].
pub fn hit_test(dl: &DisplayList, page_index: usize, x: f64, y: f64) -> Option<i64> {
    let page = dl.pages.get(page_index)?;
    resolve_point(&page.primitives, x, y)
}

/// the shared point resolver over one primitive list (a page body or one
/// HF region — both use page coordinates)
fn resolve_point(prims: &[Primitive], x: f64, y: f64) -> Option<i64> {
    let hits = text_hits(prims);

    // 1. direct hit on a text primitive's box (paint order = DOM order)
    for h in &hits {
        if h.width <= 0.0 {
            continue;
        }
        if x >= h.x && x <= h.x + h.width && y >= h.top - BAND_SLACK && y <= h.bottom + BAND_SLACK {
            return Some(position_in_run(h, x));
        }
    }

    // 2. direct hit on an image with a doc position (caret parks before it)
    for p in prims {
        let Primitive::Image(img) = p else { continue };
        let Some(ds) = img.attrs.doc_start else {
            continue;
        };
        let (ix, iy) = (img.x.as_f64().unwrap_or(0.0), img.y.as_f64().unwrap_or(0.0));
        let (iw, ih) = (img.w.as_f64().unwrap_or(0.0), img.h.as_f64().unwrap_or(0.0));
        if x >= ix && x <= ix + iw && y >= iy && y <= iy + ih {
            return Some(ds);
        }
    }

    if hits.is_empty() {
        return None;
    }

    // 3. nearest line by vertical center distance (blank-line markers included)
    let mut best_center = f64::INFINITY;
    for h in &hits {
        let center = (h.top + h.bottom) / 2.0;
        let d = (y - center).abs();
        if d < best_center {
            best_center = d;
        }
    }
    // collect the hits on that nearest band (same center within epsilon)
    let mut line: Vec<&TextHit> = Vec::new();
    for h in &hits {
        let center = (h.top + h.bottom) / 2.0;
        if ((y - center).abs() - best_center).abs() < 0.5 {
            line.push(h);
        }
    }

    // 4. nearest primitive on the line; inside -> interpolate, outside -> snap
    // to the closer edge's position
    let mut best: Option<(f64, i64)> = None;
    for h in &line {
        let (dist, pos) = if x < h.x {
            (h.x - x, h.doc_start)
        } else if x > h.x + h.width {
            (x - (h.x + h.width), h.doc_end)
        } else {
            (0.0, position_in_run(h, x))
        };
        if best.is_none() || dist < best.unwrap().0 {
            best = Some((dist, pos));
        }
    }
    best.map(|(_, pos)| pos)
}

/// region-aware hit result: which part of the page owns the point, and the
/// resolved position inside that part's PM doc. For `header`/`footer` the
/// position refers to the HF ProseMirror doc identified by `rId`, NOT the
/// body doc (the caller must route the selection to that HF editor, the way
/// `usePagesPointer` scopes clicks to `.layout-page-header|footer`).
#[derive(Serialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct RegionHit {
    pub region: HitRegion,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub r_id: Option<String>,
    pub pos: Option<i64>,
}

#[derive(Serialize, Debug, Clone, Copy, PartialEq, Eq)]
pub enum HitRegion {
    #[serde(rename = "body")]
    Body,
    #[serde(rename = "header")]
    Header,
    #[serde(rename = "footer")]
    Footer,
}

/// parse the region discriminant used at the JSON/wasm boundary
/// (`"body" | "header" | "footer"`), the string twin of [`HitRegion`]'s serde
/// rename. Shared by the JSON-arg and session-handle range-rect exports.
pub(crate) fn parse_region(region: &str) -> Result<HitRegion, String> {
    match region {
        "body" => Ok(HitRegion::Body),
        "header" => Ok(HitRegion::Header),
        "footer" => Ok(HitRegion::Footer),
        other => Err(format!("unknown region {other:?}")),
    }
}

/// a point inside an HF band's box resolves within that band (`pos` may be
/// None when the band has no positioned content — the region identification
/// alone is what routes the click into HF editing); everything else falls
/// through to the body resolver. Band membership is the vertical
/// `[y, y+height]` test on the region's box, the display-list analogue of the
/// painted `.layout-page-header` / `.layout-page-footer` hosts.
pub fn hit_test_regions(dl: &DisplayList, page_index: usize, x: f64, y: f64) -> Option<RegionHit> {
    let page = dl.pages.get(page_index)?;

    let in_band = |r: &HfRegion| -> bool {
        let top = r.y.as_f64().unwrap_or(0.0);
        let bottom = top + r.height.as_f64().unwrap_or(0.0);
        y >= top && y <= bottom
    };

    if let Some(h) = &page.header
        && in_band(h)
    {
        return Some(RegionHit {
            region: HitRegion::Header,
            r_id: Some(h.r_id.clone()),
            pos: resolve_point(&h.primitives, x, y),
        });
    }
    if let Some(f) = &page.footer
        && in_band(f)
    {
        return Some(RegionHit {
            region: HitRegion::Footer,
            r_id: Some(f.r_id.clone()),
            pos: resolve_point(&f.primitives, x, y),
        });
    }
    Some(RegionHit {
        region: HitRegion::Body,
        r_id: None,
        pos: resolve_point(&page.primitives, x, y),
    })
}

/// one highlight rectangle of a PM range, page-local coordinates
#[derive(Serialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct RangeRect {
    pub page_index: usize,
    pub x: f64,
    pub y: f64,
    pub width: f64,
    pub height: f64,
}

#[derive(Serialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct CaretRect {
    pub page_index: usize,
    pub x: f64,
    pub y: f64,
    pub height: f64,
}

/// append the highlight rects covering `[from, to)` over ONE primitive list (a
/// page body or one HF band — both page-local) to `out`, stamping `page_index`
/// on each. Shared by the body-only [`range_rects`] and the region-aware
/// [`range_rects_in_region`]: per overlapped text primitive a proportional
/// sub-span rect over the line band, a 4px sliver for blank-line markers, and
/// the full box for images inside the range.
fn collect_range_rects(
    prims: &[Primitive],
    page_index: usize,
    from: i64,
    to: i64,
    out: &mut Vec<RangeRect>,
) {
    for h in text_hits(prims) {
        // blank-line marker: zero-length span selects as a thin sliver
        if h.doc_start == h.doc_end {
            if h.doc_start >= from && h.doc_start < to {
                out.push(RangeRect {
                    page_index,
                    x: h.x,
                    y: h.top,
                    width: BLANK_LINE_SELECTION_WIDTH,
                    height: h.bottom - h.top,
                });
            }
            continue;
        }
        if h.doc_end <= from || h.doc_start >= to {
            continue;
        }
        let span = (h.doc_end - h.doc_start) as f64;
        let start_ratio = ((from - h.doc_start).max(0) as f64 / span).clamp(0.0, 1.0);
        let end_ratio =
            ((to - h.doc_start).min(h.doc_end - h.doc_start) as f64 / span).clamp(0.0, 1.0);
        let x0 = h.x + h.width * start_ratio;
        let x1 = h.x + h.width * end_ratio;
        out.push(RangeRect {
            page_index,
            x: x0,
            y: h.top,
            // degenerate overlaps keep a 1px floor like lineSpanRect
            width: (x1 - x0).max(1.0),
            height: h.bottom - h.top,
        });
    }

    for p in prims {
        let Primitive::Image(img) = p else { continue };
        let (Some(ds), Some(de)) = (img.attrs.doc_start, img.attrs.doc_end) else {
            continue;
        };
        if de <= from || ds >= to {
            continue;
        }
        out.push(RangeRect {
            page_index,
            x: img.x.as_f64().unwrap_or(0.0),
            y: img.y.as_f64().unwrap_or(0.0),
            width: img.w.as_f64().unwrap_or(0.0),
            height: img.h.as_f64().unwrap_or(0.0),
        });
    }
}

/// highlight rectangles for a PM range in the BODY doc across all pages (port of
/// getSelectionRectsFromDom). The `from`/`to` refer to the body PM doc; HF
/// regions (a different PM doc) are never consulted. Region-aware callers use
/// [`range_rects_in_region`].
pub fn range_rects(dl: &DisplayList, from: i64, to: i64) -> Vec<RangeRect> {
    range_rects_in_region(dl, HitRegion::Body, None, from, to)
}

pub fn caret_rect(dl: &DisplayList, pos: i64) -> Option<CaretRect> {
    if let Some(rect) = range_rects(dl, pos, pos.saturating_add(1))
        .into_iter()
        .next()
    {
        return Some(CaretRect {
            page_index: rect.page_index,
            x: rect.x,
            y: rect.y,
            height: rect.height,
        });
    }
    let rect = range_rects(dl, pos.saturating_sub(1), pos)
        .into_iter()
        .next_back()?;
    Some(CaretRect {
        page_index: rect.page_index,
        x: rect.x + rect.width,
        y: rect.y,
        height: rect.height,
    })
}

/// highlight rectangles for a PM range resolved inside a specific page region —
/// the selection-geometry twin of [`hit_test_regions`]'s scoping.
///
/// For [`HitRegion::Body`] this is exactly [`range_rects`] (`r_id` is ignored —
/// the body has one PM doc). For [`HitRegion::Header`] / [`HitRegion::Footer`]
/// the `from`/`to` refer to the header/footer ProseMirror doc identified by
/// `r_id`, so only bands whose `rId` matches are consulted — the display-list
/// analogue of scoping selection to `.layout-page-header` /
/// `.layout-page-footer` for the active HF part. The SAME HF doc is painted on
/// every page carrying the part, so a match emits one rect-set per such page
/// (each stamped with its own `page_index`); the caller renders the page it is
/// editing, exactly as the DOM HF overlay picks the nearest painted host.
///
/// `r_id = None` matches any band of the region kind — reserved for callers that
/// don't disambiguate variants; passing the active part's rId is preferred so a
/// first-page vs default variant on another page never contributes stray rects.
pub fn range_rects_in_region(
    dl: &DisplayList,
    region: HitRegion,
    r_id: Option<&str>,
    from: i64,
    to: i64,
) -> Vec<RangeRect> {
    let (from, to) = (from.min(to), from.max(to));
    let mut rects = Vec::new();
    if from == to {
        return rects;
    }

    let matches = |band: &HfRegion| r_id.is_none_or(|id| id == band.r_id);
    for (page_index, page) in dl.pages.iter().enumerate() {
        let prims: Option<&[Primitive]> = match region {
            HitRegion::Body => Some(&page.primitives),
            HitRegion::Header => page
                .header
                .as_ref()
                .filter(|h| matches(h))
                .map(|h| h.primitives.as_slice()),
            HitRegion::Footer => page
                .footer
                .as_ref()
                .filter(|f| matches(f))
                .map(|f| f.primitives.as_slice()),
        };
        if let Some(prims) = prims {
            collect_range_rects(prims, page_index, from, to, &mut rects);
        }
    }

    rects
}

// ---------------------------------------------------------------------------
// JSON boundary (native-testable; the wasm exports in lib.rs wrap these)
// ---------------------------------------------------------------------------

/// `hit_test` over serialized inputs; returns `"null"` or the position as JSON
pub fn hit_test_json(
    display_list: &str,
    page_index: usize,
    x: f64,
    y: f64,
) -> Result<String, String> {
    let dl: DisplayList = serde_json::from_str(display_list).map_err(|e| format!("parse: {e}"))?;
    match hit_test(&dl, page_index, x, y) {
        Some(pos) => Ok(pos.to_string()),
        None => Ok("null".to_string()),
    }
}

pub fn vertical_move_json(
    display_list: &str,
    position: i64,
    direction: &str,
    goal_x: f64,
) -> Result<String, String> {
    let dl: DisplayList = serde_json::from_str(display_list).map_err(|e| format!("parse: {e}"))?;
    let direction = VerticalDirection::parse(direction)?;
    serde_json::to_string(&vertical_move(
        &dl,
        position,
        direction,
        goal_x.is_finite().then_some(goal_x),
    ))
    .map_err(|e| format!("serialize: {e}"))
}

/// `range_rects` over serialized inputs; returns a JSON array of rects
pub fn range_rects_json(display_list: &str, from: i64, to: i64) -> Result<String, String> {
    let dl: DisplayList = serde_json::from_str(display_list).map_err(|e| format!("parse: {e}"))?;
    serde_json::to_string(&range_rects(&dl, from, to)).map_err(|e| format!("serialize: {e}"))
}

/// `range_rects_in_region` over serialized inputs. `region` is
/// `"body" | "header" | "footer"`; `r_id` is the HF part's relationship id
/// (ignored for `body`; empty string ⇒ match any band of the kind). Returns a
/// JSON array of rects, or a `parse:`/`unknown region` error string.
pub fn range_rects_region_json(
    display_list: &str,
    region: &str,
    r_id: &str,
    from: i64,
    to: i64,
) -> Result<String, String> {
    let dl: DisplayList = serde_json::from_str(display_list).map_err(|e| format!("parse: {e}"))?;
    let region = parse_region(region)?;
    let r_id = if r_id.is_empty() { None } else { Some(r_id) };
    serde_json::to_string(&range_rects_in_region(&dl, region, r_id, from, to))
        .map_err(|e| format!("serialize: {e}"))
}

/// `hit_test_regions` over serialized inputs; returns
/// `{"region":"body"|"header"|"footer","rId"?,"pos":n|null}` or `"null"`
/// for an out-of-range page
pub fn hit_test_regions_json(
    display_list: &str,
    page_index: usize,
    x: f64,
    y: f64,
) -> Result<String, String> {
    let dl: DisplayList = serde_json::from_str(display_list).map_err(|e| format!("parse: {e}"))?;
    match hit_test_regions(&dl, page_index, x, y) {
        Some(hit) => serde_json::to_string(&hit).map_err(|e| format!("serialize: {e}")),
        None => Ok("null".to_string()),
    }
}
