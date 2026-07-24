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
//! Text bands derive from the CSS font size. Caret stops use UTF-16 grapheme
//! boundaries for text primitives and shaped cluster boundaries for glyph
//! primitives, with physical edges mapped through the run direction.

use crate::display_list::{DisplayList, DocAttrs, HfRegion, Primitive, TableCellRef};
use serde::Serialize;
use std::collections::BTreeMap;
use unicode_segmentation::UnicodeSegmentation;

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

/// a positioned text-bearing primitive flattened to shared hit geometry.
struct TextHit<'a> {
    attrs: &'a DocAttrs,
    x: f64,
    width: f64,
    baseline: f64,
    top: f64,
    bottom: f64,
    doc_start: i64,
    doc_end: i64,
    caret_stops: Vec<CaretStop>,
}

#[derive(Clone, Copy)]
struct CaretStop {
    x: f64,
    position: i64,
}

fn is_rtl(attrs: &DocAttrs, rtl: Option<bool>) -> bool {
    rtl == Some(true) || attrs.bidi_level.is_some_and(|level| level % 2 == 1)
}

fn doc_position_at_utf16(doc_start: i64, doc_end: i64, utf16_offset: i64, utf16_len: i64) -> i64 {
    if utf16_offset <= 0 {
        doc_start
    } else if utf16_offset >= utf16_len {
        doc_end
    } else {
        (doc_start + utf16_offset).min(doc_end)
    }
}

fn text_caret_stops(
    text: &str,
    x: f64,
    width: f64,
    rtl: bool,
    doc_start: i64,
    doc_end: i64,
) -> Vec<CaretStop> {
    let mut utf16_boundaries = Vec::with_capacity(text.graphemes(true).count() + 1);
    let mut utf16_len = 0_i64;
    utf16_boundaries.push(utf16_len);
    for cluster in text.graphemes(true) {
        utf16_len += cluster.encode_utf16().count() as i64;
        utf16_boundaries.push(utf16_len);
    }
    let cluster_count = utf16_boundaries.len().saturating_sub(1);
    if cluster_count == 0 || width <= 0.0 {
        return vec![CaretStop {
            x,
            position: if rtl { doc_end } else { doc_start },
        }];
    }
    (0..=cluster_count)
        .map(|visual_index| {
            let logical_index = if rtl {
                cluster_count - visual_index
            } else {
                visual_index
            };
            CaretStop {
                x: x + width * visual_index as f64 / cluster_count as f64,
                position: doc_position_at_utf16(
                    doc_start,
                    doc_end,
                    utf16_boundaries[logical_index],
                    utf16_len,
                ),
            }
        })
        .collect()
}

fn glyph_caret_stops(
    text: &str,
    glyphs: &[crate::display_list::PlacedGlyph],
    rtl: bool,
    doc_start: i64,
    doc_end: i64,
) -> Vec<CaretStop> {
    let mut bounds: BTreeMap<usize, (f64, f64)> = BTreeMap::new();
    for glyph in glyphs {
        let start = glyph.x.min(glyph.x + glyph.advance);
        let end = glyph.x.max(glyph.x + glyph.advance);
        let entry = bounds
            .entry(glyph.cluster as usize)
            .or_insert((f64::INFINITY, f64::NEG_INFINITY));
        entry.0 = entry.0.min(start);
        entry.1 = entry.1.max(end);
    }
    let clusters: Vec<(usize, f64, f64)> = bounds
        .into_iter()
        .filter_map(|(byte, (left, right))| {
            text.is_char_boundary(byte).then_some((byte, left, right))
        })
        .collect();
    let utf16_len = text.encode_utf16().count() as i64;
    let mut stops = Vec::with_capacity(clusters.len() * 2);
    for (index, (byte_start, left, right)) in clusters.iter().copied().enumerate() {
        let byte_end = clusters
            .get(index + 1)
            .map_or(text.len(), |(byte, _, _)| *byte);
        let Some(prefix) = text.get(..byte_start) else {
            continue;
        };
        let Some(cluster_text) = text.get(byte_start..byte_end) else {
            continue;
        };
        let logical_start = prefix.encode_utf16().count() as i64;
        let logical_end = logical_start + cluster_text.encode_utf16().count() as i64;
        let start_position = doc_position_at_utf16(doc_start, doc_end, logical_start, utf16_len);
        let end_position = doc_position_at_utf16(doc_start, doc_end, logical_end, utf16_len);
        stops.push(CaretStop {
            x: left,
            position: if rtl { end_position } else { start_position },
        });
        stops.push(CaretStop {
            x: right,
            position: if rtl { start_position } else { end_position },
        });
    }
    stops.sort_by(|left, right| {
        left.x
            .total_cmp(&right.x)
            .then(left.position.cmp(&right.position))
    });
    stops.dedup_by(|left, right| left.x == right.x && left.position == right.position);
    stops
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
                let x = t.x.as_f64().unwrap_or(0.0);
                let width = t.width.as_f64().unwrap_or(0.0);
                out.push(TextHit {
                    attrs: &t.attrs,
                    x,
                    width,
                    baseline,
                    top: baseline - fp,
                    bottom: baseline + fp * 0.25,
                    doc_start: ds,
                    doc_end: de,
                    caret_stops: text_caret_stops(
                        &t.text,
                        x,
                        width,
                        is_rtl(&t.attrs, t.rtl),
                        ds,
                        de,
                    ),
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
                let rtl = is_rtl(&g.attrs, g.rtl);
                let mut caret_stops = glyph_caret_stops(&g.text, &g.glyphs, rtl, ds, de);
                if caret_stops.is_empty() {
                    caret_stops = text_caret_stops(&g.text, min_x, width, rtl, ds, de);
                }
                out.push(TextHit {
                    attrs: &g.attrs,
                    x: min_x,
                    width,
                    baseline,
                    top: baseline - fp,
                    bottom: baseline + fp * 0.25,
                    doc_start: ds,
                    doc_end: de,
                    caret_stops,
                });
            }
            _ => {}
        }
    }
    out
}

fn position_in_run(hit: &TextHit<'_>, x: f64) -> i64 {
    let mut best = hit.caret_stops.first().copied().unwrap_or(CaretStop {
        x: hit.x,
        position: hit.doc_start,
    });
    let mut best_distance = (x - best.x).abs();
    for stop in hit.caret_stops.iter().copied().skip(1) {
        let distance = (x - stop.x).abs();
        if distance < best_distance || (distance == best_distance && stop.x > best.x) {
            best = stop;
            best_distance = distance;
        }
    }
    best.position
}

fn x_at_position(hit: &TextHit<'_>, position: i64) -> f64 {
    hit.caret_stops
        .iter()
        .min_by(|left, right| {
            (left.position - position)
                .abs()
                .cmp(&(right.position - position).abs())
                .then(left.x.total_cmp(&right.x))
        })
        .map_or(hit.x, |stop| stop.x)
}

fn position_and_distance(hit: &TextHit<'_>, x: f64) -> (f64, i64) {
    if x < hit.x {
        (hit.x - x, position_in_run(hit, hit.x))
    } else if x > hit.x + hit.width {
        (
            x - (hit.x + hit.width),
            position_in_run(hit, hit.x + hit.width),
        )
    } else {
        (0.0, position_in_run(hit, x))
    }
}

fn position_for_hits<'hit, 'data>(
    hits: impl IntoIterator<Item = &'hit TextHit<'data>>,
    x: f64,
) -> Option<i64>
where
    'data: 'hit,
{
    let mut best: Option<(f64, i64)> = None;
    for hit in hits {
        let candidate = position_and_distance(hit, x);
        if best.is_none_or(|current| candidate.0 < current.0) {
            best = Some(candidate);
        }
    }
    best.map(|(_, position)| position)
}

fn same_line_owner(left: &TextHit<'_>, right: &TextHit<'_>) -> bool {
    match (&left.attrs.table, &right.attrs.table) {
        (Some(left), Some(right)) if left.table_id != right.table_id => return false,
        (Some(_), None) | (None, Some(_)) => return false,
        _ => {}
    }
    match (&left.attrs.cell, &right.attrs.cell) {
        (Some(left), Some(right))
            if left.row != right.row || left.col != right.col || left.cell_id != right.cell_id =>
        {
            return false;
        }
        (Some(_), None) | (None, Some(_)) => return false,
        _ => {}
    }
    if left.attrs.line_index != right.attrs.line_index
        && (left.attrs.line_index.is_some() || right.attrs.line_index.is_some())
    {
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
        position_for_hits(&self.hits, x).expect("a visual line has at least one hit")
    }

    fn distance_at_x(&self, x: f64) -> f64 {
        self.hits
            .iter()
            .map(|hit| position_and_distance(hit, x).0)
            .fold(f64::INFINITY, f64::min)
    }

    fn left(&self) -> f64 {
        self.hits
            .iter()
            .map(|hit| hit.x)
            .fold(f64::INFINITY, f64::min)
    }

    fn table_cell(&self) -> Option<(&str, &TableCellRef)> {
        let attrs = &self.hits.first()?.attrs;
        Some((&attrs.table.as_ref()?.table_id, attrs.cell.as_ref()?))
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
                    && line.hits.first().is_some_and(|previous| {
                        same_line_owner(previous, &hit)
                            && (previous.attrs.line_index.is_some()
                                || (previous.baseline - hit.baseline).abs() <= LINE_CENTER_EPSILON)
                    })
            });
            if let Some(index) = matching_line {
                lines[index].hits.push(hit);
            } else {
                lines.push(VisualLine {
                    page_index,
                    column_index,
                    hits: vec![hit],
                });
            }
        }
    }
    lines.sort_by(|left, right| {
        left.page_index
            .cmp(&right.page_index)
            .then(left.column_index.cmp(&right.column_index))
            .then(left.center().total_cmp(&right.center()))
            .then(left.left().total_cmp(&right.left()))
    });
    lines
}

fn same_cell(left: &TableCellRef, right: &TableCellRef) -> bool {
    left.row == right.row && left.col == right.col && left.cell_id == right.cell_id
}

fn cells_share_row(left: &TableCellRef, right: &TableCellRef) -> bool {
    let left_end = left.row.saturating_add(left.row_span);
    let right_end = right.row.saturating_add(right.row_span);
    left.row < right_end && right.row < left_end
}

fn table_row_target(
    lines: &[VisualLine<'_>],
    table_id: &str,
    row: u64,
    direction: VerticalDirection,
    goal_x: f64,
) -> Option<usize> {
    let candidates: Vec<usize> = lines
        .iter()
        .enumerate()
        .filter_map(|(index, line)| {
            line.table_cell()
                .filter(|(candidate_table, cell)| *candidate_table == table_id && cell.row == row)
                .map(|_| index)
        })
        .collect();
    let seed = candidates.iter().copied().min_by(|left, right| {
        lines[*left]
            .distance_at_x(goal_x)
            .total_cmp(&lines[*right].distance_at_x(goal_x))
    })?;
    let (_, selected_cell) = lines[seed].table_cell()?;
    candidates
        .into_iter()
        .filter(|index| {
            lines[*index]
                .table_cell()
                .is_some_and(|(_, cell)| same_cell(cell, selected_cell))
        })
        .min_by(|left, right| match direction {
            VerticalDirection::Down => lines[*left].center().total_cmp(&lines[*right].center()),
            VerticalDirection::Up => lines[*right].center().total_cmp(&lines[*left].center()),
        })
}

fn adjacent_visual_line(
    lines: &[VisualLine<'_>],
    current: usize,
    direction: VerticalDirection,
    goal_x: f64,
) -> Option<usize> {
    let current_table = lines[current].table_cell();
    let mut candidate = current;
    loop {
        candidate = match direction {
            VerticalDirection::Up => candidate.checked_sub(1)?,
            VerticalDirection::Down => candidate
                .checked_add(1)
                .filter(|index| *index < lines.len())?,
        };
        let candidate_table = lines[candidate].table_cell();
        match (current_table, candidate_table) {
            (Some((current_id, current_cell)), Some((candidate_id, candidate_cell)))
                if current_id == candidate_id =>
            {
                if same_cell(current_cell, candidate_cell) {
                    return Some(candidate);
                }
                if cells_share_row(current_cell, candidate_cell) {
                    continue;
                }
                return table_row_target(lines, current_id, candidate_cell.row, direction, goal_x);
            }
            (None, Some((table_id, cell))) => {
                return table_row_target(lines, table_id, cell.row, direction, goal_x);
            }
            _ => return Some(candidate),
        }
    }
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
    let target = adjacent_visual_line(&lines, current, direction, goal_x);
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
    position_for_hits(line.iter().copied(), x)
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
        let start = from.max(h.doc_start).min(h.doc_end);
        let end = to.max(h.doc_start).min(h.doc_end);
        let x0 = x_at_position(&h, start);
        let x1 = x_at_position(&h, end);
        out.push(RangeRect {
            page_index,
            x: x0.min(x1),
            y: h.top,
            // degenerate overlaps keep a 1px floor like lineSpanRect
            width: (x1 - x0).abs().max(1.0),
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
    for (page_index, page) in dl.pages.iter().enumerate() {
        let hits = text_hits(&page.primitives);
        if let Some(hit) = hits.iter().find(|hit| {
            (hit.doc_start == hit.doc_end && pos == hit.doc_start)
                || (pos >= hit.doc_start && pos < hit.doc_end)
        }) {
            return Some(CaretRect {
                page_index,
                x: x_at_position(hit, pos),
                y: hit.top,
                height: hit.bottom - hit.top,
            });
        }
        if let Some(image) = page.primitives.iter().find_map(|primitive| {
            let Primitive::Image(image) = primitive else {
                return None;
            };
            let (Some(start), Some(end)) = (image.attrs.doc_start, image.attrs.doc_end) else {
                return None;
            };
            (pos >= start && pos < end).then_some(image)
        }) {
            return Some(CaretRect {
                page_index,
                x: image.x.as_f64().unwrap_or(0.0),
                y: image.y.as_f64().unwrap_or(0.0),
                height: image.h.as_f64().unwrap_or(0.0),
            });
        }
    }
    for (page_index, page) in dl.pages.iter().enumerate().rev() {
        if let Some(image) = page.primitives.iter().rev().find_map(|primitive| {
            let Primitive::Image(image) = primitive else {
                return None;
            };
            let (Some(start), Some(end)) = (image.attrs.doc_start, image.attrs.doc_end) else {
                return None;
            };
            (pos > start && pos <= end).then_some(image)
        }) {
            return Some(CaretRect {
                page_index,
                x: image.x.as_f64().unwrap_or(0.0) + image.w.as_f64().unwrap_or(0.0),
                y: image.y.as_f64().unwrap_or(0.0),
                height: image.h.as_f64().unwrap_or(0.0),
            });
        }
        let hits = text_hits(&page.primitives);
        if let Some(hit) = hits
            .iter()
            .rev()
            .find(|hit| pos > hit.doc_start && pos <= hit.doc_end)
        {
            return Some(CaretRect {
                page_index,
                x: x_at_position(hit, pos),
                y: hit.top,
                height: hit.bottom - hit.top,
            });
        }
    }
    None
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
