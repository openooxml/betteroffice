//! Display-list hit testing and selection geometry.
//!
//! Two directions, both reading the display list alone and both working in
//! page-local coordinates: a point resolves to a document position, and a
//! document range resolves to highlight rectangles.
//!
//! A point is resolved in a fixed order. A direct hit inside a text or glyph
//! primitive's box wins first (its vertical band is padded by
//! `BAND_SLACK` so a click just above or below the glyphs still lands). Next
//! comes a direct hit on an image that carries a document position, where the
//! caret parks at its start. Failing both, the nearest line by vertical-centre
//! distance is chosen, then the nearest primitive on that line — inside a run
//! the position interpolates between caret stops, outside it snaps to the
//! closer edge.
//!
//! A run's vertical band is derived from its font size: the top sits one font
//! size above the baseline and the bottom a quarter of one below. A glyph run
//! additionally takes its horizontal extent from the real glyph geometry —
//! leftmost `x` to the trailing glyph's `x + advance` — so mixed-font lines do
//! not drift. Combining marks sit above the baseline, so the largest glyph `y`
//! is the base baseline.
//!
//! Caret stops are grapheme boundaries spread across the run's width for text
//! primitives and shaped cluster bounds for glyph runs. In an RTL run visual
//! order is inverted against logical order, so the physical left and right
//! edges map to the logical end and start respectively.
//!
//! Two primitives share a visual line when they agree on the enclosing table
//! and cell, on the paragraph's line index, and then on paragraph id, block key
//! or block id — whichever identity is present. With no identity at all,
//! adjacency of document positions decides.
//!
//! Region scoping matters because a page carries several independent
//! documents. [`hit_test_regions`] lets direct body hits override header and
//! footer bands, then tests their vertical bounds and the page's note
//! areas, resolving inside the winning one and returning the region kind with
//! the part that owns it — an `rId` for a band, a note id for a note, whose
//! story is `fn:{id}` / `en:{id}`. The position then addresses THAT document,
//! never the body. [`range_rects_in_region`] is the selection-geometry twin,
//! scoped by [`RegionScope`], and [`range_rects`] the body-only wrapper. The
//! same header/footer part paints on every page that uses it, so a scoped range
//! query emits one rect set per such page, each stamped with its own page index.
//!
//! Resolving a point also reports what it landed on ([`HoverTarget`]), which a
//! pointer cursor needs and a position cannot give: the nearest-line fallback
//! answers everywhere on a page carrying any text. The target instead names
//! what a click would act on, in the pointer path's own order — a selectable
//! image, then a run's own box, then `in_typeable_area`.

use crate::display_list::{
    DisplayBounds, DisplayList, DisplayPage, DocAttrs, HfRegion, ImagePrimitive, NoteRegion,
    Primitive, ShapePathCommand, ShapePrimitive, TableCellRef, doc_attrs, note_group_id,
};
use serde::Serialize;
use serde_json::{Number, Value};
use std::collections::{BTreeMap, HashMap};
use std::ops::Range;
use unicode_segmentation::UnicodeSegmentation;

#[cfg(test)]
thread_local! {
    static TEXT_HIT_BUILD_COUNT: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
    static LINE_OWNER_COMPARE_COUNT: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
    static CARET_STOPS_BUILD_COUNT: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
    static RANGE_RECT_PAGE_VISITS: std::cell::RefCell<Vec<usize>> = const { std::cell::RefCell::new(Vec::new()) };
}

/// Vertical slack (px) added on each side of a run's band when testing a
/// pointer, so a click in the leading still hits the line.
const BAND_SLACK: f64 = 4.0;

/// Share of a run's font size its glyphs paint above and below the baseline, about the cap
/// height and the descender depth: only there does body text take a header or footer click.
const INK_ASCENT_EM: f64 = 0.7;
const INK_DESCENT_EM: f64 = 0.2;

/// Width (px) of the selection sliver drawn for a blank line, which has no
/// glyphs of its own to highlight.
const BLANK_LINE_SELECTION_WIDTH: f64 = 4.0;

/// Tolerance (px) within which two band centres count as the same line.
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
/// Caret stops are built lazily: most queries walk every text primitive of a
/// region for range overlap but resolve caret x for only the few that
/// intersect, and stop construction is the expensive part.
struct TextHit<'a> {
    attrs: &'a DocAttrs,
    x: f64,
    width: f64,
    baseline: f64,
    top: f64,
    bottom: f64,
    doc_start: i64,
    doc_end: i64,
    stops_source: StopsSource<'a>,
    caret_stops: std::cell::OnceCell<Vec<CaretStop>>,
}

enum StopsSource<'a> {
    Text {
        text: &'a str,
        rtl: bool,
    },
    Glyphs {
        text: &'a str,
        glyphs: &'a [crate::display_list::PlacedGlyph],
        rtl: bool,
    },
}

impl TextHit<'_> {
    fn stops(&self) -> &[CaretStop] {
        self.caret_stops.get_or_init(|| {
            #[cfg(test)]
            CARET_STOPS_BUILD_COUNT.with(|count| count.set(count.get() + 1));
            match self.stops_source {
                StopsSource::Text { text, rtl } => {
                    text_caret_stops(text, self.x, self.width, rtl, self.doc_start, self.doc_end)
                }
                StopsSource::Glyphs { text, glyphs, rtl } => {
                    let stops = glyph_caret_stops(text, glyphs, rtl, self.doc_start, self.doc_end);
                    if stops.is_empty() {
                        text_caret_stops(
                            text,
                            self.x,
                            self.width,
                            rtl,
                            self.doc_start,
                            self.doc_end,
                        )
                    } else {
                        stops
                    }
                }
            }
        })
    }
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

/// Caret stops spread evenly across a run's width, one per grapheme boundary.
/// In an RTL run the visual index counts down the logical boundaries.
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

/// Caret stops taken from shaped cluster bounds, so ligatures and reordered
/// clusters land on real boundaries rather than an even split.
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
    // clusters are byte-ordered and contiguous (each ends where the next
    // starts), so UTF-16 offsets accumulate in one linear pass — a per-cluster
    // prefix rescan would be quadratic in the run length
    let mut logical_cursor = clusters.first().map_or(0, |(byte, _, _)| {
        text.get(..*byte)
            .map_or(0, |prefix| prefix.encode_utf16().count() as i64)
    });
    for (index, (byte_start, left, right)) in clusters.iter().copied().enumerate() {
        let byte_end = clusters
            .get(index + 1)
            .map_or(text.len(), |(byte, _, _)| *byte);
        let Some(cluster_text) = text.get(byte_start..byte_end) else {
            continue;
        };
        let logical_start = logical_cursor;
        let logical_end = logical_start + cluster_text.encode_utf16().count() as i64;
        logical_cursor = logical_end;
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

fn text_doc_range(primitive: &Primitive) -> Option<(i64, i64)> {
    match primitive {
        Primitive::Text(text) => Some((text.attrs.doc_start?, text.attrs.doc_end?)),
        Primitive::GlyphRun(run) if !run.glyphs.is_empty() => {
            Some((run.attrs.doc_start?, run.attrs.doc_end?))
        }
        _ => None,
    }
}

fn text_hit(primitive: &Primitive) -> Option<TextHit<'_>> {
    match primitive {
        Primitive::Text(t) => {
            let (Some(ds), Some(de)) = (t.attrs.doc_start, t.attrs.doc_end) else {
                return None;
            };
            #[cfg(test)]
            TEXT_HIT_BUILD_COUNT.with(|count| count.set(count.get() + 1));
            let fp = font_px(&t.font);
            let baseline = t.baseline_y.as_f64().unwrap_or(0.0);
            let x = t.x.as_f64().unwrap_or(0.0);
            let width = t.width.as_f64().unwrap_or(0.0);
            Some(TextHit {
                attrs: &t.attrs,
                x,
                width,
                baseline,
                top: baseline - fp,
                bottom: baseline + fp * 0.25,
                doc_start: ds,
                doc_end: de,
                stops_source: StopsSource::Text {
                    text: &t.text,
                    rtl: is_rtl(&t.attrs, t.rtl),
                },
                caret_stops: std::cell::OnceCell::new(),
            })
        }
        Primitive::GlyphRun(g) => {
            let (Some(ds), Some(de)) = (g.attrs.doc_start, g.attrs.doc_end) else {
                return None;
            };
            if g.glyphs.is_empty() {
                return None;
            }
            #[cfg(test)]
            TEXT_HIT_BUILD_COUNT.with(|count| count.set(count.get() + 1));
            let fp = g.size;
            // Bounds use glyph positions and advances; marks sit above the base baseline.
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
            Some(TextHit {
                attrs: &g.attrs,
                x: min_x,
                width,
                baseline,
                top: baseline - fp,
                bottom: baseline + fp * 0.25,
                doc_start: ds,
                doc_end: de,
                stops_source: StopsSource::Glyphs {
                    text: &g.text,
                    glyphs: &g.glyphs,
                    rtl,
                },
                caret_stops: std::cell::OnceCell::new(),
            })
        }
        _ => None,
    }
}

fn text_hits(prims: &[Primitive]) -> Vec<TextHit<'_>> {
    prims.iter().filter_map(text_hit).collect()
}

fn position_in_run(hit: &TextHit<'_>, x: f64) -> i64 {
    let stops = hit.stops();
    let mut best = stops.first().copied().unwrap_or(CaretStop {
        x: hit.x,
        position: hit.doc_start,
    });
    let mut best_distance = (x - best.x).abs();
    for stop in stops.iter().copied().skip(1) {
        let distance = (x - stop.x).abs();
        if distance < best_distance || (distance == best_distance && stop.x > best.x) {
            best = stop;
            best_distance = distance;
        }
    }
    best.position
}

fn x_at_position(hit: &TextHit<'_>, position: i64) -> f64 {
    hit.stops()
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

/// Whether two runs belong to the same visual line, checked from strongest to
/// weakest identity and ending at document adjacency.
fn same_line_owner(left: &TextHit<'_>, right: &TextHit<'_>) -> bool {
    #[cfg(test)]
    LINE_OWNER_COMPARE_COUNT.with(|count| count.set(count.get() + 1));
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

#[derive(PartialEq, Eq, Hash)]
enum LineOwner<'a> {
    Para(&'a str),
    BlockKey(&'a str),
    BlockId(&'a Number),
}

#[derive(PartialEq, Eq, Hash)]
struct LineKey<'a> {
    column_index: usize,
    table_id: Option<&'a str>,
    cell: Option<(u64, u64, Option<&'a str>)>,
    line_index: u64,
    owner: LineOwner<'a>,
}

/// Hashable line identity, mirroring [`same_line_owner`]. `None` when grouping
/// falls back to baseline proximity or doc adjacency, which only a scan settles.
fn line_key(attrs: &DocAttrs, column_index: usize) -> Option<LineKey<'_>> {
    let owner = if let Some(para_id) = attrs.para_id.as_deref() {
        LineOwner::Para(para_id)
    } else if let Some(block_key) = attrs.block_key.as_deref() {
        LineOwner::BlockKey(block_key)
    } else {
        LineOwner::BlockId(attrs.block_id.as_ref()?)
    };
    Some(LineKey {
        column_index,
        table_id: attrs.table.as_ref().map(|table| table.table_id.as_str()),
        cell: attrs
            .cell
            .as_ref()
            .map(|cell| (cell.row, cell.col, cell.cell_id.as_deref())),
        line_index: attrs.line_index?,
        owner,
    })
}

fn line_accepts_hit(line: &VisualLine<'_>, column_index: usize, hit: &TextHit<'_>) -> bool {
    line.column_index == column_index
        && line.hits.first().is_some_and(|previous| {
            same_line_owner(previous, hit)
                && (previous.attrs.line_index.is_some()
                    || (previous.baseline - hit.baseline).abs() <= LINE_CENTER_EPSILON)
        })
}

fn visual_lines(dl: &DisplayList, page_range: Range<usize>) -> Vec<VisualLine<'_>> {
    let mut lines: Vec<VisualLine<'_>> = Vec::new();
    for page_index in page_range {
        let Some(page) = dl.pages.get(page_index) else {
            continue;
        };
        // Keyed lines resolve by identity; only hits without one (no block
        // identity, or no line index) scan, and they can only ever match
        // another unkeyed line.
        let mut keyed_lines: HashMap<LineKey<'_>, usize> = HashMap::new();
        let mut unkeyed_lines: Vec<usize> = Vec::new();
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
            let key = line_key(hit.attrs, column_index);
            let matching_line = match &key {
                Some(key) => keyed_lines.get(key).copied(),
                None => unkeyed_lines
                    .iter()
                    .copied()
                    .find(|index| line_accepts_hit(&lines[*index], column_index, &hit)),
            };
            if let Some(index) = matching_line {
                lines[index].hits.push(hit);
            } else {
                match key {
                    Some(key) => {
                        keyed_lines.insert(key, lines.len());
                    }
                    None => unkeyed_lines.push(lines.len()),
                }
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

/// Where a vertical caret move lands, plus the sticky x it should keep.
#[derive(Serialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct VerticalMove {
    pub position: i64,
    pub goal_x: f64,
}

/// Moves the caret one visual line up or down, holding `goal_x` so a run of
/// moves through short lines does not creep inwards.
///
/// Only the caret's page and its immediate neighbours are scanned, so a move
/// can cross a page boundary without walking the whole document. Table cells
/// are traversed cell-first: the next line in the same cell wins, lines in
/// sibling cells of the same row are skipped over, and leaving the cell targets
/// the nearest line of the adjoining row by `goal_x`. When no line lies in the
/// requested direction the position is returned unchanged.
pub fn vertical_move(
    dl: &DisplayList,
    position: i64,
    direction: VerticalDirection,
    goal_x: Option<f64>,
) -> Option<VerticalMove> {
    let caret = caret_rect(dl, position)?;
    let goal_x = goal_x.filter(|x| x.is_finite()).unwrap_or(caret.x);
    let first_page = caret.page_index.saturating_sub(1);
    let last_page = caret.page_index.saturating_add(2).min(dl.pages.len());
    let lines = visual_lines(dl, first_page..last_page);
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

/// Body-only point resolution, in the order described on the module.
/// Region-aware callers use [`hit_test_regions`].
pub fn hit_test(dl: &DisplayList, page_index: usize, x: f64, y: f64) -> Option<i64> {
    let page = dl.pages.get(page_index)?;
    resolve_point(&page.primitives, in_typeable_area(page, x, y), x, y).pos
}

/// What a resolved point landed on: the caret position, plus the thing under
/// the pointer that earned it.
struct PointResolution {
    pos: Option<i64>,
    target: HoverTarget,
    direct: bool,
}

/// What sits under a point, for pointer-cursor feedback.
#[derive(Serialize, Debug, Clone, Copy, PartialEq, Eq)]
pub enum HoverTarget {
    /// typeable text — a run's own box, or the typeable area around it
    #[serde(rename = "text")]
    Text,
    /// a selectable picture, which a click selects instead of typing into
    #[serde(rename = "image")]
    Image,
    #[serde(rename = "none")]
    None,
}

fn in_bounds(bounds: &DisplayBounds, x: f64, y: f64) -> bool {
    let left = bounds.x.as_f64().unwrap_or(0.0);
    let top = bounds.y.as_f64().unwrap_or(0.0);
    x >= left
        && x <= left + bounds.width.as_f64().unwrap_or(0.0)
        && y >= top
        && y <= top + bounds.height.as_f64().unwrap_or(0.0)
}

/// Whether a body point lies in the page's typeable area: the authored content
/// box, so a column gutter counts like the columns it separates. Column boxes
/// are the fallback for a list without a content box, and neither degrades to
/// the runs alone.
///
/// Deliberately blind to content positioned OUTSIDE the content box — a
/// floating text box in a margin. Its glyphs still read as text (a run's own
/// box is tested first), but the blank tail of its lines does not, since no
/// container rect for it exists in the display list. That under-claims, never
/// over-claims: the cursor is an arrow where a click would still type.
fn in_typeable_area(page: &DisplayPage, x: f64, y: f64) -> bool {
    match &page.content_bounds {
        Some(bounds) => in_bounds(bounds, x, y),
        None => page.column_bounds.iter().any(|c| in_bounds(c, x, y)),
    }
}

/// Whether a point lies in a note area — the vertical `[y, y + height]` test a
/// header/footer band gets. An area stating no band, or a zero-height one, owns
/// no point: the list does not say where its notes paint, so claiming the
/// region would route a click out of the body with nowhere to send it.
fn in_note_area(area: &NoteRegion, y: f64) -> bool {
    let px = |value: &Option<Number>| value.as_ref().and_then(Number::as_f64);
    let (Some(top), Some(height)) = (px(&area.y), px(&area.height)) else {
        return false;
    };
    height > 0.0 && y >= top && y <= top + height
}

/// Whether the topmost image under the point is one a click would select.
/// Mirrors `displayListImages.ts`: reverse paint order, and no document
/// position means nothing selects it (a picture watermark), so text painted
/// over such an image stays readable through it.
fn over_selectable_image(prims: &[Primitive], x: f64, y: f64) -> bool {
    prims
        .iter()
        .rev()
        .find(|primitive| match primitive {
            Primitive::Image(img) => {
                let (ix, iy) = (img.x.as_f64().unwrap_or(0.0), img.y.as_f64().unwrap_or(0.0));
                let (iw, ih) = (img.w.as_f64().unwrap_or(0.0), img.h.as_f64().unwrap_or(0.0));
                x >= ix && x <= ix + iw && y >= iy && y <= iy + ih
            }
            Primitive::Shape(shape) => {
                if shape.attrs.inline_shape_atom != Some(true) {
                    return false;
                }
                if shape.attrs.doc_start.is_none() {
                    return false;
                }
                let (sx, sy) = (
                    shape.x.as_f64().unwrap_or(0.0),
                    shape.y.as_f64().unwrap_or(0.0),
                );
                let (sw, sh) = (
                    shape.w.as_f64().unwrap_or(0.0),
                    shape.h.as_f64().unwrap_or(0.0),
                );
                x >= sx && x <= sx + sw && y >= sy && y <= sy + sh
            }
            _ => false,
        })
        .is_some_and(|primitive| match primitive {
            Primitive::Image(img) => img.attrs.doc_start.is_some(),
            Primitive::Shape(shape) => shape.attrs.doc_start.is_some(),
            _ => false,
        })
}

/// The shared point resolver over one primitive list — a page body or a single
/// header/footer band, both of which use page coordinates. `typeable` says the
/// point is in the region's typeable area (false for a band, which has no
/// content box); it widens the target only, never the position.
fn resolve_point(prims: &[Primitive], typeable: bool, x: f64, y: f64) -> PointResolution {
    // outranks text for the CURSOR only: the pointer path selects an image
    // before it asks for a position, so position resolution keeps its own order
    let image_target = over_selectable_image(prims, x, y);
    let hits = text_hits(prims);
    // What the point earns once no run claims it. An area with no positionable
    // text is not typeable whatever it encloses: such a click is answered with
    // the document's end.
    let target = if image_target {
        HoverTarget::Image
    } else if typeable && !hits.is_empty() {
        HoverTarget::Text
    } else {
        HoverTarget::None
    };

    // 1. direct hit on a text primitive's box, in paint order
    for h in &hits {
        if h.width <= 0.0 {
            continue;
        }
        if x >= h.x && x <= h.x + h.width && y >= h.top - BAND_SLACK && y <= h.bottom + BAND_SLACK {
            return PointResolution {
                pos: Some(position_in_run(h, x)),
                target: if image_target {
                    HoverTarget::Image
                } else {
                    HoverTarget::Text
                },
                direct: true,
            };
        }
    }

    // 2. direct hit on an image with a doc position (caret parks before it).
    // Only the topmost image decides the target, so an unselectable one over
    // this leaves the area's answer standing.
    for p in prims {
        match p {
            Primitive::Image(img) => {
                let Some(ds) = img.attrs.doc_start else {
                    continue;
                };
                let (ix, iy) = (img.x.as_f64().unwrap_or(0.0), img.y.as_f64().unwrap_or(0.0));
                let (iw, ih) = (img.w.as_f64().unwrap_or(0.0), img.h.as_f64().unwrap_or(0.0));
                if x >= ix && x <= ix + iw && y >= iy && y <= iy + ih {
                    return PointResolution {
                        pos: Some(ds),
                        target,
                        direct: true,
                    };
                }
            }
            Primitive::Shape(shape) => {
                if shape.attrs.inline_shape_atom != Some(true) {
                    continue;
                }
                let Some(ds) = shape.attrs.doc_start else {
                    continue;
                };
                let (sx, sy) = (
                    shape.x.as_f64().unwrap_or(0.0),
                    shape.y.as_f64().unwrap_or(0.0),
                );
                let (sw, sh) = (
                    shape.w.as_f64().unwrap_or(0.0),
                    shape.h.as_f64().unwrap_or(0.0),
                );
                if x >= sx && x <= sx + sw && y >= sy && y <= sy + sh {
                    return PointResolution {
                        pos: Some(ds),
                        target,
                        direct: true,
                    };
                }
            }
            _ => {}
        }
    }

    if hits.is_empty() {
        return PointResolution {
            pos: None,
            target,
            direct: false,
        };
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
    PointResolution {
        pos: position_for_hits(line.iter().copied(), x),
        target,
        direct: false,
    }
}

/// Which part of the page owns a point, the position inside that part's
/// document, and what sits under the pointer. For `header` / `footer` the
/// position addresses the header/footer document identified by `rId`, and for
/// `footnote` / `endnote` the note story named by `noteId` — never the body
/// document, so the caller must route the resulting selection to that editor.
#[derive(Serialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct RegionHit {
    pub region: HitRegion,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub r_id: Option<String>,
    /// note whose story the position addresses (`fn:{id}` / `en:{id}`)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub note_id: Option<i64>,
    pub pos: Option<i64>,
    pub target: HoverTarget,
}

#[derive(Serialize, Debug, Clone, Copy, PartialEq, Eq)]
pub enum HitRegion {
    #[serde(rename = "body")]
    Body,
    #[serde(rename = "header")]
    Header,
    #[serde(rename = "footer")]
    Footer,
    #[serde(rename = "footnote")]
    Footnote,
    #[serde(rename = "endnote")]
    Endnote,
}

/// A scoped query's address: the page part, plus the instance of it that owns
/// the document the positions belong to. A header/footer part is named by its
/// `rId`, and `None` matches any band of the kind, since the same part paints
/// on every page using it. A note is named by its id, which is never optional:
/// two notes are two unrelated documents whose positions must not mix.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RegionScope<'a> {
    Body,
    Header(Option<&'a str>),
    Footer(Option<&'a str>),
    Footnote(i64),
    Endnote(i64),
}

impl RegionScope<'_> {
    fn region(self) -> HitRegion {
        match self {
            Self::Body => HitRegion::Body,
            Self::Header(_) => HitRegion::Header,
            Self::Footer(_) => HitRegion::Footer,
            Self::Footnote(_) => HitRegion::Footnote,
            Self::Endnote(_) => HitRegion::Endnote,
        }
    }
}

/// parse the scope used at the JSON/wasm boundary: `region` is the string twin
/// of [`HitRegion`]'s serde rename and `part_id` names the instance — a
/// header/footer `rId` (empty ⇒ any band of the kind) or a note id. Shared by
/// the JSON-arg and session-handle range-rect exports.
pub fn parse_region_scope<'a>(region: &str, part_id: &'a str) -> Result<RegionScope<'a>, String> {
    let r_id = (!part_id.is_empty()).then_some(part_id);
    let note_id = || {
        part_id
            .parse::<i64>()
            .map_err(|_| format!("region {region:?} needs a note id, got {part_id:?}"))
    };
    match region {
        "body" => Ok(RegionScope::Body),
        "header" => Ok(RegionScope::Header(r_id)),
        "footer" => Ok(RegionScope::Footer(r_id)),
        "footnote" => Ok(RegionScope::Footnote(note_id()?)),
        "endnote" => Ok(RegionScope::Endnote(note_id()?)),
        other => Err(format!("unknown region {other:?}")),
    }
}

/// One note's primitives inside a page's note area, with the id naming the
/// `fn:{id}` / `en:{id}` story its positions belong to.
struct NoteStory<'a> {
    id: i64,
    primitives: &'a [Primitive],
}

/// The kind string an area paints under, defaulting like the display-list
/// emitter does when the layout left it unstated. Note paint-group ids are
/// built from it, so the two must agree.
fn note_kind(area: &NoteRegion) -> &str {
    area.kind.as_deref().unwrap_or("footnote")
}

fn note_region(area: &NoteRegion) -> HitRegion {
    if note_kind(area) == "endnote" {
        HitRegion::Endnote
    } else {
        HitRegion::Footnote
    }
}

/// The stories an area stacks, in paint order: one note's primitives are
/// emitted contiguously under its paint-group id, so each story is one span of
/// them. A primitive with no attrs to carry a group — a paragraph or table
/// border line — belongs to the note it sits inside and must not end a span;
/// one naming no listed note does end it, being the area's own chrome.
fn note_stories(area: &NoteRegion) -> Vec<NoteStory<'_>> {
    /// The span being accumulated: its note, and where the span starts.
    #[derive(Clone, Copy)]
    struct OpenSpan {
        id: i64,
        start: usize,
    }

    let kind = note_kind(area);
    let mut stories = Vec::new();
    let mut open: Option<OpenSpan> = None;
    for (index, primitive) in area.primitives.iter().enumerate() {
        let Some(attrs) = doc_attrs(primitive) else {
            continue;
        };
        let id = attrs.group_id.as_deref().and_then(|group| {
            area.note_ids
                .iter()
                .copied()
                .find(|id| note_group_id(kind, *id) == group)
        });
        if id == open.map(|span| span.id) {
            continue;
        }
        if let Some(span) = open {
            stories.push(NoteStory {
                id: span.id,
                primitives: &area.primitives[span.start..index],
            });
        }
        open = id.map(|id| OpenSpan { id, start: index });
    }
    if let Some(span) = open {
        stories.push(NoteStory {
            id: span.id,
            primitives: &area.primitives[span.start..],
        });
    }
    stories
}

/// Squared distance from a point to the nearest text box of a primitive list,
/// zero inside one and infinite for a list painting no text. Both axes count:
/// a note area lays its notes out in columns it starts at the same vertical
/// position, so two of them can share a band and only `x` tells them apart.
fn text_box_distance_squared(prims: &[Primitive], x: f64, y: f64) -> f64 {
    text_hits(prims)
        .iter()
        .map(|hit| {
            let dx = (hit.x - x).max(x - (hit.x + hit.width)).max(0.0);
            let dy = (hit.top - y).max(y - hit.bottom).max(0.0);
            dx * dx + dy * dy
        })
        .fold(f64::INFINITY, f64::min)
}

/// Whether `primitive` carries a document position and paints its hit box at
/// the point, inside its clip as the canvas and raster painters draw it.
fn painted_body_hit_at(primitive: &Primitive, x: f64, y: f64) -> bool {
    let Some(attrs) = doc_attrs(primitive) else {
        return false;
    };
    if attrs.doc_start.is_none() {
        return false;
    }
    // The canvas multiplies the primitive's opacity with its group's, which it applies only
    // inside a clip: either at zero paints nothing.
    let transparent =
        |opacity: Option<&Number>| opacity.and_then(Number::as_f64).is_some_and(|o| o <= 0.0);
    let own_opacity = match primitive {
        Primitive::Image(img) => img.opacity.as_ref(),
        Primitive::Text(text) => text.opacity.as_ref(),
        Primitive::GlyphRun(run) => run.opacity.as_ref(),
        _ => None,
    };
    let group = attrs.clip_group.as_ref();
    let group_clip = group.and_then(|group| group.clip.as_ref());
    if (group_clip.is_some() && transparent(group.and_then(|group| group.opacity.as_ref())))
        || transparent(own_opacity)
        || transparent(attrs.primitive_opacity.as_ref())
    {
        return false;
    }
    let paints = match primitive {
        Primitive::Text(text) => !js_blank(&text.text) && !text_fill_none(attrs),
        Primitive::GlyphRun(run) => {
            !run.glyphs.is_empty() && !js_blank(&run.text) && !text_fill_none(attrs)
        }
        Primitive::Shape(shape) => shape_fill_paints(shape),
        _ => true,
    };
    if !paints {
        return false;
    }
    if let Some(clip) = group_clip {
        let px = |value: &Option<Number>| value.as_ref().and_then(Number::as_f64).unwrap_or(0.0);
        let (left, top) = (px(&clip.x), px(&clip.y));
        let (width, height) = (px(&clip.w).max(0.0), px(&clip.h).max(0.0));
        if width <= 0.0
            || height <= 0.0
            || x < left
            || x > left + width
            || y < top
            || y > top + height
        {
            return false;
        }
    }
    if let Some(hit) = text_hit(primitive) {
        let (paint_clip, rotation, scale) = match primitive {
            Primitive::Text(text) => (
                text.paint_clip.as_ref(),
                text.rotation_deg.as_ref(),
                text.horizontal_scale.as_ref(),
            ),
            Primitive::GlyphRun(run) => (
                run.paint_clip.as_ref(),
                run.rotation_deg.as_ref(),
                run.horizontal_scale.as_ref(),
            ),
            _ => (None, None, None),
        };
        // A turned or compressed run paints outside or short of its box.
        if rotation.and_then(Number::as_f64).unwrap_or(0.0) % 360.0 != 0.0
            || scale
                .and_then(Number::as_f64)
                .is_some_and(|scale| scale < 100.0)
        {
            return false;
        }
        if let Some(clip) = paint_clip {
            let left = clip.x.as_ref().and_then(Number::as_f64).unwrap_or(0.0);
            let width = clip.w.as_ref().and_then(Number::as_f64).unwrap_or(0.0);
            if width <= 0.0 || x < left || x > left + width {
                return false;
            }
        }
        let em = hit.baseline - hit.top;
        return hit.width > 0.0
            && x >= hit.x
            && x <= hit.x + hit.width
            && y >= hit.baseline - em * INK_ASCENT_EM
            && y <= hit.baseline + em * INK_DESCENT_EM;
    }
    let (left, top, width, height) = match primitive {
        Primitive::Image(img) => {
            let Some(rect) = image_paint_rect(img) else {
                return false;
            };
            rect
        }
        Primitive::Shape(shape) if attrs.inline_shape_atom == Some(true) => {
            let Some(rect) = shape_fill_rect(shape) else {
                return false;
            };
            rect
        }
        _ => return false,
    };
    width > 0.0 && height > 0.0 && x >= left && x <= left + width && y >= top && y <= top + height
}

/// Whether JavaScript's `trim()` leaves nothing, as the overlay tests a run's text.
fn js_blank(text: &str) -> bool {
    text.chars()
        .all(|c| c == '\u{feff}' || (c.is_whitespace() && c != '\u{85}'))
}

/// A run with no glyph fill: whether its outline paints differs by canvas path, so it covers nothing.
fn text_fill_none(attrs: &DocAttrs) -> bool {
    attrs
        .modern_effects
        .as_deref()
        .and_then(|effects| effects.pointer("/textFill/kind"))
        .and_then(Value::as_str)
        == Some("none")
}

/// The rectangle (left, top, width, height) a shape's fill paints, when its path is
/// exactly an axis-aligned rectangle turned by a multiple of 180 degrees; any other
/// path covers nothing.
fn shape_fill_rect(shape: &ShapePrimitive) -> Option<(f64, f64, f64, f64)> {
    let transform = shape.transform.as_ref();
    let rotation = transform
        .and_then(|transform| transform.rotation.as_ref())
        .and_then(Number::as_f64)
        .unwrap_or(0.0);
    if rotation % 180.0 != 0.0 {
        return None;
    }
    let commands = &shape.geometry_path;
    let end = match commands.last() {
        Some(ShapePathCommand::Close) => commands.len() - 1,
        _ => commands.len(),
    };
    let px = |value: &Number| value.as_f64().unwrap_or(0.0);
    let mut corners = Vec::with_capacity(end);
    for (index, command) in commands[..end].iter().enumerate() {
        match command {
            ShapePathCommand::Move { x, y } if index == 0 => corners.push((px(x), px(y))),
            ShapePathCommand::Line { x, y } if index > 0 => corners.push((px(x), px(y))),
            _ => return None,
        }
    }
    if corners.len() == 5 && corners.first() == corners.last() {
        corners.pop();
    }
    if corners.len() != 4 {
        return None;
    }
    let mut xs: Vec<f64> = corners.iter().map(|corner| corner.0).collect();
    let mut ys: Vec<f64> = corners.iter().map(|corner| corner.1).collect();
    for values in [&mut xs, &mut ys] {
        values.sort_by(f64::total_cmp);
        values.dedup();
    }
    let distinct = (0..4).all(|a| (a + 1..4).all(|b| corners[a] != corners[b]));
    let sides = (0..4).all(|index| {
        let ((x, y), (next_x, next_y)) = (corners[index], corners[(index + 1) % 4]);
        (x == next_x) != (y == next_y)
    });
    let ([x0, x1], [y0, y1]) = (xs.as_slice(), ys.as_slice()) else {
        return None;
    };
    if !distinct || !sides {
        return None;
    }
    let (mut left, mut top, width, height) = (*x0, *y0, x1 - x0, y1 - y0);
    // The canvas turns and flips a shape about its box's center.
    let half_turn = (rotation % 360.0).abs() == 180.0;
    let flip_h = transform.is_some_and(|transform| transform.flip_h);
    let flip_v = transform.is_some_and(|transform| transform.flip_v);
    if flip_h != half_turn {
        left = 2.0 * px(&shape.x) + px(&shape.w) - left - width;
    }
    if flip_v != half_turn {
        top = 2.0 * px(&shape.y) + px(&shape.h) - top - height;
    }
    Some((left, top, width, height))
}

/// The frame an image paints over whole, as the canvas draws it: none for a non-rectangular
/// image, a turn other than a half-turn, or a crop that leaves part of the frame bare.
fn image_paint_rect(img: &ImagePrimitive) -> Option<(f64, f64, f64, f64)> {
    let num = |value: Option<&Number>| value.and_then(Number::as_f64).filter(|v| v.is_finite());
    let shaped = img
        .attrs
        .image_shape_type
        .as_deref()
        .is_some_and(|shape| shape != "rect");
    let turned = num(img.rotation_deg.as_ref()).unwrap_or(0.0) % 180.0 != 0.0;
    let bare = img.crop.as_ref().is_some_and(|crop| {
        !crop_fills_frame(
            crop.left.as_f64(),
            crop.top.as_f64(),
            crop.right.as_f64(),
            crop.bottom.as_f64(),
        )
    });
    if shaped || turned || bare {
        return None;
    }
    let frame = img.attrs.content_frame.as_deref();
    let side = |own: Option<&Number>, outer: &Number| num(own).or(outer.as_f64()).unwrap_or(0.0);
    Some((
        side(frame.and_then(|frame| frame.x.as_ref()), &img.x),
        side(frame.and_then(|frame| frame.y.as_ref()), &img.y),
        side(frame.and_then(|frame| frame.w.as_ref()), &img.w),
        side(frame.and_then(|frame| frame.h.as_ref()), &img.h),
    ))
}

/// Whether a source crop draws over its whole frame: an outset side leaves a gutter.
fn crop_fills_frame(
    left: Option<f64>,
    top: Option<f64>,
    right: Option<f64>,
    bottom: Option<f64>,
) -> bool {
    let side = |value: Option<f64>| value.filter(|v| v.is_finite()).unwrap_or(0.0);
    let (left, top, right, bottom) = (side(left), side(top), side(right), side(bottom));
    left >= 0.0
        && top >= 0.0
        && right >= 0.0
        && bottom >= 0.0
        && left + right < 1.0
        && top + bottom < 1.0
}

/// Whether a shape's fill paints its interior, as the overlay's occlusion check decides it.
fn shape_fill_paints(shape: &ShapePrimitive) -> bool {
    let paint = shape.attrs.fill_paint.as_deref();
    let field = |key: &str| {
        paint
            .and_then(|paint| paint.get(key))
            .and_then(Value::as_str)
            .filter(|value| !value.is_empty())
    };
    match field("kind") {
        Some("none") => false,
        Some("gradient" | "pattern") => true,
        Some("picture") if field("pictureSrc").is_some() || field("pictureRelId").is_some() => {
            let opaque = paint
                .and_then(|paint| paint.get("pictureOpacity"))
                .and_then(Value::as_f64)
                .is_none_or(|opacity| opacity > 0.0);
            // An inset or a crop past the source paints only part of the shape, tiled fills
            // included: past the tile cap the canvas stretches them.
            let inset = paint
                .and_then(|paint| paint.get("pictureStretchRect"))
                .is_some_and(|rect| {
                    ["left", "top", "right", "bottom"].iter().any(|side| {
                        rect.get(side)
                            .and_then(Value::as_f64)
                            .is_some_and(|v| v > 0.0)
                    })
                });
            let crop = paint.and_then(|paint| paint.get("pictureSrcRect"));
            let side = |key: &str| crop.and_then(|rect| rect.get(key)).and_then(Value::as_f64);
            opaque
                && !inset
                && crop_fills_frame(side("left"), side("top"), side("right"), side("bottom"))
        }
        _ => paint
            .and_then(|paint| paint.get("color"))
            .and_then(Value::as_str)
            .or(shape.fill.as_deref())
            .is_some_and(|fill| !matches!(fill, "" | "transparent" | "none")),
    }
}

/// Direct body hits override header/footer bands; empty band spots activate
/// that part. Note areas resolve against the nearest note story.
pub fn hit_test_regions(dl: &DisplayList, page_index: usize, x: f64, y: f64) -> Option<RegionHit> {
    let page = dl.pages.get(page_index)?;

    let in_band = |r: &HfRegion| -> bool {
        let top = r.y.as_f64().unwrap_or(0.0);
        let bottom = top + r.height.as_f64().unwrap_or(0.0);
        y >= top && y <= bottom
    };

    for (region, band) in [
        (HitRegion::Header, &page.header),
        (HitRegion::Footer, &page.footer),
    ] {
        if let Some(band) = band
            && in_band(band)
        {
            // Only body content painted at the point takes it, and its own position answers.
            let painted: Vec<Primitive> = page
                .primitives
                .iter()
                .filter(|primitive| painted_body_hit_at(primitive, x, y))
                .cloned()
                .collect();
            let body = resolve_point(&painted, false, x, y);
            if body.direct {
                return Some(RegionHit {
                    region: HitRegion::Body,
                    r_id: None,
                    note_id: None,
                    pos: body.pos,
                    target: body.target,
                });
            }
            let resolved = resolve_point(&band.primitives, false, x, y);
            return Some(RegionHit {
                region,
                r_id: Some(band.r_id.clone()),
                note_id: None,
                pos: resolved.pos,
                target: resolved.target,
            });
        }
    }
    if let Some(area) = page.note_areas.iter().find(|area| in_note_area(area, y)) {
        let story = note_stories(area)
            .into_iter()
            .map(|story| (text_box_distance_squared(story.primitives, x, y), story))
            .min_by(|(left, _), (right, _)| left.total_cmp(right))
            .map(|(_, story)| story);
        let primitives = story.as_ref().map_or(&[][..], |story| story.primitives);
        let resolved = resolve_point(primitives, false, x, y);
        return Some(RegionHit {
            region: note_region(area),
            r_id: None,
            note_id: story.map(|story| story.id),
            pos: resolved.pos,
            target: resolved.target,
        });
    }
    let resolved = resolve_point(&page.primitives, in_typeable_area(page, x, y), x, y);
    Some(RegionHit {
        region: HitRegion::Body,
        r_id: None,
        note_id: None,
        pos: resolved.pos,
        target: resolved.target,
    })
}

/// one highlight rectangle of a document range, page-local coordinates
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
    #[cfg(test)]
    RANGE_RECT_PAGE_VISITS.with(|visits| visits.borrow_mut().push(page_index));
    let pending: Vec<(RectOwner<'_>, RangeRect)> = text_hits(prims)
        .iter()
        .filter_map(|hit| hit_rect(hit, page_index, from, to))
        .collect();
    merge_line_rects(pending, out);
    out.extend(
        prims
            .iter()
            .filter_map(atom_rect)
            .filter(|(start, end, _)| *end > from && *start < to)
            .map(|(_, _, rect)| RangeRect { page_index, ..rect }),
    );
}

/// The rect one text hit contributes to the range `[from, to)`: a proportional sub-span over
/// the line band, or a thin sliver for a blank-line marker inside it.
fn hit_rect<'a>(
    h: &TextHit<'a>,
    page_index: usize,
    from: i64,
    to: i64,
) -> Option<(RectOwner<'a>, RangeRect)> {
    // blank-line marker: zero-length span selects as a thin sliver
    if h.doc_start == h.doc_end {
        return (h.doc_start >= from && h.doc_start < to).then(|| {
            (
                rect_owner(h),
                RangeRect {
                    page_index,
                    x: h.x,
                    y: h.top,
                    width: BLANK_LINE_SELECTION_WIDTH,
                    height: h.bottom - h.top,
                },
            )
        });
    }
    if h.doc_end <= from || h.doc_start >= to {
        return None;
    }
    let start = from.max(h.doc_start).min(h.doc_end);
    let end = to.max(h.doc_start).min(h.doc_end);
    let x0 = x_at_position(h, start);
    let x1 = x_at_position(h, end);
    Some((
        rect_owner(h),
        RangeRect {
            page_index,
            x: x0.min(x1),
            y: h.top,
            // degenerate overlaps keep a 1px floor like lineSpanRect
            width: (x1 - x0).abs().max(1.0),
            height: h.bottom - h.top,
        },
    ))
}

/// The document range and box of an image or inline-shape atom primitive.
fn atom_rect(primitive: &Primitive) -> Option<(i64, i64, RangeRect)> {
    let (attrs, x, y, w, h) = match primitive {
        Primitive::Image(img) => (&img.attrs, &img.x, &img.y, &img.w, &img.h),
        Primitive::Shape(shape) if shape.attrs.inline_shape_atom == Some(true) => {
            (&shape.attrs, &shape.x, &shape.y, &shape.w, &shape.h)
        }
        _ => return None,
    };
    Some((
        attrs.doc_start?,
        attrs.doc_end?,
        RangeRect {
            page_index: 0,
            x: x.as_f64().unwrap_or(0.0),
            y: y.as_f64().unwrap_or(0.0),
            width: w.as_f64().unwrap_or(0.0),
            height: h.as_f64().unwrap_or(0.0),
        },
    ))
}

/// Highlight rects for many ranges of one page-local primitive list (a page body, one band or
/// one note), each answered as [`range_rects_in_region`] answers a single range but without
/// rescanning the list per range.
pub struct RangeRectIndex<'a> {
    page_index: usize,
    /// Text hits by start position, with the widest hit span.
    hits: Vec<TextHit<'a>>,
    widest: i64,
    atoms: Vec<(i64, i64, RangeRect)>,
    widest_atom: i64,
}

impl<'a> RangeRectIndex<'a> {
    pub fn new(prims: &'a [Primitive], page_index: usize) -> Self {
        let mut hits = text_hits(prims);
        hits.sort_by_key(|hit| hit.doc_start);
        let widest = hits
            .iter()
            .map(|hit| hit.doc_end - hit.doc_start)
            .max()
            .unwrap_or(0)
            .max(0);
        let mut atoms: Vec<_> = prims.iter().filter_map(atom_rect).collect();
        atoms.sort_by_key(|(start, ..)| *start);
        let widest_atom = atoms
            .iter()
            .map(|(start, end, _)| end - start)
            .max()
            .unwrap_or(0)
            .max(0);
        Self {
            page_index,
            hits,
            widest,
            atoms,
            widest_atom,
        }
    }

    /// The merged per-line rects covering `[from, to)`.
    pub fn rects(&self, from: i64, to: i64) -> Vec<RangeRect> {
        let mut out = Vec::new();
        if from >= to {
            return out;
        }
        let first = self
            .hits
            .partition_point(|hit| hit.doc_start < from.saturating_sub(self.widest));
        let pending: Vec<_> = self.hits[first..]
            .iter()
            .take_while(|hit| hit.doc_start < to)
            .filter_map(|hit| hit_rect(hit, self.page_index, from, to))
            .collect();
        merge_line_rects(pending, &mut out);
        let first = self
            .atoms
            .partition_point(|(start, ..)| *start < from.saturating_sub(self.widest_atom));
        out.extend(
            self.atoms[first..]
                .iter()
                .take_while(|(start, ..)| *start < to)
                .filter(|(_, end, _)| *end > from)
                .map(|(_, _, rect)| RangeRect {
                    page_index: self.page_index,
                    ..rect.clone()
                }),
        );
        out
    }
}

/// Each note of a page's note area with the primitives it paints.
pub fn note_primitives(area: &NoteRegion) -> Vec<(i64, &[Primitive])> {
    note_stories(area)
        .into_iter()
        .map(|story| (story.id, story.primitives))
        .collect()
}

const LINE_MERGE_BAND_EPSILON: f64 = 1.0;
const LINE_MERGE_GAP: f64 = 2.0;

/// Line identity for band merging, strict variant of [`same_line_owner`]:
/// rects union only within one table cell, line, and block, so bands never
/// bridge adjacent columns or cells that happen to align.
struct RectOwner<'a> {
    table_id: Option<&'a str>,
    cell: Option<(u64, u64, Option<&'a str>)>,
    line_index: Option<u64>,
    para_id: Option<&'a str>,
    block_key: Option<&'a str>,
    block_id: Option<&'a Number>,
    doc_start: i64,
    doc_end: i64,
}

impl RectOwner<'_> {
    fn matches(&self, other: &Self) -> bool {
        self.table_id == other.table_id
            && self.cell == other.cell
            && self.line_index == other.line_index
            && self.para_id == other.para_id
            && self.block_key == other.block_key
            && self.block_id == other.block_id
            && (self.para_id.is_some()
                || self.block_key.is_some()
                || self.block_id.is_some()
                || self.doc_end == other.doc_start
                || other.doc_end == self.doc_start)
    }

    fn absorb(&mut self, other: &Self) {
        self.doc_start = self.doc_start.min(other.doc_start);
        self.doc_end = self.doc_end.max(other.doc_end);
    }
}

fn rect_owner<'a>(hit: &TextHit<'a>) -> RectOwner<'a> {
    let attrs = hit.attrs;
    RectOwner {
        table_id: attrs.table.as_ref().map(|table| table.table_id.as_str()),
        cell: attrs
            .cell
            .as_ref()
            .map(|cell| (cell.row, cell.col, cell.cell_id.as_deref())),
        line_index: attrs.line_index,
        para_id: attrs.para_id.as_deref(),
        block_key: attrs.block_key.as_deref(),
        block_id: attrs.block_id.as_ref(),
        doc_start: hit.doc_start,
        doc_end: hit.doc_end,
    }
}

/// Coalesce one page's text rects into per-line bands: same-owner rects on the
/// same band that touch or nearly touch horizontally union into one. Selection
/// highlights are per line visually, so per-run granularity only multiplies
/// the rect count — a full-document selection must stay O(lines), not O(runs).
fn merge_line_rects(mut pending: Vec<(RectOwner<'_>, RangeRect)>, out: &mut Vec<RangeRect>) {
    pending.sort_by(|a, b| a.1.y.total_cmp(&b.1.y).then(a.1.x.total_cmp(&b.1.x)));
    let mut current: Option<(RectOwner<'_>, RangeRect)> = None;
    for (owner, rect) in pending {
        if let Some((held_owner, held)) = current.as_mut()
            && held_owner.matches(&owner)
            && (rect.y - held.y).abs() <= LINE_MERGE_BAND_EPSILON
            && (rect.height - held.height).abs() <= LINE_MERGE_BAND_EPSILON
            && rect.x <= held.x + held.width + LINE_MERGE_GAP
            && held.x <= rect.x + rect.width + LINE_MERGE_GAP
        {
            let left = held.x.min(rect.x);
            let right = (held.x + held.width).max(rect.x + rect.width);
            held.x = left;
            held.width = right - left;
            held_owner.absorb(&owner);
            continue;
        }
        if let Some((_, held)) = current.take() {
            out.push(held);
        }
        current = Some((owner, rect));
    }
    if let Some((_, held)) = current {
        out.push(held);
    }
}

/// Returns body-document highlight rectangles across all pages.
pub fn range_rects(dl: &DisplayList, from: i64, to: i64) -> Vec<RangeRect> {
    range_rects_in_region(dl, RegionScope::Body, from, to)
}

/// Caret box for a body-document position: the first primitive on the earliest
/// page whose half-open range covers it, or that is an empty run sitting
/// exactly at it.
pub fn caret_rect(dl: &DisplayList, pos: i64) -> Option<CaretRect> {
    for (page_index, page) in dl.pages.iter().enumerate() {
        if let Some(hit) = page.primitives.iter().find_map(|primitive| {
            let (start, end) = text_doc_range(primitive)?;
            if (start == end && pos == start) || (pos >= start && pos < end) {
                text_hit(primitive)
            } else {
                None
            }
        }) {
            return Some(CaretRect {
                page_index,
                x: x_at_position(&hit, pos),
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
        if let Some(shape) = page.primitives.iter().find_map(|primitive| {
            let Primitive::Shape(shape) = primitive else {
                return None;
            };
            if shape.attrs.inline_shape_atom != Some(true) {
                return None;
            }
            let (Some(start), Some(end)) = (shape.attrs.doc_start, shape.attrs.doc_end) else {
                return None;
            };
            (pos >= start && pos < end).then_some(shape)
        }) {
            return Some(CaretRect {
                page_index,
                x: shape.x.as_f64().unwrap_or(0.0),
                y: shape.y.as_f64().unwrap_or(0.0),
                height: shape.h.as_f64().unwrap_or(0.0),
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
        if let Some(shape) = page.primitives.iter().rev().find_map(|primitive| {
            let Primitive::Shape(shape) = primitive else {
                return None;
            };
            if shape.attrs.inline_shape_atom != Some(true) {
                return None;
            }
            let (Some(start), Some(end)) = (shape.attrs.doc_start, shape.attrs.doc_end) else {
                return None;
            };
            (pos > start && pos <= end).then_some(shape)
        }) {
            return Some(CaretRect {
                page_index,
                x: shape.x.as_f64().unwrap_or(0.0) + shape.w.as_f64().unwrap_or(0.0),
                y: shape.y.as_f64().unwrap_or(0.0),
                height: shape.h.as_f64().unwrap_or(0.0),
            });
        }
        if let Some(hit) = page.primitives.iter().rev().find_map(|primitive| {
            let (start, end) = text_doc_range(primitive)?;
            if pos > start && pos <= end {
                text_hit(primitive)
            } else {
                None
            }
        }) {
            return Some(CaretRect {
                page_index,
                x: x_at_position(&hit, pos),
                y: hit.top,
                height: hit.bottom - hit.top,
            });
        }
    }
    None
}

/// Highlight rectangles for a document range inside one page region — the
/// selection-geometry twin of [`hit_test_regions`]'s scoping.
///
/// [`RegionScope::Body`] behaves exactly like [`range_rects`], since the body
/// is a single document. Otherwise `from` and `to` address the document the
/// scope names — a header/footer part, or a note story — and only the parts
/// matching it contribute. A header/footer part paints on every page using it,
/// so a match yields one rect set per such page, each stamped with its own
/// `page_index`.
/// The document span `[start, end)` holding every body item of `page` that a
/// range query can return a rect for, or `None` when it has none: a body
/// range query that misses the span gets nothing from the page.
pub fn body_range_span(page: &DisplayPage) -> Option<(i64, i64)> {
    let mut span: Option<(i64, i64)> = None;
    let mut cover = |first: i64, second: i64| {
        let (start, end) = (first.min(second), first.max(second));
        let end = if start == end {
            start.saturating_add(1)
        } else {
            end
        };
        span = Some(span.map_or((start, end), |(low, high)| (low.min(start), high.max(end))));
    };
    for hit in text_hits(&page.primitives) {
        cover(hit.doc_start, hit.doc_end);
    }
    for (start, end, _) in page.primitives.iter().filter_map(atom_rect) {
        cover(start, end);
    }
    span
}

/// [`range_rects`] reading only the listed pages, in the order given.
pub fn range_rects_on_pages(
    dl: &DisplayList,
    pages: impl IntoIterator<Item = usize>,
    from: i64,
    to: i64,
) -> Vec<RangeRect> {
    let (from, to) = (from.min(to), from.max(to));
    let mut rects = Vec::new();
    if from == to {
        return rects;
    }
    for page_index in pages {
        if let Some(page) = dl.pages.get(page_index) {
            collect_range_rects(&page.primitives, page_index, from, to, &mut rects);
        }
    }
    rects
}

pub fn range_rects_in_region(
    dl: &DisplayList,
    scope: RegionScope<'_>,
    from: i64,
    to: i64,
) -> Vec<RangeRect> {
    let (from, to) = (from.min(to), from.max(to));
    let mut rects = Vec::new();
    if from == to {
        return rects;
    }

    let band = |band: &HfRegion, r_id: Option<&str>| r_id.is_none_or(|id| id == band.r_id);
    for (page_index, page) in dl.pages.iter().enumerate() {
        let prims: Option<&[Primitive]> = match scope {
            RegionScope::Body => Some(&page.primitives),
            RegionScope::Header(r_id) => page
                .header
                .as_ref()
                .filter(|h| band(h, r_id))
                .map(|h| h.primitives.as_slice()),
            RegionScope::Footer(r_id) => page
                .footer
                .as_ref()
                .filter(|f| band(f, r_id))
                .map(|f| f.primitives.as_slice()),
            RegionScope::Footnote(note_id) | RegionScope::Endnote(note_id) => page
                .note_areas
                .iter()
                .filter(|area| {
                    note_region(area) == scope.region() && area.note_ids.contains(&note_id)
                })
                .flat_map(note_stories)
                .find(|story| story.id == note_id)
                .map(|story| story.primitives),
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

pub fn page_window(first: f64, last: f64) -> Option<(usize, usize)> {
    if first.is_nan() || last.is_nan() {
        return None;
    }
    let lo = first.ceil().max(0.0);
    let hi = last.floor();
    if hi < lo {
        return None;
    }
    Some((lo as usize, hi as usize))
}

pub fn range_rects_on_pages_json(
    display_list: &str,
    from: i64,
    to: i64,
    first_page: usize,
    last_page: usize,
) -> Result<String, String> {
    let dl: DisplayList = serde_json::from_str(display_list).map_err(|e| format!("parse: {e}"))?;
    let pages = first_page..last_page.saturating_add(1).min(dl.pages.len());
    serde_json::to_string(&range_rects_on_pages(&dl, pages, from, to))
        .map_err(|e| format!("serialize: {e}"))
}

/// `range_rects_in_region` over serialized inputs. `region` is
/// `"body" | "header" | "footer" | "footnote" | "endnote"`; `part_id` names the
/// instance — the HF part's relationship id (empty ⇒ match any band of the
/// kind) or the note id, and is ignored for `body`. Returns a JSON array of
/// rects, or a `parse:`/`unknown region` error string.
pub fn range_rects_region_json(
    display_list: &str,
    region: &str,
    part_id: &str,
    from: i64,
    to: i64,
) -> Result<String, String> {
    let dl: DisplayList = serde_json::from_str(display_list).map_err(|e| format!("parse: {e}"))?;
    let scope = parse_region_scope(region, part_id)?;
    serde_json::to_string(&range_rects_in_region(&dl, scope, from, to))
        .map_err(|e| format!("serialize: {e}"))
}

/// `hit_test_regions` over serialized inputs; returns
/// `{"region":"body"|"header"|"footer"|"footnote"|"endnote","rId"?,"noteId"?,`
/// `"pos":n|null,"target":"text"|"image"|"none"}` or `"null"` for an
/// out-of-range page
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

#[cfg(test)]
mod tests {
    use super::*;

    fn run(x: f64, baseline: f64, width: f64, doc_start: i64) -> serde_json::Value {
        serde_json::json!({
            "kind": "text",
            "text": "hello",
            "x": x,
            "baselineY": baseline,
            "width": width,
            "font": "400 16px Calibri",
            "color": "#000000",
            "docStart": doc_start,
            "docEnd": doc_start + 5,
            "blockId": doc_start,
            "lineIndex": 0
        })
    }

    fn image(x: f64, y: f64, doc_start: Option<i64>) -> serde_json::Value {
        let mut image = serde_json::json!({
            "kind": "image", "relId": "rId1", "x": x, "y": y, "w": 60, "h": 40
        });
        if let Some(doc_start) = doc_start {
            image["docStart"] = doc_start.into();
            image["docEnd"] = (doc_start + 1).into();
        }
        image
    }

    fn page(content: serde_json::Value, primitives: Vec<serde_json::Value>) -> DisplayList {
        serde_json::from_value(serde_json::json!({
            "pages": [{
                "pageIndex": 0,
                "width": 500,
                "height": 500,
                "contentBounds": content,
                "primitives": primitives
            }]
        }))
        .unwrap()
    }

    /// A content box at x 80..420 / y 80..420, one run whose band is y 84..104
    /// (± [`BAND_SLACK`]) over x 100..150, and one inline image at
    /// x 100..160 / y 200..240.
    fn hover_fixture() -> DisplayList {
        page(
            serde_json::json!({"x": 80, "y": 80, "width": 340, "height": 340}),
            vec![run(100.0, 100.0, 50.0, 1), image(100.0, 200.0, Some(10))],
        )
    }

    #[test]
    fn body_content_in_header_footer_bands_wins_but_empty_spots_activate_the_band() {
        for (kind, top, region) in [
            ("header", 0.0, HitRegion::Header),
            ("footer", 420.0, HitRegion::Footer),
        ] {
            let mut text_box = run(300.0, top + 40.0, 50.0, 20);
            text_box["blockKey"] = "text-box-paragraph".into();
            let mut watermark = image(0.0, top, None);
            watermark["w"] = 500.into();
            watermark["h"] = 80.into();
            let mut dl = page(
                serde_json::json!({"x": 0, "y": 0, "width": 500, "height": 500}),
                vec![
                    watermark,
                    run(100.0, top + 40.0, 50.0, 1),
                    image(200.0, top + 20.0, Some(10)),
                    text_box,
                ],
            );
            dl.pages[0].watermark_primitive_count = Some(1);
            let band = serde_json::from_value(serde_json::json!({
                "rId": "rIdBand", "kind": kind, "y": top, "height": 80,
                "primitives": [run(100.0, top + 40.0, 250.0, 50)]
            }))
            .unwrap();
            if kind == "header" {
                dl.pages[0].header = Some(band);
            } else {
                dl.pages[0].footer = Some(band);
            }

            for (x, start, end, target) in [
                (120.0, 1, 6, HoverTarget::Text),
                (220.0, 10, 10, HoverTarget::Image),
                (320.0, 20, 25, HoverTarget::Text),
            ] {
                let hit = hit_test_regions(&dl, 0, x, top + 35.0).unwrap();
                assert_eq!(hit.region, HitRegion::Body);
                assert_eq!(hit.r_id, None);
                assert_eq!(hit.target, target);
                assert!(hit.pos.is_some_and(|pos| (start..=end).contains(&pos)));
            }

            let empty = hit_test_regions(&dl, 0, 400.0, top + 35.0).unwrap();
            assert_eq!(empty.region, region);
            assert_eq!(empty.r_id.as_deref(), Some("rIdBand"));
            let page = &mut dl.pages[0];
            if let Some(band) = page.header.as_mut().or(page.footer.as_mut()) {
                band.primitives.clear();
            }
            let empty = hit_test_regions(&dl, 0, 400.0, top + 35.0).unwrap();
            assert_eq!(empty.region, region);
            assert_eq!(empty.pos, None);
            assert_eq!(empty.target, HoverTarget::None);
        }
    }

    fn band_page(kind: &str, top: f64, primitives: Vec<serde_json::Value>) -> DisplayList {
        let mut dl = page(
            serde_json::json!({"x": 80, "y": 80, "width": 340, "height": 340}),
            primitives,
        );
        let band = serde_json::from_value(serde_json::json!({
            "rId": "rIdBand", "kind": kind, "y": top, "height": 80,
            "primitives": [run(100.0, top + 40.0, 50.0, 50)]
        }))
        .unwrap();
        if kind == "header" {
            dl.pages[0].header = Some(band);
        } else {
            dl.pages[0].footer = Some(band);
        }
        dl
    }

    #[test]
    fn clipped_body_text_does_not_override_header() {
        for clip in [
            serde_json::json!({"x": 100, "y": 50, "w": 50, "h": 30}),
            serde_json::json!({"x": 130, "y": 0, "w": 20, "h": 80}),
            serde_json::json!({"x": 120, "y": 0, "w": 0, "h": 80}),
            serde_json::json!({"x": 100, "y": 35, "w": 50, "h": 0}),
            serde_json::json!({"x": 100, "y": 0, "w": -50, "h": 80}),
            serde_json::json!({"x": 100, "y": 35, "w": 50, "h": -10}),
            serde_json::json!({"x": 100, "y": 0, "h": 80}),
            serde_json::json!({"x": 100, "w": 50}),
        ] {
            let mut text = run(100.0, 40.0, 50.0, 1);
            text["clipGroup"] = serde_json::json!({"clip": clip});
            let dl = band_page("header", 0.0, vec![text]);

            let hit = hit_test_regions(&dl, 0, 120.0, 35.0).unwrap();
            assert_eq!(hit.region, HitRegion::Header);
            assert_eq!(hit.r_id.as_deref(), Some("rIdBand"));
            assert_eq!(hit.pos, Some(52));
            assert_eq!(hit.target, HoverTarget::Text);
            assert_eq!(hit_test(&dl, 0, 120.0, 35.0), Some(3));

            let outside = hit_test_regions(&dl, 0, 120.0, 100.0).unwrap();
            assert_eq!(outside.region, HitRegion::Body);
            assert_eq!(outside.pos, Some(3));
        }
    }

    #[test]
    fn a_band_click_takes_the_position_of_the_body_run_painted_there() {
        let mut hidden = run(100.0, 40.0, 50.0, 1);
        hidden["clipGroup"] = serde_json::json!({
            "clip": {"x": 100, "y": 50, "w": 50, "h": 30}
        });
        let visible = run(100.0, 40.0, 50.0, 20);
        let dl = band_page("header", 0.0, vec![hidden, visible]);

        let hit = hit_test_regions(&dl, 0, 120.0, 35.0).unwrap();
        assert_eq!(hit.region, HitRegion::Body);
        assert_eq!(hit.pos, Some(22));
    }

    #[test]
    fn body_text_overrides_header_only_inside_clip() {
        let mut text = run(100.0, 40.0, 50.0, 1);
        text["clipGroup"] = serde_json::json!({
            "clip": {"x": 125, "y": 25, "w": 25, "h": 25}
        });
        let dl = band_page("header", 0.0, vec![text]);

        let excluded = hit_test_regions(&dl, 0, 110.0, 35.0).unwrap();
        assert_eq!(excluded.region, HitRegion::Header);
        assert_eq!(excluded.pos, Some(51));

        let inside = hit_test_regions(&dl, 0, 140.0, 35.0).unwrap();
        assert_eq!(inside.region, HitRegion::Body);
        assert_eq!(inside.r_id, None);
        assert_eq!(inside.pos, Some(5));
        assert_eq!(inside.target, HoverTarget::Text);
    }

    #[test]
    fn synthetic_fallback_paint_clip_limits_body_hits_in_header() {
        let glyph_run = serde_json::json!({
            "kind": "glyphRun", "fontId": 1, "size": 16, "color": "#000000",
            "text": "hello", "docStart": 1, "docEnd": 6,
            "glyphs": [{"id": 1, "x": 100, "y": 40, "cluster": 0, "advance": 50}]
        });
        for mut text in [run(100.0, 40.0, 50.0, 1), glyph_run] {
            text["paintClip"] = serde_json::json!({"x": 100, "w": 25});
            let dl = band_page("header", 0.0, vec![text.clone()]);

            let excluded = hit_test_regions(&dl, 0, 140.0, 35.0).unwrap();
            assert_eq!(excluded.region, HitRegion::Header);
            assert_eq!(excluded.r_id.as_deref(), Some("rIdBand"));
            assert_eq!(excluded.pos, Some(54));
            assert_eq!(excluded.target, HoverTarget::Text);

            let inside = hit_test_regions(&dl, 0, 110.0, 35.0).unwrap();
            assert_eq!(inside.region, HitRegion::Body);
            assert_eq!(inside.r_id, None);
            assert!(inside.pos.is_some_and(|pos| (1..=6).contains(&pos)));
            assert_eq!(inside.target, HoverTarget::Text);

            text["paintClip"] = serde_json::json!({"w": 125});
            let dl = band_page("header", 0.0, vec![text.clone()]);
            assert_eq!(
                hit_test_regions(&dl, 0, 110.0, 35.0).unwrap().region,
                HitRegion::Body
            );
            assert_eq!(
                hit_test_regions(&dl, 0, 140.0, 35.0).unwrap().region,
                HitRegion::Header
            );

            for clip in [
                serde_json::json!({"x": 100}),
                serde_json::json!({"x": 100, "w": 0}),
                serde_json::json!({"x": 100, "w": -25}),
            ] {
                text["paintClip"] = clip;
                let dl = band_page("header", 0.0, vec![text.clone()]);
                assert_eq!(
                    hit_test_regions(&dl, 0, 110.0, 35.0).unwrap().region,
                    HitRegion::Header
                );
            }
        }
    }

    #[test]
    fn body_image_overrides_footer_only_inside_clip() {
        let mut image = image(100.0, 440.0, Some(10));
        image["clipGroup"] = serde_json::json!({
            "clip": {"x": 100, "y": 460, "w": 60, "h": 20}
        });
        let dl = band_page("footer", 420.0, vec![image]);

        let excluded = hit_test_regions(&dl, 0, 120.0, 455.0).unwrap();
        assert_eq!(excluded.region, HitRegion::Footer);
        assert_eq!(excluded.r_id.as_deref(), Some("rIdBand"));
        assert_eq!(excluded.pos, Some(52));
        assert_eq!(excluded.target, HoverTarget::Text);

        let inside = hit_test_regions(&dl, 0, 120.0, 465.0).unwrap();
        assert_eq!(inside.region, HitRegion::Body);
        assert_eq!(inside.r_id, None);
        assert_eq!(inside.pos, Some(10));
        assert_eq!(inside.target, HoverTarget::Image);
    }

    #[test]
    fn body_image_cropped_past_its_source_leaves_footer_clicks_to_the_footer() {
        let cropped = |left: f64, right: f64| {
            let mut image = image(100.0, 440.0, Some(10));
            let crop = serde_json::json!({"top": 0, "right": right, "bottom": 0, "left": left});
            image["crop"] = crop;
            image
        };
        for image in [cropped(-0.5, 0.0), cropped(0.0, -0.1), cropped(0.6, 0.4)] {
            let dl = band_page("footer", 420.0, vec![image]);
            let hit = hit_test_regions(&dl, 0, 120.0, 455.0).unwrap();
            assert_eq!(hit.region, HitRegion::Footer);
            assert_eq!(hit.r_id.as_deref(), Some("rIdBand"));
        }
        let dl = band_page("footer", 420.0, vec![cropped(0.25, 0.25)]);
        assert_eq!(
            hit_test_regions(&dl, 0, 120.0, 455.0).unwrap().region,
            HitRegion::Body
        );
    }

    #[test]
    fn body_image_takes_footer_clicks_only_inside_the_rectangle_it_paints() {
        let with = |extra: serde_json::Value| {
            let mut image = image(100.0, 440.0, Some(10));
            for (key, value) in extra.as_object().unwrap() {
                image[key] = value.clone();
            }
            image
        };
        for extra in [
            serde_json::json!({"shapeType": "ellipse"}),
            serde_json::json!({"rotationDeg": 45}),
            serde_json::json!({"rotationDeg": 90}),
            serde_json::json!({"contentFrame": {"x": 130, "y": 440, "w": 30, "h": 40}}),
        ] {
            let dl = band_page("footer", 420.0, vec![with(extra)]);
            let hit = hit_test_regions(&dl, 0, 120.0, 455.0).unwrap();
            assert_eq!(hit.region, HitRegion::Footer);
            assert_eq!(hit.r_id.as_deref(), Some("rIdBand"));
        }
        for extra in [
            serde_json::json!({"shapeType": "rect"}),
            serde_json::json!({"rotationDeg": 180}),
            serde_json::json!({"contentFrame": {"x": 100, "y": 440, "w": 30, "h": 40}}),
        ] {
            let dl = band_page("footer", 420.0, vec![with(extra)]);
            assert_eq!(
                hit_test_regions(&dl, 0, 120.0, 455.0).unwrap().region,
                HitRegion::Body
            );
        }
    }

    #[test]
    fn transparent_body_image_leaves_footer_clicks_to_the_footer() {
        let visible = band_page("footer", 420.0, vec![image(100.0, 440.0, Some(10))]);
        assert_eq!(
            hit_test_regions(&visible, 0, 120.0, 455.0).unwrap().region,
            HitRegion::Body
        );

        let mut own = image(100.0, 440.0, Some(10));
        own["opacity"] = 0.into();
        let mut group = image(100.0, 440.0, Some(10));
        group["clipGroup"] = serde_json::json!({
            "clip": {"x": 0, "y": 0, "w": 500, "h": 500}, "opacity": 0
        });
        for image in [own, group] {
            let dl = band_page("footer", 420.0, vec![image]);
            let hit = hit_test_regions(&dl, 0, 120.0, 455.0).unwrap();
            assert_eq!(hit.region, HitRegion::Footer);
            assert_eq!(hit.r_id.as_deref(), Some("rIdBand"));
        }

        // The canvas applies group opacity only inside a clip.
        let mut unclipped = image(100.0, 440.0, Some(10));
        unclipped["clipGroup"] = serde_json::json!({"opacity": 0});
        let dl = band_page("footer", 420.0, vec![unclipped]);
        assert_eq!(
            hit_test_regions(&dl, 0, 120.0, 455.0).unwrap().region,
            HitRegion::Body
        );
    }

    #[test]
    fn turned_or_compressed_body_text_leaves_header_clicks_to_the_header() {
        let mut turned = run(100.0, 40.0, 50.0, 1);
        turned["rotationDeg"] = 90.into();
        let dl = band_page("header", 0.0, vec![turned]);
        let hit = hit_test_regions(&dl, 0, 120.0, 35.0).unwrap();
        assert_eq!(hit.region, HitRegion::Header);
        assert_eq!(hit.r_id.as_deref(), Some("rIdBand"));
        let mut compressed = run(100.0, 40.0, 50.0, 1);
        compressed["horizontalScale"] = 25.into();
        let dl = band_page("header", 0.0, vec![compressed]);
        assert_eq!(
            hit_test_regions(&dl, 0, 120.0, 35.0).unwrap().region,
            HitRegion::Header
        );
        let mut full_turn = run(100.0, 40.0, 50.0, 1);
        full_turn["rotationDeg"] = 360.into();
        let mut expanded = run(100.0, 40.0, 50.0, 1);
        expanded["horizontalScale"] = 150.into();
        for text in [full_turn, expanded] {
            let dl = band_page("header", 0.0, vec![text]);
            assert_eq!(
                hit_test_regions(&dl, 0, 120.0, 35.0).unwrap().region,
                HitRegion::Body
            );
        }
    }

    #[test]
    fn a_band_keeps_its_edge_where_body_text_paints_no_glyphs() {
        // Body lines start where the header ends (flush, then inside a padded band) at a 0.75em ascent.
        let flush = band_page("header", 0.0, vec![run(100.0, 92.0, 50.0, 1)]);
        let padded = band_page("header", 0.0, vec![run(100.0, 72.0, 50.0, 1)]);
        for (dl, y) in [
            (&flush, 73.0),
            (&flush, 77.0),
            (&flush, 80.0),
            (&padded, 59.0),
        ] {
            let hit = hit_test_regions(dl, 0, 120.0, y).unwrap();
            assert_eq!(hit.region, HitRegion::Header, "y {y}");
            assert_eq!(hit.r_id.as_deref(), Some("rIdBand"));
        }
        let straddling = band_page("header", 0.0, vec![run(100.0, 84.0, 50.0, 1)]);
        for (dl, y) in [(&flush, 85.0), (&padded, 66.0), (&straddling, 76.0)] {
            assert_eq!(
                hit_test_regions(dl, 0, 120.0, y).unwrap().region,
                HitRegion::Body,
                "y {y}"
            );
        }
        // The body's last line ends where the footer starts.
        let dl = band_page("footer", 420.0, vec![run(100.0, 416.0, 50.0, 1)]);
        for y in [420.0, 422.0, 424.0] {
            let hit = hit_test_regions(&dl, 0, 120.0, y).unwrap();
            assert_eq!(hit.region, HitRegion::Footer, "y {y}");
            assert_eq!(hit.r_id.as_deref(), Some("rIdBand"));
        }
    }

    #[test]
    fn body_text_that_paints_no_glyphs_leaves_header_clicks_to_the_header() {
        let glyph_run = |text: &str| {
            serde_json::json!({
                "kind": "glyphRun", "fontId": 1, "size": 16, "color": "#000000",
                "text": text, "docStart": 1, "docEnd": 6,
                "glyphs": [{"id": 1, "x": 100, "y": 40, "cluster": 0, "advance": 50}]
            })
        };
        let no_fill = serde_json::json!({"textFill": {"kind": "none"}});
        let mut unfilled_run = run(100.0, 40.0, 50.0, 1);
        unfilled_run["modernEffects"] = no_fill.clone();
        let mut unfilled_glyphs = glyph_run("hello");
        unfilled_glyphs["modernEffects"] = no_fill;
        let mut blank_run = run(100.0, 40.0, 50.0, 1);
        blank_run["text"] = "  \t".into();
        let mut no_glyphs = glyph_run("hello");
        no_glyphs["glyphs"] = serde_json::json!([]);
        let mut bom_run = run(100.0, 40.0, 50.0, 1);
        bom_run["text"] = "\u{feff} ".into();
        for text in [
            unfilled_run,
            unfilled_glyphs,
            blank_run,
            bom_run,
            glyph_run(" "),
            glyph_run("\u{feff}"),
            no_glyphs,
        ] {
            let dl = band_page("header", 0.0, vec![text]);
            let hit = hit_test_regions(&dl, 0, 120.0, 35.0).unwrap();
            assert_eq!(hit.region, HitRegion::Header);
            assert_eq!(hit.r_id.as_deref(), Some("rIdBand"));
        }
        let mut nel_run = run(100.0, 40.0, 50.0, 1);
        nel_run["text"] = "\u{85}".into();
        for text in [run(100.0, 40.0, 50.0, 1), glyph_run("hello"), nel_run] {
            let dl = band_page("header", 0.0, vec![text]);
            assert_eq!(
                hit_test_regions(&dl, 0, 120.0, 35.0).unwrap().region,
                HitRegion::Body
            );
        }
    }

    #[test]
    fn body_images_and_shapes_that_paint_nothing_leave_footer_clicks_to_the_footer() {
        let mut no_width = image(100.0, 440.0, Some(10));
        no_width["w"] = 0.into();
        let mut no_height = image(100.0, 455.0, Some(10));
        no_height["h"] = 0.into();
        let shape = |paint: Option<serde_json::Value>, fill: Option<&str>| {
            let mut shape = inline_shape_primitive(100.0, 440.0, 60.0, 40.0, 10, "shape:band");
            shape["geometryPath"] = serde_json::json!([
                {"type": "move", "x": 100, "y": 440}, {"type": "line", "x": 160, "y": 440},
                {"type": "line", "x": 160, "y": 480}, {"type": "line", "x": 100, "y": 480},
                {"type": "close"}
            ]);
            if let Some(paint) = paint {
                shape["fillPaint"] = paint;
            }
            if let Some(fill) = fill {
                shape["fill"] = fill.into();
            }
            shape
        };
        let mut no_path = shape(None, Some("#ff0000"));
        no_path["geometryPath"] = serde_json::json!([]);
        // Its box covers the click; its path does not.
        let mut triangle = shape(None, Some("#ff0000"));
        triangle["geometryPath"] = serde_json::json!([
            {"type": "move", "x": 100, "y": 440}, {"type": "line", "x": 160, "y": 480},
            {"type": "line", "x": 100, "y": 480}, {"type": "close"}
        ]);
        let mut bowtie = shape(None, Some("#ff0000"));
        bowtie["geometryPath"] = serde_json::json!([
            {"type": "move", "x": 100, "y": 440}, {"type": "line", "x": 160, "y": 480},
            {"type": "line", "x": 160, "y": 440}, {"type": "line", "x": 100, "y": 480},
            {"type": "close"}
        ]);
        let mut narrow_bowtie = shape(None, Some("#ff0000"));
        narrow_bowtie["w"] = serde_json::json!(0.011);
        narrow_bowtie["geometryPath"] = serde_json::json!([
            {"type": "move", "x": 100.002, "y": 440}, {"type": "line", "x": 100.011, "y": 480},
            {"type": "line", "x": 100.011, "y": 440}, {"type": "line", "x": 100.002, "y": 480},
            {"type": "close"}
        ]);
        let mut skewed = shape(None, Some("#ff0000"));
        skewed["w"] = serde_json::json!(0.004);
        skewed["geometryPath"] = serde_json::json!([
            {"type": "move", "x": 99.997, "y": 440}, {"type": "line", "x": 100.013, "y": 440},
            {"type": "line", "x": 100.013, "y": 480}, {"type": "line", "x": 100.009, "y": 480},
            {"type": "close"}
        ]);
        // A rectangle filling the left half of its box, over the click until flipped.
        let half = |transform: Option<serde_json::Value>| {
            let mut shape = shape(None, Some("#ff0000"));
            shape["geometryPath"] = serde_json::json!([
                {"type": "move", "x": 100, "y": 440}, {"type": "line", "x": 130, "y": 440},
                {"type": "line", "x": 130, "y": 480}, {"type": "line", "x": 100, "y": 480},
                {"type": "close"}
            ]);
            if let Some(transform) = transform {
                shape["transform"] = transform;
            }
            shape
        };
        let mut rotated = shape(None, Some("#ff0000"));
        rotated["transform"] = serde_json::json!({"rotation": 45});
        let mut turned = shape(None, Some("#ff0000"));
        turned["transform"] = serde_json::json!({"rotation": 180});
        for (x, primitive) in [
            (100.0, no_width),
            (120.0, no_height),
            (
                120.0,
                shape(Some(serde_json::json!({"kind": "none"})), Some("#ff0000")),
            ),
            (120.0, shape(None, None)),
            (120.0, shape(None, Some("transparent"))),
            (
                120.0,
                shape(
                    Some(serde_json::json!({"kind": "solid", "color": ""})),
                    Some("#ff0000"),
                ),
            ),
            (
                120.0,
                shape(Some(serde_json::json!({"kind": "picture"})), None),
            ),
            (
                120.0,
                shape(
                    Some(
                        serde_json::json!({"kind": "picture", "pictureRelId": "rId9", "pictureOpacity": 0}),
                    ),
                    None,
                ),
            ),
            (120.0, no_path),
            (140.0, triangle),
            (140.0, bowtie),
            (100.0055, narrow_bowtie),
            (
                120.0,
                shape(
                    Some(serde_json::json!({
                        "kind": "picture", "pictureRelId": "rId9",
                        "pictureStretchRect": {"left": 0.75}
                    })),
                    None,
                ),
            ),
            (
                120.0,
                shape(
                    Some(serde_json::json!({
                        "kind": "picture", "pictureRelId": "rId9",
                        "pictureSrcRect": {"left": -1}
                    })),
                    None,
                ),
            ),
            (
                120.0,
                shape(
                    Some(serde_json::json!({
                        "kind": "picture", "pictureRelId": "rId9", "pictureFillMode": "tile",
                        "pictureStretchRect": {"left": 0.75}
                    })),
                    None,
                ),
            ),
            (
                120.0,
                shape(
                    Some(serde_json::json!({
                        "kind": "picture", "pictureRelId": "rId9", "pictureFillMode": "tile",
                        "pictureSrcRect": {"left": -1}
                    })),
                    None,
                ),
            ),
            (100.002, skewed),
            (120.0, half(Some(serde_json::json!({"flipH": true})))),
            (120.0, half(Some(serde_json::json!({"rotation": 180})))),
            (120.0, rotated),
        ] {
            let dl = band_page("footer", 420.0, vec![primitive]);
            let hit = hit_test_regions(&dl, 0, x, 455.0).unwrap();
            assert_eq!(hit.region, HitRegion::Footer);
            assert_eq!(hit.r_id.as_deref(), Some("rIdBand"));
        }
        for primitive in [
            shape(None, Some("#ff0000")),
            turned,
            half(None),
            shape(
                Some(serde_json::json!({
                    "kind": "picture", "pictureRelId": "rId9",
                    "pictureStretchRect": {"left": -0.1, "top": 0}
                })),
                None,
            ),
            shape(
                Some(serde_json::json!({
                    "kind": "picture", "pictureRelId": "rId9",
                    "pictureSrcRect": {"left": 0.25, "right": 0.25}
                })),
                None,
            ),
            half(Some(serde_json::json!({"flipH": true, "rotation": 180}))),
            shape(
                Some(serde_json::json!({"kind": "solid", "color": "#00ff00"})),
                None,
            ),
            shape(Some(serde_json::json!({"kind": "gradient"})), None),
            shape(
                Some(serde_json::json!({"kind": "picture", "pictureRelId": "rId9"})),
                None,
            ),
        ] {
            let dl = band_page("footer", 420.0, vec![primitive]);
            let hit = hit_test_regions(&dl, 0, 120.0, 455.0).unwrap();
            assert_eq!(hit.region, HitRegion::Body);
            assert_eq!(hit.pos, Some(10));
        }
    }

    fn note_run(baseline: f64, doc_start: i64, group_id: &str) -> serde_json::Value {
        let mut run = run(100.0, baseline, 50.0, doc_start);
        run["groupId"] = group_id.into();
        run
    }

    /// [`hover_fixture`] plus a footnote area at y 360..420 stacking two note
    /// stories over x 100..150, with bands y 374..394 and y 394..414.
    fn note_fixture() -> DisplayList {
        let mut dl = hover_fixture();
        dl.pages[0].note_areas = serde_json::from_value(serde_json::json!([{
            "kind": "footnote",
            "y": 360,
            "height": 60,
            "noteIds": [7, 8],
            "primitives": [note_run(390.0, 1, "footnote-7"), note_run(410.0, 20, "footnote-8")]
        }]))
        .unwrap();
        dl
    }

    fn target(dl: &DisplayList, x: f64, y: f64) -> HoverTarget {
        hit_test_regions(dl, 0, x, y).unwrap().target
    }

    #[test]
    fn hover_target_reports_the_typeable_area_and_direct_hits() {
        let dl = hover_fixture();

        assert_eq!(target(&dl, 120.0, 95.0), HoverTarget::Text);
        assert_eq!(target(&dl, 120.0, 220.0), HoverTarget::Image);
        // right of a short line and below the last one — typeable, no primitive
        assert_eq!(target(&dl, 380.0, 95.0), HoverTarget::Text);
        assert_eq!(target(&dl, 120.0, 400.0), HoverTarget::Text);
        // margins and page background
        assert_eq!(target(&dl, 40.0, 95.0), HoverTarget::None);
        assert_eq!(target(&dl, 120.0, 460.0), HoverTarget::None);
    }

    /// The target discriminates where the position cannot: the nearest-line
    /// fallback answers over the margins too.
    #[test]
    fn hover_target_narrows_a_position_that_resolves_everywhere() {
        let dl = hover_fixture();
        let margin = hit_test_regions(&dl, 0, 40.0, 95.0).unwrap();
        assert!(margin.pos.is_some());
        assert_eq!(margin.target, HoverTarget::None);
    }

    /// Column boxes stand in for a list built without a content box; with
    /// neither, only the runs' own boxes answer.
    #[test]
    fn hover_target_without_a_content_box_falls_back() {
        let mut dl = hover_fixture();
        dl.pages[0].content_bounds = None;
        dl.pages[0].column_bounds = serde_json::from_value(
            serde_json::json!([{"x": 80, "y": 80, "width": 340, "height": 340}]),
        )
        .unwrap();
        assert_eq!(target(&dl, 380.0, 95.0), HoverTarget::Text);

        dl.pages[0].column_bounds.clear();
        assert_eq!(target(&dl, 120.0, 95.0), HoverTarget::Text);
        assert_eq!(target(&dl, 380.0, 95.0), HoverTarget::None);
    }

    #[test]
    fn hit_and_caret_resolve_combining_and_surrogate_clusters() {
        let mut primitive = run(100.0, 200.0, 60.0, 10);
        primitive["text"] = "a\u{301}😀b".into();
        primitive["docEnd"] = 15.into();
        let dl = page(serde_json::Value::Null, vec![primitive]);

        assert_eq!(hit_test(&dl, 0, 126.0, 195.0), Some(12));
        assert_eq!(hit_test(&dl, 0, 134.0, 195.0), Some(14));
        assert!((caret_rect(&dl, 12).unwrap().x - 120.0).abs() < 0.01);
        assert!((caret_rect(&dl, 14).unwrap().x - 140.0).abs() < 0.01);
    }

    /// Content positioned outside the content box — a floating text box in a
    /// margin — reads as text over its glyphs but not over the blank tail of
    /// its lines, which no container rect in the display list describes. The
    /// cursor under-claims there; it must never claim what a click will not do.
    #[test]
    fn hover_target_outside_the_content_box_covers_glyphs_only() {
        let dl = page(
            serde_json::json!({"x": 80, "y": 80, "width": 340, "height": 340}),
            vec![run(100.0, 100.0, 50.0, 1), run(440.0, 450.0, 50.0, 20)],
        );
        assert_eq!(target(&dl, 460.0, 445.0), HoverTarget::Text);

        let tail = hit_test_regions(&dl, 0, 520.0, 445.0).unwrap();
        assert_eq!(tail.target, HoverTarget::None);
        assert!(
            tail.pos.is_some(),
            "the blank tail still resolves a position"
        );
    }

    /// The gutter between columns is typeable — a click there lands in one of
    /// them — so the content box, not the column boxes, defines the area.
    #[test]
    fn hover_target_covers_the_gutter_between_columns() {
        let mut dl = hover_fixture();
        dl.pages[0].column_bounds = serde_json::from_value(serde_json::json!([
            {"x": 80, "y": 80, "width": 150, "height": 340},
            {"x": 270, "y": 80, "width": 150, "height": 340}
        ]))
        .unwrap();
        assert_eq!(target(&dl, 250.0, 300.0), HoverTarget::Text);
    }

    /// Note text is its own story, so a point in a note area addresses that
    /// note and its glyphs invite typing like the body's do.
    #[test]
    fn a_point_in_a_note_area_addresses_the_note_story() {
        let dl = note_fixture();

        let note = hit_test_regions(&dl, 0, 120.0, 385.0).unwrap();
        assert_eq!(note.region, HitRegion::Footnote);
        assert_eq!(note.note_id, Some(7));
        assert_eq!(note.target, HoverTarget::Text);
        assert!(matches!(note.pos, Some(pos) if (1..=6).contains(&pos)));

        // the body line above the area still answers for the body
        let body = hit_test_regions(&dl, 0, 120.0, 340.0).unwrap();
        assert_eq!(body.region, HitRegion::Body);
        assert_eq!(body.note_id, None);
        assert_eq!(body.target, HoverTarget::Text);
    }

    /// An area stating no band of its own cannot say where its notes paint, so
    /// it claims no point and the body answers — the same under-claim the
    /// typeable area makes off the content box.
    #[test]
    fn a_note_area_without_a_band_claims_no_point() {
        let mut dl = note_fixture();
        dl.pages[0].note_areas[0].y = None;
        let hit = hit_test_regions(&dl, 0, 120.0, 385.0).unwrap();
        assert_eq!(hit.region, HitRegion::Body);
        assert_eq!(hit.note_id, None);
    }

    /// One area stacks several independent stories, so a point must never
    /// borrow a position from the note above or below the one it landed in.
    #[test]
    fn stacked_notes_resolve_against_the_story_the_point_landed_in() {
        let dl = note_fixture();
        let second = hit_test_regions(&dl, 0, 120.0, 405.0).unwrap();
        assert_eq!(second.region, HitRegion::Footnote);
        assert_eq!(second.note_id, Some(8));
        assert!(matches!(second.pos, Some(pos) if (20..=25).contains(&pos)));
    }

    /// A note area carrying nothing still owns the click, like an empty
    /// header band does, but names no story to route it to.
    #[test]
    fn an_empty_note_area_owns_the_point_without_a_story() {
        let mut dl = hover_fixture();
        dl.pages[0].note_areas = serde_json::from_value(
            serde_json::json!([{"kind": "endnote", "y": 360, "height": 60}]),
        )
        .unwrap();
        let note = hit_test_regions(&dl, 0, 120.0, 390.0).unwrap();
        assert_eq!(note.region, HitRegion::Endnote);
        assert_eq!(note.note_id, None);
        assert_eq!(note.pos, None);
        assert_eq!(note.target, HoverTarget::None);
    }

    #[test]
    fn page_window_normalizes_float_bounds() {
        for (first, last, expected) in [
            (5.5, 5.5, None),
            (4.2, 5.0, Some((5, 5))),
            (f64::NAN, 5.0, None),
            (0.0, f64::NAN, None),
            (-3.0, 1.0, Some((0, 1))),
            (2.0, f64::INFINITY, Some((2, usize::MAX))),
            (5.0, 2.0, None),
            (-5.0, -1.0, None),
            (f64::NEG_INFINITY, f64::INFINITY, Some((0, usize::MAX))),
            (f64::INFINITY, f64::INFINITY, Some((usize::MAX, usize::MAX))),
            (0.0, f64::NEG_INFINITY, None),
        ] {
            assert_eq!(page_window(first, last), expected, "{first}..={last}");
        }
    }

    #[test]
    fn range_rects_on_pages_exports_normalize_float_bounds() {
        let mut dl = page(Value::Null, vec![run(100.0, 100.0, 50.0, 1)]);
        let template = dl.pages[0].clone();
        dl.pages = (0..8)
            .map(|page_index| DisplayPage {
                page_index,
                ..template.clone()
            })
            .collect();
        let json = serde_json::to_string(&dl).unwrap();
        let all: Vec<Value> =
            serde_json::from_str(&range_rects_json(&json, 2, 4).unwrap()).unwrap();
        assert_eq!(all.len(), dl.pages.len());
        let handle = crate::session::open_display_list(&json).unwrap();
        for (first, last) in [
            (5.5, 5.5),
            (4.2, 5.0),
            (1.2, 4.8),
            (f64::NAN, 5.0),
            (0.0, f64::NAN),
            (-3.0, 1.0),
            (-3.2, 2.8),
            (-5.0, -1.0),
            (2.0, f64::INFINITY),
            (5.0, 2.0),
            (f64::NEG_INFINITY, f64::INFINITY),
            (f64::INFINITY, f64::INFINITY),
            (0.0, f64::NEG_INFINITY),
            (0.0, 4294967296.0),
            (4294967296.0, f64::INFINITY),
        ] {
            let expected: Vec<Value> = all
                .iter()
                .filter(|rect| {
                    let index = rect["pageIndex"].as_u64().unwrap() as f64;
                    index >= first.ceil() && index <= last.floor()
                })
                .cloned()
                .collect();
            let by_json: Vec<Value> = serde_json::from_str(
                &crate::range_rects_on_pages_json(&json, 2.0, 4.0, first, last).unwrap(),
            )
            .unwrap();
            assert_eq!(by_json, expected, "JSON {first}..={last}");
            let by_handle: Vec<Value> = serde_json::from_str(
                &crate::range_rects_on_pages_by_handle(handle, 2.0, 4.0, first, last).unwrap(),
            )
            .unwrap();
            assert_eq!(by_handle, expected, "handle {first}..={last}");
        }
        crate::session::close_display_list(handle);
    }

    #[test]
    fn empty_page_windows_skip_display_list_and_handle_access() {
        for (first, last) in [(5.5, 5.5), (f64::NAN, 5.0), (0.0, f64::NAN), (5.0, 2.0)] {
            assert_eq!(
                crate::range_rects_on_pages_json("invalid", 2.0, 4.0, first, last).unwrap(),
                "[]"
            );
            assert_eq!(
                crate::range_rects_on_pages_by_handle(u32::MAX, 2.0, 4.0, first, last).unwrap(),
                "[]"
            );
        }
    }

    #[test]
    fn repeated_table_header_range_visits_only_page_five_zero_based() {
        let mut rows = Vec::new();
        let mut row_measures = Vec::new();
        for row in 0..21 {
            let height = if row == 0 { 20 } else { 40 };
            let cells: Vec<_> = (0..2)
                .map(|column| {
                    let start = row * 100 + column * 10 + 1;
                    serde_json::json!({
                        "id": format!("cell-r{row}c{column}"),
                        "blocks": [{
                            "kind": "paragraph", "id": format!("r{row}c{column}"),
                            "pmStart": start, "pmEnd": start + 4,
                            "runs": [{"kind": "text", "text": "head", "pmStart": start, "pmEnd": start + 4}]
                        }]
                    })
                })
                .collect();
            let cell_measures: Vec<_> = (0..2)
                .map(|_| {
                    serde_json::json!({
                        "width": 80, "height": height,
                        "blocks": [{"kind": "paragraph", "totalHeight": height,
                            "lines": [{"headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 4,
                                "width": 40, "ascent": 8, "descent": 2, "lineHeight": height}]}]
                    })
                })
                .collect();
            rows.push(serde_json::json!({"id": row, "isHeader": row == 0, "cantSplit": true, "cells": cells}));
            row_measures.push(serde_json::json!({"height": height, "cells": cell_measures}));
        }
        let mut input = serde_json::json!({
            "measured": [{
                "block": {"kind": "table", "id": "table", "columnWidths": [80, 80], "rows": rows},
                "measure": {"kind": "table", "columnWidths": [80, 80], "totalWidth": 160,
                    "totalHeight": 820, "rows": row_measures}
            }],
            "options": {"pageSize": {"w": 200, "h": 120},
                "margins": {"top": 10, "right": 10, "bottom": 10, "left": 10}}
        });
        input["layout"] =
            serde_json::from_str(&crate::layout_to_canonical_json(&input.to_string()).unwrap())
                .unwrap();
        let json = crate::display_list::build_display_list_json(&input.to_string()).unwrap();
        let dl: DisplayList = serde_json::from_str(&json).unwrap();
        assert!(dl.pages.len() >= 6);
        assert!(dl.pages[5].primitives.iter().any(|primitive| {
            let value = serde_json::to_value(primitive).unwrap();
            value["cell"]["repeatedHeader"] == true && value["docStart"] == 1
        }));
        let all = range_rects(&dl, 2, 4);
        assert_eq!(all.len(), dl.pages.len());
        let expected: Vec<_> = all
            .into_iter()
            .filter(|rect| rect.page_index == 5)
            .collect();
        assert!(!expected.is_empty());

        RANGE_RECT_PAGE_VISITS.with(|visits| visits.borrow_mut().clear());
        assert_eq!(range_rects_on_pages(&dl, [5], 2, 4), expected);
        RANGE_RECT_PAGE_VISITS.with(|visits| assert_eq!(*visits.borrow(), vec![5]));

        let expected_json = serde_json::to_string(&expected).unwrap();
        RANGE_RECT_PAGE_VISITS.with(|visits| visits.borrow_mut().clear());
        assert_eq!(
            range_rects_on_pages_json(&json, 2, 4, 5, 5).unwrap(),
            expected_json
        );
        RANGE_RECT_PAGE_VISITS.with(|visits| assert_eq!(*visits.borrow(), vec![5]));
        let handle = crate::session::open_display_list(&json).unwrap();
        RANGE_RECT_PAGE_VISITS.with(|visits| visits.borrow_mut().clear());
        assert_eq!(
            crate::session::range_rects_on_pages_by_handle(handle, 2, 4, 5, 5).unwrap(),
            expected_json
        );
        RANGE_RECT_PAGE_VISITS.with(|visits| assert_eq!(*visits.borrow(), vec![5]));
        assert_eq!(
            range_rects_on_pages_json(&json, 4, 2, 5, 5).unwrap(),
            expected_json
        );
        assert_eq!(range_rects_on_pages_json(&json, 2, 4, 6, 5).unwrap(), "[]");
        assert_eq!(
            range_rects_on_pages_json(&json, 2, 4, dl.pages.len(), usize::MAX).unwrap(),
            "[]"
        );
        crate::session::close_display_list(handle);
        assert!(crate::session::range_rects_on_pages_by_handle(handle, 2, 4, 5, 5).is_err());
    }

    #[test]
    fn range_rects_merge_same_line_runs_into_one_band() {
        // three adjacent same-line runs (a formatted or per-cluster line) and
        // one run on the next line: selecting across them yields one rect per
        // LINE, never one per run
        let owned = |x: f64, baseline: f64, doc_start: i64, line_index: u64| {
            let mut prim = run(x, baseline, 40.0, doc_start);
            prim["blockId"] = 7.into();
            prim["lineIndex"] = line_index.into();
            prim
        };
        let dl = page(
            serde_json::Value::Null,
            vec![
                owned(100.0, 200.0, 1, 0),
                owned(140.0, 200.0, 6, 0),
                owned(180.0, 200.0, 11, 0),
                owned(100.0, 230.0, 16, 1),
            ],
        );

        let rects = range_rects(&dl, 1, 21);
        assert_eq!(rects.len(), 2, "one merged band per line: {rects:?}");
        assert!((rects[0].x - 100.0).abs() < 0.01);
        assert!(
            (rects[0].width - 120.0).abs() < 0.01,
            "merged width {}",
            rects[0].width
        );
        assert!((rects[1].width - 40.0).abs() < 0.01);

        // a selection whose runs do not touch horizontally stays split
        let sparse = range_rects(&dl, 1, 3);
        assert_eq!(sparse.len(), 1);
        assert!(sparse[0].width < 40.0);
    }

    #[test]
    fn range_rects_cover_both_half_point_runs_in_one_band() {
        let sized = |x: f64, doc_start: i64, font: &str| {
            let mut prim = run(x, 200.0, 40.0, doc_start);
            prim["blockId"] = 7.into();
            prim["font"] = font.into();
            prim
        };
        let dl = page(
            serde_json::Value::Null,
            vec![
                sized(140.0, 6, "400 15.333333px Calibri"),
                sized(100.0, 1, "400 14.666667px Calibri"),
            ],
        );

        let rects = range_rects(&dl, 1, 11);
        assert_eq!(rects.len(), 1, "one merged band: {rects:?}");
        assert!((rects[0].x - 100.0).abs() < 0.01);
        assert!((rects[0].x + rects[0].width - 180.0).abs() < 0.01);
    }

    #[test]
    fn range_rects_do_not_bridge_unselected_text_between_leftward_runs() {
        let sized = |text: &str, x: f64, doc_start: i64, font: &str| {
            let mut prim = run(x, 200.0, 20.0, doc_start);
            prim["text"] = text.into();
            prim["docEnd"] = (doc_start + 2).into();
            prim["blockId"] = 7.into();
            prim["paraId"] = "paragraph-1".into();
            prim["font"] = font.into();
            prim
        };
        let dl = page(
            serde_json::Value::Null,
            vec![
                sized("ab", 100.0, 10, "400 14.666667px Calibri"),
                sized("middle", 120.0, 20, "400 14.666667px Calibri"),
                sized("cd", 140.0, 12, "400 15.333333px Calibri"),
            ],
        );

        let rects = range_rects(&dl, 10, 14);
        assert_eq!(rects.len(), 2, "separate selected bands: {rects:?}");
        assert!(
            rects
                .iter()
                .all(|rect| rect.x + rect.width <= 120.0 || rect.x >= 140.0),
            "unselected middle range must remain uncovered: {rects:?}"
        );
    }

    #[test]
    fn range_rects_do_not_merge_unowned_nonadjacent_runs() {
        let unowned = |x: f64, doc_start: i64| {
            let mut prim = run(x, 200.0, 40.0, doc_start);
            prim.as_object_mut().unwrap().remove("blockId");
            prim
        };
        let dl = page(
            serde_json::Value::Null,
            vec![unowned(100.0, 1), unowned(140.0, 20)],
        );

        let rects = range_rects(&dl, 1, 25);
        assert_eq!(
            rects.len(),
            2,
            "unowned document gaps stay split: {rects:?}"
        );
    }

    #[test]
    fn range_rects_never_merge_across_table_cells() {
        // two aligned, touching runs in adjacent cells of one table: the band
        // may not bridge the cell boundary even though the geometry allows it
        let celled = |x: f64, doc_start: i64, col: u64| {
            let mut prim = run(x, 200.0, 40.0, doc_start);
            prim["blockId"] = 7.into();
            prim["table"] = serde_json::json!({
                "tableId": "t1", "rowStart": 0, "rowEnd": 1,
                "rowCount": 1, "columnCount": 2
            });
            prim["cell"] = serde_json::json!({
                "row": 0, "col": col, "rowSpan": 1, "colSpan": 1
            });
            prim
        };
        let dl = page(
            serde_json::Value::Null,
            vec![celled(100.0, 1, 0), celled(140.0, 6, 1)],
        );

        let rects = range_rects(&dl, 1, 11);
        assert_eq!(rects.len(), 2, "one band per cell: {rects:?}");
        assert!((rects[0].width - 40.0).abs() < 0.01);
        assert!((rects[1].width - 40.0).abs() < 0.01);
    }

    #[test]
    fn indexed_range_rects_answer_every_range_like_a_single_query() {
        let owned = |x: f64, baseline: f64, doc_start: i64, line_index: u64| {
            let mut prim = run(x, baseline, 40.0, doc_start);
            prim["blockId"] = 7.into();
            prim["lineIndex"] = line_index.into();
            prim
        };
        let dl = page(
            serde_json::Value::Null,
            vec![
                owned(180.0, 200.0, 11, 0),
                owned(100.0, 200.0, 1, 0),
                owned(140.0, 200.0, 6, 0),
                owned(100.0, 230.0, 16, 1),
                image(100.0, 260.0, Some(21)),
                image(200.0, 260.0, None),
            ],
        );
        let index = RangeRectIndex::new(&dl.pages[0].primitives, 3);
        for from in 0..24 {
            for to in from..24 {
                let expected: Vec<_> = range_rects(&dl, from, to)
                    .into_iter()
                    .map(|rect| RangeRect {
                        page_index: 3,
                        ..rect
                    })
                    .collect();
                assert_eq!(index.rects(from, to), expected, "[{from}, {to})");
            }
        }
    }

    /// Selection geometry follows the same scoping: a range in a note's story
    /// highlights that note's glyphs, and the identical body range highlights
    /// the body's — the two documents never share rectangles.
    #[test]
    fn range_rects_in_a_note_cover_that_note_only() {
        let dl = note_fixture();

        let note = range_rects_in_region(&dl, RegionScope::Footnote(7), 1, 6);
        assert_eq!(note.len(), 1);
        assert!((note[0].y - 374.0).abs() < 1.0, "note rect y {}", note[0].y);

        let body = range_rects_in_region(&dl, RegionScope::Body, 1, 6);
        assert_eq!(body.len(), 1);
        assert!((body[0].y - 84.0).abs() < 1.0, "body rect y {}", body[0].y);

        // the sibling note's story shares the position range and no geometry
        let sibling = range_rects_in_region(&dl, RegionScope::Footnote(8), 1, 6);
        assert!(sibling.is_empty());
        // and an endnote scope matches no footnote area
        assert!(range_rects_in_region(&dl, RegionScope::Endnote(7), 1, 6).is_empty());
    }

    /// A picture with a document position is a click target whatever its wrap
    /// mode, and outranks text painted under it — the pointer path selects one
    /// before it asks for a position.
    #[test]
    fn hover_target_prefers_a_selectable_image_over_the_text_it_covers() {
        let dl = page(
            serde_json::json!({"x": 80, "y": 80, "width": 340, "height": 340}),
            vec![run(100.0, 100.0, 50.0, 1), image(100.0, 80.0, Some(10))],
        );
        assert_eq!(target(&dl, 120.0, 95.0), HoverTarget::Image);
    }

    /// A picture with no document position cannot be selected (a watermark),
    /// so text painted over it stays typeable.
    #[test]
    fn hover_target_reads_through_an_unpositioned_image() {
        let dl = page(
            serde_json::json!({"x": 80, "y": 80, "width": 340, "height": 340}),
            vec![image(80.0, 80.0, None), run(100.0, 100.0, 50.0, 1)],
        );
        assert_eq!(target(&dl, 120.0, 95.0), HoverTarget::Text);
        assert_eq!(target(&dl, 120.0, 110.0), HoverTarget::Text);
    }

    /// Only the TOPMOST image decides, as the click does: a watermark over a
    /// positioned picture leaves nothing to select, and the reverse selects.
    #[test]
    fn hover_target_takes_the_topmost_of_stacked_images() {
        let area = serde_json::json!({"x": 80, "y": 80, "width": 340, "height": 340});
        let covered = page(
            area.clone(),
            vec![image(100.0, 200.0, Some(10)), image(80.0, 180.0, None)],
        );
        assert_eq!(target(&covered, 120.0, 220.0), HoverTarget::None);

        let on_top = page(
            area,
            vec![image(80.0, 180.0, None), image(100.0, 200.0, Some(10))],
        );
        assert_eq!(target(&on_top, 120.0, 220.0), HoverTarget::Image);
    }

    /// With nothing positionable on the page a click jumps the caret to the
    /// document's end, so the area must not advertise typing.
    #[test]
    fn hover_target_ignores_an_area_with_no_positionable_text() {
        let dl = page(
            serde_json::json!({"x": 80, "y": 80, "width": 340, "height": 340}),
            vec![image(100.0, 200.0, None)],
        );
        let hit = hit_test_regions(&dl, 0, 300.0, 300.0).unwrap();
        assert_eq!(hit.pos, None);
        assert_eq!(hit.target, HoverTarget::None);
    }

    #[test]
    fn vertical_move_materializes_hits_only_for_neighboring_pages() {
        let pages: Vec<serde_json::Value> = (0..500)
            .map(|page_index| {
                let doc_start = page_index * 10 + 1;
                serde_json::json!({
                    "pageIndex": page_index,
                    "width": 500,
                    "height": 500,
                    "columnBounds": [{"x": 80, "y": 80, "width": 340, "height": 340}],
                    "primitives": [{
                        "kind": "text",
                        "text": "x",
                        "x": 100,
                        "baselineY": 100,
                        "width": 10,
                        "font": "400 16px Calibri",
                        "color": "#000000",
                        "docStart": doc_start,
                        "docEnd": doc_start + 1,
                        "blockId": page_index,
                        "lineIndex": 0
                    }]
                })
            })
            .collect();
        let display_list: DisplayList =
            serde_json::from_value(serde_json::json!({ "pages": pages })).unwrap();

        TEXT_HIT_BUILD_COUNT.with(|count| count.set(0));
        let movement = vertical_move(&display_list, 2501, VerticalDirection::Down, None).unwrap();
        let hit_builds = TEXT_HIT_BUILD_COUNT.with(std::cell::Cell::get);

        assert_eq!(movement.position, 2511);
        assert_eq!(hit_builds, 4);
    }

    /// Range queries walk every text primitive for overlap, but caret-stop
    /// construction (grapheme/cluster segmentation) must only run for the
    /// primitives the range actually touches.
    #[test]
    fn range_rects_build_caret_stops_only_for_overlapping_runs() {
        let runs: Vec<serde_json::Value> = (0..200)
            .map(|index| {
                let doc_start = 1 + index * 10;
                serde_json::json!({
                    "kind": "text",
                    "text": "aaaaaaaaaa",
                    "x": 100.0 + index as f64,
                    "baselineY": 200,
                    "width": 40,
                    "font": "400 16px Calibri",
                    "color": "#000000",
                    "docStart": doc_start,
                    "docEnd": doc_start + 10,
                    "blockId": index,
                    "lineIndex": index
                })
            })
            .collect();
        let dl: DisplayList = serde_json::from_value(serde_json::json!({
            "pages": [{ "pageIndex": 0, "width": 816, "height": 1056, "primitives": runs }]
        }))
        .unwrap();

        CARET_STOPS_BUILD_COUNT.with(|count| count.set(0));
        let rects = range_rects(&dl, 15, 18);
        let stop_builds = CARET_STOPS_BUILD_COUNT.with(std::cell::Cell::get);
        assert_eq!(rects.len(), 1);
        assert!(
            stop_builds <= 2,
            "a 3-char selection must not segment every run on the page ({stop_builds} builds)"
        );
    }

    /// Dense pages (fine-print tables, per-character formatting) must not cost
    /// a scan of every line built so far per primitive.
    #[test]
    fn visual_line_grouping_stays_linear_in_page_density() {
        const PRIMITIVES_PER_LINE: usize = 4;
        const PRIMITIVES_PER_PAGE: usize = 2_000;
        let lines_per_page = PRIMITIVES_PER_PAGE / PRIMITIVES_PER_LINE;
        let mut position = 1_i64;
        let mut caret = 0_i64;
        let pages: Vec<serde_json::Value> = (0..3)
            .map(|page_index| {
                let primitives: Vec<serde_json::Value> = (0..lines_per_page)
                    .flat_map(|line| {
                        (0..PRIMITIVES_PER_LINE)
                            .map(|column| {
                                let doc_start = position;
                                position += 2;
                                if page_index == 1 && line == 0 && column == 0 {
                                    caret = doc_start;
                                }
                                serde_json::json!({
                                    "kind": "text",
                                    "text": "xx",
                                    "x": 80.0 + column as f64 * 20.0,
                                    "baselineY": 80.0 + line as f64 * 12.0,
                                    "width": 20,
                                    "font": "400 16px Calibri",
                                    "color": "#000000",
                                    "docStart": doc_start,
                                    "docEnd": doc_start + 1,
                                    "blockId": line,
                                    "lineIndex": 0,
                                })
                            })
                            .collect::<Vec<_>>()
                    })
                    .collect();
                serde_json::json!({
                    "pageIndex": page_index,
                    "width": 500,
                    "height": 500,
                    "columnBounds": [{"x": 60, "y": 60, "width": 400, "height": 400}],
                    "primitives": primitives,
                })
            })
            .collect();
        let display_list: DisplayList =
            serde_json::from_value(serde_json::json!({ "pages": pages })).unwrap();

        LINE_OWNER_COMPARE_COUNT.with(|count| count.set(0));
        let movement = vertical_move(&display_list, caret, VerticalDirection::Down, None).unwrap();
        let compares = LINE_OWNER_COMPARE_COUNT.with(std::cell::Cell::get);

        assert_eq!(movement.position, caret + PRIMITIVES_PER_LINE as i64 * 2);
        assert!(
            compares <= PRIMITIVES_PER_PAGE,
            "grouping compared line owners {compares} times for {PRIMITIVES_PER_PAGE} \
             primitives per page"
        );
    }

    fn inline_shape_primitive(
        x: f64,
        y: f64,
        w: f64,
        h: f64,
        doc_start: i64,
        block_key: &str,
    ) -> serde_json::Value {
        serde_json::json!({
            "kind": "shape",
            "x": x, "y": y, "w": w, "h": h,
            "geometryPath": [
                {"type": "move", "x": 0, "y": 0},
                {"type": "line", "x": 1, "y": 1}
            ],
            "docStart": doc_start,
            "docEnd": doc_start + 1,
            "blockKey": block_key,
            "inlineShapeAtom": true
        })
    }

    fn text_atom(
        text: &str,
        x: f64,
        baseline: f64,
        width: f64,
        doc_start: i64,
    ) -> serde_json::Value {
        serde_json::json!({
            "kind": "text",
            "text": text,
            "x": x,
            "baselineY": baseline,
            "width": width,
            "font": "400 16px Calibri",
            "color": "#000000",
            "docStart": doc_start,
            "docEnd": doc_start + 1,
            "blockId": 1,
            "lineIndex": 0
        })
    }

    #[test]
    fn inline_shape_atom_hits_caret_and_range_without_images() {
        let content = serde_json::json!({"x": 80, "y": 80, "width": 340, "height": 340});
        let dl = page(
            content,
            vec![inline_shape_primitive(
                100.0,
                200.0,
                60.0,
                18.0,
                2,
                "shape:inline-test",
            )],
        );
        assert!(dl.pages[0].primitives.iter().all(|primitive| !matches!(
            primitive,
            crate::display_list::Primitive::Image(img) if img.rel_id.is_empty()
        )));
        let hit = hit_test_regions(&dl, 0, 130.0, 209.0).unwrap();
        assert_eq!(hit.pos, Some(2));
        assert_eq!(hit.target, HoverTarget::Image);
        let rects = range_rects(&dl, 2, 3);
        assert_eq!(rects.len(), 1);
        assert_eq!((rects[0].x, rects[0].y), (100.0, 200.0));
        assert_eq!((rects[0].width, rects[0].height), (60.0, 18.0));
        let before = caret_rect(&dl, 2).unwrap();
        assert_eq!((before.x, before.y, before.height), (100.0, 200.0, 18.0));
        let after = caret_rect(&dl, 3).unwrap();
        assert_eq!((after.x, after.y, after.height), (160.0, 200.0, 18.0));
    }

    #[test]
    fn inline_shape_between_text_keeps_atom_caret_at_edges() {
        let content = serde_json::json!({"x": 60, "y": 80, "width": 400, "height": 340});
        let dl = page(
            content,
            vec![
                text_atom("A", 80.0, 100.0, 10.0, 1),
                inline_shape_primitive(90.0, 82.0, 60.0, 18.0, 2, "shape:inline-test"),
                text_atom("B", 150.0, 100.0, 10.0, 3),
            ],
        );
        let hit = hit_test_regions(&dl, 0, 120.0, 91.0).unwrap();
        assert_eq!(hit.pos, Some(2));
        assert_eq!(hit.target, HoverTarget::Image);
        let rects = range_rects(&dl, 2, 3);
        assert!(
            rects
                .iter()
                .any(|rect| rect.x == 90.0 && rect.width == 60.0)
        );
        let before = caret_rect(&dl, 2).unwrap();
        assert_eq!((before.x, before.y), (90.0, 82.0));
        let end = caret_rect(&dl, 3).unwrap();
        assert_eq!(end.x, 150.0);
    }

    #[test]
    fn standalone_shape_without_atom_flag_stays_unselectable() {
        let content = serde_json::json!({"x": 80, "y": 80, "width": 340, "height": 340});
        let mut standalone = inline_shape_primitive(100.0, 200.0, 60.0, 18.0, 9, "shape:block");
        standalone
            .as_object_mut()
            .unwrap()
            .remove("inlineShapeAtom");
        let dl = page(content, vec![standalone]);
        let hit = hit_test_regions(&dl, 0, 130.0, 209.0).unwrap();
        assert_ne!(hit.target, HoverTarget::Image);
        assert!(range_rects(&dl, 9, 10).is_empty());
    }

    #[test]
    fn unselectable_image_on_top_blocks_selectable_image_and_shape() {
        let content = serde_json::json!({"x": 80, "y": 80, "width": 340, "height": 340});
        let covered_image = page(
            content.clone(),
            vec![image(100.0, 200.0, Some(10)), image(100.0, 200.0, None)],
        );
        assert_eq!(target(&covered_image, 130.0, 209.0), HoverTarget::None);
        let covered_shape = page(
            content.clone(),
            vec![
                inline_shape_primitive(100.0, 200.0, 60.0, 18.0, 2, "shape:inline-test"),
                image(100.0, 200.0, None),
            ],
        );
        assert_eq!(target(&covered_shape, 130.0, 209.0), HoverTarget::None);
        let shape_on_top = page(
            content,
            vec![
                image(100.0, 200.0, None),
                inline_shape_primitive(100.0, 200.0, 60.0, 18.0, 2, "shape:inline-test"),
            ],
        );
        assert_eq!(target(&shape_on_top, 130.0, 209.0), HoverTarget::Image);
    }

    #[test]
    fn grouped_child_without_doc_stays_parent_only_atom() {
        let content = serde_json::json!({"x": 80, "y": 80, "width": 340, "height": 340});
        let parent = inline_shape_primitive(100.0, 200.0, 60.0, 18.0, 2, "shape:inline-test");
        let child = serde_json::json!({
            "kind": "shape",
            "x": 105.0, "y": 204.0, "w": 10.0, "h": 8.0,
            "geometryPath": [
                {"type": "move", "x": 0, "y": 0},
                {"type": "line", "x": 1, "y": 1}
            ],
            "blockKey": "shape:inline-test:child:0",
            "inlineShapeAtom": true
        });
        let dl = page(content, vec![parent, child]);
        let hit = hit_test_regions(&dl, 0, 110.0, 208.0).unwrap();
        assert_eq!(hit.pos, Some(2));
        assert_eq!(hit.target, HoverTarget::Image);
        let rects = range_rects(&dl, 2, 3);
        assert_eq!(rects.len(), 1);
        assert_eq!((rects[0].x, rects[0].y), (100.0, 200.0));
        assert_eq!((rects[0].width, rects[0].height), (60.0, 18.0));
    }
}
