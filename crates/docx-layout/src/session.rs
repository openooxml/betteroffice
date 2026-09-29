//! Stateful display-list session handles: parse a [`DisplayList`] ONCE, then run
//! many hit-test / range-rect queries against it by handle with zero
//! re-serialization.
//!
//! Every interactive query (a click, a drag-mousemove) otherwise re-sends the
//! whole display-list JSON and Rust re-parses it. [`open_display_list`] parses
//! and stores the [`DisplayList`] behind a small monotonic handle; the by-handle
//! query entry points ([`hit_test_regions_by_handle`], [`range_rects_by_handle`])
//! source the parsed list from the handle map and reuse the exact hit/range
//! logic the JSON-arg exports call, so results are byte-identical — this is pure
//! perf. [`close_display_list`] drops a handle.
//!
//! Lifecycle / memory: the `createDisplayListQueries` facade
//! opens one handle per display-list build and closes it on dispose/replacement,
//! so at most one handle is live per editor at steady state. As a backstop
//! against a leaked handle (a facade that forgot to close, or a JS
//! FinalizationRegistry that has not run yet), the map is capped at
//! [`MAX_SESSIONS`]: opening past the cap evicts the oldest handle, so the store
//! can never grow unbounded. WASM is single-threaded, so a `thread_local`
//! doubles as the module-global store; native tests get an isolated store per
//! test thread.

use std::cell::RefCell;
use std::collections::VecDeque;
use std::collections::{HashMap, HashSet};

use crate::display_list::{DisplayList, DisplayPage};
use crate::hit::{
    RegionScope, VerticalDirection, body_range_span, hit_test_regions, parse_region_scope,
    range_rects_in_region, range_rects_on_pages, vertical_move,
};

/// Upper bound on concurrently-open handles. The facade keeps exactly one live,
/// so this only ever trips when a caller leaks handles; the oldest is then
/// evicted so the map stays bounded.
pub const MAX_SESSIONS: usize = 8;

/// What the store knows of a page's [`body_range_span`].
#[derive(Clone, Copy)]
enum BodySpan {
    Unknown,
    Exact(Option<(i64, i64)>),
    /// Covers the span, which a position shift moved after it was read.
    Widened(i64, i64),
}

/// A stored display list and the body spans its range queries have read, one
/// per page.
struct Stored {
    list: DisplayList,
    spans: RefCell<Vec<BodySpan>>,
}

impl Stored {
    fn new(list: DisplayList) -> Self {
        let spans = RefCell::new(vec![BodySpan::Unknown; list.pages.len()]);
        Self { list, spans }
    }

    /// Body range rects over `[from, to)`, reading only the pages whose body
    /// span it meets.
    fn body_range_rects(&self, from: i64, to: i64) -> Vec<crate::hit::RangeRect> {
        let (low, high) = (from.min(to), from.max(to));
        if low == high {
            return Vec::new();
        }
        let mut spans = self.spans.borrow_mut();
        let mut pages = Vec::new();
        for (index, page) in self.list.pages.iter().enumerate() {
            let span = match spans[index] {
                BodySpan::Exact(span) => span,
                BodySpan::Widened(start, end) if end <= low || start >= high => None,
                _ => {
                    let span = body_range_span(page);
                    spans[index] = BodySpan::Exact(span);
                    span
                }
            };
            if span.is_some_and(|(start, end)| start < high && end > low) {
                pages.push(index);
            }
        }
        range_rects_on_pages(&self.list, pages, from, to)
    }
}

/// The handle registry: parsed display lists keyed by handle id, plus the
/// insertion order used for oldest-first eviction.
struct Sessions {
    map: HashMap<u32, Stored>,
    /// handle ids in insertion order (front = oldest); the eviction queue
    order: VecDeque<u32>,
    /// monotonic id source; never hands out 0 (a reserved "no handle" sentinel)
    next_id: u32,
}

impl Sessions {
    fn new() -> Self {
        Sessions {
            map: HashMap::new(),
            order: VecDeque::new(),
            next_id: 1,
        }
    }

    fn open(&mut self, dl: DisplayList) -> u32 {
        let id = self.next_id;
        // wrap back to 1 (skip 0) after u32::MAX opens — collisions with a live
        // handle are impossible in practice given the small cap
        self.next_id = self.next_id.checked_add(1).unwrap_or(1);

        // leak backstop: evict oldest handles until there is room for this one
        while self.order.len() >= MAX_SESSIONS {
            match self.order.pop_front() {
                Some(old) => {
                    self.map.remove(&old);
                }
                None => break,
            }
        }
        self.map.insert(id, Stored::new(dl));
        self.order.push_back(id);
        id
    }

    fn get(&self, handle: u32) -> Option<&DisplayList> {
        self.map.get(&handle).map(|stored| &stored.list)
    }

    fn close(&mut self, handle: u32) {
        if self.map.remove(&handle).is_some()
            && let Some(pos) = self.order.iter().position(|&h| h == handle)
        {
            self.order.remove(pos);
        }
    }
}

thread_local! {
    static SESSIONS: RefCell<Sessions> = RefCell::new(Sessions::new());
}

/// Parse a display list once and store it behind a fresh handle id.
/// `Err` carries a `parse: ...` reason for malformed JSON (same shape as the
/// JSON-arg exports), so the caller can fall back to the JSON-arg path.
pub fn open_display_list(json: &str) -> Result<u32, String> {
    let dl: DisplayList = serde_json::from_str(json).map_err(|e| format!("parse: {e}"))?;
    Ok(SESSIONS.with(|s| s.borrow_mut().open(dl)))
}

/// Drop a handle so its parsed display list is freed. Idempotent — closing an
/// unknown/already-closed handle is a no-op.
pub fn close_display_list(handle: u32) {
    SESSIONS.with(|s| s.borrow_mut().close(handle));
}

/// Page-delta update payload for [`update_display_list`]: the next page array
/// is assembled from retained pages (`reuse`: `[next_index, previous_index]`
/// pairs) plus freshly parsed replacements (`replace`: `[next_index, page]`).
/// Every one of the `total` slots must be filled exactly once. With `keep`,
/// the update only replaces pages in place: `total` is the stored page count,
/// `replace` lists distinct slots, and every other slot keeps its page.
#[derive(serde::Deserialize)]
#[serde(rename_all = "camelCase")]
struct DisplayListUpdate {
    total: usize,
    #[serde(default)]
    contract_version: Option<u32>,
    #[serde(default)]
    keep: bool,
    #[serde(default)]
    reuse: Vec<(usize, usize)>,
    #[serde(default)]
    replace: Vec<(usize, DisplayPage)>,
    /// Retained pages whose doc positions moved: `[next_index, previous_index,
    /// run_lists, anchor_lists?]`, where each run list is applied in order and a run is
    /// `[start, count, mask, delta]` over the canonical primitive order (body,
    /// note separators+primitives per area, header, footer). Mask bits: 1
    /// docStart, 2 docEnd, 4 fragmentDocStart, 8 fragmentDocEnd, 16 inline
    /// widget pos — the owned frame-delta shift contract. `anchor_lists`, when
    /// present, holds one list per run list of `[area, note, start, end]` note
    /// anchors, each set after its run list.
    #[serde(default)]
    shift: Vec<ShiftEntry>,
}

type ShiftRun = (usize, usize, u8, i64);
type NoteAnchor = (usize, usize, Option<i64>, Option<i64>);

#[derive(serde::Deserialize)]
struct ShiftEntry(
    usize,
    usize,
    Vec<Vec<ShiftRun>>,
    #[serde(default)] Vec<Vec<NoteAnchor>>,
);

/// Apply a page-delta update to a stored display list, so an incremental
/// rebuild re-parses only its changed pages instead of the whole list. On any
/// inconsistency the handle is CLOSED before returning `Err`, so a caller's
/// fallback path can never query a half-updated list.
pub fn update_display_list(handle: u32, update_json: &str) -> Result<(), String> {
    let result = serde_json::from_str::<DisplayListUpdate>(update_json)
        .map_err(|e| format!("parse: {e}"))
        .and_then(|update| {
            SESSIONS.with(|s| {
                let mut sessions = s.borrow_mut();
                let stored = sessions
                    .map
                    .get_mut(&handle)
                    .ok_or_else(|| format!("unknown display-list handle {handle}"))?;
                apply_display_list_update(&mut stored.list, stored.spans.get_mut(), update)
            })
        });
    // Any failure (including a malformed payload) closes the handle: the
    // caller has already assumed ownership transfer, and its fallback is a
    // fresh open — never a query against this possibly-stale handle.
    if result.is_err() {
        close_display_list(handle);
    }
    result
}

const SHIFT_DOC_START: u8 = 1 << 0;
const SHIFT_DOC_END: u8 = 1 << 1;
const SHIFT_FRAGMENT_START: u8 = 1 << 2;
const SHIFT_FRAGMENT_END: u8 = 1 << 3;
const SHIFT_INLINE_WIDGET: u8 = 1 << 4;
/// Shift only the masked fields a primitive has (frame-delta run flag).
const SHIFT_PRESENT_ONLY: u8 = 1 << 7;

fn primitive_attrs_mut(
    primitive: &mut crate::display_list::Primitive,
) -> &mut crate::display_list::DocAttrs {
    use crate::display_list::Primitive;
    match primitive {
        Primitive::Text(value) => &mut value.attrs,
        Primitive::GlyphRun(value) => &mut value.attrs,
        Primitive::Rect(value) => &mut value.attrs,
        Primitive::Line(value) => &mut value.attrs,
        Primitive::Image(value) => &mut value.attrs,
        Primitive::Shape(value) => &mut value.attrs,
        Primitive::Decoration(value) => &mut value.attrs,
    }
}

fn shift_position_field(
    mask: u8,
    bit: u8,
    field: &mut Option<i64>,
    name: &str,
    delta: i64,
) -> Result<(), String> {
    if mask & bit == 0 || (mask & SHIFT_PRESENT_ONLY != 0 && field.is_none()) {
        return Ok(());
    }
    let value = field.ok_or_else(|| format!("position shift requires retained {name}"))?;
    *field = Some(
        value
            .checked_add(delta)
            .ok_or_else(|| format!("position shift overflows {name}"))?,
    );
    Ok(())
}

fn shift_primitive_positions(
    primitive: &mut crate::display_list::Primitive,
    mask: u8,
    delta: i64,
) -> Result<(), String> {
    let attrs = primitive_attrs_mut(primitive);
    shift_position_field(
        mask,
        SHIFT_DOC_START,
        &mut attrs.doc_start,
        "docStart",
        delta,
    )?;
    shift_position_field(mask, SHIFT_DOC_END, &mut attrs.doc_end, "docEnd", delta)?;
    shift_position_field(
        mask,
        SHIFT_FRAGMENT_START,
        &mut attrs.fragment_doc_start,
        "fragmentDocStart",
        delta,
    )?;
    shift_position_field(
        mask,
        SHIFT_FRAGMENT_END,
        &mut attrs.fragment_doc_end,
        "fragmentDocEnd",
        delta,
    )?;
    if mask & SHIFT_INLINE_WIDGET != 0
        && (mask & SHIFT_PRESENT_ONLY == 0 || attrs.inline_sdt_widget.is_some())
    {
        let widget = attrs
            .inline_sdt_widget
            .as_mut()
            .ok_or_else(|| "position shift requires retained inline widget metadata".to_owned())?;
        widget.pos = widget
            .pos
            .checked_add(delta)
            .ok_or_else(|| "position shift overflows inline widget pos".to_owned())?;
    }
    Ok(())
}

/// Replays one owned-frame position-shift run list onto a retained page, in
/// the same canonical primitive order the encoder used: body primitives, each
/// note area's separators then primitives, header, footer.
fn shift_page_positions(page: &mut DisplayPage, runs: &[ShiftRun]) -> Result<(), String> {
    const FIELDS: u8 = SHIFT_DOC_START
        | SHIFT_DOC_END
        | SHIFT_FRAGMENT_START
        | SHIFT_FRAGMENT_END
        | SHIFT_INLINE_WIDGET;
    if runs
        .iter()
        .any(|run| run.2 & FIELDS == 0 || run.2 & !(FIELDS | SHIFT_PRESENT_ONLY) != 0)
    {
        return Err("position shift run has an invalid field mask".to_owned());
    }
    let mut index = 0usize;
    let mut cursor = 0usize;
    let mut apply = |primitive: &mut crate::display_list::Primitive| -> Result<(), String> {
        while cursor < runs.len() && runs[cursor].0 + runs[cursor].1 <= index {
            cursor += 1;
        }
        if let Some(run) = runs.get(cursor)
            && index >= run.0
            && index < run.0 + run.1
        {
            shift_primitive_positions(primitive, run.2, run.3)?;
        }
        index += 1;
        Ok(())
    };
    for primitive in &mut page.primitives {
        apply(primitive)?;
    }
    for area in &mut page.note_areas {
        for primitive in &mut area.separator_primitives {
            apply(primitive)?;
        }
        for primitive in &mut area.primitives {
            apply(primitive)?;
        }
    }
    if let Some(header) = &mut page.header {
        for primitive in &mut header.primitives {
            apply(primitive)?;
        }
    }
    if let Some(footer) = &mut page.footer {
        for primitive in &mut footer.primitives {
            apply(primitive)?;
        }
    }
    if runs.last().is_some_and(|run| run.0 + run.1 > index) {
        return Err("position shift range exceeds page primitive count".to_owned());
    }
    Ok(())
}

fn apply_display_list_update(
    dl: &mut DisplayList,
    spans: &mut Vec<BodySpan>,
    update: DisplayListUpdate,
) -> Result<(), String> {
    if update.keep {
        return replace_pages_in_place(dl, spans, update);
    }
    let slots = update
        .reuse
        .len()
        .saturating_add(update.replace.len())
        .saturating_add(update.shift.len());
    if slots != update.total {
        return Err("update slots do not cover the page total exactly".to_owned());
    }
    let mut previous: Vec<Option<DisplayPage>> = dl.pages.drain(..).map(Some).collect();
    let previous_spans = std::mem::take(spans);
    let span_of = |index: usize| {
        previous_spans
            .get(index)
            .copied()
            .unwrap_or(BodySpan::Unknown)
    };
    let mut next: Vec<Option<DisplayPage>> = Vec::new();
    next.resize_with(update.total, || None);
    let mut next_spans = vec![BodySpan::Unknown; update.total];
    for (next_index, previous_index) in update.reuse {
        if let Some(span) = next_spans.get_mut(next_index) {
            *span = span_of(previous_index);
        }
        let page = previous
            .get_mut(previous_index)
            .and_then(Option::take)
            .ok_or_else(|| format!("reused page {previous_index} is missing"))?;
        let slot = next
            .get_mut(next_index)
            .ok_or_else(|| format!("page target {next_index} out of range"))?;
        if slot.is_some() {
            return Err(format!("duplicate page target {next_index}"));
        }
        *slot = Some(page);
    }
    for (next_index, page) in update.replace {
        let slot = next
            .get_mut(next_index)
            .ok_or_else(|| format!("page target {next_index} out of range"))?;
        if slot.is_some() {
            return Err(format!("duplicate page target {next_index}"));
        }
        *slot = Some(page);
    }
    for ShiftEntry(next_index, previous_index, run_lists, anchor_lists) in update.shift {
        if !anchor_lists.is_empty() && anchor_lists.len() != run_lists.len() {
            return Err("note anchor lists do not match the run lists".to_owned());
        }
        let mut page = previous
            .get_mut(previous_index)
            .and_then(Option::take)
            .ok_or_else(|| format!("shifted page {previous_index} is missing"))?;
        for (step, runs) in run_lists.iter().enumerate() {
            shift_page_positions(&mut page, runs)?;
            for &(area, note, start, end) in anchor_lists.get(step).into_iter().flatten() {
                let note = page
                    .note_areas
                    .get_mut(area)
                    .and_then(|area| area.notes.get_mut(note))
                    .ok_or_else(|| "note anchor shift references an unknown note".to_owned())?;
                note.anchor_doc_start = start;
                note.anchor_doc_end = end;
            }
        }
        if let Some(span) = next_spans.get_mut(next_index) {
            *span = widened(span_of(previous_index), &run_lists);
        }
        let slot = next
            .get_mut(next_index)
            .ok_or_else(|| format!("page target {next_index} out of range"))?;
        if slot.is_some() {
            return Err(format!("duplicate page target {next_index}"));
        }
        *slot = Some(page);
    }
    dl.pages = next
        .into_iter()
        .enumerate()
        .map(|(index, page)| page.ok_or_else(|| format!("page {index} missing from update")))
        .collect::<Result<Vec<_>, _>>()?;
    dl.contract_version = update.contract_version;
    *spans = next_spans;
    Ok(())
}

/// A `keep` update: each replaced slot gets its new page, and every other
/// slot keeps the page it holds, so loading a few pages costs those pages.
fn replace_pages_in_place(
    dl: &mut DisplayList,
    spans: &mut [BodySpan],
    update: DisplayListUpdate,
) -> Result<(), String> {
    if !update.reuse.is_empty() || !update.shift.is_empty() {
        return Err("an update that keeps pages in place only replaces them".to_owned());
    }
    if update.total != dl.pages.len() {
        return Err("an update that keeps pages in place cannot change the page total".to_owned());
    }
    let mut targets = HashSet::new();
    for (index, _) in &update.replace {
        if *index >= update.total {
            return Err(format!("page target {index} out of range"));
        }
        if !targets.insert(*index) {
            return Err(format!("duplicate page target {index}"));
        }
    }
    for (index, page) in update.replace {
        dl.pages[index] = page;
        spans[index] = BodySpan::Unknown;
    }
    dl.contract_version = update.contract_version;
    Ok(())
}

/// `span` after its page's positions moved by the shift runs' deltas.
fn widened(span: BodySpan, run_lists: &[Vec<(usize, usize, u8, i64)>]) -> BodySpan {
    let (mut start, mut end) = match span {
        BodySpan::Unknown | BodySpan::Exact(None) => return span,
        BodySpan::Exact(Some((start, end))) | BodySpan::Widened(start, end) => (start, end),
    };
    // Each run list moves a position by at most its extreme deltas.
    for runs in run_lists {
        let deltas = runs.iter().map(|run| run.3);
        start = start.saturating_add(deltas.clone().min().unwrap_or(0).min(0));
        end = end.saturating_add(deltas.max().unwrap_or(0).max(0));
    }
    BodySpan::Widened(start, end)
}

/// Region-aware hit test against a stored display list — the by-handle twin of
/// [`crate::hit::hit_test_regions_json`]. `Err` when the handle is unknown
/// (closed or evicted); the caller falls back to the JSON-arg export.
pub fn hit_test_regions_by_handle(
    handle: u32,
    page_index: usize,
    x: f64,
    y: f64,
) -> Result<String, String> {
    SESSIONS.with(|s| {
        let sessions = s.borrow();
        let dl = sessions
            .get(handle)
            .ok_or_else(|| format!("unknown display-list handle {handle}"))?;
        match hit_test_regions(dl, page_index, x, y) {
            Some(hit) => serde_json::to_string(&hit).map_err(|e| format!("serialize: {e}")),
            None => Ok("null".to_string()),
        }
    })
}

pub fn vertical_move_by_handle(
    handle: u32,
    position: i64,
    direction: &str,
    goal_x: f64,
) -> Result<String, String> {
    SESSIONS.with(|s| {
        let sessions = s.borrow();
        let dl = sessions
            .get(handle)
            .ok_or_else(|| format!("unknown display-list handle {handle}"))?;
        let direction = match direction {
            "up" => VerticalDirection::Up,
            "down" => VerticalDirection::Down,
            other => return Err(format!("unknown vertical direction {other:?}")),
        };
        serde_json::to_string(&vertical_move(
            dl,
            position,
            direction,
            goal_x.is_finite().then_some(goal_x),
        ))
        .map_err(|e| format!("serialize: {e}"))
    })
}

/// Range rects against a stored display list — the by-handle twin of
/// [`crate::hit::range_rects_json`]. `Err` on an unknown handle.
pub fn range_rects_by_handle(handle: u32, from: i64, to: i64) -> Result<String, String> {
    SESSIONS.with(|s| {
        let sessions = s.borrow();
        let stored = sessions
            .map
            .get(&handle)
            .ok_or_else(|| format!("unknown display-list handle {handle}"))?;
        serde_json::to_string(&stored.body_range_rects(from, to))
            .map_err(|e| format!("serialize: {e}"))
    })
}

/// Region-aware range rects against a stored display list — the by-handle twin
/// of [`crate::hit::range_rects_region_json`]. `region` is
/// `"body" | "header" | "footer" | "footnote" | "endnote"`; `part_id` scopes
/// header/footer to one HF part (empty ⇒ match any) and names the note id for
/// a note region. `Err` on an unknown handle or an unparseable region.
pub fn range_rects_region_by_handle(
    handle: u32,
    region: &str,
    part_id: &str,
    from: i64,
    to: i64,
) -> Result<String, String> {
    SESSIONS.with(|s| {
        let sessions = s.borrow();
        let stored = sessions
            .map
            .get(&handle)
            .ok_or_else(|| format!("unknown display-list handle {handle}"))?;
        let rects = match parse_region_scope(region, part_id)? {
            RegionScope::Body => stored.body_range_rects(from, to),
            scope => range_rects_in_region(&stored.list, scope, from, to),
        };
        serde_json::to_string(&rects).map_err(|e| format!("serialize: {e}"))
    })
}

/// Number of currently-open handles (test/observability helper).
#[cfg(test)]
pub fn open_count() -> usize {
    SESSIONS.with(|s| s.borrow().map.len())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::hit::{hit_test_regions_json, range_rects_json, vertical_move_json};

    // a minimal one-page display list with a single positioned text primitive,
    // enough to exercise both a real hit and a range rect
    const SAMPLE: &str = r##"{
        "pages": [{
            "pageIndex": 0,
            "width": 816,
            "height": 1056,
            "primitives": [{
                "kind": "text",
                "text": "Hello",
                "x": 100,
                "baselineY": 200,
                "width": 50,
                "font": "400 16px Arial",
                "color": "#000000",
                "docStart": 1,
                "docEnd": 6
            }]
        }]
    }"##;

    fn drain() {
        // close everything so a test starts from a clean per-thread store
        for h in 1..=(MAX_SESSIONS as u32 * 4) {
            close_display_list(h);
        }
    }

    #[test]
    fn by_handle_results_are_byte_identical_to_json_arg() {
        drain();
        let handle = open_display_list(SAMPLE).expect("opens");

        // a direct hit inside the run, an edge snap, and a whole-run range — each
        // by-handle result must equal the JSON-arg result verbatim
        for (x, y) in [(120.0, 195.0), (400.0, 195.0), (100.0, 999.0)] {
            let by_handle = hit_test_regions_by_handle(handle, 0, x, y).unwrap();
            let by_json = hit_test_regions_json(SAMPLE, 0, x, y).unwrap();
            assert_eq!(by_handle, by_json, "hit ({x},{y}) differs");
        }
        for (from, to) in [(1, 6), (2, 4), (0, 0)] {
            let by_handle = range_rects_by_handle(handle, from, to).unwrap();
            let by_json = range_rects_json(SAMPLE, from, to).unwrap();
            assert_eq!(by_handle, by_json, "range ({from},{to}) differs");
        }
        assert_eq!(
            vertical_move_by_handle(handle, 3, "down", f64::NAN).unwrap(),
            vertical_move_json(SAMPLE, 3, "down", f64::NAN).unwrap()
        );

        // an out-of-range page returns "null", same as the JSON-arg export
        assert_eq!(
            hit_test_regions_by_handle(handle, 9, 1.0, 1.0).unwrap(),
            "null"
        );

        close_display_list(handle);
    }

    #[test]
    fn closed_and_unknown_handles_error() {
        drain();
        let handle = open_display_list(SAMPLE).expect("opens");
        assert!(hit_test_regions_by_handle(handle, 0, 120.0, 195.0).is_ok());

        close_display_list(handle);
        assert!(
            hit_test_regions_by_handle(handle, 0, 120.0, 195.0).is_err(),
            "querying a closed handle errors so the caller falls back"
        );
        assert!(range_rects_by_handle(handle, 1, 6).is_err());
        assert!(
            hit_test_regions_by_handle(999_999, 0, 1.0, 1.0).is_err(),
            "an unknown handle errors"
        );

        // close is idempotent
        close_display_list(handle);
    }

    #[test]
    fn malformed_json_reports_a_parse_error() {
        drain();
        let err = open_display_list("{ not a display list").unwrap_err();
        assert!(err.starts_with("parse: "), "reason: {err}");
    }

    #[test]
    fn distinct_handles_are_handed_out_monotonically() {
        drain();
        let a = open_display_list(SAMPLE).unwrap();
        let b = open_display_list(SAMPLE).unwrap();
        assert!(b > a, "monotonic ids: {a} then {b}");
        assert_ne!(a, 0, "0 is reserved as a no-handle sentinel");
        close_display_list(a);
        close_display_list(b);
    }

    #[test]
    fn page_delta_update_matches_a_fresh_open() {
        drain();
        let two_pages = |second_text: &str| {
            format!(
                r##"{{"pages": [
                    {{"pageIndex": 0, "width": 816, "height": 1056, "primitives": [{{
                        "kind": "text", "text": "Hello", "x": 100, "baselineY": 200,
                        "width": 50, "font": "400 16px Arial", "color": "#000000",
                        "docStart": 1, "docEnd": 6
                    }}]}},
                    {{"pageIndex": 1, "width": 816, "height": 1056, "primitives": [{{
                        "kind": "text", "text": "{second_text}", "x": 100, "baselineY": 200,
                        "width": 50, "font": "400 16px Arial", "color": "#000000",
                        "docStart": 7, "docEnd": 12
                    }}]}}
                ]}}"##
            )
        };
        let handle = open_display_list(&two_pages("world")).expect("opens");
        let replacement: serde_json::Value =
            serde_json::from_str(&two_pages("patch!")).expect("list json");
        let update = serde_json::json!({
            "total": 2,
            "reuse": [[0, 0]],
            "replace": [[1, replacement["pages"][1]]],
        });
        update_display_list(handle, &update.to_string()).expect("updates");

        let fresh = open_display_list(&two_pages("patch!")).expect("opens");
        for (from, to) in [(1, 6), (7, 12), (0, 0)] {
            assert_eq!(
                range_rects_by_handle(handle, from, to).unwrap(),
                range_rects_by_handle(fresh, from, to).unwrap(),
                "range ({from},{to}) differs from a fresh open"
            );
        }
        close_display_list(handle);
        close_display_list(fresh);
    }

    #[test]
    fn an_update_that_keeps_pages_replaces_only_its_slots() {
        drain();
        let pages = |texts: [&str; 3]| {
            let page = |index: usize, text: &str| {
                serde_json::json!({
                    "pageIndex": index, "width": 816, "height": 1056,
                    "primitives": [{
                        "kind": "text", "text": text, "x": 100, "baselineY": 200,
                        "width": 50, "font": "400 16px Arial", "color": "#000000",
                        "docStart": 1 + index * 6, "docEnd": 6 + index * 6
                    }]
                })
            };
            serde_json::json!({
                "pages": texts.iter().enumerate().map(|(i, text)| page(i, text)).collect::<Vec<_>>()
            })
        };
        let handle =
            open_display_list(&pages(["Hello", "world", "again"]).to_string()).expect("opens");
        let next = pages(["Hello", "patch!", "again"]);
        let update = serde_json::json!({
            "total": 3,
            "keep": true,
            "replace": [[1, next["pages"][1]]],
        });
        update_display_list(handle, &update.to_string()).expect("updates");

        let fresh = open_display_list(&next.to_string()).expect("opens");
        for (from, to) in [(1, 6), (7, 12), (13, 18), (1, 18), (0, 0)] {
            assert_eq!(
                range_rects_by_handle(handle, from, to).unwrap(),
                range_rects_by_handle(fresh, from, to).unwrap(),
                "range ({from},{to}) differs from a fresh open"
            );
        }
        close_display_list(fresh);

        let refused = |update: serde_json::Value| {
            let handle =
                open_display_list(&pages(["Hello", "world", "again"]).to_string()).expect("opens");
            assert!(
                update_display_list(handle, &update.to_string()).is_err(),
                "{update}"
            );
            assert!(
                range_rects_by_handle(handle, 1, 6).is_err(),
                "a refused update closes the handle"
            );
        };
        let page = next["pages"][1].clone();
        refused(serde_json::json!({ "total": 2, "keep": true, "replace": [[1, page]] }));
        refused(serde_json::json!({ "total": 3, "keep": true, "replace": [[3, page]] }));
        refused(serde_json::json!({ "total": 3, "keep": true, "replace": [[1, page], [1, page]] }));
        refused(
            serde_json::json!({ "total": 3, "keep": true, "reuse": [[0, 0]], "replace": [[1, page]] }),
        );
        close_display_list(handle);
    }

    #[test]
    fn shift_update_matches_a_fresh_open_of_shifted_positions() {
        drain();
        let two_pages = |first_start: i64, second_start: i64| {
            format!(
                r##"{{"pages": [
                    {{"pageIndex": 0, "width": 816, "height": 1056, "primitives": [{{
                        "kind": "text", "text": "Hello", "x": 100, "baselineY": 200,
                        "width": 50, "font": "400 16px Arial", "color": "#000000",
                        "docStart": {first_start}, "docEnd": {first_end}
                    }}]}},
                    {{"pageIndex": 1, "width": 816, "height": 1056, "primitives": [{{
                        "kind": "text", "text": "World", "x": 100, "baselineY": 200,
                        "width": 50, "font": "400 16px Arial", "color": "#000000",
                        "docStart": {second_start}, "docEnd": {second_end}
                    }}]}}
                ]}}"##,
                first_end = first_start + 5,
                second_end = second_start + 5,
            )
        };
        let handle = open_display_list(&two_pages(1, 7)).expect("opens");

        // shift page 1's doc positions by +3 (mask 15: doc + fragment fields;
        // fragments are absent here so only doc positions move) — applied as
        // two sequential run lists (+2 then +1) to cover multi-revision replay
        let update = serde_json::json!({
            "total": 2,
            "reuse": [[0, 0]],
            "shift": [[1, 1, [[[0, 1, 3, 2]], [[0, 1, 3, 1]]]]],
        });
        update_display_list(handle, &update.to_string()).expect("updates");

        let fresh = open_display_list(&two_pages(1, 10)).expect("opens");
        for (from, to) in [(1, 6), (10, 15), (7, 12), (0, 0)] {
            assert_eq!(
                range_rects_by_handle(handle, from, to).unwrap(),
                range_rects_by_handle(fresh, from, to).unwrap(),
                "range ({from},{to}) differs from a fresh open"
            );
        }
        close_display_list(handle);
        close_display_list(fresh);
    }

    #[test]
    fn body_ranges_read_after_updates_match_a_fresh_open() {
        drain();
        let pages = |starts: &[i64], marker: Option<i64>| {
            let text = |start: i64| {
                format!(
                    r##"{{"kind": "text", "text": "Hello", "x": 100, "baselineY": 200,
                        "width": 50, "font": "400 16px Arial", "color": "#000000",
                        "docStart": {start}, "docEnd": {end}}}"##,
                    end = start + 5
                )
            };
            let pages: Vec<String> = starts
                .iter()
                .enumerate()
                .map(|(index, start)| {
                    let mut primitives = vec![text(*start)];
                    if index == 1
                        && let Some(at) = marker
                    {
                        primitives.push(format!(
                            r##"{{"kind": "text", "text": "", "x": 100, "baselineY": 240,
                                "width": 0, "font": "400 16px Arial", "color": "#000000",
                                "docStart": {at}, "docEnd": {at}}}"##
                        ));
                    }
                    format!(
                        r##"{{"pageIndex": {index}, "width": 816, "height": 1056,
                            "primitives": [{}]}}"##,
                        primitives.join(",")
                    )
                })
                .collect();
            format!(r##"{{"pages": [{}]}}"##, pages.join(","))
        };
        let every_range = |handle: u32, fresh: u32| {
            for from in 0..40 {
                for to in [from, from + 1, from + 3, from + 12] {
                    assert_eq!(
                        range_rects_by_handle(handle, from, to).unwrap(),
                        range_rects_by_handle(fresh, from, to).unwrap(),
                        "range ({from},{to})"
                    );
                }
            }
        };
        let handle = open_display_list(&pages(&[1, 8, 16], Some(14))).expect("opens");
        let fresh = open_display_list(&pages(&[1, 8, 16], Some(14))).expect("opens");
        every_range(handle, fresh);

        let update = serde_json::json!({
            "total": 3,
            "reuse": [[0, 0]],
            "shift": [[1, 1, [[[0, 2, 3, -4]]]], [2, 2, [[[0, 1, 3, 9]]]]],
        });
        update_display_list(handle, &update.to_string()).expect("updates");
        let shifted = open_display_list(&pages(&[1, 4, 25], Some(10))).expect("opens");
        every_range(handle, shifted);

        let replacement: serde_json::Value =
            serde_json::from_str(&pages(&[2, 4, 25], None)).expect("list json");
        let update = serde_json::json!({
            "total": 3,
            "reuse": [[1, 1], [2, 2]],
            "replace": [[0, replacement["pages"][0]]],
        });
        update_display_list(handle, &update.to_string()).expect("updates");
        let replaced = open_display_list(&pages(&[2, 4, 25], Some(10))).expect("opens");
        every_range(handle, replaced);
        for handle in [handle, fresh, shifted, replaced] {
            close_display_list(handle);
        }
    }

    #[test]
    fn a_present_only_shift_run_moves_what_exact_runs_move() {
        drain();
        let list = |text_start: i64, rect: i64, widget: i64, header: i64| {
            format!(
                r##"{{"pages": [{{"pageIndex": 0, "width": 816, "height": 1056,
                    "primitives": [
                        {{"kind": "text", "text": "Hello", "x": 100, "baselineY": 200,
                          "width": 50, "font": "400 16px Arial", "color": "#000000",
                          "docStart": {text_start}, "docEnd": {text_end}}},
                        {{"kind": "rect", "x": 0, "y": 0, "w": 5, "h": 5, "fill": "#000000",
                          "fragmentDocStart": {rect}, "fragmentDocEnd": {rect_end}}},
                        {{"kind": "rect", "x": 0, "y": 9, "w": 5, "h": 5, "fill": "#000000"}},
                        {{"kind": "text", "text": "w", "x": 10, "baselineY": 300,
                          "width": 5, "font": "400 16px Arial", "color": "#000000",
                          "docStart": {widget}, "docEnd": {widget_end},
                          "inlineSdtWidget": {{"pos": {widget}, "kind": "checkbox", "groupId": "g"}}}}
                    ],
                    "header": {{"kind": "header", "rId": "rId1", "y": 0, "height": 40, "primitives": [
                        {{"kind": "text", "text": "H", "x": 10, "baselineY": 20,
                          "width": 5, "font": "400 16px Arial", "color": "#000000",
                          "docStart": {header}, "docEnd": {header_end}}}
                    ]}}
                }}]}}"##,
                text_end = text_start + 5,
                rect_end = rect + 9,
                widget_end = widget + 1,
                header_end = header + 1,
            )
        };
        let stored = |handle: u32| SESSIONS.with(|s| s.borrow().get(handle).cloned().unwrap());
        for delta in [6_i64, -4] {
            let exact = open_display_list(&list(10, 9, 20, 3)).expect("opens");
            let present_only = open_display_list(&list(10, 9, 20, 3)).expect("opens");
            let exact_runs =
                serde_json::json!([[0, 1, 3, delta], [1, 1, 12, delta], [3, 1, 19, delta]]);
            let present_runs = serde_json::json!([[0, 4, 0x9f, delta]]);
            for (handle, runs) in [(exact, exact_runs), (present_only, present_runs)] {
                let update = serde_json::json!({"total": 1, "shift": [[0, 0, [runs]]]});
                update_display_list(handle, &update.to_string()).expect("updates");
            }
            assert_eq!(stored(exact), stored(present_only), "delta {delta}");
            let fresh = open_display_list(&list(10 + delta, 9 + delta, 20 + delta, 3)).unwrap();
            assert_eq!(stored(present_only), stored(fresh));
            for handle in [exact, present_only, fresh] {
                close_display_list(handle);
            }
        }

        let handle = open_display_list(&list(10, 9, 20, 3)).expect("opens");
        let flag_alone = serde_json::json!({"total": 1, "shift": [[0, 0, [[[0, 1, 0x80, 1]]]]]});
        assert!(update_display_list(handle, &flag_alone.to_string()).is_err());
        let handle = open_display_list(&list(10, 9, 20, 3)).expect("opens");
        let absent = serde_json::json!({"total": 1, "shift": [[0, 0, [[[2, 1, 1, 1]]]]]});
        assert!(
            update_display_list(handle, &absent.to_string()).is_err(),
            "an exact run still needs its fields"
        );
    }

    #[test]
    fn note_anchor_shifts_match_a_fresh_open() {
        drain();
        let text = |start: i64| {
            serde_json::json!({"kind": "text", "text": "x", "x": 10, "baselineY": 20,
                "width": 5, "font": "400 16px Arial", "color": "#000000",
                "docStart": start, "docEnd": start + 1})
        };
        let area = |kind: &str, anchors: &[Option<i64>]| {
            let notes: Vec<_> = anchors
                .iter()
                .enumerate()
                .map(|(index, anchor)| {
                    let mut note = serde_json::json!({"id": index + 1, "label": "1"});
                    if let Some(anchor) = anchor {
                        note["anchorDocStart"] = serde_json::json!(anchor);
                        note["anchorDocEnd"] = serde_json::json!(anchor + 1);
                    }
                    note
                })
                .collect();
            serde_json::json!({"kind": kind, "primitives": [text(2)],
                "noteIds": (1..=anchors.len()).collect::<Vec<_>>(), "notes": notes})
        };
        // A footnote page, an endnote page whose anchors span the document,
        // and a page with both.
        let list = |body: i64, foot: i64, ends: [Option<i64>; 2]| {
            let page = |index: usize, primitives, areas| {
                serde_json::json!({"pageIndex": index, "width": 816, "height": 1056,
                    "primitives": primitives, "noteAreas": areas})
            };
            serde_json::json!({"pages": [
                page(0, serde_json::json!([text(body)]), serde_json::json!([area("footnote", &[Some(foot)])])),
                page(1, serde_json::json!([]), serde_json::json!([area("endnote", &ends)])),
                page(2, serde_json::json!([text(body + 50)]), serde_json::json!([
                    area("footnote", &[Some(foot + 50)]),
                    area("endnote", &ends),
                ])),
            ]})
            .to_string()
        };
        let stored = |handle: u32| SESSIONS.with(|s| s.borrow().get(handle).cloned().unwrap());

        let handle = open_display_list(&list(10, 11, [Some(3), Some(40)])).expect("opens");
        let update = serde_json::json!({"total": 3, "shift": [
            [0, 0, [[[0, 1, 3, 5]], []], [[[0, 0, 16, 17]], [[0, 0, 18, 19]]]],
            [1, 1, [[], []], [[[0, 1, 45, 46]], [[0, 0, null, null]]]],
            [2, 2, [[[0, 1, 3, 5]], []], [[[0, 0, 66, 67], [1, 1, 45, 46]], [[1, 0, null, null]]]],
        ]});
        update_display_list(handle, &update.to_string()).expect("updates");
        let mut expected: serde_json::Value =
            serde_json::from_str(&list(15, 18, [None, Some(45)])).unwrap();
        expected["pages"][2]["noteAreas"][0]["notes"][0]["anchorDocStart"] = 66.into();
        expected["pages"][2]["noteAreas"][0]["notes"][0]["anchorDocEnd"] = 67.into();
        let fresh = open_display_list(&expected.to_string()).unwrap();
        assert_eq!(stored(handle), stored(fresh));
        close_display_list(fresh);
        close_display_list(handle);

        for shift in [
            serde_json::json!([0, 0, [[], []], [[[0, 0, 1, 2]]]]),
            serde_json::json!([0, 0, [[]], [[[1, 0, 1, 2]]]]),
            serde_json::json!([0, 0, [[]], [[[0, 1, 1, 2]]]]),
        ] {
            let handle = open_display_list(&list(10, 11, [Some(3), Some(40)])).expect("opens");
            let update =
                serde_json::json!({"total": 3, "reuse": [[1, 1], [2, 2]], "shift": [shift]});
            assert!(
                update_display_list(handle, &update.to_string()).is_err(),
                "{shift}"
            );
            assert!(hit_test_regions_by_handle(handle, 0, 1.0, 1.0).is_err());
        }
    }

    #[test]
    fn a_span_widened_to_the_position_limits_still_covers_its_page() {
        drain();
        let handle = open_display_list(SAMPLE).expect("opens");
        range_rects_by_handle(handle, 1, 2).unwrap();
        let update = serde_json::json!({
            "total": 1,
            "shift": [[0, 0, [[[0, 1, 3, i64::MIN]], [[0, 1, 3, -1]]]]],
        });
        update_display_list(handle, &update.to_string()).expect("updates");
        let shifted = range_rects_by_handle(handle, i64::MIN, i64::MIN + 1).unwrap();
        assert_ne!(shifted, "[]");
        close_display_list(handle);
    }

    #[test]
    fn shift_update_beyond_primitive_count_closes_the_handle() {
        drain();
        let handle = open_display_list(SAMPLE).expect("opens");
        let bad = serde_json::json!({
            "total": 1,
            "shift": [[0, 0, [[[0, 2, 3, 1]]]]],
        });
        assert!(update_display_list(handle, &bad.to_string()).is_err());
        assert!(
            hit_test_regions_by_handle(handle, 0, 120.0, 195.0).is_err(),
            "a failed shift update closes the handle so callers fall back"
        );
    }

    #[test]
    fn inconsistent_page_delta_update_closes_the_handle() {
        drain();
        let handle = open_display_list(SAMPLE).expect("opens");
        let bad = serde_json::json!({ "total": 2, "reuse": [[0, 0]], "replace": [] });
        assert!(update_display_list(handle, &bad.to_string()).is_err());
        assert!(
            hit_test_regions_by_handle(handle, 0, 120.0, 195.0).is_err(),
            "a failed update closes the handle so callers fall back"
        );
    }

    #[test]
    fn map_is_capped_and_evicts_the_oldest_handle() {
        drain();
        let mut handles = Vec::new();
        // open one past the cap; the store must never exceed MAX_SESSIONS
        for _ in 0..(MAX_SESSIONS + 1) {
            handles.push(open_display_list(SAMPLE).unwrap());
            assert!(open_count() <= MAX_SESSIONS, "store stays bounded");
        }
        assert_eq!(open_count(), MAX_SESSIONS);

        // the oldest handle was evicted (querying it now errors); the newest lives
        assert!(
            hit_test_regions_by_handle(handles[0], 0, 120.0, 195.0).is_err(),
            "oldest handle evicted at capacity"
        );
        assert!(
            hit_test_regions_by_handle(*handles.last().unwrap(), 0, 120.0, 195.0).is_ok(),
            "newest handle still resolves"
        );

        for h in handles {
            close_display_list(h);
        }
    }
}
