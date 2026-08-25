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
use std::collections::HashMap;
use std::collections::VecDeque;

use crate::display_list::{DisplayList, DisplayPage};
use crate::hit::{
    VerticalDirection, hit_test_regions, parse_region_scope, range_rects, range_rects_in_region,
    vertical_move,
};

/// Upper bound on concurrently-open handles. The facade keeps exactly one live,
/// so this only ever trips when a caller leaks handles; the oldest is then
/// evicted so the map stays bounded.
pub const MAX_SESSIONS: usize = 8;

/// The handle registry: parsed display lists keyed by handle id, plus the
/// insertion order used for oldest-first eviction.
struct Sessions {
    map: HashMap<u32, DisplayList>,
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
        self.map.insert(id, dl);
        self.order.push_back(id);
        id
    }

    fn get(&self, handle: u32) -> Option<&DisplayList> {
        self.map.get(&handle)
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
/// Every one of the `total` slots must be filled exactly once.
#[derive(serde::Deserialize)]
#[serde(rename_all = "camelCase")]
struct DisplayListUpdate {
    total: usize,
    #[serde(default)]
    contract_version: Option<u32>,
    #[serde(default)]
    reuse: Vec<(usize, usize)>,
    #[serde(default)]
    replace: Vec<(usize, DisplayPage)>,
    /// Retained pages whose doc positions moved: `[next_index, previous_index,
    /// run_lists]`, where each run list is applied in order and a run is
    /// `[start, count, mask, delta]` over the canonical primitive order (body,
    /// note separators+primitives per area, header, footer). Mask bits: 1
    /// docStart, 2 docEnd, 4 fragmentDocStart, 8 fragmentDocEnd, 16 inline
    /// widget pos — the owned frame-delta shift contract.
    #[serde(default)]
    shift: Vec<(usize, usize, Vec<Vec<(usize, usize, u8, i64)>>)>,
}

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
                let dl = sessions
                    .map
                    .get_mut(&handle)
                    .ok_or_else(|| format!("unknown display-list handle {handle}"))?;
                apply_display_list_update(dl, update)
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
    if mask & bit == 0 {
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
    if mask & SHIFT_INLINE_WIDGET != 0 {
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
fn shift_page_positions(
    page: &mut DisplayPage,
    runs: &[(usize, usize, u8, i64)],
) -> Result<(), String> {
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
    update: DisplayListUpdate,
) -> Result<(), String> {
    let slots = update
        .reuse
        .len()
        .saturating_add(update.replace.len())
        .saturating_add(update.shift.len());
    if slots != update.total {
        return Err("update slots do not cover the page total exactly".to_owned());
    }
    let mut previous: Vec<Option<DisplayPage>> = dl.pages.drain(..).map(Some).collect();
    let mut next: Vec<Option<DisplayPage>> = Vec::new();
    next.resize_with(update.total, || None);
    for (next_index, previous_index) in update.reuse {
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
    for (next_index, previous_index, run_lists) in update.shift {
        let mut page = previous
            .get_mut(previous_index)
            .and_then(Option::take)
            .ok_or_else(|| format!("shifted page {previous_index} is missing"))?;
        for runs in &run_lists {
            shift_page_positions(&mut page, runs)?;
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
    Ok(())
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
        let dl = sessions
            .get(handle)
            .ok_or_else(|| format!("unknown display-list handle {handle}"))?;
        serde_json::to_string(&range_rects(dl, from, to)).map_err(|e| format!("serialize: {e}"))
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
        let dl = sessions
            .get(handle)
            .ok_or_else(|| format!("unknown display-list handle {handle}"))?;
        let scope = parse_region_scope(region, part_id)?;
        serde_json::to_string(&range_rects_in_region(dl, scope, from, to))
            .map_err(|e| format!("serialize: {e}"))
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
