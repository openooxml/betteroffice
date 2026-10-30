#![allow(
    clippy::cloned_ref_to_slice_refs,
    clippy::collapsible_if,
    clippy::collapsible_match,
    clippy::doc_lazy_continuation,
    clippy::excessive_precision,
    clippy::field_reassign_with_default,
    clippy::if_same_then_else,
    clippy::inconsistent_digit_grouping,
    clippy::items_after_test_module,
    clippy::large_enum_variant,
    clippy::manual_contains,
    clippy::manual_is_multiple_of,
    clippy::manual_pattern_char_comparison,
    clippy::manual_repeat_n,
    clippy::manual_unwrap_or,
    clippy::map_clone,
    clippy::needless_lifetimes,
    clippy::needless_range_loop,
    clippy::neg_cmp_op_on_partial_ord,
    clippy::obfuscated_if_else,
    clippy::too_many_arguments,
    clippy::trim_split_whitespace,
    clippy::type_complexity,
    clippy::unnecessary_filter_map,
    clippy::unnecessary_lazy_evaluations,
    clippy::unnecessary_sort_by
)]

//! DOCX pagination and display-list construction.
//!
//! Two stages, in order:
//!
//! 1. **Pagination** — [`compute_layout`] takes `{ measured, options }` and
//!    walks the measured blocks into a [`types::Layout`]: pages of positioned
//!    fragments, with spacing collapse, page and column splits, section
//!    geometry, floats and per-page footnote reservations resolved. The
//!    walk lives in [`place`], the cursor in [`page_flow`], and the per-feature
//!    decisions in [`hooks`].
//! 2. **Display compilation** — the `build_display_list*` family turns a
//!    finished `Layout` plus the same measured input into a
//!    [`display_list::DisplayList`]: renderer-agnostic paint primitives per
//!    page, each carrying the document positions and block identity that
//!    interaction needs. Compilation only ever reads a `Layout`; the paginator
//!    never builds primitives itself.
//!
//! [`hit`] answers point and range queries over a finished display list, and
//! [`session`] does the same by handle so an interactive path parses the list
//! once instead of per event.
//!
//! An input engaging a feature the engine does not cover raises
//! [`LayoutError::Unsupported`], which crosses the JSON boundary as the string
//! `"UNSUPPORTED"` so the host can fall back rather than paint wrong geometry.
//!
//! This crate is the workspace's only `wasm_bindgen` site, so the `ooxml-text`
//! measurement surface is re-exported here as well and the wasm-visible font
//! registry lives in this file. Everything below the wrappers is pure and
//! native-testable.
//!
//! Font ids index one measurement font store. A host running several
//! documents in one module gives each a [`MeasureFonts`] and holds
//! [`MeasureFonts::enter`] around every call that registers fonts, measures,
//! lays out, builds a display list or reads glyph outlines. Callers that never
//! create one use the module's own store. Once any [`MeasureFonts`] is alive,
//! using fonts with none entered returns an error from module-level wasm
//! exports and panics for internal callers that miss `enter`.

mod anchor;
pub mod canonical;
mod extent_key;
mod fmath;
pub mod hooks;
mod line_fit;
pub mod page_flow;
pub mod paragraph_spacing;
pub mod place;
pub mod placement;
pub mod prescan;
pub mod progress;
pub mod regions;
pub mod resolve_lines;
mod transcode;
pub mod types;

pub mod break_policy;
pub mod cell_layout;
pub mod column_balancing;
pub mod display_list;
pub mod floating_objects;
pub mod footnotes;
pub mod header_footer;
pub mod hf_bands;
pub mod hit;
pub mod keep_together;
pub mod measure_blocks;
pub mod section_breaks;
pub mod sequence_fields;
pub mod session;
pub mod table_grid;
pub mod table_row_break;

mod leading_inline_box;
mod typed_measure;
#[cfg(any(test, feature = "test-support"))]
pub mod work_counters;

use std::collections::BTreeMap;
use std::hash::{Hash, Hasher};
use wasm_bindgen::prelude::*;

#[cfg(target_arch = "wasm32")]
#[wasm_bindgen]
extern "C" {
    #[wasm_bindgen(js_namespace = console, js_name = error)]
    fn console_error(message: &str);
    #[wasm_bindgen(js_namespace = console, js_name = warn)]
    fn console_warn(message: &str);
}

/// Reports content the engine left out rather than failing the document on:
/// a console warning in wasm, stderr natively.
pub(crate) fn warn_degraded(message: &str) {
    #[cfg(target_arch = "wasm32")]
    console_warn(&format!("[docx-wasm] {message}"));
    #[cfg(not(target_arch = "wasm32"))]
    eprintln!("[docx-layout] {message}");
}

/// wasm panics abort, so a trap reaches JS as a bare `RuntimeError:
/// unreachable executed`. Log the panic first so it stays diagnosable. Runs on
/// module init for every wasm core that links this crate (layout and edit).
#[cfg(target_arch = "wasm32")]
#[cfg_attr(not(feature = "wasm-query"), wasm_bindgen(start))]
pub fn install_panic_hook() {
    static INSTALLED: std::sync::Once = std::sync::Once::new();
    INSTALLED.call_once(|| {
        std::panic::set_hook(Box::new(|info| {
            console_error(&format!("[docx-wasm] panic: {info}"));
        }));
    });
}

/// Why the engine refused an input.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum LayoutError {
    /// The input engages a feature the engine doesn't cover yet — the
    /// caller must degrade gracefully.
    Unsupported(String),
    /// The input violates the layout contract.
    Invalid(String),
}

impl std::fmt::Display for LayoutError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            LayoutError::Unsupported(reason) => write!(f, "unsupported: {reason}"),
            LayoutError::Invalid(reason) => write!(f, "invalid: {reason}"),
        }
    }
}

impl std::error::Error for LayoutError {}

/// Run pagination from a typed native input.
pub fn compute_layout_input(input: &mut types::Input) -> Result<types::Layout, LayoutError> {
    place::layout_document(input)
}

/// Parse the `{ measured, options }` envelope and run the placement walk.
pub fn compute_layout(input: &str) -> Result<types::Layout, LayoutError> {
    let mut parsed: types::Input =
        serde_json::from_str(input).map_err(|e| LayoutError::Invalid(format!("parse: {e}")))?;
    compute_layout_input(&mut parsed)
}

/// Build a body display list from typed pagination state.
pub fn build_display_list(
    pagination: &types::Input,
    layout: &types::Layout,
) -> Result<display_list::DisplayList, String> {
    build_display_list_value_from_resident(pagination, layout, "{}")
}

/// `{ measured, options }` JSON in, `Layout` JSON out. An `Err` of
/// `"UNSUPPORTED"` means the input needs a host fallback; any other `Err`
/// describes a malformed envelope.
pub fn layout_to_json(input: &str) -> Result<String, String> {
    let layout = compute_layout(input).map_err(|e| match e {
        LayoutError::Unsupported(_) => "UNSUPPORTED".to_string(),
        LayoutError::Invalid(reason) => {
            if reason.starts_with("parse: ") {
                reason
            } else {
                "UNSUPPORTED".to_string()
            }
        }
    })?;
    serde_json::to_string(&layout).map_err(|e| format!("serialize: {e}"))
}

/// `{ measured, options }` in, canonical golden `Layout` string out (see
/// [`canonical::serialize_layout`]). The native golden harness compares this
/// byte-for-byte against `__golden__/golden/<scenario>.json`.
pub fn layout_to_canonical_json(input: &str) -> Result<String, LayoutError> {
    let layout = compute_layout(input)?;
    Ok(canonical::serialize_layout(&layout))
}

/// wasm wrapper over [`layout_to_json`].
#[cfg_attr(not(feature = "wasm-query"), wasm_bindgen)]
pub fn layout_document_json(input: &str) -> Result<String, JsValue> {
    with_module_fonts(|| layout_to_json(input)).map_err(|e| JsValue::from_str(&e))
}

/// Builds a display list from a `{ measured, options, layout }` JSON envelope,
/// threading the measurement fonts in use. With fonts registered, runs
/// whose font chains resolve are shaped into glyph runs; with none registered
/// every run stays a text primitive.
pub fn build_display_list_value(input: &str) -> Result<display_list::DisplayList, String> {
    with_measure_fonts(|store| {
        display_list::build_display_list_value_with_fonts(input, &store.borrow())
    })
}

/// Resident counterpart to [`build_display_list_value`]. Large measured and
/// layout values stay inside the editing wasm; `extras` carries only the
/// display-specific fields not owned by pagination.
pub fn build_display_list_value_from_resident(
    pagination: &types::Input,
    layout: &types::Layout,
    extras: &str,
) -> Result<display_list::DisplayList, String> {
    with_measure_fonts(|store| {
        display_list::build_display_list_value_from_resident_with_fonts(
            pagination,
            layout,
            extras,
            &store.borrow(),
        )
    })
}

pub fn build_display_list_value_from_resident_observed(
    pagination: &types::Input,
    layout: &types::Layout,
    extras: &str,
    observe_phase: &mut impl FnMut(),
) -> Result<display_list::DisplayList, String> {
    with_measure_fonts(|store| {
        display_list::build_display_list_value_from_resident_with_fonts_observed(
            pagination,
            layout,
            extras,
            &store.borrow(),
            observe_phase,
        )
    })
}

/// Build a display list while retaining its parsed compatibility input for
/// page-scoped refreshes on subsequent resident edits.
pub fn build_resident_display_list_observed(
    pagination: &types::Input,
    layout: &types::Layout,
    extras: &str,
    observe_phase: &mut impl FnMut(),
) -> Result<
    (
        display_list::ResidentDisplayInput,
        display_list::DisplayList,
    ),
    String,
> {
    with_measure_fonts(|store| {
        display_list::build_resident_display_list_with_fonts_observed(
            pagination,
            layout,
            extras,
            &store.borrow(),
            observe_phase,
        )
    })
}

/// [`build_resident_display_list_observed`] building only the pages `build`
/// selects; the rest are unbuilt placeholders.
pub fn build_resident_display_list_partial_observed(
    pagination: &types::Input,
    layout: &types::Layout,
    extras: &str,
    build: &dyn Fn(usize) -> bool,
    observe_phase: &mut impl FnMut(),
) -> Result<
    (
        display_list::ResidentDisplayInput,
        display_list::DisplayList,
    ),
    String,
> {
    with_measure_fonts(|store| {
        display_list::build_resident_display_list_partial_with_fonts_observed(
            pagination,
            layout,
            extras,
            &store.borrow(),
            build,
            observe_phase,
        )
    })
}

pub use display_list::release_resident_display_pages;

/// Build unbuilt pages of an engine-owned display list; returns the pages built.
pub fn build_resident_display_pages(
    pagination: &types::Input,
    layout: &types::Layout,
    resident: &mut display_list::ResidentDisplayInput,
    list: &mut display_list::DisplayList,
    pages: &[usize],
) -> Result<Vec<usize>, String> {
    with_measure_fonts(|store| {
        display_list::build_resident_display_pages_with_fonts(
            pagination,
            layout,
            &store.borrow(),
            resident,
            list,
            pages,
        )
    })
}

#[doc(hidden)]
pub fn build_resident_display_pages_with_fingerprints(
    pagination: &types::Input,
    layout: &types::Layout,
    resident: &mut display_list::ResidentDisplayInput,
    list: &mut display_list::DisplayList,
    pages: &[usize],
    fingerprints: &[u128],
) -> Result<Vec<usize>, String> {
    with_measure_fonts(|store| {
        display_list::build_resident_display_pages_with_fonts_and_fingerprints(
            pagination,
            layout,
            &store.borrow(),
            resident,
            list,
            pages,
            Some(fingerprints),
        )
    })
}

/// Incremental resident display build. Only the pagination-dirtied page range
/// is recompiled; converged suffix pages are retained with body positions
/// patched from stable block-id deltas.
pub fn build_display_list_value_from_resident_incremental(
    pagination: &types::Input,
    layout: &types::Layout,
    extras: &str,
    previous: &display_list::DisplayList,
    rebuilt_page_start: usize,
    rebuilt_page_end: usize,
    position_deltas: &std::collections::HashMap<String, i64>,
) -> Result<display_list::DisplayList, String> {
    with_measure_fonts(|store| {
        display_list::build_display_list_value_from_resident_incremental_with_fonts(
            pagination,
            layout,
            extras,
            &store.borrow(),
            previous,
            rebuilt_page_start,
            rebuilt_page_end,
            position_deltas,
        )
    })
}

/// Update an engine-owned display arena without cloning clean pages.
pub fn update_display_list_value_from_resident_incremental(
    pagination: &types::Input,
    layout: &types::Layout,
    extras: &str,
    previous: &mut display_list::DisplayList,
    rebuilt_page_start: usize,
    rebuilt_page_end: usize,
    position_deltas: &std::collections::HashMap<String, i64>,
) -> Result<bool, String> {
    with_measure_fonts(|store| {
        display_list::update_display_list_value_from_resident_incremental_with_fonts(
            pagination,
            layout,
            extras,
            &store.borrow(),
            previous,
            rebuilt_page_start,
            rebuilt_page_end,
            position_deltas,
        )
    })
}

pub fn update_display_list_value_from_resident_incremental_observed(
    pagination: &types::Input,
    layout: &types::Layout,
    extras: &str,
    previous: &mut display_list::DisplayList,
    rebuilt_page_start: usize,
    rebuilt_page_end: usize,
    position_deltas: &std::collections::HashMap<String, i64>,
    observe_phase: &mut impl FnMut(),
) -> Result<bool, String> {
    with_measure_fonts(|store| {
        display_list::update_display_list_value_from_resident_incremental_with_fonts_observed(
            pagination,
            layout,
            extras,
            &store.borrow(),
            previous,
            rebuilt_page_start,
            rebuilt_page_end,
            position_deltas,
            observe_phase,
        )
    })
}

/// Page-scoped update using the engine's retained parsed display input.
#[allow(clippy::too_many_arguments)]
pub fn update_resident_display_list_incremental_observed(
    pagination: &types::Input,
    layout: &types::Layout,
    resident: &mut display_list::ResidentDisplayInput,
    previous: &mut display_list::DisplayList,
    rebuilt_page_start: usize,
    rebuilt_page_end: usize,
    extra_pages: &[usize],
    position_deltas: &std::collections::HashMap<String, i64>,
    observe_phase: &mut impl FnMut(),
) -> Result<bool, String> {
    with_measure_fonts(|store| {
        display_list::update_resident_display_list_incremental_with_fonts_observed(
            pagination,
            layout,
            &store.borrow(),
            resident,
            previous,
            rebuilt_page_start,
            rebuilt_page_end,
            extra_pages,
            position_deltas,
            observe_phase,
        )
    })
}

/// [`update_resident_display_list_incremental_observed`] that builds only the
/// rebuilt pages `build` selects, replacing the others with unbuilt placeholders.
#[allow(clippy::too_many_arguments)]
pub fn update_resident_display_list_incremental_partial_observed(
    pagination: &types::Input,
    layout: &types::Layout,
    resident: &mut display_list::ResidentDisplayInput,
    previous: &mut display_list::DisplayList,
    rebuilt_page_start: usize,
    rebuilt_page_end: usize,
    extra_pages: &[usize],
    position_deltas: &std::collections::HashMap<String, i64>,
    build: &dyn Fn(usize) -> bool,
    observe_phase: &mut impl FnMut(),
) -> Result<bool, String> {
    with_measure_fonts(|store| {
        display_list::update_resident_display_list_incremental_partial_with_fonts_observed(
            pagination,
            layout,
            &store.borrow(),
            resident,
            previous,
            rebuilt_page_start,
            rebuilt_page_end,
            extra_pages,
            position_deltas,
            build,
            observe_phase,
        )
    })
}

/// Page-scoped display update reporting retained suffix position shifts.
#[allow(clippy::too_many_arguments)]
pub fn update_resident_display_list_incremental_partial_shifts_observed(
    pagination: &types::Input,
    layout: &types::Layout,
    resident: &mut display_list::ResidentDisplayInput,
    previous: &mut display_list::DisplayList,
    rebuilt_page_start: usize,
    rebuilt_page_end: usize,
    extra_pages: &[usize],
    position_deltas: &std::collections::HashMap<String, i64>,
    build: &dyn Fn(usize) -> bool,
    observe_phase: &mut impl FnMut(),
) -> Result<Option<display_list::IncrementalDisplayShifts>, String> {
    with_measure_fonts(|store| {
        display_list::update_resident_display_list_incremental_partial_with_fonts_shifts(
            pagination,
            layout,
            &store.borrow(),
            resident,
            previous,
            rebuilt_page_start,
            rebuilt_page_end,
            extra_pages,
            position_deltas,
            build,
            observe_phase,
        )
    })
}

#[doc(hidden)]
#[allow(clippy::too_many_arguments)]
pub fn update_resident_display_list_incremental_partial_cached_shifts_observed(
    pagination: &types::Input,
    layout: &types::Layout,
    resident: &mut display_list::ResidentDisplayInput,
    previous: &mut display_list::DisplayList,
    rebuilt_page_start: usize,
    rebuilt_page_end: usize,
    extra_pages: &[usize],
    position_deltas: &std::collections::HashMap<String, i64>,
    build: &dyn Fn(usize) -> bool,
    observe_phase: &mut impl FnMut(),
    fingerprints: &[u128],
) -> Result<Option<display_list::IncrementalDisplayShifts>, String> {
    with_measure_fonts(|store| {
        display_list::update_resident_display_list_incremental_partial_with_fonts_cached(
            pagination,
            layout,
            &store.borrow(),
            resident,
            previous,
            rebuilt_page_start,
            rebuilt_page_end,
            extra_pages,
            position_deltas,
            build,
            observe_phase,
            Some(fingerprints),
        )
    })
}

/// wasm compatibility wrapper. Resident engine users call
/// [`build_display_list_value`] and keep the typed result.
#[cfg_attr(not(feature = "wasm-query"), wasm_bindgen)]
pub fn build_display_list_json(input: &str) -> Result<String, JsValue> {
    with_module_fonts(|| {
        build_display_list_value(input).and_then(|display_list| {
            serde_json::to_string(&display_list).map_err(|e| format!("serialize: {e}"))
        })
    })
    .map_err(|e| JsValue::from_str(&e))
}

/// wasm wrapper over [`hit::hit_test_json`]: display-list JSON + page-local
/// point in, document position (or `null`) as JSON out.
#[cfg_attr(not(feature = "wasm-query"), wasm_bindgen)]
pub fn hit_test_json(
    display_list: &str,
    page_index: u32,
    x: f64,
    y: f64,
) -> Result<String, JsValue> {
    hit::hit_test_json(display_list, page_index as usize, x, y).map_err(|e| JsValue::from_str(&e))
}

#[wasm_bindgen]
pub fn vertical_move_json(
    display_list: &str,
    position: f64,
    direction: &str,
    goal_x: f64,
) -> Result<String, JsValue> {
    hit::vertical_move_json(display_list, position as i64, direction, goal_x)
        .map_err(|e| JsValue::from_str(&e))
}

/// wasm wrapper over [`hit::range_rects_json`]: display-list JSON + document range
/// in, JSON array of page-local rects out.
#[wasm_bindgen]
pub fn range_rects_json(display_list: &str, from: f64, to: f64) -> Result<String, JsValue> {
    hit::range_rects_json(display_list, from as i64, to as i64).map_err(|e| JsValue::from_str(&e))
}

/// @internal
#[wasm_bindgen]
pub fn range_rects_on_pages_json(
    display_list: &str,
    from: f64,
    to: f64,
    first_page: f64,
    last_page: f64,
) -> Result<String, JsValue> {
    let Some((first_page, last_page)) = hit::page_window(first_page, last_page) else {
        return Ok("[]".to_string());
    };
    hit::range_rects_on_pages_json(display_list, from as i64, to as i64, first_page, last_page)
        .map_err(|e| JsValue::from_str(&e))
}

/// wasm wrapper over [`hit::range_rects_region_json`]: region-aware range rects.
/// `region` is `"body" | "header" | "footer" | "footnote" | "endnote"`;
/// `part_id` scopes a header/footer to one HF part (empty for body / match-any)
/// and names the note id for a note region. The `from`/`to` refer to that
/// region's doc. The plain `range_rects_json` export stays body-only.
#[wasm_bindgen]
pub fn range_rects_region_json(
    display_list: &str,
    region: &str,
    part_id: &str,
    from: f64,
    to: f64,
) -> Result<String, JsValue> {
    hit::range_rects_region_json(display_list, region, part_id, from as i64, to as i64)
        .map_err(|e| JsValue::from_str(&e))
}

/// wasm wrapper over [`hit::hit_test_regions_json`]: region-aware hit test —
/// `{"region":"body"|"header"|"footer"|"footnote"|"endnote","rId"?,"noteId"?,`
/// `"pos":n|null,"target":"text"|"image"|"none"}` (or `"null"` for an
/// out-of-range page). The plain `hit_test_json` export stays body-only.
#[wasm_bindgen]
pub fn hit_test_regions_json(
    display_list: &str,
    page_index: u32,
    x: f64,
    y: f64,
) -> Result<String, JsValue> {
    hit::hit_test_regions_json(display_list, page_index as usize, x, y)
        .map_err(|e| JsValue::from_str(&e))
}

// ---------------------------------------------------------------------------
// session-handle query surface (crate::session)
//
// The JSON-arg exports above re-parse the whole display list every call, which
// dominates the per-event cost on the interactive canvas paths (click, drag).
// These exports parse the display list ONCE ([`open_display_list`]) and answer
// many queries by handle with zero re-serialization, reusing the same hit/range
// logic so results are byte-identical. The JSON-arg exports stay for
// back-compat and as the graceful fallback when a handle is unavailable.
// ---------------------------------------------------------------------------

/// wasm wrapper over [`session::open_display_list`]: parse a display list once
/// and return a handle the by-handle query exports reuse (no per-query
/// re-parse). The caller frees it with [`close_display_list`]. `Err` on
/// malformed JSON — the caller then stays on the JSON-arg path.
#[wasm_bindgen]
pub fn open_display_list(display_list: &str) -> Result<u32, JsValue> {
    session::open_display_list(display_list).map_err(|e| JsValue::from_str(&e))
}

/// wasm wrapper over [`session::close_display_list`]: drop a handle so its
/// parsed display list is freed. Idempotent.
#[wasm_bindgen]
pub fn close_display_list(handle: u32) {
    session::close_display_list(handle);
}

/// wasm wrapper over [`session::update_display_list`]: apply a page-delta
/// update to a stored display list so an incremental rebuild re-parses only
/// its changed pages. `Err` closes the handle first, so the caller's fallback
/// (a fresh [`open_display_list`]) can never race a half-updated list.
#[wasm_bindgen]
pub fn update_display_list(handle: u32, update: &str) -> Result<(), JsValue> {
    session::update_display_list(handle, update).map_err(|e| JsValue::from_str(&e))
}

/// wasm wrapper over [`session::hit_test_regions_by_handle`]: region-aware hit
/// test against a stored display list. `Err` on an unknown/closed handle so the
/// caller can fall back to [`hit_test_regions_json`].
#[wasm_bindgen]
pub fn hit_test_regions_by_handle(
    handle: u32,
    page_index: u32,
    x: f64,
    y: f64,
) -> Result<String, JsValue> {
    session::hit_test_regions_by_handle(handle, page_index as usize, x, y)
        .map_err(|e| JsValue::from_str(&e))
}

#[wasm_bindgen]
pub fn vertical_move_by_handle(
    handle: u32,
    position: f64,
    direction: &str,
    goal_x: f64,
) -> Result<String, JsValue> {
    session::vertical_move_by_handle(handle, position as i64, direction, goal_x)
        .map_err(|e| JsValue::from_str(&e))
}

/// wasm wrapper over [`session::range_rects_by_handle`]: range rects against a
/// stored display list. `Err` on an unknown/closed handle so the caller can
/// fall back to [`range_rects_json`].
#[wasm_bindgen]
pub fn range_rects_by_handle(handle: u32, from: f64, to: f64) -> Result<String, JsValue> {
    session::range_rects_by_handle(handle, from as i64, to as i64)
        .map_err(|e| JsValue::from_str(&e))
}

/// @internal
#[wasm_bindgen]
pub fn range_rects_on_pages_by_handle(
    handle: u32,
    from: f64,
    to: f64,
    first_page: f64,
    last_page: f64,
) -> Result<String, JsValue> {
    let Some((first_page, last_page)) = hit::page_window(first_page, last_page) else {
        return Ok("[]".to_string());
    };
    session::range_rects_on_pages_by_handle(handle, from as i64, to as i64, first_page, last_page)
        .map_err(|e| JsValue::from_str(&e))
}

/// wasm wrapper over [`session::range_rects_region_by_handle`]: region-aware
/// range rects against a stored display list. `region` is
/// `"body" | "header" | "footer" | "footnote" | "endnote"`; `part_id` scopes
/// header/footer to one HF part and names the note id for a note region.
/// `Err` on an unknown/closed handle so the caller can fall back to
/// [`range_rects_region_json`].
#[wasm_bindgen]
pub fn range_rects_region_by_handle(
    handle: u32,
    region: &str,
    part_id: &str,
    from: f64,
    to: f64,
) -> Result<String, JsValue> {
    session::range_rects_region_by_handle(handle, region, part_id, from as i64, to as i64)
        .map_err(|e| JsValue::from_str(&e))
}

type SharedFontStore = std::rc::Rc<std::cell::RefCell<ooxml_text::FontStore>>;

thread_local! {
    /// Measurement fonts for free exports outside an entered session.
    static MODULE_FONTS: SharedFontStore = SharedFontStore::default();
    /// The entered [`MeasureFonts`], if any.
    static ENTERED_FONTS: std::cell::RefCell<Option<SharedFontStore>> =
        const { std::cell::RefCell::new(None) };
    /// How many [`MeasureFonts`] are alive.
    static SESSION_FONTS: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
}

fn with_module_fonts<T>(run: impl FnOnce() -> T) -> T {
    if ENTERED_FONTS.with(|entered| entered.borrow().is_some()) {
        return run();
    }
    let store = MODULE_FONTS.with(std::rc::Rc::clone);
    let previous = ENTERED_FONTS.with(|entered| entered.replace(Some(store)));
    let _scope = MeasureFontsScope(previous);
    run()
}

/// Runs `read` against the entered measurement font store, else the module's
/// own. Panics when a [`MeasureFonts`] is alive but none is entered: a caller
/// forgot [`MeasureFonts::enter`] and would otherwise read the wrong store.
fn with_measure_fonts<T>(read: impl FnOnce(&std::cell::RefCell<ooxml_text::FontStore>) -> T) -> T {
    let store = ENTERED_FONTS
        .with(|entered| entered.borrow().clone())
        .unwrap_or_else(|| {
            assert_eq!(
                SESSION_FONTS.with(std::cell::Cell::get),
                0,
                "measurement fonts used outside MeasureFonts::enter while a session owns a store"
            );
            MODULE_FONTS.with(std::rc::Rc::clone)
        });
    read(&store)
}

/// A measurement font store of one session's own. Its ids start at zero and
/// stay valid whatever other stores register or clear. Internal font access
/// requires [`MeasureFonts::enter`] while any session store is alive.
pub struct MeasureFonts(SharedFontStore);

impl Default for MeasureFonts {
    fn default() -> Self {
        SESSION_FONTS.with(|count| count.set(count.get() + 1));
        Self(SharedFontStore::default())
    }
}

impl Drop for MeasureFonts {
    fn drop(&mut self) {
        let _ = SESSION_FONTS.try_with(|count| {
            count.set(count.get() - 1);
            if count.get() == 0 {
                measure_blocks::clear_extent_cache();
            }
        });
    }
}

impl MeasureFonts {
    /// Makes this the store that registration, measurement, display building
    /// and glyph outlines use until the returned scope drops, which puts the
    /// replaced store back.
    pub fn enter(&self) -> MeasureFontsScope {
        let previous =
            ENTERED_FONTS.with(|entered| entered.replace(Some(std::rc::Rc::clone(&self.0))));
        MeasureFontsScope(previous)
    }
}

/// The span a [`MeasureFonts`] is in use; see [`MeasureFonts::enter`].
#[must_use]
pub struct MeasureFontsScope(Option<SharedFontStore>);

impl Drop for MeasureFontsScope {
    fn drop(&mut self) {
        let previous = self.0.take();
        let _ = ENTERED_FONTS.try_with(|entered| entered.replace(previous));
    }
}

/// Register a font for measurement from raw sfnt bytes; returns the font id
/// that `measure_paragraph_json` inputs reference in their `fontChains`.
/// Malformed bytes (attacker-controlled embedded fonts) are rejected as an
/// error at this boundary, mirroring `FontStore::register`.
#[cfg_attr(not(feature = "wasm-query"), wasm_bindgen)]
pub fn register_measure_font(bytes: &[u8]) -> Result<u32, JsValue> {
    with_module_fonts(|| register_measure_font_bytes(bytes)).map_err(|e| JsValue::from_str(&e))
}

/// Register a measurement view of `base` carrying the vertical metrics and
/// advance pitch Word measures `requested_family` with, and return its id;
/// returns `base` unchanged for a family with no known metrics. Hosts call
/// this for a face they substituted, and put the result at the head of that
/// family's chain. Pagination, the display list and glyph outlines all read
/// this one store, so a widened view measures and paints at one pitch.
#[cfg_attr(not(feature = "wasm-query"), wasm_bindgen)]
pub fn register_substitute_measure_font(base: u32, requested_family: &str) -> Result<u32, JsValue> {
    with_module_fonts(|| register_substitute_measure_font_id(base, requested_family))
        .map_err(|e| JsValue::from_str(&e))
}

/// Registers a substitute view in the current measurement font store.
pub fn register_substitute_measure_font_id(
    base: u32,
    requested_family: &str,
) -> Result<u32, String> {
    with_measure_fonts(|store| {
        store
            .borrow_mut()
            .register_named_substitute(ooxml_text::FontId::from_u32(base), requested_family)
            .map(|id| id.to_u32())
            .map_err(|e| e.to_string())
    })
}

/// Drop every registered measurement font (ids restart at 0). Callers must
/// re-register before the next `measure_paragraph_json`.
pub fn clear_measure_fonts() {
    with_measure_fonts(|store| {
        *store.borrow_mut() = ooxml_text::FontStore::new();
    });
    measure_blocks::clear_extent_cache();
}

/// Drop every registered measurement font (ids restart at 0). Callers must
/// re-register before the next `measure_paragraph_json`.
#[cfg_attr(
    not(feature = "wasm-query"),
    wasm_bindgen(js_name = clear_measure_fonts)
)]
pub fn clear_measure_fonts_wasm() -> Result<(), JsValue> {
    with_module_fonts(clear_measure_fonts);
    Ok(())
}

/// Unique id of the current measurement font store, for caches keyed by store
/// contents. Changes whenever [`clear_measure_fonts`] installs a new store.
pub fn measure_store_id() -> u64 {
    with_measure_fonts(|store| store.borrow().id())
}

/// Which fonts measurement reads: the current store and how many fonts it holds. A cleared store
/// is a new one and fonts are only ever added, so measurements taken under equal generations
/// used the same fonts.
pub fn measure_fonts_generation() -> (u64, usize) {
    with_measure_fonts(|store| {
        let store = store.borrow();
        (store.id(), store.font_count())
    })
}

pub fn measure_font_cache_identity(chains: &BTreeMap<String, Vec<u32>>) -> (u64, u64) {
    let (store, fonts) = measure_fonts_generation();
    let mut missing = Vec::new();
    let mut all_chains_empty = true;
    for &id in chains.values().flatten() {
        all_chains_empty = false;
        if id as usize >= fonts {
            missing.push(id);
        }
    }
    missing.sort_unstable();
    missing.dedup();
    let mut hash = std::hash::DefaultHasher::new();
    (fonts == 0).hash(&mut hash);
    all_chains_empty.hash(&mut hash);
    missing.hash(&mut hash);
    (store, hash.finish())
}

/// Runs `run` against an empty measurement font store of its own, then puts the
/// store in use back unchanged: fonts `run` registers never reach another
/// session, and caches keyed by [`measure_store_id`] never mix the two.
pub fn with_private_measure_fonts<T>(run: impl FnOnce() -> T) -> T {
    let private = MeasureFonts::default();
    let _scope = private.enter();
    run()
}

/// Registers raw sfnt bytes in the current measurement font store.
pub fn register_measure_font_bytes(bytes: &[u8]) -> Result<u32, String> {
    with_measure_fonts(|store| {
        store
            .borrow_mut()
            .register_slice(bytes)
            .map(|id| id.to_u32())
            .map_err(|error| error.to_string())
    })
}

/// Reads what measurement font `id` measures with: the bytes it shapes and a descriptor of the
/// metrics, fallback advance scale and widths layered over them. `None` for an unknown id.
pub fn with_measure_face<T>(id: u32, read: impl FnOnce(&[u8], &str) -> T) -> Option<T> {
    with_measure_fonts(|store| {
        let store = store.borrow();
        let (bytes, metrics, scale) = store.measured_face(ooxml_text::FontId::from_u32(id)).ok()?;
        let widths = store.family_widths(ooxml_text::FontId::from_u32(id)).ok()?;
        let descriptor = match widths {
            Some(widths) => format!("{metrics:?}/{scale}/{widths:?}"),
            None => format!("{metrics:?}/{scale}"),
        };
        Some(read(bytes, &descriptor))
    })
}

/// Measures a paragraph: measurement input JSON in, `ParagraphExtent` JSON
/// out. An `Err` whose message starts with `"UNSUPPORTED"` means the caller
/// must fall back to browser measurement for that block.
#[cfg_attr(not(feature = "wasm-query"), wasm_bindgen)]
pub fn measure_paragraph_json(input: &str) -> Result<String, JsValue> {
    with_module_fonts(|| measure_paragraph_json_resident(input)).map_err(|e| JsValue::from_str(&e))
}

/// Native/resident-engine form of [`measure_paragraph_json`]. Keeping the
/// string error avoids constructing a wasm `JsValue` when the editing engine
/// remeasures a dirty paragraph inside a larger operation.
pub fn measure_paragraph_json_resident(input: &str) -> Result<String, String> {
    let envelope: serde_json::Value =
        serde_json::from_str(input).map_err(|error| format!("invalid: parse: {error}"))?;
    let mut input: ooxml_text::measure::MeasureInput =
        serde_json::from_str(input).map_err(|error| format!("invalid: parse: {error}"))?;
    input.compat.space_shrink_limit =
        Some(docx_space_shrink_limit(&envelope["compat"], &input.block));
    let extent = with_measure_fonts(|store| {
        ooxml_text::measure::measure_paragraph(&store.borrow(), &input)
            .map_err(|error| error.to_string())
    })?;
    serde_json::to_string(&extent).map_err(|error| format!("invalid: serialize: {error}"))
}

pub(crate) fn docx_space_shrink_limit(
    compat: &serde_json::Value,
    block: &ooxml_text::measure::BlockIn,
) -> f32 {
    let attrs = block.attrs.as_ref();
    let runs = &block.runs;
    let has_text = runs.iter().any(|run| {
        run.kind != "lineBreak"
            && run
                .text
                .as_deref()
                .or(run.fallback.as_deref())
                .is_some_and(|text| !text.is_empty())
    });
    let measured_latin = !attrs.is_some_and(|attrs| {
        attrs.bidi
            || attrs
                .justification_mode
                .as_deref()
                .is_some_and(|mode| mode != "both")
    }) && runs.iter().all(|run| {
        let text = run
            .text
            .as_deref()
            .or(run.fallback.as_deref())
            .unwrap_or("");
        if has_text && (run.kind == "lineBreak" || text.is_empty()) {
            return true;
        }
        !run.rtl
            && run.kind != "lineBreak"
            && text.chars().all(|ch| {
                ch <= '\u{024f}'
                    || ch.is_whitespace()
                    || ('\u{0300}'..='\u{036f}').contains(&ch)
                    || ('\u{1ab0}'..='\u{1aff}').contains(&ch)
                    || ('\u{1dc0}'..='\u{1dff}').contains(&ch)
                    || ('\u{2000}'..='\u{206f}').contains(&ch)
                    || ('\u{20d0}'..='\u{20ff}').contains(&ch)
                    || ('\u{fe20}'..='\u{fe2f}').contains(&ch)
            })
    });
    if !measured_latin || compat["compatibilityMode"].as_u64().unwrap_or(12) >= 15 {
        0.25
    } else {
        0.0
    }
}

/// Typed form of [`measure_paragraph_json_resident`] against the same font
/// store, skipping both JSON round trips.
pub(crate) fn measure_paragraph_typed_resident(
    request: &ooxml_text::MeasureRequest<'_>,
) -> Result<ooxml_text::ParagraphExtentOut, ooxml_text::MeasureError> {
    with_measure_fonts(|store| ooxml_text::measure_paragraph_typed(&store.borrow(), request))
}

/// [`ooxml_text::min_content_width_typed`] against the resident font store.
pub(crate) fn min_content_width_typed_resident(
    request: &ooxml_text::MeasureRequest<'_>,
) -> Result<f32, ooxml_text::MeasureError> {
    with_measure_fonts(|store| ooxml_text::min_content_width_typed(&store.borrow(), request))
}

/// wasm wrapper over [`ooxml_text::FontStore::outline_glyph_json`]: the outline
/// of a registered font's glyph, in font design units, as JSON:
/// `{"upem":2048,"cmds":[{"t":"M","x":..,"y":..},{"t":"L","x":..,"y":..},
/// {"t":"Q","cx":..,"cy":..,"x":..,"y":..},
/// {"t":"C","c1x":..,"c1y":..,"c2x":..,"c2y":..,"x":..,"y":..},{"t":"Z"}]}`.
/// The canvas caches this per `(fontId, glyphId)` and scales by `size/upem`,
/// flipping y at draw time. `cmds` is empty for a blank glyph (space).
#[cfg_attr(not(feature = "wasm-query"), wasm_bindgen)]
pub fn outline_glyph_json(font_id: u32, glyph_id: u32) -> Result<String, JsValue> {
    with_module_fonts(|| outline_glyph_json_resident(font_id, glyph_id))
        .map_err(|e| JsValue::from_str(&e))
}

/// Reads a glyph outline from the current measurement font store.
pub fn outline_glyph_json_resident(font_id: u32, glyph_id: u32) -> Result<String, String> {
    let glyph_id = u16::try_from(glyph_id)
        .map_err(|_| format!("glyph id {glyph_id} out of range for this font"))?;
    with_measure_fonts(|store| {
        store
            .borrow()
            .outline_glyph_json(ooxml_text::FontId::from_u32(font_id), glyph_id)
    })
    .map_err(|e| e.to_string())
}

/// [`ooxml_text::echo_glyph_set`] against the resident font store, as JSON.
pub fn echo_glyph_set_json(
    font_id: u32,
    size: f32,
    kern: bool,
    alphabet: &str,
) -> Result<String, String> {
    if !size.is_finite() || size <= 0.0 {
        return Err(format!("echo glyph size {size} is not positive"));
    }
    let alphabet: Vec<char> = alphabet.chars().collect();
    let set = with_measure_fonts(|store| {
        ooxml_text::echo_glyph_set(
            &store.borrow(),
            ooxml_text::FontId::from_u32(font_id),
            size,
            kern,
            &alphabet,
        )
    })
    .map_err(|error| error.to_string())?;
    serde_json::to_string(&set).map_err(|error| error.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;

    // Reuses the ooxml-text crate's vendored Liberation Sans fixture (the
    // metric-compatible Arial stand-in) to exercise the outline wasm surface.
    const LIBERATION_SANS: &[u8] =
        include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

    #[test]
    fn plain_faces_and_segoe_substitute_views_keep_separate_widths() {
        clear_measure_fonts();
        let base = register_measure_font_bytes(LIBERATION_SANS).unwrap();
        for family in ["Segoe UI Light", "Segoe UI Semilight"] {
            let view = register_substitute_measure_font(base, family).unwrap();
            assert_ne!(view, base);
            with_measure_fonts(|store| {
                let store = store.borrow();
                assert_eq!(
                    store
                        .advance_scale(ooxml_text::FontId::from_u32(base))
                        .unwrap(),
                    1.0
                );
                let metrics = ooxml_text::family_metrics::family_metrics(
                    &family.to_lowercase(),
                    false,
                    false,
                )
                .unwrap();
                let expected = ooxml_text::family_metrics::family_advance(metrics, 'm').unwrap();
                let advance = store
                    .advance_width(ooxml_text::FontId::from_u32(view), 'm')
                    .unwrap()
                    .unwrap();
                let upem = store
                    .metrics(ooxml_text::FontId::from_u32(view))
                    .unwrap()
                    .units_per_em;
                assert!((advance / f32::from(upem) - expected).abs() < 1e-6);
            });
        }
        clear_measure_fonts();
    }

    #[test]
    fn arial_narrow_substitute_measures_paragraphs_at_its_table_widths() {
        clear_measure_fonts();
        let base = register_measure_font_bytes(LIBERATION_SANS).unwrap();
        let view = register_substitute_measure_font(base, "Arial Narrow").unwrap();
        assert_ne!(view, base);
        let text = "the quick brown fox jumps over minimum widths";
        let width = |font| {
            let input = serde_json::json!({
                "block": {
                    "kind": "paragraph",
                    "runs": [{"kind": "text", "text": text, "fontFamily": "Arial Narrow"}]
                },
                "maxWidth": 1000.0,
                "fontChains": {"arial narrow|0|0": [font]},
                "defaults": {"fontSize": 12.0, "fontFamily": "Arial Narrow"}
            });
            let output = measure_paragraph_json_resident(&input.to_string()).unwrap();
            let output: serde_json::Value = serde_json::from_str(&output).unwrap();
            assert_eq!(output["lines"].as_array().unwrap().len(), 1);
            output["lines"][0]["width"].as_f64().unwrap()
        };
        let metrics =
            ooxml_text::family_metrics::family_metrics("arial narrow", false, false).unwrap();
        let expected = text
            .chars()
            .map(|ch| ooxml_text::family_metrics::family_advance(metrics, ch).unwrap())
            .sum::<f32>()
            * 16.0;
        let measured = width(view);
        let base_width = width(base);
        assert!((measured - f64::from(expected)).abs() < 1e-3);
        assert!((measured / base_width - 0.82).abs() < 0.01);
        clear_measure_fonts();
    }

    #[test]
    fn measurement_face_descriptors_distinguish_width_tables() {
        clear_measure_fonts();
        let base = register_measure_font_bytes(LIBERATION_SANS).unwrap();
        let view = register_substitute_measure_font(base, "Arial Narrow").unwrap();
        let scaled = with_measure_fonts(|store| {
            let mut store = store.borrow_mut();
            let view = ooxml_text::FontId::from_u32(view);
            let mut requested =
                ooxml_text::word_fonts::requested_line_metrics("Arial Narrow").unwrap();
            requested.advance_scale = store.advance_scale(view).unwrap();
            let scaled = store
                .register_substitute(ooxml_text::FontId::from_u32(base), requested)
                .unwrap();
            assert_eq!(
                store.measured_face(view).unwrap(),
                store.measured_face(scaled).unwrap()
            );
            scaled.to_u32()
        });
        assert_ne!(
            with_measure_face(view, |_, descriptor| descriptor.to_owned()).unwrap(),
            with_measure_face(scaled, |_, descriptor| descriptor.to_owned()).unwrap()
        );
        clear_measure_fonts();
    }

    #[test]
    fn outline_glyph_json_round_trips_through_the_font_store() {
        clear_measure_fonts();
        let font_id = register_measure_font(LIBERATION_SANS).expect("register fixture font");

        // 'A' is glyph 36 in LiberationSans (cross-checked in ooxml-text tests).
        let json = outline_glyph_json(font_id, 36).expect("outline json");
        let v: serde_json::Value = serde_json::from_str(&json).expect("valid json");

        assert_eq!(v["upem"], 2048);
        let cmds = v["cmds"].as_array().expect("cmds array");
        assert!(!cmds.is_empty(), "'A' has contours");
        assert_eq!(cmds[0]["t"], "M");
        assert!(cmds[0]["x"].is_number() && cmds[0]["y"].is_number());
        assert_eq!(cmds.last().unwrap()["t"], "Z");

        clear_measure_fonts();
    }

    #[test]
    fn echo_glyph_set_json_reads_the_resident_font_store() {
        clear_measure_fonts();
        let font_id = register_measure_font_bytes(LIBERATION_SANS).expect("register fixture font");
        let json = echo_glyph_set_json(font_id, 16.0, false, "Ax").expect("echo json");
        let v: serde_json::Value = serde_json::from_str(&json).expect("valid json");
        assert_eq!(v["glyphIds"][0], 36);
        assert_eq!(v["advances"].as_array().map(Vec::len), Some(2));
        assert!(v["horizontal"].as_f64().is_some_and(|value| value > 0.0));
        assert!(echo_glyph_set_json(font_id, 0.0, false, "A").is_err());
        clear_measure_fonts();
    }

    #[test]
    fn entered_font_stores_nest_and_restore() {
        let holds = |id| with_measure_face(id, |_, _| ()).is_some();
        clear_measure_fonts();
        register_measure_font_bytes(LIBERATION_SANS).unwrap();
        register_measure_font_bytes(LIBERATION_SANS).unwrap();
        {
            let (a, b) = (MeasureFonts::default(), MeasureFonts::default());
            let _a = a.enter();
            assert!(!holds(0));
            assert_eq!(register_measure_font_bytes(LIBERATION_SANS), Ok(0));
            {
                let _b = b.enter();
                assert!(!holds(0));
                {
                    let _a = a.enter();
                    assert!(holds(0));
                }
                assert!(!holds(0));
            }
            assert!(holds(0));
        }
        assert!(holds(1));
        let unentered = MeasureFonts::default();
        assert!(std::panic::catch_unwind(|| holds(1)).is_err());
        drop(unentered);
        assert!(holds(1));
    }

    fn font_measurement_input(font_id: u32) -> String {
        serde_json::json!({
            "block": {"kind": "paragraph", "runs": [{"kind": "text", "text": "Font stores"}]},
            "maxWidth": 400.0,
            "fontChains": {"liberation sans|0|0": [font_id]},
            "defaults": {"fontSize": 12.0, "fontFamily": "Liberation Sans"}
        })
        .to_string()
    }

    #[test]
    fn free_font_exports_use_module_fonts_beside_an_unentered_session() {
        clear_measure_fonts_wasm().unwrap();
        let session = MeasureFonts::default();
        let (session_generation, session_measurement, session_outline) = {
            let _scope = session.enter();
            assert_eq!(register_measure_font_bytes(LIBERATION_SANS), Ok(0));
            (
                measure_fonts_generation(),
                measure_paragraph_json_resident(&font_measurement_input(0)).unwrap(),
                outline_glyph_json_resident(0, 36).unwrap(),
            )
        };
        assert_eq!(register_measure_font(LIBERATION_SANS).unwrap(), 0);
        let module_id = with_module_fonts(measure_store_id);
        assert_ne!(module_id, session_generation.0);
        assert_eq!(with_module_fonts(measure_fonts_generation), (module_id, 1));
        let substitute = register_substitute_measure_font(0, "Segoe UI Light").unwrap();
        assert_eq!(substitute, 1);
        let measured = measure_paragraph_json(&font_measurement_input(substitute)).unwrap();
        assert!(!measured.is_empty());
        assert_eq!(outline_glyph_json(0, 36).unwrap(), session_outline);
        clear_measure_fonts_wasm().unwrap();
        assert_ne!(with_module_fonts(measure_store_id), module_id);
        assert_eq!(with_module_fonts(measure_fonts_generation).1, 0);
        {
            let _scope = session.enter();
            assert_eq!(measure_fonts_generation(), session_generation);
            assert_eq!(
                measure_paragraph_json_resident(&font_measurement_input(0)).unwrap(),
                session_measurement
            );
            assert_eq!(outline_glyph_json_resident(0, 36).unwrap(), session_outline);
            assert_eq!(register_measure_font_bytes(LIBERATION_SANS), Ok(1));
        }
        assert!(ENTERED_FONTS.with(|entered| entered.borrow().is_none()));
    }

    #[test]
    fn free_font_exports_preserve_an_entered_session_store() {
        clear_measure_fonts_wasm().unwrap();
        assert_eq!(register_measure_font(LIBERATION_SANS).unwrap(), 0);
        let module_generation = measure_fonts_generation();
        let module_measurement = measure_paragraph_json(&font_measurement_input(0)).unwrap();
        let module_outline = outline_glyph_json(0, 36).unwrap();
        let session = MeasureFonts::default();
        {
            let _scope = session.enter();
            let session_id = measure_store_id();
            assert_ne!(session_id, module_generation.0);
            assert_eq!(register_measure_font(LIBERATION_SANS).unwrap(), 0);
            assert_eq!(measure_fonts_generation(), (session_id, 1));
            assert_eq!(
                register_substitute_measure_font(0, "Segoe UI Light").unwrap(),
                1
            );
            assert_eq!(measure_fonts_generation(), (session_id, 2));
            let input = font_measurement_input(1);
            assert_eq!(
                measure_paragraph_json(&input).unwrap(),
                measure_paragraph_json_resident(&input).unwrap()
            );
            assert_eq!(
                outline_glyph_json(1, 36).unwrap(),
                outline_glyph_json_resident(1, 36).unwrap()
            );
            clear_measure_fonts_wasm().unwrap();
            assert_ne!(measure_store_id(), session_id);
            assert_eq!(measure_fonts_generation().1, 0);
            assert_eq!(register_measure_font(LIBERATION_SANS).unwrap(), 0);
        }
        assert_eq!(
            with_module_fonts(measure_fonts_generation),
            module_generation
        );
        assert_eq!(
            measure_paragraph_json(&font_measurement_input(0)).unwrap(),
            module_measurement
        );
        assert_eq!(outline_glyph_json(0, 36).unwrap(), module_outline);
        clear_measure_fonts_wasm().unwrap();
    }

    #[test]
    fn module_font_scope_restores_after_a_panic() {
        let session = MeasureFonts::default();
        let result = std::panic::catch_unwind(|| {
            with_module_fonts(|| panic!("scope unwind"));
        });
        assert!(result.is_err());
        assert!(ENTERED_FONTS.with(|entered| entered.borrow().is_none()));
        let _scope = session.enter();
        let id = measure_store_id();
        with_module_fonts(|| assert_eq!(measure_store_id(), id));
        assert_eq!(measure_store_id(), id);
    }

    fn options_json() -> serde_json::Value {
        serde_json::json!({
            "pageSize": {"w": 816.0, "h": 1056.0},
            "margins": {"top": 96.0, "right": 96.0, "bottom": 96.0, "left": 96.0},
            "pageGap": 20.0
        })
    }

    fn para(id: u32, text: &str, pm_start: u32, height: f64) -> serde_json::Value {
        let pm_end = pm_start + text.len() as u32;
        serde_json::json!({
            "block": {
                "kind": "paragraph",
                "id": id,
                "runs": [{"kind": "text", "text": text, "pmStart": pm_start, "pmEnd": pm_end}],
                "pmStart": pm_start,
                "pmEnd": pm_end + 1
            },
            "measure": {
                "kind": "paragraph",
                "totalHeight": height,
                "lines": [{
                    "headRun": 0, "headChar": 0, "tailRun": 0,
                    "tailChar": text.len(),
                    "width": 120.0, "ascent": height * 0.8, "descent": height * 0.2,
                    "lineHeight": height
                }]
            }
        })
    }

    #[test]
    fn stacks_paragraphs_on_one_page() {
        let input = serde_json::json!({
            "measured": [para(0, "First paragraph", 1, 24.0), para(1, "Second paragraph", 18, 24.0)],
            "options": options_json()
        })
        .to_string();
        let out = layout_to_json(&input).unwrap();
        let v: serde_json::Value = serde_json::from_str(&out).unwrap();
        assert_eq!(v["pages"].as_array().unwrap().len(), 1);
        let frags = v["pages"][0]["fragments"].as_array().unwrap();
        assert_eq!(frags.len(), 2);
        assert_eq!(frags[0]["y"], 96.0);
        assert_eq!(frags[1]["y"], 120.0);
        assert_eq!(frags[0]["width"], 624.0);
    }

    #[test]
    fn overflows_to_a_second_page() {
        let mut measured = Vec::new();
        for i in 0..10u32 {
            measured.push(para(i, "Paragraph", i * 15 + 1, 100.0));
        }
        let input =
            serde_json::json!({ "measured": measured, "options": options_json() }).to_string();
        let out = layout_to_json(&input).unwrap();
        let v: serde_json::Value = serde_json::from_str(&out).unwrap();
        assert_eq!(v["pages"].as_array().unwrap().len(), 2);
    }

    #[test]
    fn unsupported_kind_errors_for_fallback() {
        let input = serde_json::json!({
            "measured": [{
                "block": {"kind": "somethingNew", "id": 0},
                "measure": {"kind": "somethingNew"}
            }],
            "options": options_json()
        })
        .to_string();
        assert_eq!(layout_to_json(&input).unwrap_err(), "UNSUPPORTED");
    }

    #[test]
    fn places_a_simple_table() {
        let input = serde_json::json!({
            "measured": [{
                "block": {
                    "kind": "table", "id": 0, "columnWidths": [100.0],
                    "rows": [{"id": 1, "cells": [{"id": 2, "blocks": [
                        {"kind": "paragraph", "id": 3, "runs": []}
                    ]}]}]
                },
                "measure": {
                    "kind": "table", "columnWidths": [100.0],
                    "totalWidth": 100.0, "totalHeight": 24.0,
                    "rows": [{"height": 24.0, "cells": [{
                        "blocks": [{"kind": "paragraph", "lines": [], "totalHeight": 24.0}],
                        "width": 100.0, "height": 24.0
                    }]}]
                }
            }],
            "options": options_json()
        })
        .to_string();
        let out = layout_to_json(&input).unwrap();
        let v: serde_json::Value = serde_json::from_str(&out).unwrap();
        assert_eq!(v["pages"].as_array().unwrap().len(), 1);
        let frag = &v["pages"][0]["fragments"][0];
        assert_eq!(frag["kind"], "table");
        assert_eq!(frag["y"], 96.0);
        assert_eq!(frag["height"], 24.0);
        assert_eq!(frag["rowStart"], 0);
        assert_eq!(frag["rowEnd"], 1);
    }

    #[test]
    fn floating_table_engages_the_floating_table_hook() {
        let input = serde_json::json!({
            "measured": [{
                "block": {"kind": "table", "id": 0, "rows": [], "floating": {}},
                "measure": {"kind": "table", "rows": [], "columnWidths": [], "totalWidth": 0.0, "totalHeight": 0.0}
            }],
            "options": options_json()
        })
        .to_string();
        assert_eq!(layout_to_json(&input).unwrap_err(), "UNSUPPORTED");
        assert!(matches!(
            compute_layout(&input),
            Err(LayoutError::Unsupported(_))
        ));
    }

    #[test]
    fn empty_document_still_yields_page_one() {
        let input = serde_json::json!({ "measured": [], "options": options_json() }).to_string();
        let out = layout_to_json(&input).unwrap();
        let v: serde_json::Value = serde_json::from_str(&out).unwrap();
        assert_eq!(v["pages"].as_array().unwrap().len(), 1);
        assert_eq!(v["pages"][0]["number"], 1);
    }
}
