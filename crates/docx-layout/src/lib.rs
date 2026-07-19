// Ported from openooxml/docx, which did not gate on clippy style lints;
// burning these down is tracked follow-up work, not a merge blocker.
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

//! The pure layout-core (pagination) in Rust, compiled to WASM.
//!
//! A twin of the TypeScript `layoutCore` behind the same contract:
//! `(MeasuredBlock[], LayoutOptions) -> Layout`, marshaled as JSON. It ports
//! INCREMENTALLY — this spine paginates paragraph flow (spacing collapse,
//! page/column splits, PM ranges, resolved lines), explicit page/column
//! breaks, multi-column options, inline/anchored images, text boxes, and
//! per-page footnote reservations. Features still owned by the TypeScript
//! engine sit behind the stubs in [`hooks`]; engaging one returns
//! `Unsupported`, so the JS seam falls back. Coverage grows one hook at a
//! time, each checkpointed against the golden corpus
//! (`crates/docx-layout/tests/goldens.rs` natively,
//! `packages/core/src/layout/pagination/__golden__/rustParity.test.ts` via
//! wasm).
//!
//! Module map (TS source → Rust module):
//! - `pagination/types.ts` → [`types`]
//! - `pagination/pageFlow.ts` → [`page_flow`]
//! - `pagination/prescan.ts` → [`prescan`]
//! - `pagination/resolveLineSegments.ts` → [`resolve_lines`]
//! - `pagination/index.ts` (place walk) → [`place`]
//! - `__golden__/serializeLayout.ts` → [`canonical`]
//! - not-yet-ported feature seams → [`hooks`]
//!
//! wasm_bindgen stays confined to this file; everything below the
//! [`layout_to_json`] / [`layout_to_canonical_json`] boundary is pure and
//! native-testable.

pub mod canonical;
pub mod hooks;
pub mod page_flow;
pub mod paragraph_spacing;
pub mod place;
pub mod prescan;
pub mod regions;
pub mod resolve_lines;
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
pub mod session;
pub mod table_grid;
pub mod table_row_break;

use wasm_bindgen::prelude::*;

/// Why the engine refused an input.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum LayoutError {
    /// The input engages a feature this port doesn't cover yet — the caller
    /// must fall back to the reference (TypeScript) engine.
    Unsupported(String),
    /// The input violates the layout contract (the TS engine would throw).
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

/// Parse the `{ measured, options }` envelope and run the placement walk.
pub fn compute_layout(input: &str) -> Result<types::Layout, LayoutError> {
    let mut parsed: types::Input =
        serde_json::from_str(input).map_err(|e| LayoutError::Invalid(format!("parse: {e}")))?;
    place::layout_document(&mut parsed)
}

/// Pure JSON boundary: `{ measured, options }` in, `Layout` JSON out. `Err`
/// carries a reason; `"UNSUPPORTED"` means the input needs the reference
/// engine. Native-testable (no JsValue).
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
#[wasm_bindgen]
pub fn layout_document_json(input: &str) -> Result<String, JsValue> {
    layout_to_json(input).map_err(|e| JsValue::from_str(&e))
}

/// wasm wrapper over [`display_list::build_display_list_json_with_fonts`]:
/// `{ measured, options, layout }` JSON in, `DisplayList` JSON out. Threads the
/// module-global measurement fonts (the same store `measure_paragraph_json`
/// reads), so text runs whose `fontChains` resolve are shaped into
/// `GlyphRunPrimitive`s; when no fonts are registered (browser measurement)
/// every run falls back to `TextRunPrimitive`, byte-identical to before.
pub fn build_display_list_value(input: &str) -> Result<display_list::DisplayList, String> {
    MEASURE_FONTS
        .with(|store| display_list::build_display_list_value_with_fonts(input, &store.borrow()))
}

/// Resident counterpart to [`build_display_list_value`]. Large measured and
/// layout values stay inside the editing wasm; `extras` carries only the
/// display-specific fields not owned by pagination.
pub fn build_display_list_value_from_resident(
    pagination: &types::Input,
    layout: &types::Layout,
    extras: &str,
) -> Result<display_list::DisplayList, String> {
    MEASURE_FONTS.with(|store| {
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
    MEASURE_FONTS.with(|store| {
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
    MEASURE_FONTS.with(|store| {
        display_list::build_resident_display_list_with_fonts_observed(
            pagination,
            layout,
            extras,
            &store.borrow(),
            observe_phase,
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
    MEASURE_FONTS.with(|store| {
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
    MEASURE_FONTS.with(|store| {
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
    MEASURE_FONTS.with(|store| {
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
pub fn update_resident_display_list_incremental_observed(
    pagination: &types::Input,
    layout: &types::Layout,
    resident: &mut display_list::ResidentDisplayInput,
    previous: &mut display_list::DisplayList,
    rebuilt_page_start: usize,
    rebuilt_page_end: usize,
    position_deltas: &std::collections::HashMap<String, i64>,
    observe_phase: &mut impl FnMut(),
) -> Result<bool, String> {
    MEASURE_FONTS.with(|store| {
        display_list::update_resident_display_list_incremental_with_fonts_observed(
            pagination,
            layout,
            &store.borrow(),
            resident,
            previous,
            rebuilt_page_start,
            rebuilt_page_end,
            position_deltas,
            observe_phase,
        )
    })
}

/// wasm compatibility wrapper. Resident engine users call
/// [`build_display_list_value`] and keep the typed result.
#[wasm_bindgen]
pub fn build_display_list_json(input: &str) -> Result<String, JsValue> {
    build_display_list_value(input)
        .and_then(|display_list| {
            serde_json::to_string(&display_list).map_err(|e| format!("serialize: {e}"))
        })
        .map_err(|e| JsValue::from_str(&e))
}

/// wasm wrapper over [`hit::hit_test_json`]: display-list JSON + page-local
/// point in, PM position (or `null`) as JSON out.
#[wasm_bindgen]
pub fn hit_test_json(
    display_list: &str,
    page_index: u32,
    x: f64,
    y: f64,
) -> Result<String, JsValue> {
    hit::hit_test_json(display_list, page_index as usize, x, y).map_err(|e| JsValue::from_str(&e))
}

/// wasm wrapper over [`hit::range_rects_json`]: display-list JSON + PM range
/// in, JSON array of page-local rects out.
#[wasm_bindgen]
pub fn range_rects_json(display_list: &str, from: f64, to: f64) -> Result<String, JsValue> {
    hit::range_rects_json(display_list, from as i64, to as i64).map_err(|e| JsValue::from_str(&e))
}

/// wasm wrapper over [`hit::range_rects_region_json`]: region-aware range rects.
/// `region` is `"body" | "header" | "footer"`; `r_id` scopes a header/footer to
/// one HF part (empty for body / match-any). The `from`/`to` refer to that
/// region's PM doc. The legacy `range_rects_json` export stays body-only.
#[wasm_bindgen]
pub fn range_rects_region_json(
    display_list: &str,
    region: &str,
    r_id: &str,
    from: f64,
    to: f64,
) -> Result<String, JsValue> {
    hit::range_rects_region_json(display_list, region, r_id, from as i64, to as i64)
        .map_err(|e| JsValue::from_str(&e))
}

/// wasm wrapper over [`hit::hit_test_regions_json`]: region-aware hit test —
/// `{"region":"body"|"header"|"footer","rId"?,"pos":n|null}` (or `"null"` for
/// an out-of-range page). The legacy `hit_test_json` export stays body-only.
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

/// wasm wrapper over [`session::range_rects_by_handle`]: range rects against a
/// stored display list. `Err` on an unknown/closed handle so the caller can
/// fall back to [`range_rects_json`].
#[wasm_bindgen]
pub fn range_rects_by_handle(handle: u32, from: f64, to: f64) -> Result<String, JsValue> {
    session::range_rects_by_handle(handle, from as i64, to as i64)
        .map_err(|e| JsValue::from_str(&e))
}

/// wasm wrapper over [`session::range_rects_region_by_handle`]: region-aware
/// range rects against a stored display list. `region` is
/// `"body" | "header" | "footer"`; `r_id` scopes header/footer to one HF part.
/// `Err` on an unknown/closed handle so the caller can fall back to
/// [`range_rects_region_json`].
#[wasm_bindgen]
pub fn range_rects_region_by_handle(
    handle: u32,
    region: &str,
    r_id: &str,
    from: f64,
    to: f64,
) -> Result<String, JsValue> {
    session::range_rects_region_by_handle(handle, region, r_id, from as i64, to as i64)
        .map_err(|e| JsValue::from_str(&e))
}

// ---------------------------------------------------------------------------
// text measurement surface (ooxml-text)
//
// This crate is the workspace's only wasm-bindgen site, so the `ooxml-text`
// measurement pipeline is exported from here as the same kind of thin
// wrapper as the layout entry points above. All logic lives in
// `ooxml_text::measure`; this file only owns the wasm-visible font registry.
// ---------------------------------------------------------------------------

thread_local! {
    /// Fonts registered for measurement. WASM is single-threaded, so a
    /// thread_local doubles as the module-global store; native tests get an
    /// isolated store per test thread.
    static MEASURE_FONTS: std::cell::RefCell<ooxml_text::FontStore> =
        std::cell::RefCell::new(ooxml_text::FontStore::new());
}

/// Register a font for measurement from raw sfnt bytes; returns the font id
/// that `measure_paragraph_json` inputs reference in their `fontChains`.
/// Malformed bytes (attacker-controlled embedded fonts) are rejected as an
/// error at this boundary, mirroring `FontStore::register`.
#[wasm_bindgen]
pub fn register_measure_font(bytes: &[u8]) -> Result<u32, JsValue> {
    MEASURE_FONTS.with(|store| {
        store
            .borrow_mut()
            .register(bytes.to_vec())
            .map(|id| id.to_u32())
            .map_err(|e| JsValue::from_str(&e.to_string()))
    })
}

/// Drop every registered measurement font (ids restart at 0). Callers must
/// re-register before the next `measure_paragraph_json`.
#[wasm_bindgen]
pub fn clear_measure_fonts() {
    MEASURE_FONTS.with(|store| {
        *store.borrow_mut() = ooxml_text::FontStore::new();
    });
}

/// wasm wrapper over [`ooxml_text::measure_paragraph_json`]: measurement
/// input JSON in, `ParagraphExtent` JSON out. An `Err` whose message starts
/// with `"UNSUPPORTED"` means the caller must fall back to browser
/// measurement for that block.
#[wasm_bindgen]
pub fn measure_paragraph_json(input: &str) -> Result<String, JsValue> {
    measure_paragraph_json_resident(input).map_err(|e| JsValue::from_str(&e))
}

/// Native/resident-engine form of [`measure_paragraph_json`]. Keeping the
/// string error avoids constructing a wasm `JsValue` when the editing engine
/// remeasures a dirty paragraph inside a larger operation.
pub fn measure_paragraph_json_resident(input: &str) -> Result<String, String> {
    MEASURE_FONTS.with(|store| ooxml_text::measure_paragraph_json(&store.borrow(), input))
}

/// wasm wrapper over [`ooxml_text::FontStore::outline_glyph_json`]: the outline
/// of a registered font's glyph, in font design units, as JSON:
/// `{"upem":2048,"cmds":[{"t":"M","x":..,"y":..},{"t":"L","x":..,"y":..},
/// {"t":"Q","cx":..,"cy":..,"x":..,"y":..},
/// {"t":"C","c1x":..,"c1y":..,"c2x":..,"c2y":..,"x":..,"y":..},{"t":"Z"}]}`.
/// The canvas caches this per `(fontId, glyphId)` and scales by `size/upem`,
/// flipping y at draw time. `cmds` is empty for a blank glyph (space).
#[wasm_bindgen]
pub fn outline_glyph_json(font_id: u32, glyph_id: u32) -> Result<String, JsValue> {
    // Glyph ids are u16 in the sfnt spec; a value past that is out of range for
    // any font, so reject at the boundary rather than truncating.
    let glyph_id = u16::try_from(glyph_id).map_err(|_| {
        JsValue::from_str(&format!("glyph id {glyph_id} out of range for this font"))
    })?;
    MEASURE_FONTS
        .with(|store| {
            store
                .borrow()
                .outline_glyph_json(ooxml_text::FontId::from_u32(font_id), glyph_id)
        })
        .map_err(|e| JsValue::from_str(&e.to_string()))
}

#[cfg(test)]
mod tests {
    use super::*;

    // Reuses the ooxml-text crate's vendored Liberation Sans fixture (the
    // metric-compatible Arial stand-in) to exercise the outline wasm surface.
    const LIBERATION_SANS: &[u8] =
        include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

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
