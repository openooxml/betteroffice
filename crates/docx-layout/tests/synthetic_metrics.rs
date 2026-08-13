//! Geometry of the no-face measurement fallback (issue #185).
//!
//! Published packages register no measurement font unless the host supplies
//! one, so the measure engine refuses every paragraph and `measure_blocks`
//! substitutes an estimated extent. The canvas still paints each run with a
//! real face, so the estimate has to leave every run a slot at least as wide
//! as its own estimated advance — otherwise consecutive runs overprint. A
//! tracked replacement puts two runs (the deletion and its insertion) on one
//! paragraph, which is where the collision becomes visible.

use docx_layout::display_list::Primitive;
use docx_layout::measure_blocks::{MeasurementConfig, measure_blocks};
use docx_layout::types::{Input, LayoutBlock, MeasuredBlock};

const CONTENT_WIDTH: f64 = 624.0;
const FONT_SIZE_PT: f64 = 11.0;
const FONT_SIZE_PX: f64 = FONT_SIZE_PT * 96.0 / 72.0;
/// The per-code-unit advance the fallback bills, mirrored from the engine.
const ESTIMATED_ADVANCE: f64 = FONT_SIZE_PX * 0.5;

const INSERTED: &str = "Vielen Dank für Ihre Bestellung Nr. 4711. Wir freuen uns, Ihnen bestätigen zu können, dass wir Ihre Zahlung erhalten haben.";
const DELETED: &str =
    "Thank you for your order #4711. We are pleased to confirm that we have received your payment.";

fn utf16_len(text: &str) -> f64 {
    text.encode_utf16().count() as f64
}

/// One paragraph holding a tracked deletion and its replacement insertion.
fn revision_paragraph() -> serde_json::Value {
    let inserted_end = 1.0 + utf16_len(INSERTED);
    serde_json::json!({
        "kind": "paragraph",
        "id": 0,
        "runs": [
            {
                "kind": "text",
                "text": INSERTED,
                "pmStart": 1.0,
                "pmEnd": inserted_end,
                "fontSize": FONT_SIZE_PT,
                "isInsertion": true,
                "changeAuthor": "Translator",
                "changeRevisionId": 11
            },
            {
                "kind": "text",
                "text": DELETED,
                "pmStart": inserted_end,
                "pmEnd": inserted_end + utf16_len(DELETED),
                "fontSize": FONT_SIZE_PT,
                "isDeletion": true,
                "changeAuthor": "Translator",
                "changeRevisionId": 11
            }
        ],
        "pmStart": 0.0,
        "pmEnd": inserted_end + utf16_len(DELETED) + 1.0
    })
}

fn layout_page() -> serde_json::Value {
    serde_json::json!({
        "pages": [{
            "size": { "w": 816.0, "h": 1056.0 },
            "margins": { "top": 96.0, "right": 96.0, "bottom": 96.0, "left": 96.0 },
            "number": 1
        }]
    })
}

/// Text primitives as `(baseline, x, width, text)`, in paint order.
fn text_runs(list: &docx_layout::display_list::DisplayList) -> Vec<(f64, f64, f64, String)> {
    list.pages[0]
        .primitives
        .iter()
        .filter_map(|primitive| match primitive {
            Primitive::Text(run) => Some((
                run.baseline_y.as_f64().unwrap_or_default(),
                run.x.as_f64().unwrap_or_default(),
                run.width.as_f64().unwrap_or_default(),
                run.text.clone(),
            )),
            _ => None,
        })
        .collect()
}

fn display_list_without_faces() -> docx_layout::display_list::DisplayList {
    docx_layout::clear_measure_fonts();
    let block: LayoutBlock = serde_json::from_value(revision_paragraph()).expect("block parses");
    let mut blocks = vec![block];
    let config = MeasurementConfig {
        font_chains: Default::default(),
        defaults: serde_json::json!({ "fontSize": FONT_SIZE_PT, "fontFamily": "Calibri" }),
        compat: serde_json::Value::Null,
        authoritative_shaping: true,
    };
    let extents = measure_blocks(&mut blocks, CONTENT_WIDTH, &config).expect("fallback measures");
    let mut input = Input {
        measured: blocks
            .into_iter()
            .zip(extents)
            .map(|(block, measure)| MeasuredBlock { block, measure })
            .collect(),
        options: serde_json::from_value(serde_json::json!({ "layout": layout_page() }))
            .unwrap_or_default(),
    };
    let layout = docx_layout::compute_layout_input(&mut input).expect("paginates");
    docx_layout::build_display_list(&input, &layout).expect("display list builds")
}

/// Without a face the estimate is all the engine has, so every run must be
/// allotted the width that estimate implies. Billing a line less than its own
/// estimate — the old `min(content_width, …)` clamp — shrinks every slot on it
/// by the overflow ratio, and the next run starts inside the previous one.
#[test]
fn tracked_replacement_runs_do_not_overprint_without_a_measured_face() {
    let list = display_list_without_faces();
    let runs = text_runs(&list);
    assert!(!runs.is_empty(), "the paragraph paints text");

    for (baseline, x, width, text) in &runs {
        let estimated = utf16_len(text) * ESTIMATED_ADVANCE;
        assert!(
            *width + 0.01 >= estimated,
            "run {text:?} at ({x}, {baseline}) got a {width}px slot for {estimated}px of estimated text",
        );
    }

    let mut by_baseline: std::collections::BTreeMap<i64, Vec<(f64, f64, String)>> =
        Default::default();
    for (baseline, x, _, text) in runs {
        by_baseline
            .entry((baseline * 100.0).round() as i64)
            .or_default()
            .push((x, utf16_len(&text) * ESTIMATED_ADVANCE, text));
    }
    for (baseline, mut row) in by_baseline {
        row.sort_by(|left, right| left.0.total_cmp(&right.0));
        for pair in row.windows(2) {
            let (x, estimated, text) = &pair[0];
            let (next_x, _, next_text) = &pair[1];
            assert!(
                x + estimated <= next_x + 0.01,
                "on baseline {}: {text:?} spans [{x}, {}] and overprints {next_text:?} at {next_x}",
                baseline as f64 / 100.0,
                x + estimated,
            );
        }
    }
}

/// The single unwrapped line the fallback used to emit ran the paragraph off
/// the page instead of breaking it.
#[test]
fn the_fallback_breaks_a_long_paragraph_into_column_wide_lines() {
    let list = display_list_without_faces();
    let runs = text_runs(&list);
    let baselines: std::collections::BTreeSet<i64> = runs
        .iter()
        .map(|(baseline, ..)| (baseline * 100.0).round() as i64)
        .collect();
    assert!(
        baselines.len() > 1,
        "a paragraph {} code units long must wrap inside a {CONTENT_WIDTH}px column, got one line",
        utf16_len(INSERTED) + utf16_len(DELETED),
    );

    let right_edge = 96.0 + CONTENT_WIDTH;
    for (baseline, x, width, text) in &runs {
        assert!(
            x + width <= right_edge + 0.01,
            "run {text:?} on baseline {baseline} ends at {} past the {right_edge}px column edge",
            x + width,
        );
    }
}
