//! No-face measurement fallback geometry (issue #185).

use docx_layout::display_list::Primitive;
use docx_layout::measure_blocks::{MeasurementConfig, measure_blocks};
use docx_layout::types::{BlockExtent, Input, LayoutBlock, MeasuredBlock, ParagraphExtent};

const CONTENT_WIDTH: f64 = 624.0;
const FONT_SIZE_PT: f64 = 11.0;
const FONT_SIZE_PX: f64 = FONT_SIZE_PT * 96.0 / 72.0;
const ESTIMATED_ADVANCE: f64 = FONT_SIZE_PX * 0.5;
const LIBERATION: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

const INSERTED: &str = "Vielen Dank für Ihre Bestellung Nr. 4711. Wir freuen uns, Ihnen bestätigen zu können, dass wir Ihre Zahlung erhalten haben.";
const DELETED: &str =
    "Thank you for your order #4711. We are pleased to confirm that we have received your payment.";

fn utf16_len(text: &str) -> usize {
    text.encode_utf16().count()
}

fn estimated_width(text: &str, font_size_pt: f64) -> f64 {
    utf16_len(text) as f64 * font_size_pt * 96.0 / 72.0 * 0.5
}

fn fallback_config() -> MeasurementConfig {
    MeasurementConfig {
        font_chains: Default::default(),
        defaults: serde_json::json!({
            "fontSize": FONT_SIZE_PT,
            "fontFamily": "Liberation Sans"
        }),
        compat: serde_json::Value::Null,
        authoritative_shaping: true,
    }
}

fn revision_paragraph() -> serde_json::Value {
    let inserted_end = 1 + utf16_len(INSERTED);
    serde_json::json!({
        "kind": "paragraph",
        "id": 0,
        "runs": [
            {
                "kind": "text",
                "text": INSERTED,
                "pmStart": 1,
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
        "pmStart": 0,
        "pmEnd": inserted_end + utf16_len(DELETED) + 1
    })
}

fn mixed_size_paragraph() -> serde_json::Value {
    serde_json::json!({
        "kind": "paragraph",
        "id": 0,
        "runs": [
            {
                "kind": "text",
                "text": "X",
                "fontFamily": "Liberation Sans",
                "fontSize": 72.0
            },
            {
                "kind": "text",
                "text": "iiiiiiiiii",
                "fontFamily": "Liberation Sans",
                "fontSize": FONT_SIZE_PT
            }
        ],
        "attrs": {
            "defaultFontFamily": "Liberation Sans",
            "defaultFontSize": FONT_SIZE_PT
        }
    })
}

fn measure_paragraph(
    paragraph: serde_json::Value,
    width: f64,
    config: &MeasurementConfig,
) -> Result<ParagraphExtent, String> {
    let block: LayoutBlock = serde_json::from_value(paragraph).expect("block parses");
    let mut blocks = vec![block];
    let mut extents = measure_blocks(&mut blocks, width, config)?;
    match extents.pop() {
        Some(BlockExtent::Paragraph(extent)) => Ok(extent),
        _ => panic!("paragraph extent expected"),
    }
}

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

fn display_list_without_faces(
    paragraph: serde_json::Value,
    width: f64,
) -> docx_layout::display_list::DisplayList {
    docx_layout::clear_measure_fonts();
    let block: LayoutBlock = serde_json::from_value(paragraph).expect("block parses");
    let mut blocks = vec![block];
    let extents =
        measure_blocks(&mut blocks, width, &fallback_config()).expect("fallback measures");
    let mut input = Input {
        measured: blocks
            .into_iter()
            .zip(extents)
            .map(|(block, measure)| MeasuredBlock { block, measure })
            .collect(),
        options: serde_json::from_value(serde_json::json!({
            "pageSize": { "w": width + 192.0, "h": 1056.0 },
            "margins": { "top": 96.0, "right": 96.0, "bottom": 96.0, "left": 96.0 }
        }))
        .expect("options parse"),
    };
    let layout = docx_layout::compute_layout_input(&mut input).expect("paginates");
    docx_layout::build_display_list(&input, &layout).expect("display list builds")
}

#[test]
fn tracked_replacement_runs_do_not_overprint_without_a_measured_face() {
    let list = display_list_without_faces(revision_paragraph(), CONTENT_WIDTH);
    let runs = text_runs(&list);
    assert!(!runs.is_empty(), "the paragraph paints text");

    for (baseline, x, width, text) in &runs {
        let estimated = estimated_width(text, FONT_SIZE_PT);
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
            .push((x, estimated_width(&text, FONT_SIZE_PT), text));
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

#[test]
fn fallback_breaks_a_long_paragraph_into_column_wide_lines() {
    let list = display_list_without_faces(revision_paragraph(), CONTENT_WIDTH);
    let runs = text_runs(&list);
    let baselines: std::collections::BTreeSet<i64> = runs
        .iter()
        .map(|(baseline, ..)| (baseline * 100.0).round() as i64)
        .collect();
    assert!(
        baselines.len() > 1,
        "a paragraph {} scalars long must wrap inside a {CONTENT_WIDTH}px column",
        INSERTED.chars().count() + DELETED.chars().count(),
    );

    let right_edge = 96.0 + CONTENT_WIDTH;
    for (baseline, x, width, text) in &runs {
        assert!(
            x + width <= right_edge + 0.01,
            "run {text:?} on baseline {baseline} ends at {} past {right_edge}",
            x + width,
        );
    }
}

#[test]
fn fallback_rejects_more_than_one_hundred_thousand_lines() {
    docx_layout::clear_measure_fonts();
    let paragraph = serde_json::json!({
        "kind": "paragraph",
        "id": 0,
        "runs": [{
            "kind": "text",
            "text": "a".repeat(100_001),
            "fontSize": FONT_SIZE_PT
        }]
    });
    let error = match measure_paragraph(paragraph, 1.0, &fallback_config()) {
        Ok(extent) => panic!("expected line cap, got {} lines", extent.lines.len()),
        Err(error) => error,
    };
    assert!(
        error.contains("too many lines (> 100000)"),
        "unexpected error: {error}"
    );
}

#[test]
fn astral_and_bmp_runs_use_the_same_utf16_unit() {
    let paragraph = serde_json::json!({
        "kind": "paragraph",
        "id": 0,
        "runs": [
            {
                "kind": "text",
                "text": "😀",
                "pmStart": 1,
                "pmEnd": 3,
                "fontSize": FONT_SIZE_PT
            },
            {
                "kind": "text",
                "text": "a",
                "pmStart": 3,
                "pmEnd": 4,
                "fontSize": FONT_SIZE_PT
            }
        ],
        "pmStart": 0,
        "pmEnd": 5
    });
    let list = display_list_without_faces(paragraph, 100.0);
    let runs = text_runs(&list);
    assert_eq!(runs.len(), 2);
    assert!((runs[0].2 - ESTIMATED_ADVANCE * 2.0).abs() < 0.01);
    assert!((runs[1].2 - ESTIMATED_ADVANCE).abs() < 0.01);
    assert!((runs[1].1 - runs[0].1 - ESTIMATED_ADVANCE * 2.0).abs() < 0.01);
}

#[test]
fn mixed_font_sizes_stay_close_to_a_registered_face_control() {
    docx_layout::clear_measure_fonts();
    let fallback = measure_paragraph(mixed_size_paragraph(), 120.0, &fallback_config())
        .expect("fallback measures");
    let fallback_runs = text_runs(&display_list_without_faces(mixed_size_paragraph(), 120.0));

    docx_layout::clear_measure_fonts();
    let font_id = docx_layout::register_measure_font(LIBERATION).expect("font registers");
    let control_config: MeasurementConfig = serde_json::from_value(serde_json::json!({
        "fontChains": { "liberation sans|0|0": [font_id] },
        "defaults": { "fontSize": FONT_SIZE_PT, "fontFamily": "Liberation Sans" }
    }))
    .expect("config parses");
    let control = measure_paragraph(mixed_size_paragraph(), 120.0, &control_config)
        .expect("control measures");
    docx_layout::clear_measure_fonts();

    assert_eq!(control.lines.len(), 1);
    assert_eq!(fallback.lines.len(), 2);
    assert_eq!(fallback.lines[0].tail_run, 1);
    assert_eq!(fallback.lines[0].tail_char, 9);
    assert!((fallback.lines[0].width - 114.0).abs() < 0.01);
    assert!(fallback.total_height <= control.total_height * 1.2);
    assert_eq!(fallback_runs.len(), 3);
    assert_eq!(fallback_runs[0].3, "X");
    assert!((fallback_runs[0].2 - estimated_width("X", 72.0)).abs() < 0.01);
    assert_eq!(fallback_runs[1].3, "iiiiiiiii");
    assert!((fallback_runs[1].2 - estimated_width("iiiiiiiii", FONT_SIZE_PT)).abs() < 0.01);
    assert_eq!(fallback_runs[2].3, "i");
    assert!((fallback_runs[2].2 - ESTIMATED_ADVANCE).abs() < 0.01);
}
