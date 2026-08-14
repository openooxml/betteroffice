//! No-face measurement fallback geometry (issue #185).

use docx_layout::display_list::Primitive;
use docx_layout::measure_blocks::{MeasurementConfig, measure_blocks};
use docx_layout::types::{BlockExtent, Input, LayoutBlock, MeasuredBlock, ParagraphExtent};
use std::collections::BTreeMap;

const CONTENT_WIDTH: f64 = 624.0;
const FONT_SIZE_PT: f64 = 11.0;
const FONT_SIZE_PX: f64 = FONT_SIZE_PT * 96.0 / 72.0;
const LIBERATION: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const NOTO_SANS_SC: &[u8] =
    include_bytes!("../../../packages/fonts-cjk/assets/NotoSansSC-Regular.otf");

const INSERTED: &str = "Vielen Dank für Ihre Bestellung Nr. 4711. Wir freuen uns, Ihnen bestätigen zu können, dass wir Ihre Zahlung erhalten haben.";
const DELETED: &str =
    "Thank you for your order #4711. We are pleased to confirm that we have received your payment.";

fn utf16_len(text: &str) -> usize {
    text.encode_utf16().count()
}

fn fallback_width(text: &str, font_size_pt: f64) -> f64 {
    text.chars().count() as f64 * font_size_pt * 96.0 / 72.0
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

fn registered_config(family: &str, font_id: u32) -> MeasurementConfig {
    let mut font_chains = BTreeMap::new();
    font_chains.insert(format!("{}|0|0", family.to_lowercase()), vec![font_id]);
    MeasurementConfig {
        font_chains,
        defaults: serde_json::json!({
            "fontSize": FONT_SIZE_PT,
            "fontFamily": family
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

fn cjk_paragraph(text: &str) -> serde_json::Value {
    serde_json::json!({
        "kind": "paragraph",
        "id": 0,
        "runs": [{
            "kind": "text",
            "text": text,
            "fontFamily": "Noto Sans SC",
            "fontSize": FONT_SIZE_PT
        }],
        "attrs": {
            "defaultFontFamily": "Noto Sans SC",
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

fn display_list(
    paragraph: serde_json::Value,
    width: f64,
    config: &MeasurementConfig,
) -> docx_layout::display_list::DisplayList {
    let block: LayoutBlock = serde_json::from_value(paragraph).expect("block parses");
    let mut blocks = vec![block];
    let extents = measure_blocks(&mut blocks, width, config).expect("paragraph measures");
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

fn display_list_without_faces(
    paragraph: serde_json::Value,
    width: f64,
) -> docx_layout::display_list::DisplayList {
    docx_layout::clear_measure_fonts();
    display_list(paragraph, width, &fallback_config())
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

#[test]
fn tracked_replacement_runs_tile_without_a_measured_face() {
    let list = display_list_without_faces(revision_paragraph(), CONTENT_WIDTH);
    let runs = text_runs(&list);
    assert_eq!(runs.len(), 2);
    assert!(list.pages[0].primitives.iter().all(|primitive| {
        !matches!(primitive, Primitive::Text(run) if run.paint_clip.is_none())
    }));
    assert!((runs[0].0 - runs[1].0).abs() < 0.01);

    for (_, _, width, text) in &runs {
        let slot = fallback_width(text, FONT_SIZE_PT);
        assert!(
            (*width - slot).abs() < 0.01,
            "run {text:?} got a {width}px slot instead of {slot}px"
        );
    }
    assert!(runs[0].1 + runs[0].2 <= runs[1].1 + 0.01);
}

#[test]
fn synthetic_line_clips_shaped_and_unresolved_runs_to_their_slots() {
    docx_layout::clear_measure_fonts();
    let font_id = docx_layout::register_measure_font(LIBERATION).expect("font registers");
    let paragraph = serde_json::json!({
        "kind": "paragraph",
        "id": 0,
        "runs": [
            {
                "kind": "text",
                "text": "@@@@@@@@@@@@@@@@@@@@",
                "pmStart": 1,
                "pmEnd": 21,
                "fontFamily": "Liberation Sans",
                "fontSize": 12.0
            },
            {
                "kind": "text",
                "text": "next",
                "pmStart": 21,
                "pmEnd": 25,
                "fontFamily": "Unregistered Face",
                "fontSize": 12.0
            }
        ],
        "attrs": {
            "defaultFontFamily": "Liberation Sans",
            "defaultFontSize": 12.0
        },
        "pmStart": 0,
        "pmEnd": 26
    });
    let block: LayoutBlock = serde_json::from_value(paragraph).expect("block parses");
    let mut blocks = vec![block];
    let extents = measure_blocks(
        &mut blocks,
        CONTENT_WIDTH,
        &registered_config("Liberation Sans", font_id),
    )
    .expect("paragraph measures");
    let BlockExtent::Paragraph(extent) = &extents[0] else {
        panic!("paragraph extent expected");
    };
    assert_eq!(extent.lines[0].synthetic_fallback, Some(true));

    let mut input = Input {
        measured: blocks
            .into_iter()
            .zip(extents)
            .map(|(block, measure)| MeasuredBlock { block, measure })
            .collect(),
        options: serde_json::from_value(serde_json::json!({
            "pageSize": { "w": 816.0, "h": 1056.0 },
            "margins": { "top": 96.0, "right": 96.0, "bottom": 96.0, "left": 96.0 }
        }))
        .expect("options parse"),
    };
    let layout = docx_layout::compute_layout_input(&mut input).expect("paginates");
    let extras = serde_json::json!({
        "fontChains": { "liberation sans|0|0": [font_id] }
    })
    .to_string();
    let list = docx_layout::build_display_list_value_from_resident(&input, &layout, &extras)
        .expect("display list builds");
    docx_layout::clear_measure_fonts();

    let glyph_run = list.pages[0]
        .primitives
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::GlyphRun(run) if run.text.starts_with('@') => Some(run),
            _ => None,
        })
        .expect("registered run shapes");
    let text_run = list.pages[0]
        .primitives
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::Text(run) if run.text == "next" => Some(run),
            _ => None,
        })
        .expect("unregistered run stays browser text");
    let glyph_clip = glyph_run.paint_clip.as_ref().expect("glyph slot clip");
    let text_clip = text_run.paint_clip.as_ref().expect("text slot clip");
    let glyph_clip_right = glyph_clip.x.as_ref().unwrap().as_f64().unwrap()
        + glyph_clip.w.as_ref().unwrap().as_f64().unwrap();
    let shaped_right = glyph_run
        .glyphs
        .iter()
        .map(|glyph| glyph.x + glyph.advance)
        .fold(f64::NEG_INFINITY, f64::max);

    assert!(shaped_right > glyph_clip_right);
    assert!((glyph_clip_right - text_run.x.as_f64().unwrap()).abs() < 0.01);
    assert_eq!(glyph_clip.y, None);
    assert_eq!(glyph_clip.h, None);
    assert_eq!(text_clip.y, None);
    assert_eq!(text_clip.h, None);
}

#[test]
fn conservative_slots_exceed_a_registered_face_control() {
    docx_layout::clear_measure_fonts();
    let fallback_extent = measure_paragraph(mixed_size_paragraph(), 120.0, &fallback_config())
        .expect("fallback measures");
    let fallback_runs = text_runs(&display_list(
        mixed_size_paragraph(),
        120.0,
        &fallback_config(),
    ));

    let font_id = docx_layout::register_measure_font(LIBERATION).expect("font registers");
    let control_config = registered_config("Liberation Sans", font_id);
    let control_extent = measure_paragraph(mixed_size_paragraph(), 120.0, &control_config)
        .expect("control measures");
    let control_runs = text_runs(&display_list(
        mixed_size_paragraph(),
        120.0,
        &control_config,
    ));
    docx_layout::clear_measure_fonts();

    assert_eq!(fallback_extent.lines.len(), 1);
    assert_eq!(control_extent.lines.len(), 1);
    assert!(fallback_extent.lines[0].width > 120.0);
    assert_eq!(fallback_runs.len(), 2);
    assert_eq!(
        control_runs
            .iter()
            .map(|run| run.3.as_str())
            .collect::<String>(),
        "Xiiiiiiiiii"
    );
    let control_x_width = control_runs[0].2;
    let control_i_width: f64 = control_runs[1..].iter().map(|run| run.2).sum();
    assert!(fallback_runs[0].2 + 0.01 >= control_x_width);
    assert!(fallback_runs[1].2 + 0.01 >= control_i_width);
    assert!(fallback_runs[0].1 + control_x_width <= fallback_runs[1].1 + 0.01);
}

#[test]
fn fallback_never_splits_grapheme_clusters() {
    let paragraph = serde_json::json!({
        "kind": "paragraph",
        "id": 0,
        "runs": [{
            "kind": "text",
            "text": "e\u{301}x",
            "fontFamily": "Liberation Sans",
            "fontSize": FONT_SIZE_PT
        }]
    });
    docx_layout::clear_measure_fonts();
    let fallback =
        measure_paragraph(paragraph.clone(), 8.5, &fallback_config()).expect("fallback measures");

    let font_id = docx_layout::register_measure_font(LIBERATION).expect("font registers");
    let control = measure_paragraph(
        paragraph,
        8.5,
        &registered_config("Liberation Sans", font_id),
    )
    .expect("control measures");
    docx_layout::clear_measure_fonts();

    assert_eq!(fallback.lines.len(), 1);
    assert_eq!(
        (fallback.lines[0].head_char, fallback.lines[0].tail_char),
        (0, 3)
    );
    assert_eq!(control.lines.len(), 2);
    assert_eq!(
        (control.lines[0].head_char, control.lines[0].tail_char),
        (0, 2)
    );
    assert_eq!(
        (control.lines[1].head_char, control.lines[1].tail_char),
        (2, 3)
    );

    let zwj = "👩‍👩‍👧‍👧x";
    let zwj_fallback = measure_paragraph(
        serde_json::json!({
            "kind": "paragraph",
            "id": 0,
            "runs": [{ "kind": "text", "text": zwj, "fontSize": FONT_SIZE_PT }]
        }),
        8.5,
        &fallback_config(),
    )
    .expect("ZWJ fallback measures");
    assert_eq!(zwj_fallback.lines.len(), 1);
    assert_eq!(zwj_fallback.lines[0].tail_char, utf16_len(zwj));
}

#[test]
fn bmp_and_astral_scalars_receive_the_same_slot() {
    let paragraph = serde_json::json!({
        "kind": "paragraph",
        "id": 0,
        "runs": [
            {
                "kind": "text",
                "text": "中",
                "pmStart": 1,
                "pmEnd": 2,
                "fontSize": FONT_SIZE_PT
            },
            {
                "kind": "text",
                "text": "𠀀",
                "pmStart": 2,
                "pmEnd": 4,
                "fontSize": FONT_SIZE_PT
            }
        ],
        "pmStart": 0,
        "pmEnd": 5
    });
    let runs = text_runs(&display_list_without_faces(paragraph, 100.0));
    assert_eq!(runs.len(), 2);
    assert!((runs[0].2 - FONT_SIZE_PX).abs() < 0.01);
    assert!((runs[1].2 - FONT_SIZE_PX).abs() < 0.01);
    assert!((runs[1].1 - runs[0].1 - FONT_SIZE_PX).abs() < 0.01);
}

#[test]
fn cjk_fallback_stays_unwrapped_against_a_registered_face_control() {
    let paragraph = cjk_paragraph("中中中中中中中中中中中中");
    docx_layout::clear_measure_fonts();
    let fallback =
        measure_paragraph(paragraph.clone(), 100.0, &fallback_config()).expect("fallback measures");

    let font_id = docx_layout::register_measure_font(NOTO_SANS_SC).expect("font registers");
    let control = measure_paragraph(
        paragraph,
        100.0,
        &registered_config("Noto Sans SC", font_id),
    )
    .expect("control measures");
    docx_layout::clear_measure_fonts();

    assert_eq!(fallback.lines.len(), 1);
    assert!((fallback.lines[0].width - FONT_SIZE_PX * 12.0).abs() < 0.01);
    assert_eq!(control.lines.len(), 2);
    let control_width: f64 = control.lines.iter().map(|line| line.width).sum();
    assert!((fallback.lines[0].width - control_width).abs() < 0.01);
}

#[test]
fn indented_fallback_accepts_overflow_instead_of_guessing_wraps() {
    let paragraph = serde_json::json!({
        "kind": "paragraph",
        "id": 0,
        "runs": [{
            "kind": "text",
            "text": "中中中中中中中中中中中中",
            "fontSize": FONT_SIZE_PT
        }],
        "attrs": {
            "defaultFontSize": FONT_SIZE_PT,
            "indent": { "left": 20.0, "right": 20.0 },
            "listMarker": "•"
        }
    });
    docx_layout::clear_measure_fonts();
    let fallback =
        measure_paragraph(paragraph, 100.0, &fallback_config()).expect("fallback measures");

    assert_eq!(fallback.lines.len(), 1);
    assert!(fallback.lines[0].width > 100.0);
    assert!(fallback.lines[0].width > 60.0);
}
