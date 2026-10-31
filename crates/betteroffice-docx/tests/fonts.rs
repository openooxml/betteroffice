mod common;

#[cfg(feature = "fonts-latin")]
use betteroffice_docx::fonts::{BundledFonts, FontProvider};
#[cfg(feature = "fonts")]
use betteroffice_docx::fonts::{FontDefinition, FontSource};
use betteroffice_docx::{Document, LayoutInput, OpenOptions, Primitive};
use serde_json::json;

#[cfg(feature = "fonts")]
const APP_FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn options() -> OpenOptions {
    let options = OpenOptions::default();
    #[cfg(feature = "fonts")]
    let options = OpenOptions {
        font_source: FontSource::Backstop,
        ..options
    };
    options
}

fn input() -> LayoutInput {
    let mut value: serde_json::Value = serde_json::from_str(include_str!(
        "../../docx-layout/tests/fixtures/single-page-multi-paragraph.input.json"
    ))
    .unwrap();
    value["measured"].as_array_mut().unwrap().truncate(2);
    for (index, block) in value["measured"]
        .as_array_mut()
        .unwrap()
        .iter_mut()
        .enumerate()
    {
        block["block"]["runs"][0]["fontFamily"] = json!("Calibri");
        block["block"]["runs"][0]["bold"] = json!(index == 1);
    }
    serde_json::from_value(value).unwrap()
}

fn glyphs(document: &Document) -> Vec<Vec<(u32, f64)>> {
    document
        .layout(input())
        .unwrap()
        .display_list
        .pages
        .iter()
        .flat_map(|page| &page.primitives)
        .filter_map(|primitive| match primitive {
            Primitive::GlyphRun(run) => Some(
                run.glyphs
                    .iter()
                    .map(|glyph| (glyph.id, glyph.advance))
                    .collect(),
            ),
            _ => None,
        })
        .collect()
}

#[cfg(feature = "fonts")]
fn wrapping_input() -> LayoutInput {
    let mut value = serde_json::to_value(input()).unwrap();
    value["measured"].as_array_mut().unwrap().truncate(1);
    let text = "A paragraph that wraps across several lines with real font metrics. ".repeat(20);
    value["measured"][0]["block"]["runs"][0]["text"] = json!(text);
    value["measured"][0]["block"]["runs"][0]["pmEnd"] = json!(text.len() + 1);
    value["measured"][0]["block"]["pmEnd"] = json!(text.len() + 2);
    serde_json::from_value(value).unwrap()
}

#[cfg(feature = "fonts")]
fn measure_wrapping_input(font: &[u8]) -> LayoutInput {
    let mut input = wrapping_input();
    let mut store = ooxml_text::FontStore::new();
    let id = store.register_slice(font).unwrap();
    let config = docx_layout::measure_blocks::MeasurementConfig {
        font_chains: [("calibri|0|0".into(), vec![id.to_u32()])].into(),
        defaults: json!({"fontFamily": "Calibri", "fontSize": 11}),
        authoritative_shaping: true,
        ..Default::default()
    };
    docx_layout::with_measure_font_store(&mut store, || {
        let measured = &mut input.measured[0];
        measured.measure =
            docx_layout::measure_blocks::measure_block(&mut measured.block, 624.0, &config)
                .unwrap();
    });
    assert!(!docx_layout::measure_blocks::measured_synthetically(
        input.measured.iter().map(|measured| &measured.measure)
    ));
    input
}

#[cfg(feature = "fonts")]
fn paragraph_line_count(layout: &docx_layout::types::Layout) -> usize {
    layout
        .pages
        .iter()
        .flat_map(|page| &page.fragments)
        .filter_map(|fragment| match fragment {
            docx_layout::types::Fragment::Paragraph(paragraph) => {
                Some(paragraph.to_line - paragraph.from_line)
            }
            _ => None,
        })
        .sum()
}

#[cfg(feature = "fonts-latin")]
#[test]
fn default_layout_wraps_like_native_measurement() {
    let face = BundledFonts::default()
        .resolve_family("Calibri", false, false)
        .unwrap()
        .load()
        .unwrap();
    let mut reference = measure_wrapping_input(&face);
    let expected = docx_layout::compute_layout_input(&mut reference).unwrap();
    let expected_lines = paragraph_line_count(&expected);
    assert!(expected_lines > 1);
    let document = Document::open(&common::sample_docx()).unwrap();
    let actual = document.layout(wrapping_input()).unwrap();
    assert_eq!(paragraph_line_count(&actual.layout), expected_lines);
}

#[cfg(all(target_arch = "wasm32", feature = "fonts-latin"))]
#[test]
fn default_font_layout_is_available_on_wasm() {
    let document = Document::open(&common::sample_docx()).unwrap();
    assert!(!glyphs(&document).is_empty());
}

#[cfg(feature = "fonts")]
#[test]
fn floating_image_preserves_supplied_exclusion_widths() {
    let input: LayoutInput = serde_json::from_value(json!({
        "measured": [{
            "block": {
                "kind": "paragraph", "id": 0, "pmStart": 1, "pmEnd": 27,
                "runs": [
                    {"kind": "image", "src": "rId5", "width": 120, "height": 80,
                     "wrapType": "square", "cssFloat": "left", "pmStart": 1, "pmEnd": 2,
                     "position": {
                         "horizontal": {"relativeTo": "column", "posOffset": 0},
                         "vertical": {"relativeTo": "paragraph", "posOffset": 0}
                     }},
                    {"kind": "text", "text": "wrapped textwrapped text", "fontFamily": "Calibri", "fontSize": 8,
                     "pmStart": 2, "pmEnd": 26}
                ]
            },
            "measure": {"kind": "paragraph", "totalHeight": 48, "lines": [
                {"headRun": 1, "headChar": 0, "tailRun": 1, "tailChar": 12,
                 "width": 76, "ascent": 12, "descent": 4, "lineHeight": 24, "leftOffset": 132},
                {"headRun": 1, "headChar": 12, "tailRun": 1, "tailChar": 24,
                 "width": 76, "ascent": 12, "descent": 4, "lineHeight": 24, "leftOffset": 132}
            ]}
        }],
        "options": {"pageSize": {"w": 400, "h": 400},
                    "margins": {"top": 96, "right": 96, "bottom": 96, "left": 96}}
    }))
    .unwrap();
    let mut document = Document::open_with_options(&common::sample_docx(), &options()).unwrap();
    document
        .register_font("Calibri", false, false, APP_FONT)
        .unwrap();
    let result = document.layout(input).unwrap();
    assert_eq!(paragraph_line_count(&result.layout), 2);
    let runs: Vec<_> = result
        .display_list
        .pages
        .iter()
        .flat_map(|page| &page.primitives)
        .filter_map(|primitive| match primitive {
            Primitive::GlyphRun(run) => Some(run),
            _ => None,
        })
        .collect();
    assert_eq!(runs.len(), 2);
    for run in runs {
        assert!(run.glyphs.first().unwrap().x >= 228.0);
        let end = run.glyphs.last().unwrap();
        assert!(end.x + end.advance <= 304.001);
    }
}

#[cfg(feature = "fonts")]
#[test]
fn floating_image_in_table_cell_preserves_supplied_measurement() {
    let input: LayoutInput = serde_json::from_value(json!({
        "measured": [{
            "block": {"kind": "table", "id": 0, "rows": [{"id": 2, "cells": [{"id": 3, "blocks": [{
                "kind": "paragraph", "id": 1, "pmStart": 2, "pmEnd": 28,
                "runs": [
                    {"kind": "image", "src": "rId5", "width": 120, "height": 80,
                     "wrapType": "square", "cssFloat": "left", "pmStart": 2, "pmEnd": 3,
                     "position": {
                         "horizontal": {"relativeTo": "column", "posOffset": 0},
                         "vertical": {"relativeTo": "paragraph", "posOffset": 0}
                     }},
                    {"kind": "text", "text": "wrapped textwrapped text", "fontFamily": "Calibri", "fontSize": 8,
                     "pmStart": 3, "pmEnd": 27}
                ]
            }]}]}]},
            "measure": {"kind": "table", "columnWidths": [208.0], "totalWidth": 208.0, "totalHeight": 80.0, "rows": [
                {"height": 80.0, "cells": [{"width": 208.0, "height": 80.0, "blocks": [{"kind": "paragraph", "totalHeight": 48, "lines": [
                    {"headRun": 1, "headChar": 0, "tailRun": 1, "tailChar": 12,
                     "width": 76, "ascent": 12, "descent": 4, "lineHeight": 24, "leftOffset": 132},
                    {"headRun": 1, "headChar": 12, "tailRun": 1, "tailChar": 24,
                     "width": 76, "ascent": 12, "descent": 4, "lineHeight": 24, "leftOffset": 132}
                ]}]}]}
            ]}
        }],
        "options": {"pageSize": {"w": 400, "h": 400},
                    "margins": {"top": 96, "right": 96, "bottom": 96, "left": 96}}
    }))
    .unwrap();
    let expected = docx_layout::compute_layout_input(&mut input.clone()).unwrap();
    let mut document = Document::open_with_options(&common::sample_docx(), &options()).unwrap();
    document
        .register_font("Calibri", false, false, APP_FONT)
        .unwrap();
    let result = document.layout(input).unwrap();
    assert_eq!(
        serde_json::to_value(&result.layout).unwrap(),
        serde_json::to_value(&expected).unwrap()
    );
}

#[cfg(feature = "fonts")]
#[test]
fn failed_native_measurement_preserves_supplied_wrapping() {
    let mut input = measure_wrapping_input(APP_FONT);
    let expected = docx_layout::compute_layout_input(&mut input).unwrap();
    let expected_lines = paragraph_line_count(&expected);
    assert!(expected_lines > 1);
    let mut document = Document::open_with_options(&common::sample_docx(), &options()).unwrap();
    document
        .register_font("Unrelated", false, false, APP_FONT)
        .unwrap();
    let actual = document.layout(input).unwrap();
    assert_eq!(paragraph_line_count(&actual.layout), expected_lines);
}

#[cfg(feature = "fonts-latin")]
#[test]
fn default_calibri_uses_regular_and_bold_carlito() {
    let bytes = common::sample_docx();
    let document = Document::open(&bytes).unwrap();
    let mut expected = Document::open_with_options(&bytes, &options()).unwrap();
    for bold in [false, true] {
        let face = BundledFonts::default()
            .resolve_family("Calibri", bold, false)
            .unwrap();
        assert_eq!(face.family, "Carlito");
        expected
            .register_font("Calibri", bold, false, &face.load().unwrap())
            .unwrap();
    }
    let actual = glyphs(&document);
    assert!(!actual.is_empty());
    assert!(actual.iter().all(|glyphs| !glyphs.is_empty()));
    assert_eq!(actual, glyphs(&expected));
}

#[cfg(feature = "fonts")]
#[test]
fn registered_and_option_faces_win_over_the_source() {
    let bytes = common::sample_docx();
    let mut document = Document::open(&bytes).unwrap();
    #[cfg(feature = "fonts-latin")]
    assert!(!glyphs(&document).is_empty());
    let mut expected = Document::open_with_options(&bytes, &options()).unwrap();
    for bold in [false, true] {
        document
            .register_font("Calibri", bold, false, APP_FONT)
            .unwrap();
        expected
            .register_font("Calibri", bold, false, APP_FONT)
            .unwrap();
    }
    assert_eq!(glyphs(&document), glyphs(&expected));
    let from_options = Document::open_with_options(
        &bytes,
        &OpenOptions {
            fonts: [false, true]
                .into_iter()
                .map(|bold| FontDefinition {
                    family: "Calibri".into(),
                    data: APP_FONT.to_vec(),
                    bold,
                    italic: false,
                })
                .collect(),
            ..Default::default()
        },
    )
    .unwrap();
    assert_eq!(glyphs(&from_options), glyphs(&expected));
}

#[cfg(feature = "fonts")]
#[test]
fn embedded_faces_win_over_registered_faces_and_the_source() {
    let bytes = common::sample_docx();
    let mut parts = common::parts_of(&bytes);
    let embedded = include_bytes!("../../docx-raster/tests/assets/Carlito-Regular.ttf");
    parts.extend([
        ("word/fontTable.xml".into(), br#"<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:font w:name="Calibri"><w:embedRegular r:id="face"/><w:embedBold r:id="face"/></w:font></w:fonts>"#.to_vec()),
        ("word/_rels/fontTable.xml.rels".into(), br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="face" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/face.ttf"/></Relationships>"#.to_vec()),
        ("word/fonts/face.ttf".into(), embedded.to_vec()),
    ]);
    let mut document = Document::open(&ooxml_opc::rezip_parts(&parts).unwrap()).unwrap();
    let mut expected = Document::open_with_options(&bytes, &options()).unwrap();
    for bold in [false, true] {
        document
            .register_font("Calibri", bold, false, APP_FONT)
            .unwrap();
        expected
            .register_font("Calibri", bold, false, embedded)
            .unwrap();
    }
    assert_eq!(glyphs(&document), glyphs(&expected));
}

#[test]
fn backstop_preserves_fontless_layout() {
    let document = Document::open_with_options(&common::sample_docx(), &options()).unwrap();
    assert!(glyphs(&document).is_empty());
    assert!(
        document.layout(input()).unwrap().display_list.pages[0]
            .primitives
            .iter()
            .any(|primitive| matches!(primitive, Primitive::Text(_)))
    );
}

#[cfg(not(any(feature = "fonts-latin", feature = "fonts-intl", feature = "fonts-cjk")))]
#[test]
fn open_without_pack_features_preserves_fontless_layout() {
    let document = Document::open(&common::sample_docx()).unwrap();
    assert!(glyphs(&document).is_empty());
}
