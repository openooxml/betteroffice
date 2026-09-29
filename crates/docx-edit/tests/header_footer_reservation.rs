//! Word widens a page's body margins by the header and footer that page
//! shows: the first-page band only on a `w:titlePg` section's first page, the
//! even band only on pages whose displayed number is even under
//! `w:evenAndOddHeaders`, and nothing for a missing first or even band. Each
//! fixture has 280 one-line paragraphs at an exact 12pt pitch on A4, 58 to a
//! page with a one-line band and 53 with an eight-line one. The page breaks
//! were exported from Word 16.113; `continuous-title-page-at-page-break` only
//! checks that an edit repaginates as a fresh layout does.

use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn fixture(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests/fixtures/header-footer-reservation")
            .join(format!("{name}.docx")),
    )
    .unwrap()
}

fn request(bytes: &[u8]) -> String {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    let sections: Vec<_> = package
        .document
        .sections
        .unwrap_or_default()
        .into_iter()
        .map(|section| json!({"properties": section.properties}))
        .collect();
    json!({
        "bodyStory": "body", "renderEnv": {},
        "regions": {"sections": sections, "settings": package.settings},
        "measurement": {"fontChains": {"arial|0|0": [font]}, "defaults": {"fontFamily": "Arial", "fontSize": 12}}
    })
    .to_string()
}

/// The first and last body line on each page.
fn page_lines(bytes: &[u8]) -> Vec<(u32, u32)> {
    let engine = EngineSession::new(76500);
    seed_from_docx(engine.doc(), bytes).unwrap();
    let output: Value = serde_json::from_str(
        &engine
            .layout_document_with_regions_json(&request(bytes))
            .unwrap(),
    )
    .unwrap();
    output["layout"]["pages"]
        .as_array()
        .unwrap()
        .iter()
        .map(|page| {
            let lines: Vec<u32> = page["fragments"]
                .as_array()
                .unwrap()
                .iter()
                .flat_map(|fragment| {
                    fragment["resolvedLines"]
                        .as_array()
                        .cloned()
                        .unwrap_or_default()
                })
                .flat_map(|line| line["segments"].as_array().cloned().unwrap_or_default())
                .filter_map(|segment| {
                    segment["text"]
                        .as_str()?
                        .strip_prefix("Line ")?
                        .parse()
                        .ok()
                })
                .collect();
            (lines[0], *lines.last().unwrap())
        })
        .collect()
}

#[test]
fn a_tall_first_page_footer_shortens_only_the_first_page() {
    assert_eq!(
        page_lines(&fixture("title-page-footer")),
        [(1, 53), (54, 111), (112, 169), (170, 227), (228, 280)]
    );
}

#[test]
fn a_tall_first_page_header_shortens_only_the_first_page() {
    assert_eq!(
        page_lines(&fixture("title-page-header")),
        [(1, 53), (54, 111), (112, 169), (170, 227), (228, 280)]
    );
}

#[test]
fn a_blank_first_page_band_leaves_the_first_page_its_full_height() {
    assert_eq!(
        page_lines(&fixture("title-page-blank-first")),
        [
            (1, 58),
            (59, 111),
            (112, 164),
            (165, 217),
            (218, 270),
            (271, 280)
        ]
    );
}

#[test]
fn an_even_header_reserves_nothing_without_even_and_odd_headers() {
    assert_eq!(
        page_lines(&fixture("even-header-ignored")),
        [(1, 58), (59, 116), (117, 174), (175, 232), (233, 280)]
    );
}

#[test]
fn an_even_header_shortens_only_even_pages_under_even_and_odd_headers() {
    assert_eq!(
        page_lines(&fixture("even-and-odd-headers")),
        [(1, 58), (59, 111), (112, 169), (170, 222), (223, 280)]
    );
}

#[test]
fn a_first_page_footer_reserves_nothing_without_a_title_page() {
    assert_eq!(
        page_lines(&fixture("first-without-title-page")),
        [(1, 58), (59, 116), (117, 174), (175, 232), (233, 280)]
    );
}

#[test]
fn a_missing_even_band_is_blank_under_even_and_odd_headers() {
    assert_eq!(
        page_lines(&fixture("even-and-odd-missing-even")),
        [
            (1, 53),
            (54, 111),
            (112, 164),
            (165, 222),
            (223, 275),
            (276, 280)
        ]
    );
}

#[test]
fn the_displayed_page_number_selects_the_even_band() {
    assert_eq!(
        page_lines(&fixture("even-and-odd-restart-at-two")),
        [
            (1, 53),
            (54, 111),
            (112, 164),
            (165, 222),
            (223, 275),
            (276, 280)
        ]
    );
}

#[test]
fn an_edit_repaginates_with_the_bands_its_pages_show() {
    for (name, position) in [
        ("title-page-footer", 149 * 9),
        ("even-and-odd-headers", 149 * 9),
        ("even-and-odd-restart-at-two", 149 * 9),
        ("continuous-title-page-at-page-break", 39 * 9),
    ] {
        let bytes = fixture(name);
        let request = request(&bytes);
        let engine = EngineSession::new(76501);
        seed_from_docx(engine.doc(), &bytes).unwrap();
        engine
            .layout_document_with_regions_retained_json(&request)
            .unwrap();
        let before = engine.stats();
        engine
            .doc()
            .insert_text(
                &docx_edit::EditCtx::local("", ""),
                docx_edit::Position::new("body", position),
                "x",
                docx_edit::FormatPolicy::Inherit,
            )
            .unwrap();
        let incremental = engine
            .layout_document_with_regions_retained_json(&request)
            .unwrap();
        assert_eq!(
            engine.stats().incremental_pagination_calls - before.incremental_pagination_calls,
            1,
            "{name}"
        );
        let fresh = EngineSession::new(76502);
        fresh
            .doc()
            .apply_update_v1(&engine.doc().encode_state_as_update_v1())
            .unwrap();
        assert_eq!(
            fresh
                .layout_document_with_regions_retained_json(&request)
                .unwrap(),
            incremental,
            "{name}"
        );
    }
}
