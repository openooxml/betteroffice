use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn fixture(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests/fixtures/header-float-wrap")
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

fn layout(bytes: &[u8]) -> Value {
    let engine = EngineSession::new(76500);
    seed_from_docx(engine.doc(), bytes).unwrap();
    serde_json::from_str(
        &engine
            .layout_document_with_regions_json(&request(bytes))
            .unwrap(),
    )
    .unwrap()
}

/// The first and last body line on each page.
fn page_lines(bytes: &[u8]) -> Vec<(u32, u32)> {
    let output = layout(bytes);
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

fn page_starts(bytes: &[u8]) -> Vec<u32> {
    page_lines(bytes)
        .into_iter()
        .map(|(first, _)| first)
        .collect()
}

fn body_boxes(bytes: &[u8]) -> Vec<Vec<(f64, f64)>> {
    layout(bytes)["layout"]["pages"]
        .as_array()
        .unwrap()
        .iter()
        .map(|page| {
            page["fragments"]
                .as_array()
                .unwrap()
                .iter()
                .filter(|fragment| fragment["resolvedLines"].is_array())
                .map(|fragment| {
                    let top = fragment["y"].as_f64().unwrap() * 0.75;
                    let bottom = top + fragment["height"].as_f64().unwrap() * 0.75;
                    (top, bottom)
                })
                .collect()
        })
        .collect()
}

#[test]
fn body_text_clears_a_top_and_bottom_header_float() {
    assert_eq!(
        page_starts(&fixture("header-float-top-and-bottom")),
        [1, 50, 99, 148, 197, 246]
    );
    for page in body_boxes(&fixture("header-float-top-and-bottom")) {
        assert!((page[0].0 - 171.4).abs() < 0.1);
        assert!(page.iter().all(|&(top, _)| top >= 171.4 - 0.1));
    }
}

#[test]
fn body_text_clears_a_full_width_square_header_float() {
    assert_eq!(
        page_starts(&fixture("header-float-square-full-width")),
        [1, 50, 99, 148, 197, 246]
    );
}

#[test]
fn a_header_float_without_wrapping_leaves_body_flow_unchanged() {
    assert_eq!(
        page_starts(&fixture("header-float-wrap-none")),
        [1, 59, 117, 175, 233]
    );
}

#[test]
fn a_header_float_behind_text_leaves_body_flow_unchanged() {
    assert_eq!(
        page_starts(&fixture("header-float-behind-text")),
        [1, 59, 117, 175, 233]
    );
}

#[test]
fn a_narrow_square_header_float_leaves_body_flow_unchanged() {
    assert_eq!(
        page_starts(&fixture("header-float-narrow-square")),
        [1, 59, 117, 175, 233]
    );
}

#[test]
fn body_text_clears_a_first_page_header_float_only_on_page_one() {
    assert_eq!(
        page_starts(&fixture("first-page-header-float")),
        [1, 50, 108, 166, 224]
    );
}

#[test]
fn body_text_clears_a_page_relative_header_float() {
    assert_eq!(
        page_starts(&fixture("header-float-page-relative")),
        [1, 51, 101, 151, 201, 251]
    );
    for page in body_boxes(&fixture("header-float-page-relative")) {
        assert!((page[0].0 - 160.0).abs() < 0.1);
    }
}

#[test]
fn body_text_flows_above_and_below_a_top_and_bottom_footer_float() {
    assert_eq!(
        page_starts(&fixture("footer-float-top-and-bottom")),
        [1, 50, 99, 148, 197, 246]
    );
    let pages = body_boxes(&fixture("footer-float-top-and-bottom"));
    for page in &pages[..pages.len() - 1] {
        assert!((page[0].0 - 72.0).abs() < 0.1);
        assert!((page.last().unwrap().0 - 750.25).abs() < 0.1);
        assert!(
            page.iter()
                .all(|&(top, bottom)| bottom <= 650.25 + 0.1 || top >= 750.25 - 0.1),
            "{page:?}"
        );
    }
}

#[test]
fn an_edit_repaginates_with_the_float_bands_its_pages_show() {
    for name in ["header-float-top-and-bottom", "footer-float-top-and-bottom"] {
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
                docx_edit::Position::new("body", 149 * 9),
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
