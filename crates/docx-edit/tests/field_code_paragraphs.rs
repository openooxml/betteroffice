//! A paragraph mark inside a complex field's code is part of the code, so
//! Word hides it and joins the paragraph with the next one. Each fixture
//! holds IF fields whose code spans paragraph marks, followed by their
//! results, on A4 at an exact 12pt pitch. The page breaks and lines were
//! exported from Word 16.113.

use docx_edit::structured::{
    Anchor, PageDiagnosticCode, PageExportOptions, RevisionView, StorySelection,
};
use docx_edit::{EditTextView, EngineSession, ReadParagraphsRequest, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn fixture(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests/fixtures/field-code-paragraphs")
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
    json!({
        "bodyStory": "body", "renderEnv": {},
        "regions": {"sections": [{"properties": package.document.final_section_properties}], "settings": package.settings},
        "measurement": {"fontChains": {"arial|0|0": [font]}, "defaults": {"fontFamily": "Arial", "fontSize": 12}}
    })
    .to_string()
}

fn seeded(bytes: &[u8]) -> EngineSession {
    let engine = EngineSession::new(76500);
    seed_from_docx(engine.doc(), bytes).unwrap();
    engine
}

/// Each page's body lines, as text.
fn pages(bytes: &[u8]) -> Vec<Vec<String>> {
    let output: Value = serde_json::from_str(
        &seeded(bytes)
            .layout_document_with_regions_json(&request(bytes))
            .unwrap(),
    )
    .unwrap();
    output["layout"]["pages"]
        .as_array()
        .unwrap()
        .iter()
        .map(|page| {
            page["fragments"]
                .as_array()
                .unwrap()
                .iter()
                .flat_map(|fragment| {
                    fragment["resolvedLines"]
                        .as_array()
                        .cloned()
                        .unwrap_or_default()
                })
                .map(|line| {
                    line["segments"]
                        .as_array()
                        .cloned()
                        .unwrap_or_default()
                        .iter()
                        .filter_map(|segment| segment["text"].as_str().map(str::to_owned))
                        .collect::<String>()
                })
                .collect()
        })
        .collect()
}

fn page_ranges(pages: &[Vec<String>]) -> Vec<(String, String)> {
    pages
        .iter()
        .map(|lines| (lines[0].clone(), lines.last().unwrap().clone()))
        .collect()
}

#[test]
fn a_header_field_code_over_paragraph_marks_takes_one_line() {
    let pages = pages(&fixture("header-field-code-paragraphs"));
    let starts: Vec<_> = page_ranges(&pages)
        .into_iter()
        .map(|(first, _)| first)
        .collect();
    assert_eq!(
        starts,
        [
            "Line 001", "Line 055", "Line 109", "Line 163", "Line 217", "Line 271"
        ]
    );
}

#[test]
fn a_body_field_code_over_paragraph_marks_joins_its_paragraphs() {
    let pages = pages(&fixture("body-field-code-paragraphs"));
    assert!(
        pages[0]
            .iter()
            .any(|line| line.trim_end() == "Line 041 yes")
    );
    let starts: Vec<_> = page_ranges(&pages)
        .into_iter()
        .map(|(first, _)| first)
        .collect();
    assert_eq!(
        starts,
        ["Line 001", "Line 059", "Line 117", "Line 175", "Line 233"]
    );
}

#[test]
fn adjacent_chained_and_nested_field_codes_join_their_paragraphs() {
    for (name, joined) in [
        ("body-field-code-adjacent", "Line 041 yes"),
        ("body-field-code-chained", "Line 041 yes and again"),
        ("body-field-code-nested", "Line 041 yes"),
    ] {
        let pages = pages(&fixture(name));
        assert!(
            pages[0].iter().any(|line| line.trim_end() == joined),
            "{name}: {:?}",
            &pages[0][38..44]
        );
        let starts: Vec<_> = page_ranges(&pages)
            .into_iter()
            .map(|(first, _)| first)
            .collect();
        assert_eq!(
            starts,
            ["Line 001", "Line 059", "Line 117", "Line 175", "Line 233"],
            "{name}"
        );
    }
}

#[test]
fn a_field_code_over_paragraph_marks_in_a_table_cell_joins_its_paragraphs() {
    let pages = pages(&fixture("cell-field-code-paragraphs"));
    let starts: Vec<_> = page_ranges(&pages)
        .into_iter()
        .map(|(first, _)| first)
        .collect();
    assert_eq!(
        starts,
        ["Line 001", "Line 059", "Line 117", "Line 175", "Line 233"]
    );
}

#[test]
fn the_page_map_places_both_joined_paragraphs() {
    let bytes = fixture("body-field-code-paragraphs");
    let engine = seeded(&bytes);
    engine
        .layout_document_with_regions_retained_json(&request(&bytes))
        .unwrap();
    let content = engine
        .export_structured_with_pages(&PageExportOptions {
            stories: Some(vec![StorySelection::Body]),
            ..PageExportOptions::new(RevisionView::Accepted)
        })
        .unwrap()
        .content;
    let paragraphs = engine
        .doc()
        .read_paragraphs(&ReadParagraphsRequest {
            story: None,
            para_ids: None,
            view: EditTextView::Accepted,
        })
        .unwrap()
        .paragraphs;
    let owner = paragraphs
        .iter()
        .position(|paragraph| paragraph.text.starts_with("Line 041"))
        .unwrap();
    for paragraph in [&paragraphs[owner], &paragraphs[owner + 3]] {
        let block = content.structured.stories[0]
            .blocks
            .iter()
            .find(|block| {
                matches!(&block.anchor, Anchor::Paragraph { para_id, .. } if *para_id == paragraph.para_id)
            })
            .unwrap();
        assert!(
            content
                .layout
                .fragments
                .iter()
                .any(|fragment| fragment.node_id == block.id),
            "{:?} has a page fragment",
            paragraph.text
        );
        assert!(
            !content.layout.diagnostics.iter().any(|diagnostic| {
                diagnostic.code == PageDiagnosticCode::NotLaidOut
                    && diagnostic.node_id.as_deref() == Some(block.id.as_str())
            }),
            "{:?} is laid out",
            paragraph.text
        );
    }
}
