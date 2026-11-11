//! Link annotation collection and PDF serialization.
//!
//! The crate declares `mod link;` privately, so the integration test compiles
//! the real module (and the shared coordinate helpers it uses) directly.

use docx_layout::display_list::DisplayPage;
use pdf_writer::{Pdf, Ref};
use serde_json::json;

#[allow(dead_code)]
#[path = "../src/units.rs"]
mod units;

/// The only error the coordinate helpers construct; the test crate has no PDF
/// export error type of its own.
#[derive(Debug)]
pub enum PdfError {
    InvalidNumber(&'static str),
}

#[path = "../src/link.rs"]
mod link;

use link::{LinkAction, LinkAnnotation};

const EPSILON: f32 = 1e-3;

#[track_caller]
fn assert_rect(actual: [f32; 4], expected: [f32; 4]) {
    for (actual, expected) in actual.iter().zip(expected) {
        assert!(
            (actual - expected).abs() < EPSILON,
            "rect {actual:?} != expected {expected:?}"
        );
    }
}

fn page() -> DisplayPage {
    serde_json::from_value(json!({
        "pageIndex": 0,
        "width": 612,
        "height": 792,
        "primitives": [
            {
                "kind": "text",
                "text": "External",
                "x": 96,
                "baselineY": 120,
                "width": 100,
                "font": "16px serif",
                "color": "#000000",
                "href": "https://example.com",
                "linkTitle": "Example link",
            },
            {
                "kind": "image",
                "relId": "rIdImage",
                "x": 200,
                "y": 200,
                "w": 100,
                "h": 50,
                "href": "https://images.example/figure",
            },
            {
                "kind": "text",
                "text": "Intro",
                "x": 96,
                "baselineY": 160,
                "width": 48,
                "font": "12px serif",
                "color": "#000000",
                "href": "#intro",
            },
            {
                "kind": "text",
                "text": "Fallback",
                "x": 96,
                "baselineY": 200,
                "width": 64,
                "font": "12px serif",
                "color": "#000000",
                "href": "",
                "linkDocLocation": "https://fallback.example/doc",
            },
            {
                "kind": "glyphRun",
                "fontId": 0,
                "size": 16,
                "color": "#000000",
                "text": "Shaped",
                "glyphs": [
                    {"id": 1, "x": 100, "y": 240, "cluster": 0, "advance": 10},
                    {"id": 2, "x": 110, "y": 240, "cluster": 3, "advance": 12},
                ],
                "href": "#shaped",
            },
            {
                "kind": "text",
                "text": "Hidden",
                "x": 96,
                "baselineY": 280,
                "width": 48,
                "font": "12px serif",
                "color": "#000000",
                "href": "https://hidden.example",
                "hidden": true,
            },
            {
                "kind": "text",
                "text": "Nowhere",
                "x": 96,
                "baselineY": 320,
                "width": 48,
                "font": "12px serif",
                "color": "#000000",
                "href": "",
            },
        ],
        "header": {
            "rId": "rIdHeader",
            "kind": "header",
            "y": 36,
            "height": 36,
            "primitives": [
                {
                    "kind": "text",
                    "text": "Header link",
                    "x": 400,
                    "baselineY": 50,
                    "width": 80,
                    "font": "12px serif",
                    "color": "#000000",
                    "href": "https://header.example",
                },
            ],
        },
        "footer": {
            "rId": "rIdFooter",
            "kind": "footer",
            "y": 720,
            "height": 36,
            "primitives": [
                {
                    "kind": "text",
                    "text": "Footer note",
                    "x": 300,
                    "baselineY": 740,
                    "width": 60,
                    "font": "12px serif",
                    "color": "#000000",
                    "href": "#footer-note",
                },
            ],
        },
        "noteAreas": [
            {
                "primitives": [
                    {
                        "kind": "text",
                        "text": "Note",
                        "x": 96,
                        "baselineY": 700,
                        "width": 48,
                        "font": "10px serif",
                        "color": "#000000",
                        "href": "#note",
                    },
                ],
                "separatorPrimitives": [
                    {
                        "kind": "rect",
                        "x": 96,
                        "y": 680,
                        "w": 200,
                        "h": 1,
                        "fill": "#000000",
                        "href": "https://separator.example",
                    },
                ],
            },
        ],
    }))
    .expect("display page fixture")
}

#[test]
fn collects_body_header_footer_and_note_links() {
    let links = link::collect_links(&page());

    assert_eq!(links.len(), 4, "{links:#?}");
    assert!(links.iter().all(
        |link| !matches!(&link.action, LinkAction::Uri(uri) if uri == "https://hidden.example")
    ));

    let external = links
        .iter()
        .find(|link| link.action == LinkAction::Uri("https://example.com".into()))
        .expect("body text link");
    assert_eq!(external.title.as_deref(), Some("Example link"));
    assert_rect(external.rect, [72.0, 501.0, 147.0, 513.6]);

    let image = links
        .iter()
        .find(|link| link.action == LinkAction::Uri("https://images.example/figure".into()))
        .expect("image link");
    assert_eq!(image.title, None);
    assert_rect(image.rect, [150.0, 406.5, 225.0, 444.0]);

    assert!(
        links
            .iter()
            .all(|link| !matches!(link.action, LinkAction::Uri(ref uri) if uri.starts_with('#')))
    );
    assert!(
        links
            .iter()
            .any(|link| link.action == LinkAction::Uri("https://header.example".into()))
    );
    assert!(
        links
            .iter()
            .any(|link| link.action == LinkAction::Uri("https://separator.example".into()))
    );
}

#[test]
fn writes_link_annotations() {
    let mut pdf = Pdf::new();
    let page = Ref::new(2);

    link::write_annotation(
        &mut pdf,
        Ref::new(1),
        page,
        &LinkAnnotation {
            rect: [10.0, 20.0, 30.0, 40.0],
            action: LinkAction::Uri("https://example.com/a(b".into()),
            title: Some("Example".into()),
        },
    );

    let bytes = pdf.finish();
    let text = String::from_utf8_lossy(&bytes);
    assert!(text.contains("/Type /Annot"), "{text}");
    assert!(text.contains("/Subtype /Link"), "{text}");
    assert!(text.contains("/Rect [10 20 30 40]"), "{text}");
    assert!(text.contains("/P 2 0 R"), "{text}");
    assert!(text.contains("/T (Example)"), "{text}");
    assert!(text.contains("/S /URI"), "{text}");
    assert!(text.contains("/URI (https://example.com/a\\(b)"), "{text}");
    assert!(!text.contains("/S /GoTo"), "{text}");
}
