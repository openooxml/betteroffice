//! Link annotation collection and PDF serialization.
//!
//! The crate declares `mod link;` privately, so the integration test compiles
//! the real module (and the shared coordinate helpers it uses) directly.

use std::path::Path;

use docx_layout::display_list::{
    DisplayList, DisplayPage, DocAttrs, GlyphRunPrimitive, PlacedGlyph, Primitive,
    ShapePathCommand, ShapePrimitive, ShapeTransformPrimitive, TextRunPrimitive,
};
use ooxml_text::FontStore;
use pdf_writer::{Pdf, Ref};
use serde_json::{Number, json};
use ttf_parser::Face;

#[allow(dead_code)]
#[path = "../src/units.rs"]
mod units;

/// Mirror of the crate's export error; the test crate has no PDF export error
/// type of its own. `InvalidNumber` backs the coordinate helpers, `Unsupported`
/// backs link-action rejection.
#[derive(Debug)]
pub enum PdfError {
    InvalidNumber(&'static str),
    Unsupported(&'static str),
}

/// Mirror of the crate's conformance profile; the test crate links `link.rs`
/// against this stub.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PdfConformance {
    Standard,
    PdfA2u,
}

impl PdfConformance {
    fn is_pdfa(self) -> bool {
        matches!(self, Self::PdfA2u)
    }
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
        PdfConformance::Standard,
    )
    .expect("a URI link is always representable");

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
    // Standard keeps the historical dictionary: no annotation flags, no
    // appearance stream.
    assert!(!text.contains("/F 4"), "{text}");
    assert!(!text.contains("/AP"), "{text}");
}

#[test]
fn writes_pdfa_link_annotations_with_print_flag_and_no_appearance() {
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
        PdfConformance::PdfA2u,
    )
    .expect("PDF/A-2u represents URI links without an appearance stream");

    let bytes = pdf.finish();
    let text = String::from_utf8_lossy(&bytes);
    // ISO 19005-2 6.3.2: Print set, Hidden/Invisible/NoView/ToggleNoView clear.
    assert!(text.contains("/F 4"), "{text}");
    // ISO 19005-2 6.3.3 exempts Link annotations from /AP.
    assert!(!text.contains("/AP"), "{text}");
    // Geometry, page link and the URI action are unchanged from Standard.
    assert!(text.contains("/Rect [10 20 30 40]"), "{text}");
    assert!(text.contains("/P 2 0 R"), "{text}");
    assert!(text.contains("/T (Example)"), "{text}");
    assert!(text.contains("/S /URI"), "{text}");
    assert!(text.contains("/URI (https://example.com/a\\(b)"), "{text}");
}

#[test]
fn rejects_unsupported_link_actions() {
    let mut pdf = Pdf::new();

    let error = link::write_annotation(
        &mut pdf,
        Ref::new(1),
        Ref::new(2),
        &LinkAnnotation {
            rect: [10.0, 20.0, 30.0, 40.0],
            action: LinkAction::Unsupported("JavaScript"),
            title: None,
        },
        PdfConformance::PdfA2u,
    )
    .expect_err("PDF/A-2u must not emit an unsupported action");

    assert!(
        matches!(error, PdfError::Unsupported("JavaScript")),
        "unexpected error: {error:?}"
    );
}

fn number(value: f64) -> Number {
    Number::from_f64(value).expect("finite display-list number")
}

fn font_bytes() -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../packages/fonts/assets/Roboto-Regular.ttf"),
    )
    .expect("Roboto fixture")
}

fn one_page(primitives: Vec<Primitive>) -> DisplayPage {
    DisplayPage {
        page_index: 0,
        width: number(816.0),
        height: number(1056.0),
        content_bounds: None,
        column_bounds: Vec::new(),
        section_id: None,
        section_index: None,
        section_page_index: None,
        section_page_number: None,
        page_label: None,
        primitives,
        watermark_primitive_count: None,
        background: None,
        page_borders: Vec::new(),
        header: None,
        footer: None,
        note_areas: Vec::new(),
        reservations: Vec::new(),
        table_gridlines: Vec::new(),
        unbuilt: false,
        position_span: None,
        hf_parts: None,
    }
}

fn text_run(
    text: &str,
    x: f64,
    baseline: f64,
    width: f64,
    rotation: Option<f64>,
    horizontal_scale: Option<f64>,
    href: &str,
) -> TextRunPrimitive {
    TextRunPrimitive {
        text: text.to_owned(),
        paint_text: None,
        x: number(x),
        baseline_y: number(baseline),
        width: number(width),
        paint_clip: None,
        font: "16px Roboto".to_owned(),
        scripts: None,
        color: "#000000".to_owned(),
        letter_spacing: None,
        word_spacing: None,
        rtl: None,
        opacity: None,
        rotation_deg: rotation.map(number),
        horizontal_scale: horizontal_scale.map(number),
        substitute_scale: None,
        all_caps: false,
        small_caps: false,
        hidden: false,
        text_shadow: None,
        text_outline: false,
        emphasis_mark: None,
        text_effect: None,
        attrs: DocAttrs {
            href: Some(href.to_owned()),
            ..DocAttrs::default()
        },
    }
}

fn glyph_run(
    font_id: u32,
    face: &Face<'_>,
    text: &str,
    x: f64,
    baseline: f64,
    size: f64,
    rotation: Option<f64>,
    horizontal_scale: Option<f64>,
    href: &str,
) -> GlyphRunPrimitive {
    let units_per_em = f64::from(face.units_per_em());
    let mut cursor = x;
    let mut glyphs = Vec::new();
    for (cluster, character) in text.char_indices() {
        let id = face.glyph_index(character).expect("font covers the run");
        let advance = f64::from(face.glyph_hor_advance(id).expect("horizontal advance"))
            / units_per_em
            * size;
        glyphs.push(PlacedGlyph {
            id: u32::from(id.0),
            x: cursor,
            y: baseline,
            cluster: cluster as u32,
            advance,
            logical_order: None,
            bidi_level: None,
        });
        cursor += advance;
    }
    GlyphRunPrimitive {
        font_id,
        size,
        color: "#000000".to_owned(),
        text: text.to_owned(),
        paint_text: None,
        glyphs,
        echo: None,
        paint_clip: None,
        word_spacing: None,
        rtl: None,
        opacity: None,
        rotation_deg: rotation.map(number),
        horizontal_scale: horizontal_scale.map(number),
        all_caps: false,
        small_caps: false,
        hidden: false,
        text_shadow: None,
        text_outline: false,
        emphasis_mark: None,
        text_effect: None,
        attrs: DocAttrs {
            href: Some(href.to_owned()),
            ..DocAttrs::default()
        },
    }
}

fn text_width(face: &Face<'_>, text: &str, size: f64) -> f64 {
    let units_per_em = f64::from(face.units_per_em());
    text.chars()
        .map(|character| {
            let id = face.glyph_index(character).expect("font covers the text");
            f64::from(face.glyph_hor_advance(id).expect("horizontal advance")) / units_per_em * size
        })
        .sum()
}

#[derive(Debug)]
struct ParsedLink {
    object_ref: i32,
    rect: [f64; 4],
    uri: String,
    page_ref: i32,
}

fn pdf_objects(bytes: &[u8]) -> Vec<(i32, String)> {
    String::from_utf8_lossy(bytes)
        .split("endobj")
        .filter_map(|object| {
            let (reference, _) = object.split_once(" 0 obj")?;
            let object_ref = reference.trim().parse().ok()?;
            Some((object_ref, object.to_owned()))
        })
        .collect()
}

fn parse_link_annotations(bytes: &[u8]) -> (i32, Vec<i32>, Vec<ParsedLink>) {
    let objects = pdf_objects(bytes);
    let mut page_ref = 0;
    let mut page_annotations = Vec::new();
    for (object_ref, object) in &objects {
        if object.contains("/Type /Page") && !object.contains("/Type /Pages") {
            page_ref = *object_ref;
            if let Some(start) = object.find("/Annots [") {
                let refs = &object[start + "/Annots [".len()..];
                let refs = &refs[..refs.find(']').expect("closed annots array")];
                let tokens: Vec<&str> = refs.split_whitespace().collect();
                page_annotations = tokens
                    .windows(3)
                    .filter(|window| window[1] == "0" && window[2] == "R")
                    .map(|window| window[0].parse().expect("annotation ref"))
                    .collect();
            }
        }
    }

    let mut links = Vec::new();
    for (object_ref, object) in &objects {
        if !object.contains("/Subtype /Link") {
            continue;
        }
        let rect_start = object.find("/Rect [").expect("link rect") + "/Rect [".len();
        let rect_end = rect_start + object[rect_start..].find(']').expect("closed rect");
        let values: Vec<f64> = object[rect_start..rect_end]
            .split_whitespace()
            .map(|value| value.parse().expect("rect number"))
            .collect();
        assert_eq!(values.len(), 4, "annotation rect: {object}");
        let uri_start = object.find("/URI (").expect("URI action") + "/URI (".len();
        let uri_end = uri_start + object[uri_start..].find(')').expect("closed URI");
        let page_start = object.find("/P ").expect("page reference") + "/P ".len();
        links.push(ParsedLink {
            object_ref: *object_ref,
            rect: [values[0], values[1], values[2], values[3]],
            uri: object[uri_start..uri_end].to_owned(),
            page_ref: object[page_start..]
                .split_whitespace()
                .next()
                .expect("page ref number")
                .parse()
                .expect("page ref"),
        });
    }
    (page_ref, page_annotations, links)
}

fn rotation_degrees(rotation: Option<&Number>) -> Option<f64> {
    let degrees = rotation.and_then(Number::as_f64)?;
    (degrees.abs() > f64::EPSILON).then_some(degrees)
}

fn rotate_point(x: f64, y: f64, degrees: f64, center: (f64, f64)) -> (f64, f64) {
    let (sin, cos) = degrees.to_radians().sin_cos();
    let dx = x - center.0;
    let dy = y - center.1;
    (
        center.0 + dx * cos - dy * sin,
        center.1 + dx * sin + dy * cos,
    )
}

fn bounds(points: impl IntoIterator<Item = (f64, f64)>) -> [f64; 4] {
    let mut bounds = [
        f64::INFINITY,
        f64::INFINITY,
        f64::NEG_INFINITY,
        f64::NEG_INFINITY,
    ];
    for (x, y) in points {
        bounds[0] = bounds[0].min(x);
        bounds[1] = bounds[1].min(y);
        bounds[2] = bounds[2].max(x);
        bounds[3] = bounds[3].max(y);
    }
    bounds
}

/// CSS-pixel box the painter covers for a fallback text run.
fn painted_text_box(text: &TextRunPrimitive) -> [f64; 4] {
    let x = text.x.as_f64().expect("text x");
    let baseline = text.baseline_y.as_f64().expect("text baseline");
    let width = text.width.as_f64().expect("text width");
    let size = f64::from(units::font_size_px(&text.font));
    let rect = [x, baseline - size * 0.8, x + width, baseline + size * 0.25];
    match rotation_degrees(text.rotation_deg.as_ref()) {
        Some(degrees) => {
            let center = (x + width / 2.0, baseline - size * 0.2);
            bounds(
                [
                    (rect[0], rect[1]),
                    (rect[2], rect[1]),
                    (rect[2], rect[3]),
                    (rect[0], rect[3]),
                ]
                .map(|(px, py)| rotate_point(px, py, degrees, center)),
            )
        }
        None => rect,
    }
}

/// CSS-pixel box the painter covers for a shaped run.
fn painted_glyph_box(run: &GlyphRunPrimitive) -> [f64; 4] {
    let size = run.size;
    let scale = run
        .horizontal_scale
        .as_ref()
        .and_then(Number::as_f64)
        .unwrap_or(100.0)
        / 100.0;
    let left = run
        .glyphs
        .iter()
        .map(|glyph| glyph.x)
        .fold(f64::INFINITY, f64::min);
    let right = run
        .glyphs
        .iter()
        .map(|glyph| glyph.x + glyph.advance)
        .fold(f64::NEG_INFINITY, f64::max);
    let first = run.glyphs.first().expect("non-empty glyph run");
    let pivot = ((left + right) / 2.0, first.y - size * 0.2);
    let degrees = rotation_degrees(run.rotation_deg.as_ref());
    let mut points = Vec::new();
    for glyph in &run.glyphs {
        for (x, y) in [
            (glyph.x, glyph.y - size * 0.8),
            (glyph.x + glyph.advance, glyph.y - size * 0.8),
            (glyph.x + glyph.advance, glyph.y + size * 0.25),
            (glyph.x, glyph.y + size * 0.25),
        ] {
            let scaled_x = left + (x - left) * scale;
            points.push(match degrees {
                Some(degrees) => rotate_point(scaled_x, y, degrees, pivot),
                None => (scaled_x, y),
            });
        }
    }
    bounds(points)
}

/// CSS-pixel box the painter covers for a shape whose path is its frame.
fn painted_shape_box(shape: &ShapePrimitive) -> [f64; 4] {
    let x = shape.x.as_f64().expect("shape x");
    let y = shape.y.as_f64().expect("shape y");
    let w = shape.w.as_f64().expect("shape width");
    let h = shape.h.as_f64().expect("shape height");
    let rect = [x, y, x + w, y + h];
    let rotation = shape
        .transform
        .as_ref()
        .and_then(|transform| transform.rotation.as_ref())
        .and_then(Number::as_f64);
    match rotation.filter(|degrees| degrees.abs() > f64::EPSILON) {
        Some(degrees) => {
            let center = (x + w / 2.0, y + h / 2.0);
            bounds(
                [
                    (rect[0], rect[1]),
                    (rect[2], rect[1]),
                    (rect[2], rect[3]),
                    (rect[0], rect[3]),
                ]
                .map(|(px, py)| rotate_point(px, py, degrees, center)),
            )
        }
        None => rect,
    }
}

fn to_pdf_box(box_px: [f64; 4], page_height: f64) -> [f64; 4] {
    [
        box_px[0] * 0.75,
        (page_height - box_px[3]) * 0.75,
        box_px[2] * 0.75,
        (page_height - box_px[1]) * 0.75,
    ]
}

#[track_caller]
fn assert_contains(annotation: [f64; 4], painted: [f64; 4], label: &str) {
    const SLOP: f64 = 0.02;
    assert!(
        annotation[0] <= painted[0] + SLOP
            && annotation[1] <= painted[1] + SLOP
            && annotation[2] >= painted[2] - SLOP
            && annotation[3] >= painted[3] - SLOP,
        "{label}: annotation {annotation:?} must contain painted primitive {painted:?}"
    );
}

#[test]
fn exported_links_cover_rotated_and_scaled_text_glyph_runs_and_shapes() {
    let bytes = font_bytes();
    let face = Face::parse(&bytes, 0).expect("parse Roboto");
    let mut store = FontStore::default();
    let font_id = store
        .register(bytes.clone())
        .expect("register Roboto")
        .to_u32();

    let rotated_text = text_run(
        "Rotated",
        100.0,
        220.0,
        text_width(&face, "Rotated", 16.0),
        Some(90.0),
        None,
        "https://example.com/rotated-text",
    );
    let scaled_text = text_run(
        "Scaled",
        320.0,
        220.0,
        text_width(&face, "Scaled", 16.0) * 2.0,
        None,
        Some(200.0),
        "https://example.com/scaled-text",
    );
    let rotated_glyphs = glyph_run(
        font_id,
        &face,
        "Hi",
        100.0,
        430.0,
        16.0,
        Some(90.0),
        Some(200.0),
        "https://example.com/rotated-glyphs",
    );
    let scaled_glyphs = glyph_run(
        font_id,
        &face,
        "Hi",
        320.0,
        430.0,
        16.0,
        None,
        Some(150.0),
        "https://example.com/scaled-glyphs",
    );
    let shape = ShapePrimitive {
        x: number(520.0),
        y: number(300.0),
        w: number(120.0),
        h: number(40.0),
        geometry_path: vec![
            ShapePathCommand::Move {
                x: number(520.0),
                y: number(300.0),
            },
            ShapePathCommand::Line {
                x: number(640.0),
                y: number(300.0),
            },
            ShapePathCommand::Line {
                x: number(640.0),
                y: number(340.0),
            },
            ShapePathCommand::Line {
                x: number(520.0),
                y: number(340.0),
            },
            ShapePathCommand::Close,
        ],
        fill: Some("#336699".to_owned()),
        stroke: None,
        transform: Some(ShapeTransformPrimitive {
            rotation: Some(number(30.0)),
            flip_h: false,
            flip_v: false,
        }),
        decorative: false,
        attrs: DocAttrs {
            href: Some("https://example.com/rotated-shape".to_owned()),
            ..DocAttrs::default()
        },
    };

    let display_list = DisplayList {
        contract_version: None,
        pages: vec![one_page(vec![
            Primitive::Text(rotated_text.clone()),
            Primitive::Text(scaled_text.clone()),
            Primitive::GlyphRun(rotated_glyphs.clone()),
            Primitive::GlyphRun(scaled_glyphs.clone()),
            Primitive::Shape(shape.clone()),
        ])],
    };
    let pdf = betteroffice_docx_pdf::export_pdf_with_fonts(&display_list, &store)
        .expect("export a public display list");
    let (page_ref, annots, links) = parse_link_annotations(&pdf);

    assert_eq!(links.len(), 5, "{links:#?}");
    assert_eq!(
        annots.len(),
        5,
        "page must reference every link: {annots:?}"
    );
    assert!(links.iter().all(|link| link.page_ref == page_ref));
    let mut object_refs: Vec<i32> = links.iter().map(|link| link.object_ref).collect();
    object_refs.sort_unstable();
    let mut annot_refs = annots.clone();
    annot_refs.sort_unstable();
    assert_eq!(object_refs, annot_refs);

    for (text, label) in [
        (&rotated_text, "https://example.com/rotated-text"),
        (&scaled_text, "https://example.com/scaled-text"),
    ] {
        let link = links.iter().find(|link| link.uri == label).expect(label);
        assert_contains(link.rect, to_pdf_box(painted_text_box(text), 1056.0), label);
    }
    for (run, label) in [
        (&rotated_glyphs, "https://example.com/rotated-glyphs"),
        (&scaled_glyphs, "https://example.com/scaled-glyphs"),
    ] {
        let link = links.iter().find(|link| link.uri == label).expect(label);
        assert_contains(link.rect, to_pdf_box(painted_glyph_box(run), 1056.0), label);
    }
    let link = links
        .iter()
        .find(|link| link.uri == "https://example.com/rotated-shape")
        .expect("rotated shape link");
    assert_contains(
        link.rect,
        to_pdf_box(painted_shape_box(&shape), 1056.0),
        "rotated shape",
    );
}
