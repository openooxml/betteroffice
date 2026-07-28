//! Raster contract and refusal-path tests.

use docx_layout::display_list::DisplayList;
use docx_raster::{FontChains, ImageMap, RenderResources, render_png};
use ooxml_text::{FontStore, shape};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("assets/Carlito-Regular.ttf");

fn list(width: f64, height: f64, primitives: Vec<Value>) -> DisplayList {
    serde_json::from_value(json!({
        "pages": [{
            "pageIndex": 0,
            "width": width,
            "height": height,
            "primitives": primitives
        }]
    }))
    .expect("display list")
}

fn empty_resources() -> (FontStore, FontChains, ImageMap) {
    (FontStore::new(), FontChains::new(), ImageMap::new())
}

#[test]
fn every_line_style_renders() {
    let styles = [
        "solid",
        "double",
        "dotted",
        "dashed",
        "dashDot",
        "dashDotDot",
        "triple",
        "thinThick",
        "thickThin",
        "wave",
        "doubleWave",
        "groove",
        "ridge",
        "inset",
        "outset",
    ];
    let primitives = styles
        .iter()
        .enumerate()
        .map(|(index, style)| {
            let y = 5 + index * 6;
            json!({
                "kind": "line",
                "x1": 4,
                "y1": y,
                "x2": 116,
                "y2": y,
                "strokeWidth": 1.5,
                "color": "#4472c4",
                "secondaryColor": "#9dc3e6",
                "borderStyle": style
            })
        })
        .collect();
    let scene = list(120.0, 96.0, primitives);
    let (fonts, chains, images) = empty_resources();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let png = render_png(&scene, 0, &resources).expect("render");
    assert_eq!(&png[..8], &[0x89, b'P', b'N', b'G', 0x0d, 0x0a, 0x1a, 0x0a]);
}

#[test]
fn glyph_run_uses_positioned_ids_without_reshaping_source_text() {
    let mut fonts = FontStore::new();
    let font = fonts.register(FONT.to_vec()).expect("font");
    let glyph = shape(&fonts, font, "G", 18.0, &[])
        .expect("shape")
        .into_iter()
        .next()
        .expect("glyph");
    let primitive = |text: &str| {
        json!({
            "kind": "glyphRun",
            "fontId": font.to_u32(),
            "size": 18,
            "color": "#000000",
            "text": text,
            "glyphs": [{
                "id": glyph.glyph_id,
                "x": 4,
                "y": 20,
                "cluster": 0,
                "advance": glyph.x_advance
            }]
        })
    };
    let chains = FontChains::new();
    let images = ImageMap::new();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let first = render_png(&list(28.0, 28.0, vec![primitive("G")]), 0, &resources).expect("first");
    let second =
        render_png(&list(28.0, 28.0, vec![primitive("not G")]), 0, &resources).expect("second");
    assert_eq!(first, second);
}

#[test]
fn text_requires_the_layout_font_chain() {
    let mut fonts = FontStore::new();
    fonts.register(FONT.to_vec()).expect("font");
    let chains = FontChains::new();
    let images = ImageMap::new();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let scene = list(
        30.0,
        20.0,
        vec![json!({
            "kind":"text",
            "text":"A",
            "x":2,
            "baselineY":15,
            "width":10,
            "font":"400 12px Carlito, sans-serif",
            "color":"#000000"
        })],
    );
    assert_eq!(
        render_png(&scene, 0, &resources).unwrap_err(),
        "missing font chain for `carlito|0|0`"
    );
}

#[test]
fn unsupported_shape_paint_is_refused() {
    let (fonts, chains, images) = empty_resources();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let scene = list(
        30.0,
        30.0,
        vec![json!({
            "kind":"shape",
            "x":2,
            "y":2,
            "w":26,
            "h":26,
            "geometryPath":[
                {"type":"move","x":2,"y":2},
                {"type":"line","x":28,"y":2},
                {"type":"line","x":28,"y":28},
                {"type":"close"}
            ],
            "fillPaint":{"kind":"pattern","patternPreset":"pct20"}
        })],
    );
    assert_eq!(
        render_png(&scene, 0, &resources).unwrap_err(),
        "unsupported shape fillPaint kind: pattern"
    );
}

#[test]
fn invalid_page_and_page_selector_are_errors() {
    let (fonts, chains, images) = empty_resources();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let scene = list(10.0, 10.0, Vec::new());
    assert_eq!(
        render_png(&scene, 1, &resources).unwrap_err(),
        "page ordinal 1 is out of range"
    );
    let invalid = list(0.0, 10.0, Vec::new());
    assert_eq!(
        render_png(&invalid, 0, &resources).unwrap_err(),
        "page width must be finite and positive"
    );
}

#[test]
fn small_caps_is_refused_instead_of_synthesized() {
    let mut fonts = FontStore::new();
    let font = fonts.register(FONT.to_vec()).expect("font");
    let chains = FontChains::from([("carlito|0|0".to_string(), vec![font])]);
    let images = ImageMap::new();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let scene = list(
        30.0,
        20.0,
        vec![json!({
            "kind":"text",
            "text":"abc",
            "x":2,
            "baselineY":15,
            "width":20,
            "font":"small-caps 400 12px Carlito, sans-serif",
            "color":"#000000",
            "smallCaps":true
        })],
    );
    assert_eq!(
        render_png(&scene, 0, &resources).unwrap_err(),
        "unsupported text field: smallCaps"
    );
}

#[test]
fn unknown_glyph_font_is_an_error() {
    let (fonts, chains, images) = empty_resources();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let scene = list(
        30.0,
        20.0,
        vec![json!({
            "kind":"glyphRun",
            "fontId":42,
            "size":12,
            "color":"#000000",
            "text":"A",
            "glyphs":[{"id":1,"x":2,"y":15,"cluster":0,"advance":8}]
        })],
    );
    assert_eq!(
        render_png(&scene, 0, &resources).unwrap_err(),
        "unknown FontId for this FontStore"
    );
}

/// `leaderGlyphs.font` carries a bare family name, not a CSS shorthand: the
/// producer assigns it from the same value it puts in `fontFamily`.
#[test]
fn a_tab_leader_paints_from_a_bare_family_name() {
    let mut fonts = FontStore::new();
    let font = fonts.register(FONT.to_vec()).expect("font");
    let chains = FontChains::from([("carlito|0|0".to_string(), vec![font])]);
    let images = ImageMap::new();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let list = list(
        200.0,
        40.0,
        vec![json!({
            "kind": "text",
            "text": "Chapter",
            "x": 8,
            "baselineY": 26,
            "width": 60,
            "font": "400 16px Carlito, sans-serif",
            "color": "#000000",
            "leaderGlyphs": {
                "glyph": "\u{00b7}",
                "count": 12,
                "advance": 4,
                "font": "Carlito",
                "size": 16,
                "color": "#000000"
            }
        })],
    );
    let png = render_png(&list, 0, &resources).expect("a bare leader family must render");
    assert_eq!(&png[..8], b"\x89PNG\r\n\x1a\n");
}
