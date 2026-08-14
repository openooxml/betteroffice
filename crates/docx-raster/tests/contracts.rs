//! Raster contract and refusal-path tests.

use std::io::{Cursor, Write};

use docx_layout::display_list::DisplayList;
use docx_raster::{
    FontChains, ImageMap, ImageScope, MAX_IMAGE_PIXELS, MAX_PAGE_IMAGE_PIXELS, RenderResources,
    render_page, render_png, scoped_image_key,
};
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
fn synthetic_text_run_clip_intersects_the_outer_mask() {
    let mut fonts = FontStore::new();
    let font = fonts.register(FONT.to_vec()).expect("font");
    let chains = FontChains::from([("carlito|0|0".to_string(), vec![font])]);
    let images = ImageMap::new();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let primitive = |clipped: bool| {
        let mut run = json!({
            "kind": "text",
            "text": "@@@@@@@@@@",
            "x": 4,
            "baselineY": 31,
            "width": 12,
            "font": "400 28px Carlito, sans-serif",
            "color": "#000000"
        });
        if clipped {
            run["paintClip"] = json!({"x": 4, "w": 12});
            run["clipGroup"] = json!({"clip": {"x": 8, "y": 0, "w": 72, "h": 40}});
        }
        run
    };

    let control = render_png(&list(80.0, 40.0, vec![primitive(false)]), 0, &resources)
        .expect("control render");
    let clipped = render_png(&list(80.0, 40.0, vec![primitive(true)]), 0, &resources)
        .expect("clipped render");
    let control_bounds = ink_bounds(&control).expect("control ink");
    let clipped_bounds = ink_bounds(&clipped).expect("clipped ink");

    assert!(control_bounds.0 < 8);
    assert!(control_bounds.1 >= 16);
    assert!(clipped_bounds.0 >= 8);
    assert!(clipped_bounds.1 < 16);
}

#[test]
fn synthetic_glyph_run_clip_intersects_the_outer_mask() {
    let mut fonts = FontStore::new();
    let font = fonts.register(FONT.to_vec()).expect("font");
    let shaped = shape(&fonts, font, "@@@@@@@@@@", 28.0, &[]).expect("shape");
    let mut pen = 4.0_f64;
    let glyphs: Vec<Value> = shaped
        .into_iter()
        .map(|glyph| {
            let placed = json!({
                "id": glyph.glyph_id,
                "x": pen + f64::from(glyph.x_offset),
                "y": 31.0 - f64::from(glyph.y_offset),
                "cluster": glyph.cluster,
                "advance": glyph.x_advance
            });
            pen += f64::from(glyph.x_advance);
            placed
        })
        .collect();
    let primitive = |clipped: bool| {
        let mut run = json!({
            "kind": "glyphRun",
            "fontId": font.to_u32(),
            "size": 28,
            "color": "#000000",
            "text": "@@@@@@@@@@",
            "glyphs": glyphs
        });
        if clipped {
            run["paintClip"] = json!({"x": 4, "w": 12});
            run["clipGroup"] = json!({"clip": {"x": 8, "y": 0, "w": 72, "h": 40}});
        }
        run
    };
    let chains = FontChains::new();
    let images = ImageMap::new();
    let resources = RenderResources::new(&fonts, &chains, &images);

    let control = render_png(&list(80.0, 40.0, vec![primitive(false)]), 0, &resources)
        .expect("control render");
    let clipped = render_png(&list(80.0, 40.0, vec![primitive(true)]), 0, &resources)
        .expect("clipped render");
    let control_bounds = ink_bounds(&control).expect("control ink");
    let clipped_bounds = ink_bounds(&clipped).expect("clipped ink");

    assert!(control_bounds.0 < 8);
    assert!(control_bounds.1 >= 16);
    assert!(clipped_bounds.0 >= 8);
    assert!(clipped_bounds.1 < 16);
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

/// PNG headers declaring a bomb-sized surface, with an IDAT too short to
/// decode. The declared extent is the only thing the budget may consult, so
/// these skip on the cap and never on the truncated stream.
const BOMB_40000_SQUARE: &str = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAnEAAAJxACAIAAADebplSAAAAC0lEQVR4nGNgQAUAABAAATm9j2UAAAAASUVORK5CYII=";
const BOMB_9000_SQUARE: &str = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAIygAACMoCAIAAADit+XtAAAAC0lEQVR4nGNgQAUAABAAATm9j2UAAAAASUVORK5CYII=";

fn image_list(src: &str) -> DisplayList {
    list(
        20.0,
        20.0,
        vec![json!({"kind":"image","relId":src,"x":0,"y":0,"w":20,"h":20})],
    )
}

/// A solid RGBA PNG of any extent, streamed a row at a time so the fixture
/// costs one row rather than the whole surface.
fn solid_png(width: u32, height: u32) -> Vec<u8> {
    let mut bytes = Vec::new();
    {
        let mut encoder = png::Encoder::new(&mut bytes, width, height);
        encoder.set_color(png::ColorType::Rgba);
        encoder.set_depth(png::BitDepth::Eight);
        encoder.set_compression(png::Compression::Fast);
        let mut writer = encoder.write_header().expect("png header");
        let mut stream = writer.stream_writer().expect("png stream");
        let row: Vec<u8> = [0x11_u8, 0x66, 0xcc, 0xff].repeat(width as usize);
        for _ in 0..height {
            stream.write_all(&row).expect("png row");
        }
        stream.finish().expect("png finish");
    }
    bytes
}

fn pixel(png: &[u8], x: u32, y: u32) -> [u8; 4] {
    let mut reader = png::Decoder::new(Cursor::new(png))
        .read_info()
        .expect("png header");
    let mut pixels = vec![0_u8; reader.output_buffer_size().expect("buffer size")];
    let info = reader.next_frame(&mut pixels).expect("png frame");
    let start = ((y * info.width + x) * 4) as usize;
    pixels[start..start + 4].try_into().expect("rgba pixel")
}

fn ink_bounds(png: &[u8]) -> Option<(u32, u32, u32, u32)> {
    let mut reader = png::Decoder::new(Cursor::new(png))
        .read_info()
        .expect("png header");
    let mut pixels = vec![0_u8; reader.output_buffer_size().expect("buffer size")];
    let info = reader.next_frame(&mut pixels).expect("png frame");
    let mut bounds = (u32::MAX, 0, u32::MAX, 0);
    let mut found = false;
    for y in 0..info.height {
        for x in 0..info.width {
            let start = ((y * info.width + x) * 4) as usize;
            if pixels[start..start + 4] != [255, 255, 255, 255] {
                bounds.0 = bounds.0.min(x);
                bounds.1 = bounds.1.max(x);
                bounds.2 = bounds.2.min(y);
                bounds.3 = bounds.3.max(y);
                found = true;
            }
        }
    }
    found.then_some(bounds)
}

/// An image past the pixel budget is skipped, not refused: a 9000x9000 scan is
/// not a bomb, and neither it nor a real bomb may take the page down with it.
#[test]
fn an_image_past_the_pixel_budget_is_skipped_not_an_error() {
    let (fonts, chains, images) = empty_resources();
    let resources = RenderResources::new(&fonts, &chains, &images);
    for reference in [BOMB_40000_SQUARE, BOMB_9000_SQUARE] {
        let png = render_png(&image_list(reference), 0, &resources)
            .unwrap_or_else(|error| panic!("an oversized image failed the page: {error}"));
        assert_eq!(pixel(&png, 10, 10), [255, 255, 255, 255]);
    }
}

/// 20000x100 is 2 Mpx and a few KB on the wire — a banner, not a bomb. Cost is
/// area, so a long side alone must not cost it the page.
#[test]
fn a_banner_wider_than_the_page_still_renders() {
    let (fonts, chains) = (FontStore::new(), FontChains::new());
    let images = ImageMap::from([("\u{1f}rIdBanner".to_string(), solid_png(20_000, 100))]);
    let resources = RenderResources::new(&fonts, &chains, &images);
    let png = render_png(&image_list("rIdBanner"), 0, &resources).expect("banner");
    assert_eq!(pixel(&png, 10, 10), [0x11, 0x66, 0xcc, 255]);
}

/// An unresolvable image reference is skipped, matching the canvas backend's
/// resolver. Empty `src`, an external or dangling relationship, and media
/// outside `word/media/` all reach the renderer this way, and one of them must
/// not blank the page around it.
#[test]
fn an_unresolvable_image_reference_is_skipped_not_an_error() {
    let (fonts, chains, images) = empty_resources();
    let resources = RenderResources::new(&fonts, &chains, &images);
    for reference in ["", "rId9", "http://example.invalid/logo.png"] {
        let list = list(
            20.0,
            20.0,
            vec![
                json!({"kind":"image","relId":reference,"x":0,"y":0,"w":20,"h":20}),
                json!({"kind":"rect","x":0,"y":0,"w":20,"h":20,"fill":"#ff0000"}),
            ],
        );
        let png = render_png(&list, 0, &resources)
            .unwrap_or_else(|error| panic!("`{reference}` failed the page: {error}"));
        assert_eq!(pixel(&png, 10, 10), [255, 0, 0, 255]);
    }
}

/// Two distinct headers, each declaring exactly [`MAX_IMAGE_PIXELS`], with an
/// IDAT too short to decode. Together they claim the whole page budget.
const CAP_8192_BY_4096: &str = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAIAAAABAACAIAAABVq/hcAAAAC0lEQVR4nGNgQAUAABAAATm9j2UAAAAASUVORK5CYII=";
const CAP_4096_BY_8192: &str = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAEAAAACAACAIAAADVohYSAAAAC0lEQVR4nGNgQAUAABAAATm9j2UAAAAASUVORK5CYII=";

/// The page budget is spent across the whole render, so an image that arrives
/// after it is gone is skipped however cheap it is on its own.
#[test]
fn a_page_stops_decoding_once_its_pixel_budget_is_spent() {
    const { assert!(MAX_IMAGE_PIXELS * 2 == MAX_PAGE_IMAGE_PIXELS) };
    let (fonts, chains) = (FontStore::new(), FontChains::new());
    let images = ImageMap::from([("\u{1f}rIdSmall".to_string(), solid_png(8, 8))]);
    let resources = RenderResources::new(&fonts, &chains, &images);
    let list = list(
        20.0,
        20.0,
        vec![
            json!({"kind":"image","relId":CAP_8192_BY_4096,"x":0,"y":0,"w":5,"h":5}),
            json!({"kind":"image","relId":CAP_4096_BY_8192,"x":5,"y":0,"w":5,"h":5}),
            json!({"kind":"image","relId":"rIdSmall","x":0,"y":0,"w":20,"h":20}),
        ],
    );
    let rendered = render_page(&list, 0, &resources).expect("render");
    assert_eq!(rendered.skipped_images, 3);
    assert_eq!(pixel(&rendered.bytes, 10, 10), [255, 255, 255, 255]);
}

/// `ImageMap` shipped documented and keyed by bare relationship id, and a
/// caller that predates part scoping still hands one over. A body image has to
/// keep resolving that way rather than turning into a silent hole.
#[test]
fn a_legacy_unscoped_image_key_still_resolves_in_the_body() {
    let (fonts, chains) = (FontStore::new(), FontChains::new());
    let images = ImageMap::from([("rIdImage".to_string(), solid_png(8, 8))]);
    let resources = RenderResources::new(&fonts, &chains, &images);
    let rendered = render_page(&image_list("rIdImage"), 0, &resources).expect("render");
    assert_eq!(rendered.skipped_images, 0);
    assert_eq!(pixel(&rendered.bytes, 10, 10), [0x11, 0x66, 0xcc, 255]);
}

/// The legacy key is the body's alone. Reaching one from a header is the
/// cross-part lookup scoping exists to prevent, so it stays a hole.
#[test]
fn a_legacy_unscoped_image_key_resolves_in_no_other_part() {
    let (fonts, chains) = (FontStore::new(), FontChains::new());
    let images = ImageMap::from([("rIdImage".to_string(), solid_png(8, 8))]);
    let resources = RenderResources::new(&fonts, &chains, &images);
    let scene: DisplayList = serde_json::from_value(json!({
        "pages": [{
            "pageIndex": 0,
            "width": 20,
            "height": 20,
            "primitives": [],
            "header": {
                "rId": "rId7",
                "kind": "header",
                "y": 0,
                "height": 20,
                "primitives": [
                    {"kind": "image", "relId": "rIdImage", "x": 0, "y": 0, "w": 20, "h": 20}
                ]
            }
        }]
    }))
    .expect("display list");
    let rendered = render_page(&scene, 0, &resources).expect("render");
    assert_eq!(rendered.skipped_images, 1);
    assert_eq!(pixel(&rendered.bytes, 10, 10), [255, 255, 255, 255]);
}

/// Scoping runs both ways. A body relationship id that spells a header's
/// scoped key must not reach the header's bytes through the legacy lookup.
#[test]
fn a_body_relationship_id_cannot_spell_another_parts_scoped_key() {
    let (fonts, chains) = (FontStore::new(), FontChains::new());
    let images = ImageMap::from([(
        scoped_image_key(ImageScope::HeaderFooter("rId7"), "rIdImage"),
        solid_png(8, 8),
    )]);
    let resources = RenderResources::new(&fonts, &chains, &images);
    let rendered =
        render_page(&image_list("hf:rId7\u{1f}rIdImage"), 0, &resources).expect("render");
    assert_eq!(rendered.skipped_images, 1);
    assert_eq!(pixel(&rendered.bytes, 10, 10), [255, 255, 255, 255]);
}

/// A skipped reference leaves a hole in an otherwise successful page, so the
/// count is the only thing a caller can log about it.
#[test]
fn a_render_reports_every_image_reference_it_skipped() {
    let (fonts, chains, images) = empty_resources();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let list = list(
        20.0,
        20.0,
        vec![
            json!({"kind":"image","relId":"rId9","x":0,"y":0,"w":5,"h":5}),
            json!({"kind":"image","relId":BOMB_9000_SQUARE,"x":5,"y":0,"w":5,"h":5}),
            json!({"kind":"image","relId":"data:image/png;base64,%%%%","x":10,"y":0,"w":5,"h":5}),
        ],
    );
    let rendered = render_page(&list, 0, &resources).expect("render");
    assert_eq!(rendered.skipped_images, 3);
    assert_eq!(&rendered.bytes[..8], b"\x89PNG\r\n\x1a\n");
}

/// The canvas resolver skips what will not load as well as what does not
/// resolve, so bytes that cannot be decoded are skipped rather than refused.
#[test]
fn undecodable_image_bytes_are_skipped_like_an_unresolvable_reference() {
    let (fonts, chains, images) = empty_resources();
    let resources = RenderResources::new(&fonts, &chains, &images);
    for reference in [
        "data:image/png;base64,%%%%",
        "data:image/png;base64",
        "data:image/png,notbase64",
        "data:image/png;base64,aGVsbG8=",
    ] {
        let png = render_png(&image_list(reference), 0, &resources)
            .unwrap_or_else(|error| panic!("`{reference}` failed the page: {error}"));
        assert_eq!(pixel(&png, 10, 10), [255, 255, 255, 255]);
    }
}
