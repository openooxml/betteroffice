//! Raster contract and refusal-path tests.

use std::io::{Cursor, Write};

use docx_layout::display_list::DisplayList;
use docx_raster::{
    FontChains, GlyphCache, ImageCache, ImageMap, ImageScope, MAX_IMAGE_PIXELS,
    MAX_PAGE_IMAGE_PIXELS, RenderResources, render_page, render_page_cached, render_png,
    scoped_image_key,
};
use ooxml_text::{FontStore, shape};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("assets/Carlito-Regular.ttf");
const LIBERATION: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

#[path = "../../docx-layout/tests/fixtures/base_latin.rs"]
mod base_latin;

#[test]
fn fontless_paragraph_rasterizes_with_ink() {
    let (fonts, chains, images) = empty_resources();
    let resources = RenderResources::new(&fonts, &chains, &images);
    for bold in [false, true] {
        let scene = base_latin::measured_document(
            &fonts,
            &[ooxml_text::FontId::base_latin(bold, false).to_u32()],
            "Latin Ελληνικά Кириллица ✓ ₹",
            bold,
        );
        assert!(ink_bounds(&render_png(&scene, 0, &resources).unwrap()).is_some());
    }
    assert_eq!(fonts.font_count(), 0);
}

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

/// A chain whose first face lacks Hebrew, so text mixing Latin and Hebrew
/// shapes into one segment per font.
fn fallback_resources() -> (FontStore, FontChains, ImageMap) {
    let mut fonts = FontStore::new();
    let carlito = fonts.register(FONT.to_vec()).expect("carlito");
    let liberation = fonts.register(LIBERATION.to_vec()).expect("liberation");
    assert!(!fonts.covers(carlito, '\u{05d0}').expect("coverage"));
    assert!(fonts.covers(liberation, '\u{05d0}').expect("coverage"));
    let chains = FontChains::from([("carlito|0|0".to_string(), vec![carlito, liberation])]);
    (fonts, chains, ImageMap::new())
}

fn carlito_resources() -> (FontStore, FontChains, ImageMap) {
    let mut fonts = FontStore::new();
    let id = fonts.register(FONT.to_vec()).expect("font");
    let chains = FontChains::from([("carlito|0|0".to_string(), vec![id])]);
    (fonts, chains, ImageMap::new())
}

fn leader_scene(rtl: bool, glyph: &str) -> DisplayList {
    list(
        160.0,
        48.0,
        vec![json!({
            "kind": "text",
            "text": "Tab",
            "x": 10,
            "baselineY": 34,
            "width": 130,
            "font": "400 24px Carlito, sans-serif",
            "color": "#000000",
            "leaderGlyphs": {"glyph": glyph, "count": 3, "advance": 10, "rtl": rtl}
        })],
    )
}

fn single_glyph_scene() -> DisplayList {
    list(
        60.0,
        24.0,
        vec![json!({
            "kind": "text",
            "text": "A",
            "x": 4,
            "baselineY": 16,
            "width": 20,
            "font": "400 12px Carlito, sans-serif",
            "color": "#000000"
        })],
    )
}

#[test]
fn native_nonbreaking_hyphen_paints_the_first_faces_ascii_hyphen() {
    let (fonts, chains, images) = fallback_resources();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let scene = |text| {
        list(
            180.0,
            48.0,
            vec![json!({
                "kind": "text", "text": text,
                "x": 10, "baselineY": 34, "width": 160,
                "font": "400 24px Carlito, sans-serif", "color": "#000000"
            })],
        )
    };
    let expected = render_png(&scene("co-operate"), 0, &resources).unwrap();
    assert!(ink_bounds(&expected).is_some());
    assert_eq!(
        render_png(&scene("co\u{2011}operate"), 0, &resources).unwrap(),
        expected
    );
}

#[test]
fn native_text_ignores_the_canvas_substitute_scale() {
    let mut fonts = FontStore::new();
    let base = fonts.register(LIBERATION.to_vec()).unwrap();
    let view = fonts
        .register_named_substitute(base, "Lucida Bright")
        .unwrap();
    let chains = FontChains::from([("lucida bright|0|0".to_string(), vec![view])]);
    let images = ImageMap::new();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let primitive = json!({
        "kind": "text", "text": "native", "allCaps": true,
        "x": 10, "baselineY": 30, "width": 110,
        "font": "400 16px Lucida Bright, sans-serif", "color": "#000000",
        "horizontalScale": 150
    });
    let control = list(160.0, 48.0, vec![primitive.clone()]);
    let mut browser = primitive;
    browser["substituteScale"] = json!(1.113);
    let browser = list(160.0, 48.0, vec![browser]);
    assert_eq!(
        serde_json::to_value(&browser).unwrap()["pages"][0]["primitives"][0]["substituteScale"],
        json!(1.113)
    );
    let expected = render_png(&control, 0, &resources).unwrap();
    assert!(ink_bounds(&expected).is_some());
    assert_eq!(render_png(&browser, 0, &resources).unwrap(), expected);
}

/// One glyph id, painted from whichever face the run names.
fn glyph_id_scene(font_id: u32, glyph_id: u32) -> DisplayList {
    list(
        60.0,
        24.0,
        vec![json!({
            "kind": "glyphRun",
            "fontId": font_id,
            "size": 18,
            "color": "#000000",
            "text": "A",
            "glyphs": [{"id": glyph_id, "x": 6, "y": 19, "cluster": 0, "advance": 12}]
        })],
    )
}

/// A glyph id past the outline range, so the page fails before the cache holds
/// anything.
fn gradient_scene(kind: &str) -> DisplayList {
    list(
        40.0,
        40.0,
        vec![json!({
            "kind": "shape",
            "x": 0,
            "y": 0,
            "w": 40,
            "h": 40,
            "geometryPath": [
                {"type":"move","x":0,"y":0},
                {"type":"line","x":40,"y":0},
                {"type":"line","x":40,"y":40},
                {"type":"line","x":0,"y":40},
                {"type":"close"}
            ],
            "fillPaint": {
                "kind": "gradient",
                "gradientType": kind,
                "stops": [
                    {"position":0,"color":"#ff0000"},
                    {"position":1,"color":"#0000ff"}
                ]
            }
        })],
    )
}

fn assert_paints_like_radial(kind: &str) {
    let (fonts, chains, images) = empty_resources();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let radial = render_page(&gradient_scene("radial"), 0, &resources).expect("radial");
    let page = render_page(&gradient_scene(kind), 0, &resources).expect(kind);
    assert_eq!(page.bytes, radial.bytes, "{kind}");
    let pixels = image::load_from_memory(&page.bytes)
        .expect("png")
        .into_rgba8();
    let centre = pixels.get_pixel(20, 20).0;
    assert!(centre[0] > 240 && centre[2] < 16, "centre {centre:?}");
    assert_eq!(pixels.get_pixel(0, 0).0, [0, 0, 255, 255]);
}

#[test]
fn a_rectangular_gradient_paints_like_a_radial_gradient() {
    assert_paints_like_radial("rectangular");
}

#[test]
fn a_path_gradient_paints_like_a_radial_gradient() {
    assert_paints_like_radial("path");
}

#[test]
fn an_unknown_gradient_kind_still_fails_the_page() {
    let (fonts, chains, images) = empty_resources();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let error = render_page(&gradient_scene("conic"), 0, &resources).expect_err("unknown kind");
    assert!(
        error.contains("unsupported gradient type: conic"),
        "{error}"
    );
}

fn missing_glyph_scene() -> DisplayList {
    list(
        60.0,
        24.0,
        vec![json!({
            "kind": "glyphRun",
            "fontId": 0,
            "size": 12,
            "color": "#000000",
            "text": "A",
            "glyphs": [{"id": 65_536, "x": 4, "y": 16, "cluster": 0, "advance": 8}]
        })],
    )
}

fn leader_extent(resources: &RenderResources<'_>, rtl: bool, glyph: &str) -> f64 {
    let png = render_png(&leader_scene(rtl, glyph), 0, resources).expect("leader render");
    let bounds = ink_bounds(&png).expect("leader ink");
    f64::from(bounds.1 - bounds.0)
}

/// A leader whose text splits across fallback fonts has to advance the pen
/// through every segment; resetting it at each segment overprints them.
#[test]
fn a_multi_segment_leader_advances_across_fallback_segments() {
    let (fonts, chains, images) = fallback_resources();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let single = leader_extent(&resources, false, "A");
    let mixed = leader_extent(&resources, false, "A\u{05d0}");
    let grew = mixed - single;
    assert!(
        grew > 3.0,
        "the fallback segment must extend the leader past the first segment's ink, grew {grew}"
    );
}

/// The same, reading right-to-left: the segments still lay out sequentially
/// rather than stacking on the stamp origin.
#[test]
fn an_rtl_multi_segment_leader_advances_across_fallback_segments() {
    let (fonts, chains, images) = fallback_resources();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let single = leader_extent(&resources, true, "A");
    let mixed = leader_extent(&resources, true, "A\u{05d0}");
    let grew = mixed - single;
    assert!(
        grew > 3.0,
        "the fallback segment must extend the leader past the first segment's ink, grew {grew}"
    );
}

/// Glyph ids are store-local, so a cache filled by one store must never feed
/// renders of another: it is bound to the store that filled it.
#[test]
fn a_glyph_cache_is_bound_to_the_store_that_filled_it() {
    let mut cache = GlyphCache::default();
    {
        let (fonts, chains, images) = carlito_resources();
        let resources = RenderResources::new(&fonts, &chains, &images);
        render_page_cached(
            &single_glyph_scene(),
            0,
            &resources,
            &mut cache,
            &mut ImageCache::default(),
        )
        .expect("render into an unbound cache");
    }
    let (fonts, chains, images) = carlito_resources();
    let resources = RenderResources::new(&fonts, &chains, &images);
    assert_eq!(
        render_page_cached(
            &single_glyph_scene(),
            0,
            &resources,
            &mut cache,
            &mut ImageCache::default()
        )
        .unwrap_err(),
        "glyph cache is bound to another font store"
    );
}

/// Glyph ids are face-local, so a cache spanning pages has to key on the face
/// too. The same id from a second face must paint that face's outline, not the
/// one an earlier page cached.
#[test]
fn a_shared_cache_keeps_one_glyph_id_apart_per_face() {
    let mut fonts = FontStore::new();
    let carlito = fonts.register(FONT.to_vec()).expect("carlito");
    let liberation = fonts.register(LIBERATION.to_vec()).expect("liberation");
    let chains = FontChains::new();
    let images = ImageMap::new();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let first = glyph_id_scene(carlito.to_u32(), 36);
    let second = glyph_id_scene(liberation.to_u32(), 36);
    let fresh_first = render_page(&first, 0, &resources).expect("first").bytes;
    let fresh_second = render_page(&second, 0, &resources).expect("second").bytes;
    assert_ne!(
        fresh_first, fresh_second,
        "the fixture faces must draw glyph 36 differently"
    );

    let mut cache = GlyphCache::default();
    assert_eq!(
        render_page_cached(
            &first,
            0,
            &resources,
            &mut cache,
            &mut ImageCache::default()
        )
        .expect("shared first")
        .bytes,
        fresh_first
    );
    assert_eq!(
        render_page_cached(
            &second,
            0,
            &resources,
            &mut cache,
            &mut ImageCache::default()
        )
        .expect("shared second")
        .bytes,
        fresh_second
    );
}

/// A page that fails before it caches anything leaves the cache unbound: only
/// an outline the cache actually holds can tie it to the store it came from.
#[test]
fn a_failed_page_leaves_a_glyph_cache_free_to_bind_elsewhere() {
    let mut cache = GlyphCache::default();
    {
        let (fonts, chains, images) = carlito_resources();
        let resources = RenderResources::new(&fonts, &chains, &images);
        assert_eq!(
            render_page_cached(
                &missing_glyph_scene(),
                0,
                &resources,
                &mut cache,
                &mut ImageCache::default()
            )
            .unwrap_err(),
            "glyph id 65536 exceeds the font outline range"
        );
    }
    let (fonts, chains, images) = carlito_resources();
    let resources = RenderResources::new(&fonts, &chains, &images);
    render_page_cached(
        &single_glyph_scene(),
        0,
        &resources,
        &mut cache,
        &mut ImageCache::default(),
    )
    .expect("render into a cache no failed page bound");
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

fn picture_shape(paint: serde_json::Value, fill: Option<&str>) -> serde_json::Value {
    let mut shape = json!({
        "kind":"shape",
        "x":4,
        "y":4,
        "w":20,
        "h":20,
        "geometryPath":[
            {"type":"move","x":4,"y":4},
            {"type":"line","x":24,"y":4},
            {"type":"line","x":24,"y":24},
            {"type":"line","x":4,"y":24},
            {"type":"close"}
        ],
        "fillPaint":paint,
        "stroke":{"color":"#00ff00","width":2}
    });
    if let Some(fill) = fill {
        shape["fill"] = json!(fill);
    }
    shape
}

fn picture_resources(images: &mut ImageMap) -> Vec<u8> {
    let png = solid_png(2, 2);
    images.insert("rIdPicture".to_string(), png.clone());
    png
}

#[test]
fn a_picture_fill_paints_its_picture_inside_the_shape_and_keeps_the_outline() {
    let (fonts, chains, mut images) = empty_resources();
    let source = picture_resources(&mut images);
    let resources = RenderResources::new(&fonts, &chains, &images);
    let scene = list(
        30.0,
        30.0,
        vec![picture_shape(
            json!({"kind":"picture","pictureRelId":"rIdPicture","pictureFillMode":"stretch"}),
            None,
        )],
    );
    let png = render_png(&scene, 0, &resources).unwrap();
    assert_eq!(pixel(&png, 14, 14), pixel(&source, 0, 0));
    assert_eq!(pixel(&png, 28, 28), [255, 255, 255, 255]);
    assert_eq!(pixel(&png, 4, 14)[1], 255);
    assert_ne!(pixel(&png, 4, 14), [255, 255, 255, 255]);
}

#[test]
fn a_stretch_rectangle_keeps_a_picture_fill_inside_it() {
    let (fonts, chains, mut images) = empty_resources();
    let source = picture_resources(&mut images);
    let resources = RenderResources::new(&fonts, &chains, &images);
    let scene = list(
        30.0,
        30.0,
        vec![picture_shape(
            json!({"kind":"picture","pictureRelId":"rIdPicture","pictureFillMode":"stretch",
                "pictureStretchRect":{"left":0.5,"top":0,"right":0,"bottom":0}}),
            None,
        )],
    );
    let png = render_png(&scene, 0, &resources).unwrap();
    assert_eq!(pixel(&png, 20, 14), pixel(&source, 0, 0));
    assert_eq!(pixel(&png, 9, 14), [255, 255, 255, 255]);
}

/// Four 8px quadrants, so a tile shows which way up and round it was drawn.
fn quadrant_png() -> Vec<u8> {
    let mut bytes = Vec::new();
    {
        let mut encoder = png::Encoder::new(&mut bytes, 16, 16);
        encoder.set_color(png::ColorType::Rgba);
        encoder.set_depth(png::BitDepth::Eight);
        let mut writer = encoder.write_header().expect("png header");
        let mut pixels = Vec::new();
        for y in 0..16 {
            for x in 0..16 {
                pixels.extend_from_slice(&QUADRANTS[usize::from(y >= 8)][usize::from(x >= 8)]);
            }
        }
        writer.write_image_data(&pixels).expect("png pixels");
    }
    bytes
}

const QUADRANTS: [[[u8; 4]; 2]; 2] = [
    [[255, 0, 0, 255], [0, 0, 255, 255]],
    [[0, 255, 0, 255], [255, 255, 0, 255]],
];

fn tiled_scene(size: u32, tile: serde_json::Value) -> DisplayList {
    let end = size + 4;
    list(
        f64::from(end + 8),
        f64::from(end + 8),
        vec![json!({
            "kind":"shape",
            "x":4,
            "y":4,
            "w":size,
            "h":size,
            "geometryPath":[
                {"type":"move","x":4,"y":4},
                {"type":"line","x":end,"y":4},
                {"type":"line","x":end,"y":end},
                {"type":"line","x":4,"y":end},
                {"type":"close"}
            ],
            "fillPaint":{"kind":"picture","pictureRelId":"rIdPicture","pictureFillMode":"tile","pictureTile":tile}
        })],
    )
}

#[test]
fn a_tiled_picture_fill_repeats_and_mirrors_alternate_tiles() {
    let (fonts, chains, mut images) = empty_resources();
    images.insert("rIdPicture".to_string(), quadrant_png());
    let resources = RenderResources::new(&fonts, &chains, &images);
    let columns = [(8, 0), (16, 1), (24, 0), (32, 1), (40, 0), (48, 1)];
    let mirrored = [(8, 0), (16, 1), (24, 1), (32, 0), (40, 0), (48, 1)];
    for (flip, across, down) in [
        ("none", columns, columns),
        ("x", mirrored, columns),
        ("y", columns, mirrored),
        ("xy", mirrored, mirrored),
    ] {
        let scene = tiled_scene(
            48,
            json!({"scaleX":1,"scaleY":1,"alignment":"tl","flip":flip}),
        );
        let png = render_png(&scene, 0, &resources).unwrap();
        for (y, row) in down {
            for (x, column) in across {
                assert_eq!(
                    pixel(&png, x, y),
                    QUADRANTS[row][column],
                    "{flip} at {x},{y}"
                );
            }
        }
    }
}

#[test]
fn a_tile_grid_past_the_cap_stretches_the_picture_once() {
    let (fonts, chains, mut images) = empty_resources();
    images.insert("rIdPicture".to_string(), quadrant_png());
    let resources = RenderResources::new(&fonts, &chains, &images);
    let scene = tiled_scene(
        100,
        json!({"scaleX":0.01,"scaleY":0.01,"alignment":"tl","flip":"none"}),
    );
    let png = render_png(&scene, 0, &resources).unwrap();
    assert_eq!(pixel(&png, 20, 20), QUADRANTS[0][0]);
    assert_eq!(pixel(&png, 90, 90), QUADRANTS[1][1]);
}

#[test]
fn a_picture_fill_it_cannot_resolve_keeps_the_solid_fill_and_counts_it() {
    let (fonts, chains, images) = empty_resources();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let scene = list(
        30.0,
        30.0,
        vec![picture_shape(
            json!({"kind":"picture","pictureRelId":"rIdMissing"}),
            Some("#0000ff"),
        )],
    );
    let rendered = render_page(&scene, 0, &resources).unwrap();
    assert_eq!(pixel(&rendered.bytes, 14, 14), [0, 0, 255, 255]);
    assert_eq!(rendered.skipped_images, 1);
}

#[test]
fn a_png_that_ends_without_its_closing_chunk_paints() {
    let (fonts, chains, mut images) = empty_resources();
    let source = solid_png(2, 2);
    let cut = source[..source.len() - 12].to_vec();
    assert_eq!(&source[source.len() - 8..source.len() - 4], b"IEND");
    images.insert("rIdPicture".to_string(), cut);
    let resources = RenderResources::new(&fonts, &chains, &images);
    let rendered = render_page(&image_list("rIdPicture"), 0, &resources).unwrap();
    assert_eq!(rendered.skipped_images, 0);
    assert_eq!(pixel(&rendered.bytes, 10, 10), pixel(&source, 0, 0));
}

fn ellipse_image_scene(extra: Value) -> DisplayList {
    let mut image = json!({"kind":"image","relId":"rIdPicture","x":4,"y":4,"w":40,"h":40,"shapeType":"ellipse"});
    for (key, value) in extra.as_object().expect("object") {
        image[key] = value.clone();
    }
    list(48.0, 48.0, vec![image])
}

#[test]
fn an_ellipse_image_paints_inside_its_ellipse() {
    let (fonts, chains, mut images) = empty_resources();
    let source = solid_png(2, 2);
    images.insert("rIdPicture".to_string(), source.clone());
    let resources = RenderResources::new(&fonts, &chains, &images);
    let png = render_png(&ellipse_image_scene(json!({})), 0, &resources).unwrap();
    assert_eq!(pixel(&png, 24, 24), pixel(&source, 0, 0));
    assert_eq!(pixel(&png, 24, 7), pixel(&source, 0, 0));
    assert_eq!(pixel(&png, 6, 6), [255, 255, 255, 255]);
    assert_eq!(pixel(&png, 41, 41), [255, 255, 255, 255]);
}

#[test]
fn an_ellipse_image_keeps_its_crop_inside_the_ellipse() {
    let (fonts, chains, mut images) = empty_resources();
    let source = solid_png(4, 4);
    images.insert("rIdPicture".to_string(), source.clone());
    let resources = RenderResources::new(&fonts, &chains, &images);
    let scene =
        ellipse_image_scene(json!({"crop":{"left":0.25,"top":0.25,"right":0.25,"bottom":0.25}}));
    let png = render_png(&scene, 0, &resources).unwrap();
    assert_eq!(pixel(&png, 24, 24), pixel(&source, 0, 0));
    assert_eq!(pixel(&png, 6, 6), [255, 255, 255, 255]);
}

#[test]
fn an_ellipse_image_border_follows_the_ellipse() {
    let (fonts, chains, mut images) = empty_resources();
    images.insert("rIdPicture".to_string(), solid_png(2, 2));
    let resources = RenderResources::new(&fonts, &chains, &images);
    let scene = ellipse_image_scene(json!({"border":{"width":2,"color":"#ff0000"}}));
    let png = render_png(&scene, 0, &resources).unwrap();
    assert_eq!(pixel(&png, 24, 4), [255, 0, 0, 255]);
    assert_eq!(pixel(&png, 4, 4), [255, 255, 255, 255]);
    assert_eq!(pixel(&png, 44, 44), [255, 255, 255, 255]);
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

fn carlito_text(extra: Value) -> Value {
    let mut run = json!({
        "kind":"text",
        "text":"Hex",
        "x":4,
        "baselineY":20,
        "width":24,
        "font":"400 16px Carlito, sans-serif",
        "color":"#000000"
    });
    for (key, value) in extra.as_object().expect("fields") {
        run[key] = value.clone();
    }
    run
}

#[test]
fn small_caps_text_runs_are_synthesized_and_drawn() {
    let (fonts, chains, images) = carlito_resources();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let small = render_page(
        &list(
            60.0,
            30.0,
            vec![carlito_text(json!({"text":"abc","smallCaps":true}))],
        ),
        0,
        &resources,
    )
    .expect("small caps");
    let lower = render_page(
        &list(60.0, 30.0, vec![carlito_text(json!({"text":"abc"}))]),
        0,
        &resources,
    )
    .expect("lowercase");
    let upper = render_page(
        &list(60.0, 30.0, vec![carlito_text(json!({"text":"ABC"}))]),
        0,
        &resources,
    )
    .expect("uppercase");
    assert!(small.skipped_effects.is_empty());
    let (small, lower, upper) = (
        ink_bounds(&small.bytes).expect("small ink"),
        ink_bounds(&lower.bytes).expect("lower ink"),
        ink_bounds(&upper.bytes).expect("upper ink"),
    );
    assert!(
        small.1 - small.0 < upper.1 - upper.0,
        "small caps narrower than capitals"
    );
    assert!(
        small.3 - small.2 < upper.3 - upper.2,
        "small caps shorter than capitals"
    );
    assert_ne!(small, lower);
}

#[test]
fn text_effects_this_backend_cannot_draw_are_counted_and_the_text_drawn() {
    let (fonts, chains, images) = carlito_resources();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let plain = render_page(
        &list(40.0, 30.0, vec![carlito_text(json!({}))]),
        0,
        &resources,
    )
    .expect("plain");
    let outlined = render_page(
        &list(
            40.0,
            30.0,
            vec![carlito_text(
                json!({"textOutline":true,"emphasisMark":"dot"}),
            )],
        ),
        0,
        &resources,
    )
    .expect("outlined");
    assert_eq!(outlined.bytes, plain.bytes);
    assert_eq!(
        outlined.skipped_effects,
        std::collections::BTreeMap::from([
            ("text.emphasisMark".to_string(), 1),
            ("text.textOutline".to_string(), 1)
        ])
    );
}

#[test]
fn glyph_runs_draw_the_glyphs_layout_shaped_whatever_their_casing() {
    let (fonts, chains, images) = carlito_resources();
    let font = chains["carlito|0|0"][0];
    let resources = RenderResources::new(&fonts, &chains, &images);
    let glyphs: Vec<Value> = shape(&fonts, font, "Hex", 16.0, &[])
        .expect("shape")
        .iter()
        .scan(4.0_f32, |pen, glyph| {
            let placed = json!({"id":glyph.glyph_id,"x":*pen,"y":20,"cluster":glyph.cluster,"advance":glyph.x_advance});
            *pen += glyph.x_advance;
            Some(placed)
        })
        .collect();
    let run = |extra: Value| {
        let mut run = json!({"kind":"glyphRun","fontId":font.to_u32(),"size":16,"color":"#000000","text":"Hex","glyphs":glyphs});
        for (key, value) in extra.as_object().expect("fields") {
            run[key] = value.clone();
        }
        list(40.0, 30.0, vec![run])
    };
    let plain = render_page(&run(json!({})), 0, &resources).expect("plain");
    for casing in [json!({"allCaps":true}), json!({"smallCaps":true})] {
        let page = render_page(&run(casing), 0, &resources).expect("cased");
        assert_eq!(page.bytes, plain.bytes);
        assert!(page.skipped_effects.is_empty());
    }
    let outlined = render_page(&run(json!({"textOutline":true})), 0, &resources).expect("outlined");
    assert_eq!(outlined.bytes, plain.bytes);
    assert_eq!(
        outlined.skipped_effects,
        std::collections::BTreeMap::from([("glyphRun.textOutline".to_string(), 1)])
    );
}

#[test]
fn a_text_shadow_is_drawn_below_and_right_of_the_text() {
    let (fonts, chains, images) = carlito_resources();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let plain = render_page(
        &list(40.0, 30.0, vec![carlito_text(json!({}))]),
        0,
        &resources,
    )
    .expect("plain");
    let shadowed = render_page(
        &list(
            40.0,
            30.0,
            vec![carlito_text(json!({"textShadow":"shadow"}))],
        ),
        0,
        &resources,
    )
    .expect("shadowed");
    assert!(shadowed.skipped_effects.is_empty());
    let (left, right, top, bottom) = ink_bounds(&plain.bytes).expect("plain ink");
    let mut grey = 0;
    for y in top..=(bottom + 3).min(29) {
        for x in left..=(right + 3).min(39) {
            let (before, after) = (pixel(&plain.bytes, x, y), pixel(&shadowed.bytes, x, y));
            if before == [255, 255, 255, 255] && after != before {
                assert!(
                    after[0] > 100,
                    "the shadow is a light grey, not ink: {after:?}"
                );
                grey += 1;
            }
        }
    }
    assert!(grey > 0, "the shadow paints where the text leaves white");
    let shadow_right = ink_bounds(&shadowed.bytes).expect("shadowed ink").1;
    assert!(shadow_right > right, "the shadow extends right of the text");
}

#[test]
fn a_text_shadow_neither_backend_draws_is_counted() {
    let (fonts, chains, images) = carlito_resources();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let plain = render_page(
        &list(40.0, 30.0, vec![carlito_text(json!({}))]),
        0,
        &resources,
    )
    .expect("plain");
    let embossed = render_page(
        &list(
            40.0,
            30.0,
            vec![carlito_text(json!({"textShadow":"emboss"}))],
        ),
        0,
        &resources,
    )
    .expect("embossed");
    assert_eq!(embossed.bytes, plain.bytes);
    assert_eq!(
        embossed.skipped_effects,
        std::collections::BTreeMap::from([("text.textShadow".to_string(), 1)])
    );
}

#[test]
fn a_shape_shadow_with_an_unbounded_blur_still_renders() {
    let (fonts, chains, images) = empty_resources();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let scene = list(
        40.0,
        40.0,
        vec![json!({
            "kind":"shape","x":4,"y":4,"w":20,"h":20,"fill":"#3366ff",
            "geometryPath":[
                {"type":"move","x":4,"y":4},{"type":"line","x":24,"y":4},
                {"type":"line","x":24,"y":24},{"type":"close"}
            ],
            "effects":[{"kind":"glow","color":"#ff0000","blurRadius":1e20}]
        })],
    );
    assert!(
        render_page(&scene, 0, &resources)
            .expect("glow")
            .skipped_effects
            .is_empty()
    );
}

#[test]
fn a_shape_shadow_is_drawn_and_other_shape_effects_are_counted() {
    let (fonts, chains, images) = empty_resources();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let shape = |effects: Value| {
        list(
            40.0,
            40.0,
            vec![json!({
                "kind":"shape","x":4,"y":4,"w":20,"h":20,"fill":"#3366ff",
                "geometryPath":[
                    {"type":"move","x":4,"y":4},{"type":"line","x":24,"y":4},
                    {"type":"line","x":24,"y":24},{"type":"line","x":4,"y":24},{"type":"close"}
                ],
                "effects": effects
            })],
        )
    };
    let plain = render_page(&shape(json!([])), 0, &resources).expect("plain");
    let shadowed = render_page(
        &shape(json!([
            {"kind":"outerShadow","color":"#000000","blurRadius":2,"distance":6,"direction":45},
            {"kind":"softEdge","radius":3}
        ])),
        0,
        &resources,
    )
    .expect("shadowed shape");
    assert_eq!(
        shadowed.skipped_effects,
        std::collections::BTreeMap::from([("shape.effects".to_string(), 1)])
    );
    assert_eq!(pixel(&plain.bytes, 28, 28), [255, 255, 255, 255]);
    let below_right = pixel(&shadowed.bytes, 28, 28);
    assert!(
        below_right[0] < 255 && below_right[3] == 255,
        "shadow below right: {below_right:?}"
    );
    assert_eq!(pixel(&shadowed.bytes, 14, 14), pixel(&plain.bytes, 14, 14));
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

/// A cache shared across pages has to paint exactly what a fresh cache paints:
/// the second page reuses the outlines the first one extracted.
#[test]
fn a_shared_glyph_cache_renders_every_page_identically() {
    let mut fonts = FontStore::new();
    let font = fonts.register(FONT.to_vec()).expect("font");
    let chains = FontChains::from([("carlito|0|0".to_string(), vec![font])]);
    let images = ImageMap::new();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let run = json!({
        "kind": "text",
        "text": "Page glyph",
        "x": 4,
        "baselineY": 26,
        "width": 60,
        "font": "400 16px Carlito, sans-serif",
        "color": "#000000"
    });
    let list: DisplayList = serde_json::from_value(json!({
        "pages": [
            {"pageIndex": 0, "width": 120.0, "height": 40.0, "primitives": [run.clone()]},
            {"pageIndex": 1, "width": 120.0, "height": 40.0, "primitives": [run]}
        ]
    }))
    .expect("display list");
    let mut cache = GlyphCache::default();
    for page in 0..2 {
        let shared = render_page_cached(
            &list,
            page,
            &resources,
            &mut cache,
            &mut ImageCache::default(),
        )
        .expect("shared");
        let fresh = render_page(&list, page, &resources).expect("fresh");
        assert_eq!(shared.bytes, fresh.bytes);
    }
}

/// An image cache shared across pages has to paint exactly what fresh caches
/// paint: the second page resolves the decode the first one cached.
#[test]
fn a_shared_image_cache_renders_every_page_identically() {
    let (fonts, chains) = (FontStore::new(), FontChains::new());
    let images = ImageMap::from([(
        scoped_image_key(ImageScope::Body, "rIdLogo"),
        solid_png(32, 32),
    )]);
    let resources = RenderResources::new(&fonts, &chains, &images);
    let logo = |y: f64| json!({"kind":"image","relId":"rIdLogo","x":8,"y":y,"w":32,"h":32});
    let list: DisplayList = serde_json::from_value(json!({
        "pages": [
            {"pageIndex": 0, "width": 120.0, "height": 60.0, "primitives": [logo(4.0)]},
            {"pageIndex": 1, "width": 120.0, "height": 60.0, "primitives": [logo(20.0), logo(24.0)]}
        ]
    }))
    .expect("display list");
    let mut glyphs = GlyphCache::default();
    let mut cache = ImageCache::default();
    for page in 0..2 {
        let shared =
            render_page_cached(&list, page, &resources, &mut glyphs, &mut cache).expect("shared");
        let fresh = render_page(&list, page, &resources).expect("fresh");
        assert_eq!(shared.bytes, fresh.bytes);
        assert_eq!(shared.skipped_images, fresh.skipped_images);
    }
}

/// The same bytes under two relationship ids are one decode in one cache: the
/// key is the content, not the reference.
#[test]
fn two_relationship_ids_sharing_bytes_share_the_decode() {
    let (fonts, chains) = (FontStore::new(), FontChains::new());
    let images = ImageMap::from([
        (
            scoped_image_key(ImageScope::Body, "rIdLeft"),
            solid_png(16, 16),
        ),
        (
            scoped_image_key(ImageScope::Body, "rIdRight"),
            solid_png(16, 16),
        ),
    ]);
    let resources = RenderResources::new(&fonts, &chains, &images);
    let scene = list(
        40.0,
        24.0,
        vec![
            json!({"kind":"image","relId":"rIdLeft","x":0,"y":4,"w":16,"h":16}),
            json!({"kind":"image","relId":"rIdRight","x":20,"y":4,"w":16,"h":16}),
        ],
    );
    let rendered = render_page(&scene, 0, &resources).expect("render");
    assert_eq!(rendered.skipped_images, 0);
    assert_eq!(pixel(&rendered.bytes, 8, 12), [0x11, 0x66, 0xcc, 255]);
    assert_eq!(pixel(&rendered.bytes, 28, 12), [0x11, 0x66, 0xcc, 255]);
}

#[test]
fn header_and_footer_drawings_lie_under_the_body() {
    let rect =
        |x: u32, y: u32, fill: &str| json!({"kind":"rect","x":x,"y":y,"w":20,"h":10,"fill":fill});
    let scene: DisplayList = serde_json::from_value(json!({
        "pages": [{
            "pageIndex": 0,
            "width": 60,
            "height": 40,
            "primitives": [rect(0, 10, "#1166cc"), rect(30, 10, "#000000")],
            "header": {"rId": "rIdHeader", "kind": "header", "y": 0, "height": 20,
                "primitives": [rect(0, 10, "#ffffff"), rect(30, 20, "#ffffff")]},
            "footer": {"rId": "rIdFooter", "kind": "footer", "y": 20, "height": 20,
                "primitives": [rect(30, 10, "#ffffff")]}
        }]
    }))
    .expect("display list");
    let (fonts, chains, images) = empty_resources();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let rendered = render_page(&scene, 0, &resources).expect("render");
    assert_eq!(pixel(&rendered.bytes, 10, 15), [0x11, 0x66, 0xcc, 255]);
    assert_eq!(pixel(&rendered.bytes, 40, 15), [0, 0, 0, 255]);
    assert_eq!(pixel(&rendered.bytes, 40, 25), [0xff, 0xff, 0xff, 255]);
}

#[test]
fn watermark_prefix_lies_under_header_footer_and_body_drawings() {
    let rect = |x: u32, y: u32, w: u32, fill: &str| json!({"kind":"rect","x":x,"y":y,"w":w,"h":10,"fill":fill});
    let scene: DisplayList = serde_json::from_value(json!({
        "pages": [{
            "pageIndex": 0, "width": 80, "height": 50,
            "watermarkPrimitiveCount": 2,
            "primitives": [
                rect(0, 10, 80, "#cc0000"), rect(0, 30, 80, "#cc0000"),
                rect(30, 10, 20, "#1166cc")
            ],
            "header": {"rId": "rIdHeader", "kind": "header", "y": 0, "height": 20,
                "primitives": [rect(0, 10, 50, "#ffffff")]},
            "footer": {"rId": "rIdFooter", "kind": "footer", "y": 30, "height": 20,
                "primitives": [rect(0, 30, 80, "#ffffff")]},
            "noteAreas": [{"kind": "footnote", "y": 30, "height": 10,
                "primitives": [rect(30, 30, 20, "#000000")]}]
        }]
    }))
    .expect("display list");
    let (fonts, chains, images) = empty_resources();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let rendered = render_page(&scene, 0, &resources).expect("render");
    assert_eq!(pixel(&rendered.bytes, 10, 15), [0xff, 0xff, 0xff, 255]);
    assert_eq!(pixel(&rendered.bytes, 40, 15), [0x11, 0x66, 0xcc, 255]);
    assert_eq!(pixel(&rendered.bytes, 70, 15), [0xcc, 0, 0, 255]);
    assert_eq!(pixel(&rendered.bytes, 10, 35), [0xff, 0xff, 0xff, 255]);
    assert_eq!(pixel(&rendered.bytes, 40, 35), [0, 0, 0, 255]);
}

/// A refusal written to a shared cache must not spend the next page's decode
/// budget: page two gets its own [`MAX_PAGE_IMAGE_PIXELS`] to spend, no matter
/// what page one refused.
#[test]
fn a_shared_cache_does_not_spend_the_next_pages_budget_on_refusals() {
    let (fonts, chains) = (FontStore::new(), FontChains::new());
    let images = ImageMap::from([("\u{1f}rIdSmall".to_string(), solid_png(8, 8))]);
    let resources = RenderResources::new(&fonts, &chains, &images);
    let pages: DisplayList = serde_json::from_value(json!({
        "pages": [
            {"pageIndex": 0, "width": 20.0, "height": 20.0,
             "primitives": [{"kind":"image","relId":BOMB_40000_SQUARE,"x":0,"y":0,"w":5,"h":5}]},
            {"pageIndex": 1, "width": 20.0, "height": 20.0,
             "primitives": [
                {"kind":"image","relId":BOMB_40000_SQUARE,"x":0,"y":0,"w":5,"h":5},
                {"kind":"image","relId":CAP_8192_BY_4096,"x":0,"y":0,"w":5,"h":5},
                {"kind":"image","relId":CAP_4096_BY_8192,"x":5,"y":0,"w":5,"h":5},
                {"kind":"image","relId":"rIdSmall","x":0,"y":0,"w":20,"h":20}
             ]}
        ]
    }))
    .expect("display list");
    let mut glyphs = GlyphCache::default();
    let mut cache = ImageCache::default();
    for page in 0..2 {
        let shared =
            render_page_cached(&pages, page, &resources, &mut glyphs, &mut cache).expect("shared");
        let fresh = render_page(&pages, page, &resources).expect("fresh");
        assert_eq!(shared.bytes, fresh.bytes, "page {page}");
        assert_eq!(shared.skipped_images, fresh.skipped_images, "page {page}");
    }
}

/// PNG headers declaring a bomb-sized surface, with an IDAT too short to
/// decode. The larger is past what a reduced decode may read and skips on its
/// header; the smaller reaches the reduced decode and skips on its stream.
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

#[test]
fn text_named_for_the_backstop_family_rasterizes_without_a_host_chain() {
    let (fonts, chains, images) = empty_resources();
    let resources = RenderResources::new(&fonts, &chains, &images);
    for font in [
        format!("400 24px {}, sans-serif", ooxml_text::BASE_LATIN_FAMILY),
        format!("700 24px {}, sans-serif", ooxml_text::BASE_LATIN_FAMILY),
    ] {
        let scene = list(
            80.0,
            40.0,
            vec![json!({
                "kind": "text",
                "text": "1.",
                "x": 4,
                "baselineY": 30,
                "width": 40,
                "font": font,
                "color": "#000000"
            })],
        );
        assert!(ink_bounds(&render_png(&scene, 0, &resources).unwrap()).is_some());
    }
}

#[test]
fn a_leader_font_shorthand_resolves_its_family() {
    let (fonts, chains, images) = empty_resources();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let font = format!("400 24px {}, sans-serif", ooxml_text::BASE_LATIN_FAMILY);
    let scene = list(
        160.0,
        48.0,
        vec![json!({
            "kind": "text",
            "text": "Tab",
            "x": 10,
            "baselineY": 34,
            "width": 130,
            "font": font,
            "color": "#000000",
            "leaderGlyphs": {"glyph": ".", "count": 3, "advance": 10, "font": font}
        })],
    );
    assert!(ink_bounds(&render_png(&scene, 0, &resources).unwrap()).is_some());
}

#[test]
fn scale_transforms_geometry_and_clip_masks() {
    let (fonts, chains, images) = empty_resources();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let scene = list(
        64.0,
        32.0,
        vec![json!({
            "kind": "rect", "x": 4, "y": 4, "w": 20, "h": 16, "fill": "#000000",
            "clipGroup": {"clip": {"x": 8, "y": 6, "w": 8, "h": 8}}
        })],
    );
    let first = render_page(&scene, 0, &resources).unwrap();
    let second = docx_raster::render_page_cached_scaled(
        &scene,
        0,
        &resources,
        &mut GlyphCache::default(),
        &mut ImageCache::default(),
        2.0,
    )
    .unwrap();
    assert_eq!(ink_bounds(&first.bytes), Some((8, 15, 6, 13)));
    assert_eq!(ink_bounds(&second.bytes), Some((16, 31, 12, 27)));
    assert_eq!(
        u32::from_be_bytes(second.bytes[16..20].try_into().unwrap()),
        128
    );
    assert_eq!(
        u32::from_be_bytes(second.bytes[20..24].try_into().unwrap()),
        64
    );
}

#[test]
fn scale_transforms_text_paint_clips() {
    let (fonts, chains, images) = carlito_resources();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let scene = list(
        80.0,
        40.0,
        vec![json!({
            "kind": "text", "text": "@@@@@@@@@@", "x": 4, "baselineY": 31,
            "width": 12, "font": "400 28px Carlito, sans-serif", "color": "#000000",
            "paintClip": {"x": 8, "w": 8},
            "clipGroup": {"clip": {"x": 0, "y": 0, "w": 80, "h": 40}}
        })],
    );
    let second = docx_raster::render_page_cached_scaled(
        &scene,
        0,
        &resources,
        &mut GlyphCache::default(),
        &mut ImageCache::default(),
        2.0,
    )
    .unwrap();
    let (left, right, _, _) = ink_bounds(&second.bytes).unwrap();
    assert!(left >= 16);
    assert!(right < 32);
}

#[test]
fn scale_validates_surface_limits_before_painting() {
    let (fonts, chains, images) = empty_resources();
    let resources = RenderResources::new(&fonts, &chains, &images);
    for scene in [list(10_000.0, 1.0, vec![]), list(2049.0, 2049.0, vec![])] {
        assert!(
            docx_raster::render_page_cached_scaled(
                &scene,
                0,
                &resources,
                &mut GlyphCache::default(),
                &mut ImageCache::default(),
                2.0,
            )
            .is_err()
        );
    }
    for scale in [0.0, -1.0, f32::NAN, f32::INFINITY] {
        assert!(
            docx_raster::render_page_cached_scaled(
                &list(64.0, 32.0, vec![]),
                0,
                &resources,
                &mut GlyphCache::default(),
                &mut ImageCache::default(),
                scale,
            )
            .is_err()
        );
    }
}

fn vector_picture_scene() -> DisplayList {
    list(
        48.0,
        48.0,
        vec![json!({"kind":"image","relId":"rIdVector","x":4,"y":4,"w":40,"h":40})],
    )
}

#[test]
fn an_svg_picture_paints() {
    let (fonts, chains, mut images) = empty_resources();
    images.insert(
        "rIdVector".to_string(),
        br##"<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10"><rect width="10" height="10" fill="#ff0000"/></svg>"##.to_vec(),
    );
    let resources = RenderResources::new(&fonts, &chains, &images);
    let rendered = render_page(&vector_picture_scene(), 0, &resources).unwrap();
    assert_eq!(rendered.skipped_images, 0);
    assert_eq!(pixel(&rendered.bytes, 24, 24), [255, 0, 0, 255]);
    assert_eq!(pixel(&rendered.bytes, 2, 2), [255, 255, 255, 255]);
}

#[test]
fn emf_and_wmf_pictures_paint() {
    for metafile in [
        &include_bytes!("../../ooxml-metafile/tests/fixtures/shapes.emf")[..],
        &include_bytes!("../../ooxml-metafile/tests/fixtures/shapes.wmf")[..],
    ] {
        let (fonts, chains, mut images) = empty_resources();
        images.insert("rIdVector".to_string(), metafile.to_vec());
        let resources = RenderResources::new(&fonts, &chains, &images);
        let rendered = render_page(&vector_picture_scene(), 0, &resources).unwrap();
        assert_eq!(rendered.skipped_images, 0);
        let painted = (4..44)
            .flat_map(|y| (4..44).map(move |x| (x, y)))
            .filter(|&(x, y)| pixel(&rendered.bytes, x, y) != [255, 255, 255, 255])
            .count();
        assert!(painted > 40, "{painted} painted pixels");
    }
}

#[test]
fn a_picture_this_backend_cannot_decode_is_counted_as_skipped() {
    let (fonts, chains, mut images) = empty_resources();
    images.insert("rIdVector".to_string(), b"not a picture".to_vec());
    let resources = RenderResources::new(&fonts, &chains, &images);
    let rendered = render_page(&vector_picture_scene(), 0, &resources).unwrap();
    assert_eq!(rendered.skipped_images, 1);
}

#[test]
fn an_svg_past_the_picture_limits_is_refused() {
    let (fonts, chains, mut images) = empty_resources();
    images.insert(
        "rIdVector".to_string(),
        br##"<svg xmlns="http://www.w3.org/2000/svg" width="8191" height="8191" viewBox="0 0 8191 8191"><rect width="8191" height="8191" fill="#ff0000"/></svg>"##.to_vec(),
    );
    let resources = RenderResources::new(&fonts, &chains, &images);
    let rendered = render_page(&vector_picture_scene(), 0, &resources).unwrap();
    assert_eq!(rendered.skipped_images, 1);
}

#[test]
fn an_svg_picture_fill_tiles_at_its_own_size() {
    let (fonts, chains, mut images) = empty_resources();
    images.insert(
        "rIdPicture".to_string(),
        br##"<svg xmlns="http://www.w3.org/2000/svg" width="8" height="8" viewBox="0 0 8 8"><rect width="4" height="8" fill="#ff0000"/><rect x="4" width="4" height="8" fill="#0000ff"/></svg>"##.to_vec(),
    );
    let resources = RenderResources::new(&fonts, &chains, &images);
    let scene = list(
        30.0,
        30.0,
        vec![picture_shape(
            json!({"kind":"picture","pictureRelId":"rIdPicture","pictureFillMode":"tile","pictureTile":{"alignment":"tl"}}),
            None,
        )],
    );
    let png = render_png(&scene, 0, &resources).unwrap();
    let red = pixel(&png, 6, 14);
    let blue = pixel(&png, 10, 14);
    let red_again = pixel(&png, 14, 14);
    assert!(red[0] > 200 && red[2] < 80, "{red:?}");
    assert!(blue[2] > 200 && blue[0] < 80, "{blue:?}");
    assert!(red_again[0] > 200 && red_again[2] < 80, "{red_again:?}");
}
