//! PNG golden coverage for every DOCX primitive class.

use std::path::PathBuf;

use docx_layout::display_list::DisplayList;
use docx_raster::{
    FontChains, ImageMap, ImageScope, RenderResources, measure_text, render_png, scoped_image_key,
};
use ooxml_text::{FontId, FontStore, shape};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("assets/Carlito-Regular.ttf");

fn fixture_resources() -> (FontStore, FontChains, ImageMap) {
    let mut fonts = FontStore::new();
    let id = fonts.register(FONT.to_vec()).expect("font");
    let chains = FontChains::from([
        ("carlito|0|0".to_string(), vec![id]),
        ("carlito|1|0".to_string(), vec![id]),
        ("carlito|0|1".to_string(), vec![id]),
        ("carlito|1|1".to_string(), vec![id]),
    ]);
    let images = ImageMap::from([(scoped_image_key(ImageScope::Body, "rIdImage"), source_png())]);
    (fonts, chains, images)
}

fn list(width: f64, height: f64, primitives: Vec<Value>) -> DisplayList {
    serde_json::from_value(json!({
        "contractVersion": 1,
        "pages": [{
            "pageIndex": 0,
            "width": width,
            "height": height,
            "primitives": primitives
        }]
    }))
    .expect("display list")
}

fn text_scene(fonts: &FontStore, font: FontId) -> DisplayList {
    let width = f64::from(measure_text(fonts, font, "DOCX raster", 18.0).expect("measure"));
    list(
        180.0,
        48.0,
        vec![
            json!({
                "kind": "rect",
                "x": 0,
                "y": 0,
                "w": 180,
                "h": 48,
                "fill": "#eef4ff"
            }),
            json!({
                "kind": "text",
                "text": "DOCX raster",
                "x": 12,
                "baselineY": 31,
                "width": width,
                "font": "400 18px Carlito, sans-serif",
                "color": "#17365d",
                "letterSpacing": 0.25,
                "clipGroup": {
                    "clip": {"x": 8, "y": 5, "w": 164, "h": 38}
                }
            }),
        ],
    )
}

fn glyph_scene(fonts: &FontStore, font: FontId) -> DisplayList {
    let shaped = shape(fonts, font, "Glyph run", 20.0, &[]).expect("shape");
    let mut pen = 12.0_f64;
    let glyphs: Vec<Value> = shaped
        .into_iter()
        .map(|glyph| {
            let placed = json!({
                "id": glyph.glyph_id,
                "x": pen + f64::from(glyph.x_offset),
                "y": 33.0 - f64::from(glyph.y_offset),
                "cluster": glyph.cluster,
                "advance": glyph.x_advance
            });
            pen += f64::from(glyph.x_advance);
            placed
        })
        .collect();
    list(
        180.0,
        50.0,
        vec![json!({
            "kind": "glyphRun",
            "fontId": font.to_u32(),
            "size": 20,
            "color": "#7030a0",
            "text": "Glyph run",
            "glyphs": glyphs
        })],
    )
}

fn table_scene() -> DisplayList {
    let mut primitives = Vec::new();
    for row in 0..2 {
        for column in 0..2 {
            primitives.push(json!({
                "kind": "rect",
                "x": 8 + column * 62,
                "y": 8 + row * 30,
                "w": 62,
                "h": 30,
                "fill": if (row + column) % 2 == 0 { "#d9eaf7" } else { "#fff2cc" }
            }));
        }
    }
    for primitive in [
        json!({"kind":"line","x1":8,"y1":8,"x2":132,"y2":8,"strokeWidth":2,"color":"#1f4e78","borderStyle":"double"}),
        json!({"kind":"line","x1":8,"y1":68,"x2":132,"y2":68,"strokeWidth":2,"color":"#1f4e78","borderStyle":"double"}),
        json!({"kind":"line","x1":8,"y1":8,"x2":8,"y2":68,"strokeWidth":2,"color":"#1f4e78","borderStyle":"solid"}),
        json!({"kind":"line","x1":132,"y1":8,"x2":132,"y2":68,"strokeWidth":2,"color":"#1f4e78","borderStyle":"solid"}),
        json!({"kind":"line","x1":70,"y1":8,"x2":70,"y2":68,"strokeWidth":1,"color":"#4472c4","borderStyle":"dashed"}),
        json!({"kind":"line","x1":8,"y1":38,"x2":132,"y2":38,"strokeWidth":1,"color":"#4472c4","borderStyle":"dotted"}),
    ] {
        primitives.push(primitive);
    }
    list(140.0, 76.0, primitives)
}

fn image_scene() -> DisplayList {
    list(
        110.0,
        76.0,
        vec![
            json!({"kind":"rect","x":0,"y":0,"w":110,"h":76,"fill":"#f2f2f2"}),
            json!({
                "kind": "image",
                "relId": "rIdImage",
                "x": 17,
                "y": 10,
                "w": 76,
                "h": 56,
                "rotationDeg": -5,
                "opacity": 0.9,
                "flipH": true,
                "crop": {"top":0.1,"right":0.12,"bottom":0.05,"left":0.08},
                "border": {"width":2,"color":"#404040","style":"dash"}
            }),
        ],
    )
}

fn shape_scene() -> DisplayList {
    list(
        150.0,
        84.0,
        vec![json!({
            "kind": "shape",
            "x": 12,
            "y": 10,
            "w": 126,
            "h": 64,
            "geometryPath": [
                {"type":"move","x":18,"y":60},
                {"type":"cubic","cp1x":34,"cp1y":12,"cp2x":78,"cp2y":8,"x":96,"y":42},
                {"type":"quad","cpx":112,"cpy":72,"x":132,"y":22},
                {"type":"line","x":132,"y":68},
                {"type":"line","x":18,"y":68},
                {"type":"close"}
            ],
            "fill": "#5b9bd5",
            "fillPaint": {
                "kind": "gradient",
                "gradientType": "linear",
                "angle": 25,
                "stops": [
                    {"position":0,"color":"#9dc3e6"},
                    {"position":1,"color":"#2f5597"}
                ]
            },
            "stroke": {"color":"#203864","width":2,"dash":"dashDot"},
            "transform": {"rotation":3}
        })],
    )
}

fn decoration_scene() -> DisplayList {
    list(
        170.0,
        78.0,
        vec![
            json!({"kind":"decoration","deco":"highlight","x":8,"y":7,"w":154,"h":11,"color":"#fff2cc"}),
            json!({"kind":"decoration","deco":"comment-range","x":8,"y":21,"w":154,"h":11,"color":"rgba(255, 212, 0, 0.25)"}),
            json!({"kind":"decoration","deco":"underline","x":8,"y":38,"w":154,"h":2,"color":"#2e7d32","dashed":true}),
            json!({"kind":"decoration","deco":"strike","x":8,"y":49,"w":154,"h":2,"color":"#c62828","dotted":true}),
            json!({"kind":"decoration","deco":"spell","x":8,"y":64,"w":154,"h":2,"color":"#d32f2f","style":"wave"}),
        ],
    )
}

fn source_png() -> Vec<u8> {
    let mut pixels = Vec::new();
    for y in 0..8 {
        for x in 0..8 {
            let primary = (x + y) % 2 == 0;
            pixels.extend_from_slice(if primary {
                &[68, 114, 196, 255]
            } else {
                &[255, 192, 0, 255]
            });
        }
    }
    let mut data = Vec::new();
    {
        let mut encoder = png::Encoder::new(&mut data, 8, 8);
        encoder.set_color(png::ColorType::Rgba);
        encoder.set_depth(png::BitDepth::Eight);
        encoder.set_compression(png::Compression::Fast);
        let mut writer = encoder.write_header().expect("header");
        writer.write_image_data(&pixels).expect("pixels");
    }
    data
}

fn golden_path(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("golden")
        .join(format!("{name}.png"))
}

fn check(name: &str, display_list: &DisplayList) {
    let (fonts, chains, images) = fixture_resources();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let actual = render_png(display_list, 0, &resources).expect("render");
    let path = golden_path(name);
    if std::env::var("GOLDEN_UPDATE").is_ok() {
        std::fs::create_dir_all(path.parent().expect("parent")).expect("directory");
        std::fs::write(&path, &actual).expect("golden");
        return;
    }
    let expected = std::fs::read(&path).unwrap_or_else(|_| {
        panic!(
            "missing golden {}: regenerate with `GOLDEN_UPDATE=1 cargo test -p betteroffice-docx-raster`",
            path.display()
        )
    });
    assert_eq!(
        actual, expected,
        "golden mismatch for {name}: regenerate deliberately with `GOLDEN_UPDATE=1 cargo test -p betteroffice-docx-raster`"
    );
}

#[test]
fn golden_text() {
    let (fonts, _, _) = fixture_resources();
    check("text", &text_scene(&fonts, FontId::from_u32(0)));
}

#[test]
fn golden_glyph_run() {
    let (fonts, _, _) = fixture_resources();
    check("glyph-run", &glyph_scene(&fonts, FontId::from_u32(0)));
}

#[test]
fn golden_table() {
    check("table", &table_scene());
}

#[test]
fn golden_image() {
    check("image", &image_scene());
}

#[test]
fn golden_shape() {
    check("shape", &shape_scene());
}

#[test]
fn golden_decorations() {
    check("decorations", &decoration_scene());
}

#[test]
fn output_is_byte_deterministic() {
    let (fonts, chains, images) = fixture_resources();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let scene = text_scene(&fonts, FontId::from_u32(0));
    let first = render_png(&scene, 0, &resources).expect("first");
    let second = render_png(&scene, 0, &resources).expect("second");
    assert_eq!(first, second);
}

#[test]
fn unsupported_visual_fields_are_errors() {
    let (fonts, chains, images) = fixture_resources();
    let resources = RenderResources::new(&fonts, &chains, &images);

    let filtered = list(
        20.0,
        20.0,
        vec![
            json!({"kind":"image","relId":"rIdImage","x":0,"y":0,"w":20,"h":20,"filter":"grayscale(1)"}),
        ],
    );
    assert_eq!(
        render_png(&filtered, 0, &resources).unwrap_err(),
        "unsupported image field: filter"
    );

    let shadowed = list(
        40.0,
        20.0,
        vec![json!({
            "kind":"text",
            "text":"x",
            "x":0,
            "baselineY":14,
            "width":8,
            "font":"400 12px Carlito, sans-serif",
            "color":"#000000",
            "textShadow":"shadow"
        })],
    );
    assert_eq!(
        render_png(&shadowed, 0, &resources).unwrap_err(),
        "unsupported text field: textShadow"
    );
}
