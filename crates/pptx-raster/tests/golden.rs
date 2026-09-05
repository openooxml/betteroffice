//! png golden harness: scenarios are byte-compared against committed pngs.
//! regenerate deliberately with `GOLDEN_UPDATE=1 cargo test -p betteroffice-pptx-raster`.

use std::collections::BTreeMap;
use std::path::PathBuf;
use std::sync::LazyLock;

use ooxml_text::{FontId, FontStore};
use pptx_raster::{AssetMap, Background, RenderOptions, RenderResources, render_slide};
use pptx_render::{
    CONTRACT_VERSION, CaretStop, GradientStop, GradientType, Paint, PositionedGlyph,
    PositionedTextLine, PositionedTextRun, Primitive, Stroke, SurfaceDisplayList, TextAnchor,
    TextParagraph, Transform,
};

const CARLITO: &[u8] = include_bytes!("assets/Carlito-Regular.ttf");

fn golden_path(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("golden")
        .join(format!("{name}.png"))
}

fn check(name: &str, list: &SurfaceDisplayList) {
    let (fonts, font) = font_store();
    let images = assets();
    let resources = RenderResources::new(&fonts, &images).with_label_font(Some(font));
    let rendered = render_slide(list, &resources, &RenderOptions::default()).expect("render");
    assert_eq!(
        rendered.skipped_images, 0,
        "{name} skipped an image the golden expects painted"
    );
    let actual = rendered.bytes;
    let path = golden_path(name);
    if std::env::var("GOLDEN_UPDATE").is_ok() {
        std::fs::write(&path, &actual).expect("write golden");
        return;
    }
    let expected = std::fs::read(&path).unwrap_or_else(|_| {
        panic!(
            "missing golden {}: regenerate with `GOLDEN_UPDATE=1 cargo test -p betteroffice-pptx-raster`",
            path.display()
        )
    });
    assert!(
        actual == expected,
        "golden mismatch for {name}: if intended, regenerate with `GOLDEN_UPDATE=1 cargo test -p betteroffice-pptx-raster`"
    );
}

fn font_store() -> (FontStore, FontId) {
    let mut fonts = FontStore::new();
    let id = fonts.register(CARLITO.to_vec()).expect("register carlito");
    (fonts, id)
}

/// A 2x2 checkerboard, small enough to keep the golden about the painter rather
/// than the decoder, and asymmetric so a flipped draw is visible.
static CHECKER: LazyLock<Vec<u8>> = LazyLock::new(|| {
    let pixels: [u8; 16] = [
        0xef, 0x44, 0x44, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0x1d, 0x4e, 0xd8,
        0xff,
    ];
    let mut bytes = Vec::new();
    {
        let mut encoder = png::Encoder::new(&mut bytes, 2, 2);
        encoder.set_color(png::ColorType::Rgba);
        encoder.set_depth(png::BitDepth::Eight);
        let mut writer = encoder.write_header().expect("png header");
        writer.write_image_data(&pixels).expect("png data");
    }
    bytes
});

fn assets() -> AssetMap<'static> {
    AssetMap::from([("ppt/media/image1.png", CHECKER.as_slice())])
}

fn slide(primitives: Vec<Primitive>) -> SurfaceDisplayList {
    SurfaceDisplayList {
        contract_version: CONTRACT_VERSION,
        width: 240.0,
        height: 135.0,
        background: Some(Paint::Solid {
            color: "#fdfdfd".into(),
        }),
        primitives,
    }
}

fn rect_path() -> Vec<ooxml_drawingml::GeometryPathCommand> {
    use ooxml_drawingml::GeometryPathCommand as Cmd;
    vec![
        Cmd::Move { x: 0.0, y: 0.0 },
        Cmd::Line { x: 1.0, y: 0.0 },
        Cmd::Line { x: 1.0, y: 1.0 },
        Cmd::Line { x: 0.0, y: 1.0 },
        Cmd::Close,
    ]
}

fn shape(x: f32, y: f32, w: f32, h: f32, fill: Option<Paint>, stroke: Option<Stroke>) -> Primitive {
    Primitive::Shape {
        object_id: 1,
        shape_id: Some("shape-1".into()),
        name: "rect".into(),
        x,
        y,
        w,
        h,
        geometry: "rect".into(),
        path: rect_path(),
        adjust_values: BTreeMap::new(),
        fill,
        stroke,
        transform: Transform::default(),
    }
}

/// Shapes a real run through the store so the golden covers the glyph path the
/// layout engine feeds a rasterizer, not a hand-invented one.
fn text_box(x: f32, y: f32, text: &str, size_px: f32, underline: bool) -> Primitive {
    let (fonts, font) = font_store();
    let shaped = ooxml_text::shape(&fonts, font, text, size_px, &[]).expect("shape");
    let baseline = y + size_px;
    let mut pen = x;
    let mut glyphs = Vec::with_capacity(shaped.len());
    for glyph in &shaped {
        glyphs.push(PositionedGlyph {
            glyph_id: glyph.glyph_id,
            cluster: glyph.cluster,
            x: pen,
            advance: glyph.x_advance,
            x_offset: glyph.x_offset,
            y_offset: baseline + glyph.y_offset,
        });
        pen += glyph.x_advance;
    }
    let width = pen - x;
    Primitive::TextBox {
        object_id: 2,
        shape_id: Some("text-1".into()),
        story_id: Some("story-1".into()),
        x,
        y,
        w: width + 8.0,
        h: size_px * 2.0,
        anchor: TextAnchor::Top,
        paragraphs: vec![TextParagraph {
            align: None,
            level: 0,
            runs: Vec::new(),
        }],
        lines: vec![PositionedTextLine {
            x,
            y,
            width,
            height: size_px * 1.2,
            baseline,
            start: 0,
            end: text.len() as u32,
            runs: vec![PositionedTextRun {
                text: text.to_owned(),
                start: 0,
                end: text.len() as u32,
                x,
                width,
                font_id: font.to_u32(),
                font_family: "Carlito".into(),
                font_size_px: size_px,
                bold: false,
                italic: false,
                underline,
                color: "#1b2733".into(),
                glyphs,
            }],
            caret_stops: vec![CaretStop { position: 0, x }],
        }],
        overflow: false,
        transform: Transform::default(),
    }
}

#[test]
fn golden_shapes() {
    check(
        "shapes",
        &slide(vec![
            shape(
                12.0,
                12.0,
                90.0,
                50.0,
                Some(Paint::Solid {
                    color: "#3b82f6".into(),
                }),
                Some(Stroke {
                    color: "#1e3a8a".into(),
                    width: 2.0,
                    dashed: false,
                    head_end: None,
                    tail_end: None,
                }),
            ),
            shape(
                120.0,
                12.0,
                90.0,
                50.0,
                None,
                Some(Stroke {
                    color: "#ef4444".into(),
                    width: 3.0,
                    dashed: true,
                    head_end: None,
                    tail_end: None,
                }),
            ),
        ]),
    );
}

#[test]
fn golden_gradient() {
    check(
        "gradient",
        &slide(vec![
            shape(
                10.0,
                10.0,
                100.0,
                110.0,
                Some(Paint::Gradient {
                    gradient_type: GradientType::Linear,
                    angle_deg: Some(45.0),
                    stops: vec![
                        GradientStop {
                            position: 0.0,
                            color: "#f97316".into(),
                        },
                        GradientStop {
                            position: 1.0,
                            color: "#7c3aed".into(),
                        },
                    ],
                }),
                None,
            ),
            shape(
                130.0,
                10.0,
                100.0,
                110.0,
                Some(Paint::Gradient {
                    gradient_type: GradientType::Radial,
                    angle_deg: None,
                    stops: vec![
                        GradientStop {
                            position: 0.0,
                            color: "#ffffff".into(),
                        },
                        GradientStop {
                            position: 1.0,
                            color: "#0f766e".into(),
                        },
                    ],
                }),
                None,
            ),
        ]),
    );
}

#[test]
fn golden_text() {
    check(
        "text",
        &slide(vec![
            text_box(12.0, 20.0, "Quarterly review", 22.0, false),
            text_box(12.0, 70.0, "Underlined subtitle", 14.0, true),
        ]),
    );
}

#[test]
fn golden_image() {
    check(
        "image",
        &slide(vec![Primitive::Image {
            object_id: 3,
            shape_id: Some("pic-1".into()),
            name: "picture".into(),
            x: 20.0,
            y: 20.0,
            w: 100.0,
            h: 80.0,
            asset_id: Some("ppt/media/image1.png".into()),
            crop: Default::default(),
            path: None,
            stroke: Some(Stroke {
                color: "#111827".into(),
                width: 2.0,
                dashed: false,
                head_end: None,
                tail_end: None,
            }),
            transform: Transform::default(),
        }]),
    );
}

#[test]
fn golden_rotated() {
    let mut rotated = shape(
        70.0,
        35.0,
        100.0,
        60.0,
        Some(Paint::Solid {
            color: "#0ea5e9".into(),
        }),
        None,
    );
    if let Primitive::Shape { transform, .. } = &mut rotated {
        *transform = Transform {
            rotation_deg: 30.0,
            flip_h: false,
            flip_v: true,
        };
    }
    check("rotated", &slide(vec![rotated]));
}

#[test]
fn golden_placeholder() {
    check(
        "placeholder",
        &slide(vec![Primitive::Placeholder {
            object_id: 4,
            shape_id: Some("ph-1".into()),
            name: "table".into(),
            x: 30.0,
            y: 30.0,
            w: 180.0,
            h: 70.0,
            label: Some("Table".into()),
            transform: Transform::default(),
        }]),
    );
}

#[test]
fn golden_chart() {
    check(
        "chart",
        &slide(vec![Primitive::Chart {
            object_id: 5,
            shape_id: Some("chart-1".into()),
            name: "chart".into(),
            x: 20.0,
            y: 20.0,
            w: 120.0,
            h: 90.0,
            label: "Revenue by quarter".into(),
            primitives: vec![
                shape(
                    30.0,
                    60.0,
                    24.0,
                    50.0,
                    Some(Paint::Solid {
                        color: "#22c55e".into(),
                    }),
                    None,
                ),
                shape(
                    64.0,
                    40.0,
                    24.0,
                    70.0,
                    Some(Paint::Solid {
                        color: "#16a34a".into(),
                    }),
                    None,
                ),
                // Reaches past the chart rect, so the golden proves the clip.
                shape(
                    98.0,
                    20.0,
                    80.0,
                    90.0,
                    Some(Paint::Solid {
                        color: "#15803d".into(),
                    }),
                    None,
                ),
            ],
            transform: Transform::default(),
        }]),
    );
}

#[test]
fn output_is_byte_deterministic() {
    let list = slide(vec![
        text_box(12.0, 20.0, "Deterministic", 20.0, false),
        shape(
            140.0,
            20.0,
            80.0,
            60.0,
            Some(Paint::Solid {
                color: "#3b82f6".into(),
            }),
            None,
        ),
    ]);
    let (fonts, font) = font_store();
    let images = assets();
    let resources = RenderResources::new(&fonts, &images).with_label_font(Some(font));
    let first = render_slide(&list, &resources, &RenderOptions::default()).expect("first");
    let second = render_slide(&list, &resources, &RenderOptions::default()).expect("second");
    assert_eq!(first.bytes, second.bytes);
}

#[test]
fn a_transparent_render_differs_from_the_opaque_one() {
    let mut list = slide(vec![]);
    list.background = None;
    let (fonts, _) = font_store();
    let images = assets();
    let resources = RenderResources::new(&fonts, &images);
    let opaque = render_slide(&list, &resources, &RenderOptions::default()).expect("opaque");
    let clear = render_slide(
        &list,
        &resources,
        &RenderOptions {
            background: Background::Transparent,
            ..RenderOptions::default()
        },
    )
    .expect("clear");
    assert_ne!(opaque.bytes, clear.bytes);
}
