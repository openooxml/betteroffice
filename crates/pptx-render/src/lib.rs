//! PPTX display-list compiler.

mod display_list;

pub use display_list::*;

use serde::Deserialize;
use std::collections::BTreeMap;

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ComposedSlide {
    width_px: f32,
    height_px: f32,
    #[serde(default)]
    background: Option<Paint>,
    shapes: Vec<ComposedShape>,
}

#[derive(Debug, Deserialize)]
#[serde(
    tag = "kind",
    rename_all = "camelCase",
    rename_all_fields = "camelCase"
)]
enum ComposedShape {
    Shape {
        #[serde(flatten)]
        base: ShapeBase,
        geometry: String,
        #[serde(default)]
        adjust_values: BTreeMap<String, f32>,
        #[serde(default)]
        fill: Option<Paint>,
        #[serde(default)]
        stroke: Option<ComposedStroke>,
        #[serde(default)]
        text: Option<ComposedText>,
    },
    Picture {
        #[serde(flatten)]
        base: ShapeBase,
        #[serde(default)]
        image_part_path: Option<String>,
        #[serde(default)]
        stroke: Option<ComposedStroke>,
    },
    TablePlaceholder {
        #[serde(flatten)]
        base: ShapeBase,
    },
    ChartPlaceholder {
        #[serde(flatten)]
        base: ShapeBase,
    },
    Unknown {
        #[serde(flatten)]
        base: ShapeBase,
    },
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ShapeBase {
    id: u32,
    name: String,
    rect: Rect,
    rotation_deg: f32,
    #[serde(default)]
    flip_h: bool,
    #[serde(default)]
    flip_v: bool,
}

#[derive(Debug, Deserialize)]
struct Rect {
    x: f32,
    y: f32,
    w: f32,
    h: f32,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ComposedStroke {
    color_hex: String,
    width_px: f32,
    #[serde(default)]
    dash: bool,
}

#[derive(Debug, Deserialize)]
struct ComposedText {
    paragraphs: Vec<ComposedParagraph>,
    anchor: ComposedAnchor,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
enum ComposedAnchor {
    T,
    Ctr,
    B,
}

#[derive(Debug, Deserialize)]
struct ComposedParagraph {
    #[serde(default)]
    align: Option<ComposedAlign>,
    level: u32,
    runs: Vec<ComposedRun>,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
enum ComposedAlign {
    L,
    Ctr,
    R,
    Just,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ComposedRun {
    text: String,
    font_family: String,
    font_size_pt: f32,
    #[serde(default)]
    bold: bool,
    #[serde(default)]
    italic: bool,
    #[serde(default)]
    underline: bool,
    color_hex: String,
}

pub fn compile_json(slide_json: &str) -> Result<String, String> {
    let slide: ComposedSlide = serde_json::from_str(slide_json)
        .map_err(|error| format!("invalid composed slide: {error}"))?;
    serde_json::to_string(&compile(slide))
        .map_err(|error| format!("could not serialize display list: {error}"))
}

fn compile(slide: ComposedSlide) -> SurfaceDisplayList {
    let mut primitives = Vec::with_capacity(slide.shapes.len() * 2);
    for shape in slide.shapes {
        match shape {
            ComposedShape::Shape {
                base,
                geometry,
                adjust_values,
                fill,
                stroke,
                text,
            } => {
                let transform = transform(&base);
                primitives.push(Primitive::Shape {
                    object_id: base.id,
                    name: base.name,
                    x: base.rect.x,
                    y: base.rect.y,
                    w: base.rect.w,
                    h: base.rect.h,
                    geometry,
                    adjust_values,
                    fill,
                    stroke: stroke.map(Into::into),
                    transform,
                });
                if let Some(text) = text {
                    primitives.push(text_primitive(base.id, base.rect, transform, text));
                }
            }
            ComposedShape::Picture {
                base,
                image_part_path,
                stroke,
            } => {
                let transform = transform(&base);
                primitives.push(Primitive::Image {
                    object_id: base.id,
                    name: base.name,
                    x: base.rect.x,
                    y: base.rect.y,
                    w: base.rect.w,
                    h: base.rect.h,
                    asset_id: image_part_path,
                    stroke: stroke.map(Into::into),
                    transform,
                });
            }
            ComposedShape::TablePlaceholder { base } => {
                primitives.push(placeholder(base, Some("Table")))
            }
            ComposedShape::ChartPlaceholder { base } => {
                primitives.push(placeholder(base, Some("Chart")))
            }
            ComposedShape::Unknown { base } => primitives.push(placeholder(base, None)),
        }
    }

    SurfaceDisplayList {
        contract_version: CONTRACT_VERSION,
        width: slide.width_px,
        height: slide.height_px,
        background: slide.background.or_else(|| {
            Some(Paint::Solid {
                color: "#ffffff".into(),
            })
        }),
        primitives,
    }
}

fn transform(base: &ShapeBase) -> Transform {
    Transform {
        rotation_deg: base.rotation_deg,
        flip_h: base.flip_h,
        flip_v: base.flip_v,
    }
}

fn placeholder(base: ShapeBase, label: Option<&str>) -> Primitive {
    let transform = transform(&base);
    Primitive::Placeholder {
        object_id: base.id,
        name: base.name,
        x: base.rect.x,
        y: base.rect.y,
        w: base.rect.w,
        h: base.rect.h,
        label: label.map(str::to_owned),
        transform,
    }
}

fn text_primitive(
    object_id: u32,
    rect: Rect,
    transform: Transform,
    text: ComposedText,
) -> Primitive {
    Primitive::TextBox {
        object_id,
        x: rect.x,
        y: rect.y,
        w: rect.w,
        h: rect.h,
        anchor: match text.anchor {
            ComposedAnchor::T => TextAnchor::Top,
            ComposedAnchor::Ctr => TextAnchor::Center,
            ComposedAnchor::B => TextAnchor::Bottom,
        },
        paragraphs: text
            .paragraphs
            .into_iter()
            .map(|paragraph| TextParagraph {
                align: paragraph.align.map(|align| match align {
                    ComposedAlign::L => TextAlign::Left,
                    ComposedAlign::Ctr => TextAlign::Center,
                    ComposedAlign::R => TextAlign::Right,
                    ComposedAlign::Just => TextAlign::Justify,
                }),
                level: paragraph.level,
                runs: paragraph
                    .runs
                    .into_iter()
                    .map(|run| TextRun {
                        text: run.text,
                        font_family: run.font_family,
                        font_size_pt: run.font_size_pt,
                        bold: run.bold,
                        italic: run.italic,
                        underline: run.underline,
                        color: run.color_hex,
                    })
                    .collect(),
            })
            .collect(),
        transform,
    }
}

impl From<ComposedStroke> for Stroke {
    fn from(stroke: ComposedStroke) -> Self {
        Self {
            color: stroke.color_hex,
            width: stroke.width_px,
            dashed: stroke.dash,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn compiles_shape_and_text_in_paint_order() {
        let json = r##"{
          "widthPx":1280,"heightPx":720,
          "shapes":[{
            "kind":"shape","id":7,"name":"Title",
            "rect":{"x":10,"y":20,"w":300,"h":80},
            "rotationDeg":15,"geometry":"roundRect","adjustValues":{"adj":0.2},
            "fill":{"kind":"solid","color":"#4472c4"},
            "text":{"anchor":"ctr","paragraphs":[{"align":"ctr","level":0,"runs":[{
              "text":"Hello","fontFamily":"Aptos","fontSizePt":24,"bold":true,
              "colorHex":"#ffffff"
            }]}]}
          }]
        }"##;

        let output: serde_json::Value =
            serde_json::from_str(&compile_json(json).expect("compile")).expect("display list json");
        assert_eq!(output["contractVersion"], CONTRACT_VERSION);
        assert_eq!(output["primitives"][0]["kind"], "shape");
        assert_eq!(output["primitives"][1]["kind"], "textBox");
        assert_eq!(output["primitives"][1]["objectId"], 7);
        assert_eq!(output["primitives"][1]["anchor"], "center");
    }

    #[test]
    fn rejects_invalid_composed_json() {
        let error = compile_json("{}").expect_err("missing dimensions must fail");
        assert!(error.starts_with("invalid composed slide:"));
    }

    #[test]
    fn defaults_surface_background_to_white() {
        let output: serde_json::Value = serde_json::from_str(
            &compile_json(r#"{"widthPx":10,"heightPx":10,"shapes":[]}"#).expect("compile"),
        )
        .expect("json");
        assert_eq!(output["background"]["color"], "#ffffff");
    }

    #[test]
    fn paint_type_round_trips_through_input() {
        let input = r##"{
          "widthPx":10,"heightPx":10,
          "background":{"kind":"gradient","gradientType":"linear","angleDeg":90,
            "stops":[{"position":0,"color":"#000000"},{"position":1,"color":"#ffffff"}]},
          "shapes":[]
        }"##;
        let output: serde_json::Value =
            serde_json::from_str(&compile_json(input).expect("compile")).expect("json");
        assert_eq!(output["background"]["gradientType"], "linear");
        assert_eq!(
            output["background"]["stops"][1]["position"].as_f64(),
            Some(1.0)
        );
    }
}
